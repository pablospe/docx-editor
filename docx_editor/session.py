"""Persistent Jupyter kernel session for multi-step document editing.

Keeps documents open across many small commands (AI-agent friendly) instead
of re-opening them in one-off scripts. Requires the optional extra:

    pip install docx-editor[session]

CLI (see main()):
    docx-session start | exec "code" | eval "expr" | status | stop
"""

import argparse
import ast
import json
import math
import os
import re
import signal
import subprocess
import sys
import time
from dataclasses import asdict, dataclass
from pathlib import Path
from queue import Empty
from typing import TYPE_CHECKING, Any, Literal

from .exceptions import SessionDeadError, SessionError
from .workspace import _pid_alive

if TYPE_CHECKING:
    from jupyter_client import BlockingKernelClient

DEFAULT_CONNECTION_FILE = Path.home() / ".cache" / "docx-editor" / "kernel.json"

# start_session's default budget, and the environment variable that overrides
# it. A kernel start is a fresh interpreter importing ipykernel + IPython; on a
# loaded machine (a 2-vCPU CI runner with parallel test workers, coverage
# tracing on) that import alone can take longer than the unloaded 30 s default.
DEFAULT_START_TIMEOUT = 30.0
START_TIMEOUT_ENV = "DOCX_EDITOR_SESSION_START_TIMEOUT"

# How the kernel is launched. Tests substitute a stand-in kernel here to
# exercise the start-failure paths deterministically; the "-f <file>" argument
# is appended by start_session.
_KERNEL_COMMAND = (sys.executable, "-m", "ipykernel_launcher")

# Bytes of kernel stderr kept in a start-failure diagnostic.
_STDERR_TAIL_BYTES = 4096

_EXTRA_HINT = "Session mode requires extra dependencies: pip install 'docx-editor[session]'"


def _client(connection_file: Path) -> "BlockingKernelClient":
    """Return a connected BlockingKernelClient for the session.

    Raises:
        ImportError: If the [session] extra is not installed.
        ValueError/OSError: If the connection file is corrupt or unreadable.
    """
    try:
        from jupyter_client import BlockingKernelClient
    except ImportError as e:
        raise ImportError(_EXTRA_HINT) from e

    kc = BlockingKernelClient(connection_file=str(connection_file))
    kc.load_connection_file()
    # Skip the heartbeat channel: we open a fresh short-lived client per call, and
    # its background thread races with channel teardown, spraying "Too many open
    # files" ZMQError tracebacks to stderr on every invocation.
    kc.start_channels(hb=False)
    return kc


def _kernel_alive(kc: "BlockingKernelClient", timeout: float = 5.0) -> bool:
    """True if the kernel answers a kernel_info request on the *control* channel.

    ipykernel services the control channel on a dedicated thread, so it replies
    even while an execution is in flight. The shell channel (which
    wait_for_ready() uses) is serialized behind the running execute_request, so
    there a busy kernel is indistinguishable from a dead one.
    """
    msg = kc.session.msg("kernel_info_request", {})
    kc.control_channel.send(msg)
    msg_id = msg["header"]["msg_id"]

    deadline = time.monotonic() + timeout
    while True:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return False
        try:
            reply = kc.get_control_msg(timeout=min(remaining, 0.5))
        except Empty:
            continue
        if reply.get("parent_header", {}).get("msg_id") == msg_id:
            return True


def _kernel_idle(kc: "BlockingKernelClient", timeout: float = 2.0) -> bool:
    """True if the kernel answers a kernel_info request on the *shell* channel.

    The shell channel is serialized behind any running execute_request, so a
    busy kernel leaves this request queued past the window — which is exactly
    the idle/busy signal. The queued reply eventually lands on this
    disconnected client's socket and is dropped by ZMQ.
    """
    msg = kc.session.msg("kernel_info_request", {})
    kc.shell_channel.send(msg)
    msg_id = msg["header"]["msg_id"]

    deadline = time.monotonic() + timeout
    while True:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return False
        try:
            reply = kc.get_shell_msg(timeout=min(remaining, 0.5))
        except Empty:
            continue
        if reply.get("parent_header", {}).get("msg_id") == msg_id:
            return True


def _kernel_dead(kc: "BlockingKernelClient", connection_file: Path) -> bool:
    """True only when the kernel is provably gone or unreachable.

    Fast path: the recorded PID no longer exists (reap=True detects our own
    zombie child in library use). Fallback: the control-channel probe — which
    ipykernel answers on a dedicated thread even mid-exec — gets no reply.
    _pid_alive errs toward "alive", so a live kernel never trips the fast path.
    """
    pid = _read_pid(connection_file)
    if pid is not None and not _pid_alive(pid, reap=True):
        return True
    return not _kernel_alive(kc, timeout=5.0)


def _pid_file(connection_file: Path) -> Path:
    return connection_file.with_suffix(".pid")


def _read_pid(connection_file: Path) -> int | None:
    """PID recorded for this session, or None if absent/unreadable/corrupt."""
    try:
        return int(_pid_file(connection_file).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def _start_timeout() -> float:
    """start_session's budget: DOCX_EDITOR_SESSION_START_TIMEOUT, else the 30 s default.

    Raises:
        SessionError: If the variable is set to anything but a positive number.
    """
    raw = os.environ.get(START_TIMEOUT_ENV, "").strip()
    if not raw:
        return DEFAULT_START_TIMEOUT
    try:
        value = float(raw)
        if not (math.isfinite(value) and value > 0):  # rejects nan and inf too
            raise ValueError
    except ValueError:
        raise SessionError(f"{START_TIMEOUT_ENV} must be a positive number of seconds, got {raw!r}") from None
    return value


def _drain_stderr(proc: subprocess.Popen[bytes]) -> str:
    """Tail of what a (killed and reaped) kernel wrote to stderr, without blocking.

    Invariant: nothing may read ``proc.stderr`` before this runs — the raw-fd
    read below bypasses the ``BufferedReader``, so any bytes already buffered
    there would be silently dropped. (Today the only other reader, the
    exited-during-startup path, always raises without reaching ``_abort``.)

    The pipe is switched to non-blocking first: every writer is dead, so the
    read hits EOF at once — but if a grandchild ever inherited the write end,
    EAGAIN ends the loop instead of hanging a diagnostic that exists to explain
    a hang.
    """
    if proc.stderr is None:
        return ""
    fd = proc.stderr.fileno()
    if os.name == "posix":
        os.set_blocking(fd, False)
    chunks: list[bytes] = []
    while True:
        try:
            chunk = os.read(fd, 65536)
        except OSError:  # EAGAIN: a writer still holds the pipe open
            break
        if not chunk:
            break
        chunks.append(chunk)
    return b"".join(chunks)[-_STDERR_TAIL_BYTES:].decode(errors="replace").strip()


def _describe_connection_file(connection_file: Path) -> str:
    """One line on what the kernel managed to write before the budget ran out."""
    try:
        size = connection_file.stat().st_size
    except OSError:
        return "Connection file: absent."
    try:
        info = json.loads(connection_file.read_text(encoding="utf-8"))
        return (
            f"Connection file: {size} bytes, {info['transport']} {info['ip']}, "
            f"shell={info['shell_port']} control={info['control_port']}."
        )
    except (OSError, ValueError, KeyError, TypeError) as e:
        return f"Connection file: {size} bytes, unreadable: {e}."


def _load_average() -> str:
    """Suffix "; load average N.NN" where the platform reports one, else ""."""
    if not hasattr(os, "getloadavg"):  # Windows
        return ""
    try:
        return f"; load average {os.getloadavg()[0]:.2f}"
    except OSError:
        return ""


def start_session(connection_file: Path = DEFAULT_CONNECTION_FILE, timeout: float | None = None) -> int:
    """Start a detached IPython kernel and wait until it answers.

    Args:
        connection_file: Where to write the kernel connection file
        timeout: Seconds to wait for the kernel to become ready. ``None`` (the
            default) reads ``DOCX_EDITOR_SESSION_START_TIMEOUT`` and falls back to
            30 s when it is unset or blank; an explicit argument always wins.
            The variable exists for loaded machines — CI runners, parallel
            test workers — where a fresh interpreter importing IPython takes
            far longer than it does interactively. ``docx-session start``
            has no flag for it and reaches the variable through this default.

    Returns:
        PID of the kernel process.

    Raises:
        SessionError: If a session is already running, the kernel fails to
            start, the connection file it wrote is unreadable, or
            ``DOCX_EDITOR_SESSION_START_TIMEOUT`` is not a positive number.
            A start that runs out of budget explains itself: which phase
            timed out (connection file vs. control-channel reply), whether
            the kernel was still alive, the load average, what it wrote to
            stderr, and the variable to raise.
        ImportError: If the [session] extra is not installed.
    """
    if timeout is None:
        timeout = _start_timeout()
    if is_session_running(connection_file):
        raise SessionError(f"Session already running (connection file: {connection_file})")

    connection_file.parent.mkdir(parents=True, exist_ok=True)
    connection_file.unlink(missing_ok=True)

    proc = subprocess.Popen(
        [*_KERNEL_COMMAND, "-f", str(connection_file)],
        stdout=subprocess.DEVNULL,
        # Keep stderr so a failed start (e.g. missing ipykernel) can explain itself.
        stderr=subprocess.PIPE,
        # Detach on POSIX so the kernel outlives this CLI invocation.
        start_new_session=(os.name == "posix"),
    )
    _pid_file(connection_file).write_text(str(proc.pid), encoding="utf-8")

    def _abort(summary: str) -> SessionError:
        # Sampled before the kill so the report says what the kernel was doing
        # at the timeout, not what the kill did to it.
        exit_code = proc.poll()
        proc.kill()
        proc.wait()
        kernel = "was still alive at the timeout" if exit_code is None else f"had exited with code {exit_code}"
        stderr = _drain_stderr(proc)
        report = "\n".join([
            summary,
            f"Kernel pid {proc.pid} {kernel}{_load_average()}.",
            _describe_connection_file(connection_file),
            "Kernel stderr:",
            "\n".join(f"  {line}" for line in stderr.splitlines()) if stderr else "  (empty)",
            f"Set {START_TIMEOUT_ENV} to raise the budget on a loaded machine.",
        ])
        connection_file.unlink(missing_ok=True)
        _pid_file(connection_file).unlink(missing_ok=True)
        return SessionError(report)

    # Two phases share one budget. The connection file is written as soon as
    # the sockets are bound — *before* ipykernel imports IPython and wires the
    # control channel — so its appearance is an early signal, and the second
    # phase gets whatever is left (floored at 1 s).
    t0 = time.monotonic()
    deadline = t0 + timeout
    while not (connection_file.exists() and connection_file.stat().st_size > 0):
        if proc.poll() is not None:
            stderr = proc.stderr.read().decode(errors="replace").strip() if proc.stderr else ""
            _pid_file(connection_file).unlink(missing_ok=True)
            hint = _EXTRA_HINT if "ipykernel" in stderr else ""
            detail = f"\n{stderr}" if stderr else ""
            raise SessionError(f"Kernel process exited during startup (code {proc.returncode}). {hint}{detail}".strip())
        if time.monotonic() > deadline:
            raise _abort(
                f"Kernel did not start within {timeout}s: no connection file after {time.monotonic() - t0:.1f}s."
            )
        time.sleep(0.1)
    t_file = time.monotonic() - t0

    kc = None
    try:
        # Inside the guard: the kernel writes the connection file
        # non-atomically, so a mid-write read fails to parse and must not
        # leave the detached kernel running.
        try:
            kc = _client(connection_file)
        except (ValueError, OSError) as e:
            raise SessionError(f"Kernel wrote an unreadable connection file ({connection_file}): {e}") from e
        remaining = max(1.0, deadline - time.monotonic())
        if not _kernel_alive(kc, timeout=remaining):
            raise _abort(
                f"Kernel did not become ready within {timeout}s: kernel.json appeared after {t_file:.1f}s, "
                f"then no kernel_info reply on the control channel in the remaining {remaining:.1f}s."
            )
    except BaseException:
        # Never leave a half-started kernel behind with no way to reach it.
        if proc.poll() is None:
            proc.kill()
            proc.wait()
        connection_file.unlink(missing_ok=True)
        _pid_file(connection_file).unlink(missing_ok=True)
        raise
    finally:
        # ``kc`` is None only when ``_client`` raised, and that exception always
        # propagates: the false arm never reaches the ``return`` below.
        if kc is not None:  # pragma: no branch
            kc.stop_channels()
    return proc.pid


def is_session_running(connection_file: Path = DEFAULT_CONNECTION_FILE, timeout: float = 5.0) -> bool:
    """True if a kernel is answering on this connection file.

    A *busy* kernel counts as running: the probe goes over the control channel,
    which ipykernel answers while an execution is still in flight.

    Args:
        connection_file: Kernel connection file to probe
        timeout: Seconds to wait for the kernel to answer

    Raises:
        ImportError: If the [session] extra is not installed.
    """
    if not connection_file.exists():
        return False
    try:
        kc = _client(connection_file)
    except (ValueError, OSError):
        return False  # Corrupt/truncated connection file — treat as no session.
    try:
        return _kernel_alive(kc, timeout=timeout)
    finally:
        kc.stop_channels()


def stop_session(connection_file: Path = DEFAULT_CONNECTION_FILE, timeout: float = 5.0) -> bool:
    """Shut down the kernel (graceful request, SIGTERM fallback).

    Args:
        connection_file: Kernel connection file of the session to stop
        timeout: Seconds to wait for the kernel to exit before SIGTERM

    Returns:
        True if a session existed and was stopped.

    Raises:
        ImportError: If the [session] extra is not installed.
    """
    if not connection_file.exists():
        return False

    # reply=True blocks until the kernel acks. The bare shutdown() this replaced
    # was fire-and-forget, and stop_channels() tore the socket down before the
    # request was ever flushed — so every stop fell through to the SIGTERM path.
    acknowledged = False
    try:
        kc = _client(connection_file)
    except (ValueError, OSError):
        kc = None  # Corrupt connection file — nothing to talk to; just clean up.
    if kc is not None:
        try:
            kc.shutdown(reply=True, timeout=timeout)
            acknowledged = True
        except (RuntimeError, TimeoutError, Empty):
            acknowledged = False
        finally:
            kc.stop_channels()

    # os.kill() on Windows terminates rather than signals, so the poll itself would
    # kill the kernel. There we rely on the shutdown ack alone.
    pid = _read_pid(connection_file)
    if pid is not None and os.name == "posix":
        deadline = time.monotonic() + timeout
        # reap=True: the kernel may be this process's own child (library use),
        # and only reaping detects its exit — a zombie otherwise polls as alive.
        while time.monotonic() < deadline and _pid_alive(pid, reap=True):
            time.sleep(0.05)
        # Only signal a kernel that answered us: an unacknowledged PID may be stale
        # (crash, reboot) and since recycled to an unrelated process.
        if acknowledged and _pid_alive(pid, reap=True):
            os.kill(pid, signal.SIGTERM)

    connection_file.unlink(missing_ok=True)
    _pid_file(connection_file).unlink(missing_ok=True)
    return True


@dataclass
class SessionStatus:
    """Snapshot of a session's state (see session_status())."""

    running: bool
    pid: int | None
    state: Literal["idle", "busy"] | None  # None when not running
    connection_file: Path
    stale: bool  # session files exist but the kernel is unreachable — 'stop' cleans up


def session_status(connection_file: Path = DEFAULT_CONNECTION_FILE) -> SessionStatus:
    """Inspect the session: liveness, PID, and whether the kernel is idle or busy.

    stale=True means session files exist but the kernel is not answering (it
    crashed, was killed, or the machine rebooted).

    Args:
        connection_file: Kernel connection file to inspect

    Raises:
        ImportError: If the [session] extra is not installed.
    """
    if not connection_file.exists():
        return SessionStatus(running=False, pid=None, state=None, connection_file=connection_file, stale=False)

    pid = _read_pid(connection_file)
    try:
        kc = _client(connection_file)
    except (ValueError, OSError):
        return SessionStatus(running=False, pid=pid, state=None, connection_file=connection_file, stale=True)
    try:
        if not _kernel_alive(kc):
            return SessionStatus(running=False, pid=pid, state=None, connection_file=connection_file, stale=True)
        state: Literal["idle", "busy"] = "idle" if _kernel_idle(kc) else "busy"
    finally:
        kc.stop_channels()
    return SessionStatus(running=True, pid=pid, state=state, connection_file=connection_file, stale=False)


_ANSI_RE = re.compile(r"\x1b\[[0-9;]*m")

# Machine-specific prefix before a docx_editor/ path component. The lookahead
# requires a separator immediately before docx_editor/, so a user path like
# /x/my_docx_editor/foo.py stays untouched. Any literal docx_editor/ component
# matches, though — even a user document under a directory so named — but
# stripping only ever rewrites traceback text, never structured error fields.
_INTERNAL_PATH_RE = re.compile(r"(?:[A-Za-z]:)?[\\/~]\S*?[\\/](?=docx_editor[\\/])")
# ipykernel writes each executed cell to /tmp/ipykernel_<pid>/<n>.py.
_CELL_PATH_RE = re.compile(r"(?:[A-Za-z]:)?[\\/~]\S*?[\\/]ipykernel_\d+[\\/]\d+\.py")


def _strip_internal_paths(text: str) -> str:
    """Rewrite machine-specific paths in a traceback to stable relative forms.

    Library frames become `docx_editor/...`; ipykernel cell files become
    `<session-cell>`. Keeps tracebacks compact and free of absolute repo
    paths that mean nothing to a CLI consumer.
    """
    return _CELL_PATH_RE.sub("<session-cell>", _INTERNAL_PATH_RE.sub("", text))


ExecStatus = Literal["ok", "error", "timeout", "dead"]

# Seconds of iopub silence before probing whether a mid-exec kernel is dead.
# A live-but-busy kernel answers the control probe and the silence clock resets.
_SILENCE_PROBE_AFTER = 10.0


@dataclass
class ExecResult:
    """Outcome of one exec_code() call against the session kernel.

    ``started`` records whether the kernel began executing this request at all,
    which disambiguates two statuses:

    - ``"timeout"``: with ``started=True`` your code ran and was still running
      when the clock ran out (raise ``timeout``, or make the code faster); with
      ``started=False`` the request never left the queue, because the kernel was
      still busy with an earlier command — nothing of yours executed, so wait or
      check what is hogging the session.
    - ``"ok"`` with ``started=False``: the kernel **discarded** the request
      without running it. ipykernel aborts everything still queued behind a
      command that raises, so a request sent while an earlier one was failing
      comes back clean and empty having done nothing. Re-send it.

    An ``"error"`` result, and any ``"ok"`` with ``started=True``, did execute.
    """

    status: ExecStatus
    stdout: str = ""
    stderr: str = ""
    result: str | None = None  # repr of the last expression, if any
    traceback: str | None = None  # ANSI-stripped traceback when status == "error"
    started: bool = False  # True once the kernel began executing this request


def exec_code(code: str, connection_file: Path = DEFAULT_CONNECTION_FILE, timeout: float = 120.0) -> ExecResult:
    """Execute code in the session kernel and collect its output.

    If the kernel is busy, the request queues behind the running one; `timeout`
    covers the whole wait. A timeout therefore has two flavours, told apart by
    the result's ``started`` flag: your own code ran too long (True), or it never
    got out of the queue (False).

    Args:
        code: Python source to execute in the session
        connection_file: Kernel connection file of the session
        timeout: Seconds to wait for the execution to finish

    Returns:
        ExecResult with status "ok", "error", "timeout", or "dead" (the kernel
        died mid-execution — its state is lost), and ``started`` telling whether
        the kernel began executing this request at all. Check ``started`` on an
        "ok" result too: ``ok`` with ``started=False`` means the kernel aborted
        the request behind a failing predecessor and ran nothing.

    Raises:
        FileNotFoundError: If no session connection file exists.
        SessionDeadError: If the kernel is gone or not answering.
        SessionError: If the connection file is unreadable.
        ImportError: If the [session] extra is not installed.
    """
    if not connection_file.exists():
        raise FileNotFoundError(f"No session found ({connection_file} missing). Run 'docx-session start' first.")

    try:
        kc = _client(connection_file)
    except (ValueError, OSError) as e:
        raise SessionError(f"Session connection file is unreadable ({connection_file}): {e}") from e

    try:
        if _kernel_dead(kc, connection_file):
            raise SessionDeadError(
                f"Session kernel is not responding ({connection_file}). Run 'docx-session stop' then 'start'."
            )
        # allow_stdin=False: with stdin allowed, an input() call parks the kernel on
        # an input_request this client never services, wedging the session forever.
        # Disabled, input() raises StdinNotImplementedError -> a normal error result.
        msg_id = kc.execute(code, allow_stdin=False)

        stdout_parts: list[str] = []
        stderr_parts: list[str] = []
        result: str | None = None
        traceback: str | None = None
        status: ExecStatus = "ok"
        started = False
        deadline = time.monotonic() + timeout
        last_activity = time.monotonic()

        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                final: ExecStatus = "dead" if _kernel_dead(kc, connection_file) else "timeout"
                return ExecResult(
                    status=final,
                    stdout="".join(stdout_parts),
                    stderr="".join(stderr_parts),
                    started=started,
                )
            try:
                msg = kc.get_iopub_msg(timeout=min(remaining, 1.0))
            except Empty:
                # A crashed kernel just goes silent; without this probe the loop
                # would run out the full timeout and report it as still running.
                if time.monotonic() - last_activity < _SILENCE_PROBE_AFTER:
                    continue
                if _kernel_dead(kc, connection_file):
                    return ExecResult(
                        status="dead",
                        stdout="".join(stdout_parts),
                        stderr="".join(stderr_parts),
                        started=started,
                    )
                last_activity = time.monotonic()
                continue
            # Any iopub traffic proves the kernel is alive, whoever it belongs to.
            last_activity = time.monotonic()
            if msg.get("parent_header", {}).get("msg_id") != msg_id:
                continue

            msg_type = msg["msg_type"]
            content = msg["content"]
            if msg_type == "execute_input":
                # The kernel broadcasts this the moment it starts executing OUR
                # request — the only reliable "left the queue" signal, since a
                # queued request is otherwise indistinguishable from a slow one.
                started = True
            elif msg_type == "stream":
                target = stdout_parts if content["name"] == "stdout" else stderr_parts
                target.append(content["text"])
            elif msg_type in ("execute_result", "display_data"):
                result = content.get("data", {}).get("text/plain", result)
            elif msg_type == "error":
                status = "error"
                traceback = _strip_internal_paths(_ANSI_RE.sub("", "\n".join(content["traceback"])))
            elif msg_type == "status" and content["execution_state"] == "idle":
                break

        return ExecResult(
            status=status,
            stdout="".join(stdout_parts),
            stderr="".join(stderr_parts),
            result=result,
            traceback=traceback,
            started=started,
        )
    finally:
        kc.stop_channels()


# JSON serialization must happen kernel-side: the client only ever sees mime
# bundles. The wrapper's return value is the code block's last expression, so
# the JSON rides back as the execute_result repr — a channel that user print()
# output can't corrupt. The %s slot takes repr(expr) — a valid Python string
# literal whatever quotes/backslashes/unicode the expression contains (str.format
# would collide with the template's own braces). Keep this template %-free.
# One __docx_session_eval name is left in the user namespace (overwritten each
# call) — same class of pollution as IPython's own In/Out.
_EVAL_TEMPLATE = """\
def __docx_session_eval():
    import dataclasses
    import datetime
    import json
    import pathlib
    import traceback

    def _default(o):
        # Library dataclasses (SearchResult, Revision, Comment, ...) ride out
        # as real JSON objects instead of opaque reprs. The type guard keeps a
        # bare dataclass *class* on the repr path.
        if dataclasses.is_dataclass(o) and not isinstance(o, type):
            return dataclasses.asdict(o)
        if isinstance(o, (datetime.datetime, datetime.date)):
            return o.isoformat()
        if isinstance(o, pathlib.PurePath):
            return str(o)
        raise TypeError(f"{type(o).__name__} is not JSON serializable")

    def _err_payload(exc, depth=0):
        # type/message plus every structured recovery field the exception
        # carries (actual_hash, total_occurrences, ...). A nested exception
        # (BatchOperationError.original) recurses, depth-capped so a
        # pathological self-referential chain cannot hang the wrapper.
        payload = {"type": type(exc).__name__, "message": str(exc)}
        for key, val in vars(exc).items():
            if key in payload:
                continue
            if isinstance(val, BaseException):
                payload[key] = _err_payload(val, depth + 1) if depth < 3 else repr(val)
                continue
            try:
                json.dumps(val, allow_nan=False, default=_default)
            except Exception:
                payload[key] = repr(val)
            else:
                payload[key] = val
        return payload

    try:
        value = eval(compile(%s, "<docx-session eval>", "eval"), globals())
    except Exception as e:
        # tb_next skips this wrapper's own frame, so the trace starts at
        # <docx-session eval> (for a compile() SyntaxError it is None and the
        # caret-formatted exception alone comes back). This is a plain, few-
        # hundred-char traceback, not IPython's multi-kilobyte verbose one.
        tb = "".join(traceback.format_exception(type(e), e, e.__traceback__.tb_next))
        try:
            return json.dumps({"error": _err_payload(e), "traceback": tb}, allow_nan=False, default=_default)
        except Exception:
            return json.dumps({"error": {"type": type(e).__name__, "message": str(e)}, "traceback": tb})
    try:
        # allow_nan=False: NaN/Infinity would otherwise ride out as bare tokens
        # that are not valid RFC-8259 JSON. except Exception, not just
        # (TypeError, ValueError): an exotic dataclass field may fail deepcopy
        # inside asdict — degrade to repr, never crash the transport.
        return json.dumps({"value": value, "serialized": True}, allow_nan=False, default=_default)
    except Exception:
        return json.dumps({"value": repr(value), "serialized": False})
__docx_session_eval()
"""


@dataclass
class EvalResult:
    """Outcome of one eval_code() call: an expression's value, JSON-transported."""

    status: ExecStatus
    value: Any = None
    serialized: bool = False  # False: value was not JSON-serializable; it holds a repr string
    stdout: str = ""
    stderr: str = ""
    traceback: str | None = None
    error: dict[str, Any] | None = None  # {"type", "message", <structured fields>} when the expression raised
    started: bool = False  # as in ExecResult: False on a timeout means the request never left the queue


def eval_code(expr: str, connection_file: Path = DEFAULT_CONNECTION_FILE, timeout: float = 120.0) -> EvalResult:
    """Evaluate an expression in the session kernel and return its value.

    JSON-serializable values arrive as their JSON equivalents (serialized=True;
    a tuple arrives as a list); library dataclasses (SearchResult,
    ParagraphInfo, ParagraphLocation, Revision, Comment, ...) arrive as JSON
    objects with datetimes as ISO strings. Anything else — including
    non-finite floats, which have no valid JSON form — arrives as its repr
    string (serialized=False). Statements ('x = 5') are a SyntaxError — use
    exec_code for those.

    When the expression raises, status is "error" and `error` holds
    {"type", "message", <structured recovery fields>} captured from the
    exception object kernel-side (e.g. actual_hash for a HashMismatchError),
    with a compact path-stripped traceback. Exceptions that bypass the
    kernel-side capture (KeyboardInterrupt/SystemExit, transport failures)
    still ride the exec path: `error` stays None and `traceback` holds the
    IPython-formatted trace.

    Args:
        expr: Python expression to evaluate in the session
        connection_file: Kernel connection file of the session
        timeout: Seconds to wait for the evaluation to finish

    Raises:
        FileNotFoundError: If no session connection file exists.
        SessionDeadError: If the kernel is gone or not answering.
        SessionError: If the connection file is unreadable or the value
            transport breaks.
        ImportError: If the [session] extra is not installed.
    """
    res = exec_code(_EVAL_TEMPLATE % (repr(expr),), connection_file=connection_file, timeout=timeout)
    if res.status != "ok":
        return EvalResult(
            status=res.status,
            stdout=res.stdout,
            stderr=res.stderr,
            traceback=res.traceback,
            started=res.started,
        )
    if res.result is None:
        if not res.started:
            # Not a transport fault: the kernel discarded the request without
            # running it (see ExecResult.started), so there is no reply to decode.
            raise SessionError(
                "The kernel discarded this request without running it — it was queued behind a command "
                "that raised. Nothing was evaluated; re-send it."
            )
        raise SessionError("eval transport failed: kernel returned no result for the eval wrapper")
    try:
        payload = json.loads(ast.literal_eval(res.result))
    except (ValueError, SyntaxError) as e:
        raise SessionError(f"eval transport failed: could not decode the kernel's reply: {e}") from e
    if "error" in payload:
        return EvalResult(
            status="error",
            stdout=res.stdout,
            stderr=res.stderr,
            traceback=_strip_internal_paths(payload["traceback"]),
            error=payload["error"],
            started=res.started,
        )
    return EvalResult(
        status="ok",
        value=payload["value"],
        serialized=payload["serialized"],
        stdout=res.stdout,
        stderr=res.stderr,
        started=res.started,
    )


EXIT_OK = 0
EXIT_ERROR = 1
EXIT_TIMEOUT = 2
EXIT_NO_SESSION = 3
EXIT_KERNEL_DEAD = 4

_KERNEL_DIED_MSG = (
    "Kernel died or became unreachable during execution — its state is lost. Run 'docx-session stop' then 'start'."
)


def _read_code(arg: str) -> str:
    """The code argument itself, or stdin's content when the argument is '-'."""
    return sys.stdin.read() if arg == "-" else arg


# The examples are deliberately unindented: a copy-pasted heredoc only
# terminates when the closing PY sits at column 0.
_EXEC_EPILOG = """\
Pass '-' to read the code from stdin — no shell quoting to fight; also the easy
route for code starting with '-' (argparse would read a bare "-x" as a flag):

docx-session exec - <<'PY'
for p in doc.list_paragraphs(limit=None):
    print(p)
PY"""

_EVAL_EPILOG = """\
Pass '-' to read the expression from stdin — no shell quoting to fight; also the
easy route for expressions starting with '-' (argparse would read a bare "-x" as a flag):

docx-session eval - <<'PY'
[str(p) for p in doc.list_paragraphs(limit=None) if 'deadline' in str(p)]
PY"""


def main(argv: list[str] | None = None) -> int:
    """CLI entry point for docx-session."""
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument(
        "--session-file",
        type=Path,
        default=DEFAULT_CONNECTION_FILE,
        help="Kernel connection file; distinct files give isolated parallel sessions "
        f"(default: {DEFAULT_CONNECTION_FILE})",
    )

    parser = argparse.ArgumentParser(
        prog="docx-session",
        description="Persistent Python session for multi-step .docx editing.",
    )
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("start", parents=[common], help="Start a background kernel")
    p_exec = sub.add_parser(
        "exec",
        parents=[common],
        help="Execute code in the running kernel",
        epilog=_EXEC_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p_exec.add_argument("code", help="Python code to execute ('-' reads it from stdin)")
    p_exec.add_argument("--timeout", type=float, default=120.0, help="Seconds to wait (default: 120)")
    p_eval = sub.add_parser(
        "eval",
        parents=[common],
        help="Evaluate an expression; prints one JSON envelope on stdout",
        epilog=_EVAL_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p_eval.add_argument("expr", help="Python expression to evaluate ('-' reads it from stdin)")
    p_eval.add_argument("--timeout", type=float, default=120.0, help="Seconds to wait (default: 120)")
    sub.add_parser("status", parents=[common], help="Check whether the kernel is answering")
    sub.add_parser("stop", parents=[common], help="Shut the kernel down")
    args = parser.parse_args(argv)

    try:
        return _run(args)
    except FileNotFoundError as e:
        print(e, file=sys.stderr)
        return EXIT_NO_SESSION
    except SessionDeadError as e:
        print(e, file=sys.stderr)
        return EXIT_KERNEL_DEAD
    except (SessionError, ImportError) as e:
        print(e, file=sys.stderr)
        return EXIT_ERROR


def _run(args: argparse.Namespace) -> int:
    """Dispatch one parsed subcommand."""
    if args.command == "start":
        pid = start_session(connection_file=args.session_file)
        # The kernel inherits this cwd for its whole life, so relative paths passed
        # to a later `exec` resolve against it, not against the caller's cwd.
        print(f"Session started (pid {pid}, cwd: {Path.cwd()}, connection file: {args.session_file})")
        return EXIT_OK

    if args.command == "status":
        st = session_status(connection_file=args.session_file)
        if st.running:
            print("running")
            print(f"pid: {st.pid}")
            print(f"state: {st.state}")
            print(f"connection file: {st.connection_file}")
            return EXIT_OK
        print("not running")
        if st.stale:
            # "unreachable", not "dead": stale also covers a live-but-wedged pid
            # or a corrupt connection file.
            detail = f"kernel unreachable (pid {st.pid})" if st.pid is not None else "kernel unreachable"
            print(f"stale session files present ({detail}) — run 'docx-session stop' to clean up")
        return EXIT_NO_SESSION

    if args.command == "stop":
        if stop_session(connection_file=args.session_file):
            print("stopped")
            return EXIT_OK
        print("no session")
        return EXIT_NO_SESSION

    if args.command == "eval":
        try:
            res = eval_code(_read_code(args.expr), connection_file=args.session_file, timeout=args.timeout)
        except SessionDeadError as e:
            # Keep eval's stdout contract — one JSON envelope whenever session
            # files exist — even when the kernel is already gone.
            print(e, file=sys.stderr)
            res = EvalResult(status="dead")
        else:
            if res.status == "dead":  # died mid-eval: same stderr recovery hint as exec
                print(_KERNEL_DIED_MSG, file=sys.stderr)
        print(json.dumps(asdict(res)))
        return {"ok": EXIT_OK, "error": EXIT_ERROR, "timeout": EXIT_TIMEOUT, "dead": EXIT_KERNEL_DEAD}[res.status]

    # exec
    res = exec_code(_read_code(args.code), connection_file=args.session_file, timeout=args.timeout)

    if res.stdout:
        print(res.stdout, end="")
    if res.stderr:
        print(res.stderr, end="", file=sys.stderr)
    if res.result is not None:
        print(res.result)
    if res.status == "error":
        print(res.traceback, file=sys.stderr)
        return EXIT_ERROR
    if res.status == "dead":
        print(_KERNEL_DIED_MSG, file=sys.stderr)
        return EXIT_KERNEL_DEAD
    if res.status == "timeout":
        if res.started:
            message = f"Timed out after {args.timeout}s (kernel still running; the command may still finish)."
        else:
            message = (
                f"Timed out after {args.timeout}s while still queued — your code never started. The kernel is "
                "busy with an earlier command; wait for it, raise --timeout, or check 'docx-session status'."
            )
        print(message, file=sys.stderr)
        return EXIT_TIMEOUT
    if not res.started:
        # "ok" but never dequeued: the kernel discards everything queued behind
        # a command that raises, so this returned clean without running. Exit
        # code stays EXIT_OK (changing it is a contract change), which is
        # exactly why the warning has to be loud.
        print(
            "Warning: the kernel discarded this request without running it — it was queued behind a "
            "command that raised. Nothing was executed; re-send it.",
            file=sys.stderr,
        )
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
