"""Persistent Jupyter kernel session for multi-step document editing.

Keeps documents open across many small commands (AI-agent friendly) instead
of re-opening them in one-off scripts. Requires the optional extra:

    pip install docx-editor[session]

CLI (see main()):
    docx-session start | exec "code" | status | stop
"""

import argparse
import os
import re
import signal
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from queue import Empty
from typing import TYPE_CHECKING, Literal

from .exceptions import SessionError
from .workspace import _pid_alive

if TYPE_CHECKING:
    from jupyter_client import BlockingKernelClient

DEFAULT_CONNECTION_FILE = Path.home() / ".cache" / "docx-editor" / "kernel.json"

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


def _pid_file(connection_file: Path) -> Path:
    return connection_file.with_suffix(".pid")


def _read_pid(connection_file: Path) -> int | None:
    """PID recorded for this session, or None if absent/unreadable/corrupt."""
    try:
        return int(_pid_file(connection_file).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def start_session(connection_file: Path = DEFAULT_CONNECTION_FILE, timeout: float = 30.0) -> int:
    """Start a detached IPython kernel and wait until it answers.

    Args:
        connection_file: Where to write the kernel connection file
        timeout: Seconds to wait for the kernel to become ready

    Returns:
        PID of the kernel process.

    Raises:
        SessionError: If a session is already running or the kernel fails to start.
        ImportError: If the [session] extra is not installed.
    """
    if is_session_running(connection_file):
        raise SessionError(f"Session already running (connection file: {connection_file})")

    connection_file.parent.mkdir(parents=True, exist_ok=True)
    connection_file.unlink(missing_ok=True)

    proc = subprocess.Popen(
        [sys.executable, "-m", "ipykernel_launcher", "-f", str(connection_file)],
        stdout=subprocess.DEVNULL,
        # Keep stderr so a failed start (e.g. missing ipykernel) can explain itself.
        stderr=subprocess.PIPE,
        # Detach on POSIX so the kernel outlives this CLI invocation.
        start_new_session=(os.name == "posix"),
    )
    _pid_file(connection_file).write_text(str(proc.pid), encoding="utf-8")

    def _abort(message: str) -> SessionError:
        proc.kill()
        proc.wait()
        connection_file.unlink(missing_ok=True)
        _pid_file(connection_file).unlink(missing_ok=True)
        return SessionError(message)

    deadline = time.monotonic() + timeout
    while not (connection_file.exists() and connection_file.stat().st_size > 0):
        if proc.poll() is not None:
            stderr = proc.stderr.read().decode(errors="replace").strip() if proc.stderr else ""
            _pid_file(connection_file).unlink(missing_ok=True)
            hint = _EXTRA_HINT if "ipykernel" in stderr else ""
            detail = f"\n{stderr}" if stderr else ""
            raise SessionError(f"Kernel process exited during startup (code {proc.returncode}). {hint}{detail}".strip())
        if time.monotonic() > deadline:
            raise _abort(f"Kernel did not start within {timeout}s")
        time.sleep(0.1)

    kc = _client(connection_file)
    try:
        if not _kernel_alive(kc, timeout=max(1.0, deadline - time.monotonic())):
            raise _abort(f"Kernel did not become ready within {timeout}s")
    except BaseException:
        # Never leave a half-started kernel behind with no way to reach it.
        if proc.poll() is None:
            proc.kill()
            proc.wait()
        connection_file.unlink(missing_ok=True)
        _pid_file(connection_file).unlink(missing_ok=True)
        raise
    finally:
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


_ANSI_RE = re.compile(r"\x1b\[[0-9;]*m")


@dataclass
class ExecResult:
    """Outcome of one exec_code() call against the session kernel."""

    status: Literal["ok", "error", "timeout"]
    stdout: str = ""
    stderr: str = ""
    result: str | None = None  # repr of the last expression, if any
    traceback: str | None = None  # ANSI-stripped traceback when status == "error"


def exec_code(code: str, connection_file: Path = DEFAULT_CONNECTION_FILE, timeout: float = 120.0) -> ExecResult:
    """Execute code in the session kernel and collect its output.

    If the kernel is busy, the request queues behind the running one; `timeout`
    covers the whole wait.

    Args:
        code: Python source to execute in the session
        connection_file: Kernel connection file of the session
        timeout: Seconds to wait for the execution to finish

    Returns:
        ExecResult with status "ok", "error", or "timeout".

    Raises:
        FileNotFoundError: If no session connection file exists.
        SessionError: If the kernel is not answering.
        ImportError: If the [session] extra is not installed.
    """
    if not connection_file.exists():
        raise FileNotFoundError(f"No session found ({connection_file} missing). Run 'docx-session start' first.")

    try:
        kc = _client(connection_file)
    except (ValueError, OSError) as e:
        raise SessionError(f"Session connection file is unreadable ({connection_file}): {e}") from e

    try:
        if not _kernel_alive(kc, timeout=10.0):
            raise SessionError(
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
        status: Literal["ok", "error", "timeout"] = "ok"
        deadline = time.monotonic() + timeout

        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return ExecResult(status="timeout", stdout="".join(stdout_parts), stderr="".join(stderr_parts))
            try:
                msg = kc.get_iopub_msg(timeout=min(remaining, 1.0))
            except Empty:
                continue
            if msg.get("parent_header", {}).get("msg_id") != msg_id:
                continue

            msg_type = msg["msg_type"]
            content = msg["content"]
            if msg_type == "stream":
                target = stdout_parts if content["name"] == "stdout" else stderr_parts
                target.append(content["text"])
            elif msg_type in ("execute_result", "display_data"):
                result = content.get("data", {}).get("text/plain", result)
            elif msg_type == "error":
                status = "error"
                traceback = _ANSI_RE.sub("", "\n".join(content["traceback"]))
            elif msg_type == "status" and content["execution_state"] == "idle":
                break

        return ExecResult(
            status=status,
            stdout="".join(stdout_parts),
            stderr="".join(stderr_parts),
            result=result,
            traceback=traceback,
        )
    finally:
        kc.stop_channels()


EXIT_OK = 0
EXIT_ERROR = 1
EXIT_TIMEOUT = 2
EXIT_NO_SESSION = 3


def main(argv: list[str] | None = None) -> int:
    """CLI entry point for docx-session."""
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument(
        "--session-file",
        type=Path,
        default=DEFAULT_CONNECTION_FILE,
        help=f"Kernel connection file (default: {DEFAULT_CONNECTION_FILE})",
    )

    parser = argparse.ArgumentParser(
        prog="docx-session",
        description="Persistent Python session for multi-step .docx editing.",
    )
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("start", parents=[common], help="Start a background kernel")
    p_exec = sub.add_parser("exec", parents=[common], help="Execute code in the running kernel")
    p_exec.add_argument("code", help="Python code to execute")
    p_exec.add_argument("--timeout", type=float, default=120.0, help="Seconds to wait (default: 120)")
    sub.add_parser("status", parents=[common], help="Check whether the kernel is answering")
    sub.add_parser("stop", parents=[common], help="Shut the kernel down")
    args = parser.parse_args(argv)

    try:
        return _run(args)
    except FileNotFoundError as e:
        print(e, file=sys.stderr)
        return EXIT_NO_SESSION
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
        if is_session_running(connection_file=args.session_file):
            print("running")
            return EXIT_OK
        print("not running")
        return EXIT_NO_SESSION

    if args.command == "stop":
        if stop_session(connection_file=args.session_file):
            print("stopped")
            return EXIT_OK
        print("no session")
        return EXIT_NO_SESSION

    # exec
    res = exec_code(args.code, connection_file=args.session_file, timeout=args.timeout)

    if res.stdout:
        print(res.stdout, end="")
    if res.stderr:
        print(res.stderr, end="", file=sys.stderr)
    if res.result is not None:
        print(res.result)
    if res.status == "error":
        print(res.traceback, file=sys.stderr)
        return EXIT_ERROR
    if res.status == "timeout":
        print(
            f"Timed out after {args.timeout}s (kernel still running; the command may still finish).",
            file=sys.stderr,
        )
        return EXIT_TIMEOUT
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
