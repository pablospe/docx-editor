"""Persistent Jupyter kernel session for multi-step document editing.

Keeps documents open across many small commands (AI-agent friendly) instead
of re-opening them in one-off scripts. Requires the optional extra:

    pip install docx-editor[session]

CLI (see main()):
    docx-session start | exec "code" | status | stop
"""

import os
import re
import signal
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from queue import Empty

DEFAULT_CONNECTION_FILE = Path.home() / ".cache" / "docx-editor" / "kernel.json"

_EXTRA_HINT = "Session mode requires extra dependencies: pip install 'docx-editor[session]'"


def _client(connection_file: Path):
    """Return a connected BlockingKernelClient for the session."""
    try:
        from jupyter_client import BlockingKernelClient
    except ImportError as e:
        raise ImportError(_EXTRA_HINT) from e

    kc = BlockingKernelClient(connection_file=str(connection_file))
    kc.load_connection_file()
    kc.start_channels()
    return kc


def _pid_file(connection_file: Path) -> Path:
    return connection_file.with_suffix(".pid")


def _pid_alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True


def start_session(connection_file: Path = DEFAULT_CONNECTION_FILE, timeout: float = 30.0) -> int:
    """Start a detached IPython kernel and wait until it answers.

    Returns:
        PID of the kernel process.

    Raises:
        RuntimeError: If a session is already running or the kernel fails to start.
    """
    if is_session_running(connection_file):
        raise RuntimeError(f"Session already running (connection file: {connection_file})")

    connection_file.parent.mkdir(parents=True, exist_ok=True)
    connection_file.unlink(missing_ok=True)

    proc = subprocess.Popen(
        [sys.executable, "-m", "ipykernel_launcher", "-f", str(connection_file)],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        # Detach on POSIX so the kernel outlives this CLI invocation.
        start_new_session=(os.name == "posix"),
    )
    _pid_file(connection_file).write_text(str(proc.pid), encoding="utf-8")

    deadline = time.monotonic() + timeout
    while not (connection_file.exists() and connection_file.stat().st_size > 0):
        if proc.poll() is not None:
            raise RuntimeError(f"Kernel process exited during startup (code {proc.returncode})")
        if time.monotonic() > deadline:
            proc.kill()
            raise RuntimeError(f"Kernel did not start within {timeout}s")
        time.sleep(0.1)

    kc = _client(connection_file)
    try:
        kc.wait_for_ready(timeout=max(1.0, deadline - time.monotonic()))
    finally:
        kc.stop_channels()
    return proc.pid


def is_session_running(connection_file: Path = DEFAULT_CONNECTION_FILE, timeout: float = 2.0) -> bool:
    """True if a kernel is answering on this connection file."""
    if not connection_file.exists():
        return False
    kc = _client(connection_file)
    try:
        kc.wait_for_ready(timeout=timeout)
        return True
    except RuntimeError:
        return False
    finally:
        kc.stop_channels()


def stop_session(connection_file: Path = DEFAULT_CONNECTION_FILE) -> bool:
    """Shut down the kernel (graceful request, SIGTERM fallback).

    Returns:
        True if a session existed and was stopped.
    """
    if not connection_file.exists():
        return False

    kc = _client(connection_file)
    try:
        kc.shutdown()
    finally:
        kc.stop_channels()

    pid_file = _pid_file(connection_file)
    if pid_file.exists():
        pid = int(pid_file.read_text(encoding="utf-8"))
        deadline = time.monotonic() + 5.0
        while time.monotonic() < deadline and _pid_alive(pid):
            time.sleep(0.1)
        if _pid_alive(pid):
            os.kill(pid, signal.SIGTERM)

    connection_file.unlink(missing_ok=True)
    pid_file.unlink(missing_ok=True)
    return True


_ANSI_RE = re.compile(r"\x1b\[[0-9;]*m")


@dataclass
class ExecResult:
    """Outcome of one exec_code() call against the session kernel."""

    status: str  # "ok" | "error" | "timeout"
    stdout: str = ""
    stderr: str = ""
    result: str | None = None  # repr of the last expression, if any
    traceback: str | None = None  # ANSI-stripped traceback when status == "error"


def exec_code(code: str, connection_file: Path = DEFAULT_CONNECTION_FILE, timeout: float = 120.0) -> ExecResult:
    """Execute code in the session kernel and collect its output.

    Raises:
        FileNotFoundError: If no session connection file exists.
    """
    if not connection_file.exists():
        raise FileNotFoundError(f"No session found ({connection_file} missing). Run 'docx-session start' first.")

    kc = _client(connection_file)
    try:
        kc.wait_for_ready(timeout=10.0)
        msg_id = kc.execute(code)

        stdout_parts: list[str] = []
        stderr_parts: list[str] = []
        result: str | None = None
        traceback: str | None = None
        status = "ok"
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
