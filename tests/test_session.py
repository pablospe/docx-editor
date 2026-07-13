"""Tests for the persistent session module (docx_editor/session.py)."""

import subprocess
import sys
import threading
import time

import pytest

pytest.importorskip("jupyter_client")
pytest.importorskip("ipykernel")

from docx_editor.exceptions import SessionError  # noqa: E402
from docx_editor.session import (  # noqa: E402
    exec_code,
    is_session_running,
    main,
    start_session,
    stop_session,
)


@pytest.fixture(scope="module")
def session_conn(tmp_path_factory):
    """One kernel shared by the read-only tests in this module."""
    conn = tmp_path_factory.mktemp("session") / "kernel.json"
    start_session(conn)
    yield conn
    stop_session(conn)


def test_start_session_creates_connection_file(session_conn):
    assert session_conn.exists()
    assert is_session_running(session_conn)


def test_start_session_twice_raises(session_conn):
    with pytest.raises(SessionError, match="already running"):
        start_session(session_conn)


def test_is_session_running_false_without_connection_file(tmp_path):
    assert is_session_running(tmp_path / "nope.json") is False


def test_stop_session(tmp_path):
    conn = tmp_path / "kernel.json"
    start_session(conn)
    assert is_session_running(conn)
    assert stop_session(conn) is True
    assert conn.exists() is False
    assert is_session_running(conn) is False


def test_stop_session_without_session_returns_false(tmp_path):
    assert stop_session(tmp_path / "nope.json") is False


def test_exec_returns_expression_result(session_conn):
    res = exec_code("1 + 1", connection_file=session_conn)
    assert res.status == "ok"
    assert res.result == "2"


def test_exec_state_persists_between_calls(session_conn):
    assert exec_code("x = 41", connection_file=session_conn).status == "ok"
    res = exec_code("x + 1", connection_file=session_conn)
    assert res.result == "42"


def test_exec_captures_stdout(session_conn):
    res = exec_code("print('hello session')", connection_file=session_conn)
    assert res.status == "ok"
    assert "hello session" in res.stdout
    assert res.result is None


def test_exec_error_returns_traceback_and_session_survives(session_conn):
    res = exec_code("1 / 0", connection_file=session_conn)
    assert res.status == "error"
    assert res.traceback is not None
    assert "ZeroDivisionError" in res.traceback
    assert "\x1b[" not in res.traceback  # ANSI codes stripped
    # Session survives the exception:
    assert exec_code("2 + 2", connection_file=session_conn).result == "4"


def test_exec_without_session_raises(tmp_path):
    with pytest.raises(FileNotFoundError, match="docx-session start"):
        exec_code("1 + 1", connection_file=tmp_path / "nope.json")


def test_exec_timeout(tmp_path):
    # Own kernel: the timed-out sleep would queue behind later tests otherwise.
    conn = tmp_path / "kernel.json"
    start_session(conn)
    try:
        res = exec_code("import time; time.sleep(30)", connection_file=conn, timeout=2.0)
        assert res.status == "timeout"
    finally:
        stop_session(conn)


def test_exec_docx_editing_workflow(session_conn, temp_docx):
    """End-to-end: a document stays open across separate exec calls."""
    r1 = exec_code(
        f"from docx_editor import Document; doc = Document.open({str(temp_docx)!r}, author='Session')",
        connection_file=session_conn,
    )
    assert r1.status == "ok", r1.traceback
    r2 = exec_code("paras = doc.list_paragraphs(); len(paras)", connection_file=session_conn)
    assert r2.status == "ok"
    assert r2.result is not None
    assert int(r2.result) > 0
    r3 = exec_code("doc.close()", connection_file=session_conn)
    assert r3.status == "ok"


class TestBusyKernel:
    """A busy kernel must stay distinguishable from a dead one.

    The liveness probe rides the control channel; ipykernel serializes the *shell*
    channel behind the running execute_request, so a shell-based probe reports a
    busy kernel as dead — which let `start` spawn a second kernel over a live one
    and orphan it, still holding the user's open document.
    """

    @pytest.fixture
    def busy_conn(self, tmp_path):
        conn = tmp_path / "kernel.json"
        start_session(conn)
        # Fire an execution and leave it running for the duration of the test.
        t = threading.Thread(target=exec_code, args=("import time; time.sleep(10)", conn), daemon=True)
        t.start()
        time.sleep(1.5)  # let the kernel actually enter the busy state
        try:
            yield conn
        finally:
            stop_session(conn)

    def test_busy_kernel_reports_running(self, busy_conn):
        assert is_session_running(busy_conn) is True

    def test_start_refuses_to_clobber_busy_kernel(self, busy_conn):
        with pytest.raises(SessionError, match="already running"):
            start_session(busy_conn)

    def test_exec_queues_behind_busy_kernel(self, busy_conn):
        # Must not raise "Kernel didn't respond in 10 seconds" — it queues instead.
        res = exec_code("1 + 1", connection_file=busy_conn, timeout=30.0)
        assert res.status == "ok"
        assert res.result == "2"


def test_exec_stdin_does_not_wedge_session(tmp_path):
    """input() must fail cleanly, not park the kernel on an unanswered stdin request."""
    conn = tmp_path / "kernel.json"
    start_session(conn)
    try:
        res = exec_code("input('name? ')", connection_file=conn, timeout=15.0)
        assert res.status == "error"
        assert res.traceback is not None
        # The session survives and is immediately usable.
        assert exec_code("7 * 6", connection_file=conn, timeout=15.0).result == "42"
    finally:
        stop_session(conn)


def test_stop_session_is_prompt(tmp_path):
    """Graceful shutdown must be acknowledged, not silently dropped.

    The old code fired shutdown() then tore the socket down before it flushed, so
    every stop fell through to the 5s SIGTERM fallback.
    """
    conn = tmp_path / "kernel.json"
    start_session(conn)
    elapsed = time.monotonic()
    assert stop_session(conn) is True
    elapsed = time.monotonic() - elapsed
    assert elapsed < 3.0, f"stop_session took {elapsed:.2f}s — shutdown ack was not honored"


def test_stop_session_survives_corrupt_pid_file(tmp_path):
    """A truncated pid file must not crash stop or strand the state files."""
    conn = tmp_path / "kernel.json"
    start_session(conn)
    conn.with_suffix(".pid").write_text("", encoding="utf-8")
    assert stop_session(conn) is True
    assert conn.exists() is False
    assert conn.with_suffix(".pid").exists() is False


def test_main_full_lifecycle(tmp_path, capsys):
    conn = tmp_path / "kernel.json"
    sf = ["--session-file", str(conn)]

    assert main(["start", *sf]) == 0
    assert "Session started" in capsys.readouterr().out

    assert main(["status", *sf]) == 0
    assert "running" in capsys.readouterr().out

    assert main(["exec", "print('via cli'); 10 * 2", *sf]) == 0
    out = capsys.readouterr().out
    assert "via cli" in out
    assert "20" in out

    assert main(["exec", "1 / 0", *sf]) == 1
    assert "ZeroDivisionError" in capsys.readouterr().err

    assert main(["stop", *sf]) == 0
    assert main(["status", *sf]) == 3


def test_main_exec_without_session(tmp_path, capsys):
    assert main(["exec", "1 + 1", "--session-file", str(tmp_path / "nope.json")]) == 3
    assert "docx-session start" in capsys.readouterr().err


def test_module_entrypoint_runs():
    proc = subprocess.run(
        [sys.executable, "-m", "docx_editor.session", "--help"],
        capture_output=True,
        text=True,
    )
    assert proc.returncode == 0
    assert "exec" in proc.stdout
