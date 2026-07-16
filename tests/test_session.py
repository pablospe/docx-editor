"""Tests for the persistent session module (docx_editor/session.py)."""

import io
import json
import subprocess
import sys
import threading
import time

import pytest

pytest.importorskip("jupyter_client")
pytest.importorskip("ipykernel")

from docx_editor.exceptions import SessionDeadError, SessionError  # noqa: E402
from docx_editor.session import (  # noqa: E402
    eval_code,
    exec_code,
    is_session_running,
    main,
    session_status,
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


class TestEval:
    """eval_code(): expression values come back as JSON, not display reprs."""

    def test_eval_simple_expression(self, session_conn):
        res = eval_code("1 + 1", connection_file=session_conn)
        assert res.status == "ok"
        assert res.value == 2
        assert res.serialized is True

    def test_eval_round_trips_unicode_and_quotes(self, session_conn):
        value = {"text": "double \" and single ' quotes — ünïcode", "nested": [1, {"k": [True, None]}]}
        # repr(value) is a valid Python expression that stresses the repr-embedding transport.
        res = eval_code(repr(value), connection_file=session_conn)
        assert res.status == "ok", res.traceback
        assert res.value == value
        assert res.serialized is True

    def test_eval_large_payload_round_trips(self, session_conn):
        """Guards the repr transport against any pretty-printer truncation/wrapping."""
        res = eval_code("list(range(10000))", connection_file=session_conn)
        assert res.status == "ok"
        assert res.value == list(range(10000))

    def test_eval_non_serializable_falls_back_to_repr(self, session_conn):
        res = eval_code("object()", connection_file=session_conn)
        assert res.status == "ok"
        assert res.serialized is False
        assert isinstance(res.value, str)
        assert "object" in res.value

    def test_eval_non_finite_floats_fall_back_to_repr(self, session_conn):
        """NaN/Infinity have no RFC-8259 form — they must not leak into the envelope."""
        res = eval_code("[float('nan'), float('inf')]", connection_file=session_conn)
        assert res.status == "ok"
        assert res.serialized is False
        assert res.value == "[nan, inf]"

    def test_eval_statement_is_a_syntax_error(self, session_conn):
        res = eval_code("some_var = 5", connection_file=session_conn)
        assert res.status == "error"
        assert res.traceback is not None
        assert "SyntaxError" in res.traceback

    def test_eval_captures_side_effect_stdout(self, session_conn):
        res = eval_code("print('noise') or 7", connection_file=session_conn)
        assert res.status == "ok"
        assert res.value == 7
        assert "noise" in res.stdout

    def test_eval_sees_state_from_prior_exec(self, session_conn):
        assert exec_code("eval_state = {'a': 1}", connection_file=session_conn).status == "ok"
        res = eval_code("eval_state", connection_file=session_conn)
        assert res.value == {"a": 1}

    def test_main_eval_prints_json_envelope(self, session_conn, capsys):
        assert main(["eval", "2 + 3", "--session-file", str(session_conn)]) == 0
        envelope = json.loads(capsys.readouterr().out)
        assert envelope["status"] == "ok"
        assert envelope["value"] == 5
        assert envelope["serialized"] is True

    def test_main_eval_error_envelope(self, session_conn, capsys):
        assert main(["eval", "1 / 0", "--session-file", str(session_conn)]) == 1
        envelope = json.loads(capsys.readouterr().out)
        assert envelope["status"] == "error"
        assert "ZeroDivisionError" in envelope["traceback"]

    def test_main_eval_without_session(self, tmp_path, capsys):
        assert main(["eval", "1 + 1", "--session-file", str(tmp_path / "nope.json")]) == 3
        captured = capsys.readouterr()
        assert captured.out == ""  # no envelope without a session
        assert "docx-session start" in captured.err


class TestSessionStatus:
    """session_status(): richer detail than the boolean is_session_running()."""

    def test_status_idle_session(self, session_conn):
        st = session_status(session_conn)
        assert st.running is True
        assert st.state == "idle"
        assert st.pid == int(session_conn.with_suffix(".pid").read_text(encoding="utf-8"))
        assert st.connection_file == session_conn
        assert st.stale is False

    def test_status_no_session(self, tmp_path):
        st = session_status(tmp_path / "nope.json")
        assert st.running is False
        assert st.pid is None
        assert st.state is None
        assert st.stale is False

    def test_main_status_prints_details(self, session_conn, capsys):
        assert main(["status", "--session-file", str(session_conn)]) == 0
        out = capsys.readouterr().out
        assert out.splitlines()[0] == "running"
        assert "pid: " in out
        assert "state: idle" in out
        assert f"connection file: {session_conn}" in out


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

    def test_busy_kernel_status_reports_busy(self, busy_conn, capsys):
        st = session_status(busy_conn)
        assert st.running is True
        assert st.state == "busy"
        assert st.stale is False
        assert main(["status", "--session-file", str(busy_conn)]) == 0
        assert "state: busy" in capsys.readouterr().out


class TestKernelDeath:
    """A kernel that dies mid-exec must be reported dead, not 'still running'."""

    @pytest.fixture
    def dead_conn(self, tmp_path):
        """A session whose kernel SIGKILLed itself mid-exec."""
        conn = tmp_path / "kernel.json"
        start_session(conn)
        started = time.monotonic()
        res = exec_code("import os; os.kill(os.getpid(), 9)", connection_file=conn, timeout=30.0)
        elapsed = time.monotonic() - started
        assert res.status == "dead"
        # The silence probe (or pid fast path) must beat the 30s timeout by far.
        assert elapsed < 20.0, f"death detection took {elapsed:.1f}s"
        yield conn
        stop_session(conn)

    def test_dead_kernel_library_surface(self, dead_conn):
        with pytest.raises(SessionDeadError, match="docx-session stop"):
            exec_code("1 + 1", connection_file=dead_conn)
        st = session_status(dead_conn)
        assert st.running is False
        assert st.stale is True
        # stop still cleans the stale files up:
        assert stop_session(dead_conn) is True
        assert dead_conn.exists() is False
        assert dead_conn.with_suffix(".pid").exists() is False

    def test_dead_kernel_cli_surface(self, dead_conn, capsys):
        sf = ["--session-file", str(dead_conn)]

        assert main(["exec", "1 + 1", *sf]) == 4
        assert "docx-session stop" in capsys.readouterr().err

        assert main(["eval", "1 + 1", *sf]) == 4
        captured = capsys.readouterr()
        assert json.loads(captured.out)["status"] == "dead"

        assert main(["status", *sf]) == 3
        out = capsys.readouterr().out
        assert out.splitlines()[0] == "not running"
        assert "stale session files present" in out

        assert main(["stop", *sf]) == 0

    def test_eval_mid_exec_death_prints_recovery_hint(self, tmp_path, capsys):
        """Dying mid-eval must give the same stderr hint as the pre-checked dead path."""
        conn = tmp_path / "kernel.json"
        start_session(conn)
        try:
            expr = "__import__('os').kill(__import__('os').getpid(), 9)"
            assert main(["eval", expr, "--session-file", str(conn), "--timeout", "30"]) == 4
            captured = capsys.readouterr()
            assert json.loads(captured.out)["status"] == "dead"
            assert "docx-session stop" in captured.err
        finally:
            stop_session(conn)


class TestStdinCode:
    """exec/eval accept '-' to read the code from stdin — no shell quoting to fight."""

    def test_main_exec_stdin_multiline_mixed_quotes(self, session_conn, capsys, monkeypatch):
        code = "\n".join([
            "a = 'single'",
            'b = "double"',
            'print(f"{a} {b}")',
        ])
        monkeypatch.setattr("sys.stdin", io.StringIO(code))
        assert main(["exec", "-", "--session-file", str(session_conn)]) == 0
        assert "single double" in capsys.readouterr().out

    def test_main_eval_stdin(self, session_conn, capsys, monkeypatch):
        monkeypatch.setattr("sys.stdin", io.StringIO("{'k': 'v'}\n"))
        assert main(["eval", "-", "--session-file", str(session_conn)]) == 0
        envelope = json.loads(capsys.readouterr().out)
        assert envelope["value"] == {"k": "v"}

    def test_subprocess_exec_stdin(self, session_conn):
        """End-to-end through a real pipe, mirroring the documented heredoc pattern."""
        code = "sp_a = 'via'\nsp_b = \"stdin\"\nprint(sp_a, sp_b)\n"
        proc = subprocess.run(
            [sys.executable, "-m", "docx_editor.session", "exec", "-", "--session-file", str(session_conn)],
            input=code,
            capture_output=True,
            text=True,
        )
        assert proc.returncode == 0, proc.stderr
        assert "via stdin" in proc.stdout


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
