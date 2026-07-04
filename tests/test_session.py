"""Tests for the persistent session module (docx_editor/session.py)."""

import pytest

pytest.importorskip("jupyter_client")
pytest.importorskip("ipykernel")

from docx_editor.session import (  # noqa: E402
    exec_code,
    is_session_running,
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
    with pytest.raises(RuntimeError, match="already running"):
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
    assert int(r2.result) > 0
    r3 = exec_code("doc.close()", connection_file=session_conn)
    assert r3.status == "ok"
