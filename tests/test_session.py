"""Tests for the persistent session module (docx_editor/session.py)."""

import pytest

pytest.importorskip("jupyter_client")
pytest.importorskip("ipykernel")

from docx_editor.session import (  # noqa: E402
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
