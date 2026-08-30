"""Tests for the persistent session module (docx_editor/session.py)."""

import io
import json
import os
import re
import shutil
import signal
import socket
import subprocess
import sys
import textwrap
import threading
import time
import types
from contextlib import contextmanager
from pathlib import Path
from queue import Empty

import pytest

pytest.importorskip("jupyter_client")
pytest.importorskip("ipykernel")

from conftest import sweep_leaked_kernels  # noqa: E402

import docx_editor.session as session_mod  # noqa: E402
from docx_editor.exceptions import SessionDeadError, SessionError  # noqa: E402
from docx_editor.session import (  # noqa: E402
    EXIT_ERROR,
    START_TIMEOUT_ENV,
    ExecResult,
    _client,
    _kernel_alive,
    _start_timeout,
    eval_code,
    exec_code,
    is_session_running,
    main,
    session_status,
    start_session,
    stop_session,
)
from docx_editor.workspace import _pid_alive  # noqa: E402

# Start waits and busy-kernel handshakes in this module are bounded by the
# kernel start budget, so a loaded machine raises them all through
# DOCX_EDITOR_SESSION_START_TIMEOUT (CI sets 60). A few fixed windows remain
# where the tested behaviour is itself wall-clock (the 10 s silence probe in
# TestKernelDeath, the stdin and silent-exec bounds).
START_BUDGET = _start_timeout()
# For the one behaviour that is inherently wall-clock — "my own code overran":
# an idle kernel must dequeue a request within this long. 2 s locally, 12 s on CI.
OVERRUN_TIMEOUT = max(2.0, START_BUDGET / 10)


@contextmanager
def _occupied(conn: Path, code: str):
    """Kernel at ``conn`` is provably executing ``code`` for the whole block.

    The hog is sent on a throwaway client and the block starts only once the
    kernel has broadcast ``execute_input`` for it — the same "left the queue"
    signal ``exec_code`` uses for ``started``. A sleep in its place encodes a
    guess about dequeue latency that a loaded machine falsifies: the next
    request may be dequeued first, or the kernel may not be busy yet.
    """
    hog = _client(conn)
    try:
        # Control round trip first (as exec_code does): the reply proves the
        # client's sockets are connected, so the iopub broadcast is not missed.
        assert _kernel_alive(hog, timeout=START_BUDGET), "kernel stopped answering"
        msg_id = hog.execute(code)
        deadline = time.monotonic() + START_BUDGET
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                pytest.fail(f"kernel never began executing the hog within {START_BUDGET}s")
            try:
                msg = hog.get_iopub_msg(timeout=min(remaining, 1.0))
            except Empty:
                continue
            if msg["msg_type"] == "execute_input" and msg["parent_header"].get("msg_id") == msg_id:
                break
        yield hog
    finally:
        hog.stop_channels()


def _wait_for_reply(hog) -> None:
    """Block until the hog's execute_reply lands — its code has finished.

    The hog client sent exactly one shell request, so the first execute_reply
    on its shell channel is that one's.
    """
    deadline = time.monotonic() + START_BUDGET
    while True:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            pytest.fail(f"hog never finished within {START_BUDGET}s")
        try:
            msg = hog.get_shell_msg(timeout=min(remaining, 1.0))
        except Empty:
            continue
        if msg["msg_type"] == "execute_reply":
            return


@pytest.fixture(scope="module")
def session_conn(tmp_path_factory):
    """One kernel shared by the read-only tests in this module.

    Deliberately no retry: pytest re-raises a failed module fixture for every
    later test in the module, and a start that exhausts the budget explains
    itself in that error. Retrying would hide the diagnosis.
    """
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


def _launch_fake_kernel(monkeypatch, script: str, *args: str) -> list[subprocess.Popen]:
    """Make ``start_session`` spawn ``python -c script *args`` in place of ipykernel.

    Patches the module attribute, not ``subprocess.Popen`` itself: conftest's
    kernel reaper wraps the real class for the whole session. Returns the list
    the spawned process is appended to; the caller reaps it with ``_reap``.
    """
    import docx_editor.session as session_mod

    spawned: list[subprocess.Popen] = []

    def fake_popen(argv, **kwargs):
        proc = subprocess.Popen([sys.executable, "-c", script, *args], **kwargs)
        spawned.append(proc)
        return proc

    fake_subprocess = types.SimpleNamespace(Popen=fake_popen, DEVNULL=subprocess.DEVNULL, PIPE=subprocess.PIPE)
    monkeypatch.setattr(session_mod, "subprocess", fake_subprocess)
    return spawned


def _reap(spawned: list[subprocess.Popen]) -> None:
    for proc in spawned:
        if proc.poll() is None:
            proc.kill()
            proc.wait()


def _free_ports(n: int) -> list[int]:
    """``n`` distinct ports nobody is listening on right now."""
    socks = [socket.socket() for _ in range(n)]
    for sock in socks:
        sock.bind(("127.0.0.1", 0))
    ports = [sock.getsockname()[1] for sock in socks]
    for sock in socks:
        sock.close()
    return ports


def test_start_session_unreadable_connection_file_stops_kernel(tmp_path, monkeypatch):
    """A kernel whose connection file cannot be parsed is killed, not leaked (ROADMAP.md #78).

    jupyter_client writes kernel.json non-atomically, so the startup loop's
    "exists and non-empty" check can pass on a half-written file. The stand-in
    kernel writes exactly such a file and then idles like a live kernel would.
    """
    conn = tmp_path / "kernel.json"
    script = "import sys, time; open(sys.argv[1], 'w').write('{\"shell_port\": 1'); time.sleep(60)"
    spawned = _launch_fake_kernel(monkeypatch, script, str(conn))

    try:
        with pytest.raises(SessionError, match="unreadable connection file"):
            start_session(conn, timeout=5.0)
        assert len(spawned) == 1
        assert spawned[0].wait(timeout=5) is not None
        assert not conn.exists()
        assert not conn.with_suffix(".pid").exists()
    finally:
        _reap(spawned)


def test_start_session_unreadable_connection_file_kernel_already_exited(tmp_path, monkeypatch):
    """The cleanup guard must not kill a kernel that has already exited on its own.

    The stand-in writes a truncated connection file, then waits for a "go"
    file before exiting; a wrapper around ``_client`` creates that file and
    waits for the exit before the real ``_client`` fails on the truncated
    JSON. The guard therefore runs with ``proc.poll()`` already set — the arm
    the idling stand-in above never reaches — and must only remove the files.
    """
    import docx_editor.session as session_mod

    conn = tmp_path / "kernel.json"
    go = tmp_path / "go"
    script = textwrap.dedent(
        """
        import os, sys, time
        open(sys.argv[1], "w").write('{"shell_port": 1')
        while not os.path.exists(sys.argv[2]):
            time.sleep(0.05)
        """
    )
    spawned = _launch_fake_kernel(monkeypatch, script, str(conn), str(go))
    real_client = session_mod._client

    def client_after_kernel_exit(connection_file):
        go.touch()
        spawned[0].wait(timeout=10)
        return real_client(connection_file)

    monkeypatch.setattr(session_mod, "_client", client_after_kernel_exit)

    try:
        with pytest.raises(SessionError, match="unreadable connection file"):
            start_session(conn, timeout=5.0)
        assert spawned[0].returncode == 0  # exited by itself; the guard did not kill it
        assert not conn.exists()
        assert not conn.with_suffix(".pid").exists()
    finally:
        _reap(spawned)


def test_start_session_kernel_never_ready_stops_kernel(tmp_path, monkeypatch):
    """A kernel that writes a valid connection file but never answers is killed, not leaked.

    The file names ports nobody listens on: ``_client`` connects (ZeroMQ
    connects lazily) and ``_kernel_answers_shell`` times out, so
    ``start_session`` must report the kernel as not ready, stop the
    still-running process and remove both files.
    """
    conn = tmp_path / "kernel.json"
    shell, iopub, stdin, control, hb = _free_ports(5)
    info = {
        "transport": "tcp",
        "ip": "127.0.0.1",
        "key": "",
        "signature_scheme": "hmac-sha256",
        "kernel_name": "",
        "shell_port": shell,
        "iopub_port": iopub,
        "stdin_port": stdin,
        "control_port": control,
        "hb_port": hb,
    }
    script = "import sys, time; open(sys.argv[1], 'w').write(sys.argv[2]); time.sleep(60)"
    spawned = _launch_fake_kernel(monkeypatch, script, str(conn), json.dumps(info))

    try:
        with pytest.raises(SessionError, match="did not become ready"):
            start_session(conn, timeout=1.0)
        assert spawned[0].wait(timeout=5) != 0  # killed by the guard, not a natural exit
        assert not conn.exists()
        assert not conn.with_suffix(".pid").exists()
    finally:
        _reap(spawned)


# A kernel stand-in that answers kernel_info on the *shell* channel only. Its
# control channel is bound but never serviced — exactly what ipykernel looks
# like to us when it drops a control request that arrived mid-initialization
# (ROADMAP.md #82). Bind precedes writing the connection file, as in a real
# kernel; a port collision therefore surfaces as a bind traceback on stderr and
# a non-zero exit, not as a mystery timeout.
_SHELL_ONLY_KERNEL = textwrap.dedent(
    """
    import json, sys, zmq
    from jupyter_client.session import Session

    conn_file, info = sys.argv[1], json.loads(sys.argv[2])
    ctx = zmq.Context()
    shell = ctx.socket(zmq.ROUTER)
    shell.bind(f"tcp://127.0.0.1:{info['shell_port']}")
    control = ctx.socket(zmq.ROUTER)
    control.bind(f"tcp://127.0.0.1:{info['control_port']}")
    stdin = ctx.socket(zmq.ROUTER)
    stdin.bind(f"tcp://127.0.0.1:{info['stdin_port']}")
    iopub = ctx.socket(zmq.PUB)
    iopub.bind(f"tcp://127.0.0.1:{info['iopub_port']}")

    with open(conn_file, "w") as fh:
        json.dump(info, fh)

    session = Session(key=b"")
    while True:
        idents, frames = session.feed_identities(shell.recv_multipart())
        msg = session.deserialize(frames, content=False)
        if msg["header"]["msg_type"] == "kernel_info_request":
            session.send(
                shell,
                "kernel_info_reply",
                content={"status": "ok", "protocol_version": "5.3"},
                parent=msg,
                ident=idents,
            )
    """
)


def test_start_session_ready_when_control_never_answers(tmp_path, monkeypatch):
    """Readiness is decided on the shell channel, so a silent control channel does not hang it.

    ipykernel registers the control callback before it assigns
    ``_control_lock``, so a probe sent the instant kernel.json appears can be
    dropped kernel-side; the one request we sent is then never answered and the
    whole remaining budget burns (ROADMAP.md #82). The stand-in reproduces that
    end state deterministically: shell answers, control never does.
    """
    conn = tmp_path / "kernel.json"
    shell, iopub, stdin, control, hb = _free_ports(5)
    info = {
        "transport": "tcp",
        "ip": "127.0.0.1",
        "key": "",
        "signature_scheme": "hmac-sha256",
        "kernel_name": "",
        "shell_port": shell,
        "iopub_port": iopub,
        "stdin_port": stdin,
        "control_port": control,
        "hb_port": hb,
    }
    spawned = _launch_fake_kernel(monkeypatch, _SHELL_ONLY_KERNEL, str(conn), json.dumps(info))

    try:
        pid = start_session(conn, timeout=START_BUDGET)
        assert pid == spawned[0].pid
        assert conn.exists()
        assert conn.with_suffix(".pid").exists()
    finally:
        # conftest's reaper only records spawns whose argv names
        # ipykernel_launcher, and _launch_fake_kernel spawns `python -c`, so it
        # never sees this stand-in at all. Clean up by hand.
        _reap(spawned)
        conn.unlink(missing_ok=True)
        conn.with_suffix(".pid").unlink(missing_ok=True)


def test_start_session_survives_control_probe_during_init(tmp_path):
    """A real kernel still starts when a control probe lands during its initialization.

    A guard, not a reproduction: the window is a single assignment inside
    ipykernel, so an intruding probe may well miss it on any given run — the
    deterministic half is
    ``test_start_session_ready_when_control_never_answers``. What this pins is
    that when the window *is* hit, start still succeeds and the kernel is
    usable afterwards.
    """
    conn = tmp_path / "kernel.json"
    stop = threading.Event()

    def probe_control_at_birth():
        while not stop.is_set():
            try:
                if conn.stat().st_size == 0:
                    time.sleep(0.0005)
                    continue
                kc = _client(conn)
            except (OSError, ValueError):
                # No connection file yet, or a half-written one — the two
                # states _client documents. Retry: giving up here would leave
                # the guard silently probing nothing.
                time.sleep(0.0005)
                continue
            try:
                kc.control_channel.send(kc.session.msg("kernel_info_request", {}))
                # stop_channels() closes the socket with linger=0, which
                # discards anything ZMQ has not yet handed to the kernel — and
                # on loopback the connect handshake is usually still in flight
                # right after send(). Hold the socket open until the start has
                # finished, or the probe this test exists to fire never
                # arrives.
                stop.wait(timeout=START_BUDGET)
            finally:
                kc.stop_channels()
            return

    intruder = threading.Thread(target=probe_control_at_birth, daemon=True)
    intruder.start()
    try:
        assert start_session(conn, timeout=START_BUDGET) > 0
        assert is_session_running(conn)
        assert eval_code("1 + 1", conn).value == 2
    finally:
        stop.set()
        intruder.join(timeout=5)
        assert not intruder.is_alive(), "intruder thread did not finish"
        stop_session(conn)


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
        code = f"import time; time.sleep({OVERRUN_TIMEOUT + 30})"
        res = exec_code(code, connection_file=conn, timeout=OVERRUN_TIMEOUT)
        assert res.status == "timeout"
        assert res.started is True  # our own code ran and overran
    finally:
        stop_session(conn)


def test_exec_ok_reports_started(session_conn):
    assert exec_code("1 + 1", connection_file=session_conn).started is True


def test_exec_error_reports_started(session_conn):
    """An exception still means the code executed."""
    assert exec_code("1 / 0", connection_file=session_conn).started is True


def test_exec_timeout_while_queued_reports_not_started(tmp_path):
    """A timeout waiting in the QUEUE is distinguishable from a timeout of your
    own running code (ISSUES.md #52): nothing of ours executed.

    Own kernel, occupied by a fire-and-forget sleep sent on a separate client,
    so our request cannot leave the queue before the clock runs out.
    """
    conn = tmp_path / "kernel.json"
    start_session(conn)
    try:
        with _occupied(conn, "import time; time.sleep(30)"):
            res = exec_code("1 + 1", connection_file=conn, timeout=2.0)
            assert res.status == "timeout"
            assert res.started is False  # never left the queue
            assert res.result is None
    finally:
        stop_session(conn)


def test_cli_distinguishes_the_two_timeouts(tmp_path, capsys):
    """Each timeout flavour gets its own stderr advice; exit code stays 2."""
    from docx_editor.session import EXIT_TIMEOUT

    conn = tmp_path / "kernel.json"
    start_session(conn)
    try:
        # Own code overran: an idle kernel dequeues it within the scaled timeout.
        sleep = f"import time; time.sleep({OVERRUN_TIMEOUT + 30})"
        code = main(["exec", sleep, "--session-file", str(conn), "--timeout", str(OVERRUN_TIMEOUT)])
        assert code == EXIT_TIMEOUT
        assert "kernel still running" in capsys.readouterr().err

        # Still queued: "kernel still running" above proved the sleep is
        # executing and it outlasts this call, so the request cannot dequeue.
        code = main(["exec", "1 + 1", "--session-file", str(conn), "--timeout", "2"])
        assert code == EXIT_TIMEOUT
        err = capsys.readouterr().err
        assert "still queued" in err
        assert "never started" in err
    finally:
        stop_session(conn)


def test_request_discarded_behind_a_raising_command(tmp_path):
    """An "ok" result with started=False ran NOTHING.

    ipykernel aborts everything still queued behind a command that raises, so a
    request sent while an earlier one was failing comes back clean and empty
    having done nothing. Without ``started`` that is indistinguishable from a
    successful silent statement — the caller would believe its edit applied.
    """
    conn = tmp_path / "kernel.json"
    start_session(conn)
    try:
        with _occupied(conn, "import time; time.sleep(3); raise RuntimeError('boom')") as hog:
            res = exec_code("discarded = 'I RAN'", connection_file=conn, timeout=30)
            assert res.status == "ok"
            assert res.started is False  # the tell: never dequeued
            assert res.stdout == ""

            # The reply lands just before ipykernel's post-error abort turn
            # (_abort_queues, one event-loop turn at the default
            # stop_on_error_timeout of 0); the eval below builds a fresh
            # client first, which comfortably outlasts it.
            _wait_for_reply(hog)  # the failing command is done
        assert eval_code("globals().get('discarded')", connection_file=conn).value is None
    finally:
        stop_session(conn)


def test_cli_warns_when_the_kernel_discards_the_request(tmp_path, capsys):
    """Exit code stays 0 (contract), so the warning has to carry the signal."""
    from docx_editor.session import EXIT_OK

    conn = tmp_path / "kernel.json"
    start_session(conn)
    try:
        with _occupied(conn, "import time; time.sleep(3); raise RuntimeError('boom')"):
            code = main(["exec", "discarded = 1", "--session-file", str(conn), "--timeout", "30"])
            assert code == EXIT_OK
            assert "discarded this request" in capsys.readouterr().err
    finally:
        stop_session(conn)


def test_eval_names_the_discard_instead_of_blaming_the_transport(tmp_path, capsys):
    """A discarded eval has no reply to decode — but that is not a transport
    fault, and saying so sends the caller debugging the wrong thing."""
    conn = tmp_path / "kernel.json"
    start_session(conn)
    try:
        with _occupied(conn, "import time; time.sleep(3); raise RuntimeError('boom')"):
            code = main(["eval", "1 + 1", "--session-file", str(conn), "--timeout", "30"])
            captured = capsys.readouterr()
            assert code == EXIT_ERROR
            assert "discarded this request" in captured.err
            assert "transport failed" not in captured.err
            assert captured.out == ""  # no envelope: nothing was evaluated
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

    def test_eval_transport_no_result_raises(self, tmp_path, monkeypatch):
        """A kernel reply with no execute_result is a transport bug, not a user error.

        ``started=True`` is what makes it one: the wrapper ran and still produced
        nothing. The same shape with ``started=False`` means the kernel discarded
        the request, which gets its own message (see
        test_eval_names_the_discard_instead_of_blaming_the_transport).
        """
        from docx_editor import session as session_mod

        monkeypatch.setattr(
            session_mod, "exec_code", lambda *a, **k: ExecResult(status="ok", result=None, started=True)
        )
        with pytest.raises(SessionError, match="returned no result"):
            eval_code("1 + 1", connection_file=tmp_path / "unused.json")

    def test_eval_transport_undecodable_reply_raises(self, tmp_path, monkeypatch):
        from docx_editor import session as session_mod

        monkeypatch.setattr(session_mod, "exec_code", lambda *a, **k: ExecResult(status="ok", result="{not a literal"))
        with pytest.raises(SessionError, match="could not decode"):
            eval_code("1 + 1", connection_file=tmp_path / "unused.json")


@pytest.fixture(scope="class")
def eval_doc(session_conn, tmp_path_factory):
    """A document opened as `doc` inside the shared kernel, closed at class teardown."""
    src = Path(__file__).parent / "test_data" / "simple.docx"
    path = tmp_path_factory.mktemp("evaldoc") / "doc.docx"
    shutil.copy(src, path)
    r = exec_code(
        f"from docx_editor import Document, EditOperation; doc = Document.open({str(path)!r}, author='Eval')",
        connection_file=session_conn,
    )
    assert r.status == "ok", r.traceback
    yield session_conn
    exec_code("doc.close()", connection_file=session_conn)


class TestEvalLibraryTypes:
    """Library dataclasses must arrive as real JSON objects, not opaque reprs.

    simple.docx: P1 "Sample Test Document" / P2 "The quick brown fox jumps
    over the lazy dog." / P3 "This is a sample document for testing the
    editing features." / P4 "A well-structured document helps ensure
    comprehensive test coverage."
    """

    def test_search_result_arrives_as_json(self, eval_doc):
        res = eval_code("doc.find_text('quick')", connection_file=eval_doc)
        assert res.status == "ok", res.traceback
        assert res.serialized is True
        sr = res.value
        assert sr["text"] == "quick"
        assert isinstance(sr["start"], int)
        assert isinstance(sr["end"], int)
        assert sr["paragraph_ref"].startswith("P2#")
        assert sr["paragraph_occurrence"] == 0
        assert sr["spans_revision"] is False
        assert sr["paragraph_index"] == 2

    def test_find_all_arrives_as_json_list(self, eval_doc):
        res = eval_code("doc.find_all('document')", connection_file=eval_doc)
        assert res.status == "ok", res.traceback
        assert res.serialized is True
        assert [m["paragraph_index"] for m in res.value] == [3, 4]
        assert all(m["text"] == "document" for m in res.value)

    def test_paragraph_info_arrives_as_json(self, eval_doc):
        res = eval_code("doc.get_paragraph(1)", connection_file=eval_doc)
        assert res.status == "ok", res.traceback
        assert res.serialized is True
        assert res.value["index"] == 1
        assert res.value["ref"].startswith("P1#")
        assert res.value["text"] == "Sample Test Document"

    def test_paragraph_location_arrives_as_json(self, eval_doc):
        res = eval_code("doc.get_paragraph_location(doc.get_paragraph(2).ref)", connection_file=eval_doc)
        assert res.status == "ok", res.traceback
        assert res.serialized is True
        loc = res.value
        assert loc["table"] is None
        assert loc["list"] is None
        assert loc["heading_path"] == []  # tuple arrives as a list
        assert loc["section"] == 1

    def test_revision_arrives_as_json_with_iso_date(self, eval_doc):
        r = exec_code("doc.replace('quick', 'swift', paragraph=doc.get_paragraph(2).ref)", connection_file=eval_doc)
        assert r.status == "ok", r.traceback
        res = eval_code("doc.list_revisions()", connection_file=eval_doc)
        assert res.status == "ok", res.traceback
        assert res.serialized is True
        assert len(res.value) > 0
        for rev in res.value:
            assert rev["type"] in ("insertion", "deletion")
            assert rev["author"] == "Eval"
            assert rev["date"] is None or "T" in rev["date"]  # ISO string, not a repr
            assert isinstance(rev["contains_ids"], list)

    def test_comment_arrives_as_json_with_nested_replies(self, eval_doc):
        r = exec_code(
            "cid = doc.add_comment('lazy dog', 'Check this.'); doc.reply_to_comment(cid, 'Agreed.')",
            connection_file=eval_doc,
        )
        assert r.status == "ok", r.traceback
        res = eval_code("doc.list_comments()", connection_file=eval_doc)
        assert res.status == "ok", res.traceback
        assert res.serialized is True
        comment = res.value[0]
        assert comment["text"] == "Check this."
        assert comment["resolved"] is False
        assert comment["date"] is None or "T" in comment["date"]
        assert comment["replies"][0]["text"] == "Agreed."  # nested dataclass, deep-converted


class TestEvalErrorEnvelope:
    """Expression errors carry {type, message, <recovery fields>} in `error`."""

    def test_hash_mismatch_structured_fields(self, eval_doc):
        res = eval_code("doc.replace('Sample', 'X', paragraph='P1#0000')", connection_file=eval_doc)
        assert res.status == "error"
        assert res.error is not None
        assert res.error["type"] == "HashMismatchError"
        assert res.error["paragraph_index"] == 1
        assert res.error["expected_hash"] == "0000"
        assert re.fullmatch(r"[0-9a-f]{4}", res.error["actual_hash"])
        assert "Sample Test Document" in res.error["paragraph_preview"]

    def test_text_not_found_fields(self, eval_doc):
        res = eval_code(
            "doc.replace('no-such-text', 'x', paragraph=doc.get_paragraph(3).ref)",
            connection_file=eval_doc,
        )
        assert res.status == "error"
        assert res.error is not None
        assert res.error["type"] == "TextNotFoundError"
        assert res.error["search_text"] == "no-such-text"
        assert res.error["paragraph_ref"].startswith("P3#")
        assert "sample document" in res.error["paragraph_preview"]

    def test_occurrence_overflow_fields(self, eval_doc):
        res = eval_code(
            "doc.replace('document', 'doc', paragraph=doc.get_paragraph(3).ref, occurrence=9)",
            connection_file=eval_doc,
        )
        assert res.status == "error"
        assert res.error is not None
        assert res.error["type"] == "TextNotFoundError"
        assert res.error["occurrence"] == 9
        assert res.error["total_occurrences"] == 1

    def test_path_field_coerced_to_string(self, eval_doc):
        res = eval_code("Document.open('/nope/missing.docx')", connection_file=eval_doc)
        assert res.status == "error"
        assert res.error is not None
        assert res.error["type"] == "DocumentNotFoundError"
        assert res.error["path"] == "/nope/missing.docx"

    def test_batch_error_nests_original_exception(self, eval_doc):
        res = eval_code(
            "doc.batch_edit([EditOperation.replace('quick', 'fast', paragraph='P2#0000')])",
            connection_file=eval_doc,
        )
        assert res.status == "error"
        assert res.error is not None
        assert res.error["type"] == "BatchOperationError"
        assert res.error["operation_index"] == 0
        original = res.error["original"]
        assert original["type"] == "HashMismatchError"
        assert original["expected_hash"] == "0000"
        assert re.fullmatch(r"[0-9a-f]{4}", original["actual_hash"])

    def test_non_library_error_has_no_stray_fields(self, eval_doc):
        res = eval_code("1 / 0", connection_file=eval_doc)
        assert res.status == "error"
        assert res.error == {"type": "ZeroDivisionError", "message": "division by zero"}

    def test_session_survives_enveloped_error(self, eval_doc):
        assert eval_code("doc.replace('x', 'y', paragraph='P1#0000')", connection_file=eval_doc).status == "error"
        res = eval_code("len(doc.list_paragraphs())", connection_file=eval_doc)
        assert res.status == "ok"
        assert res.value == 4

    def test_main_eval_error_envelope_has_structured_fields(self, eval_doc, capsys):
        expr = "doc.replace('Sample', 'X', paragraph='P1#0000')"
        assert main(["eval", expr, "--session-file", str(eval_doc)]) == 1
        envelope = json.loads(capsys.readouterr().out)
        assert envelope["status"] == "error"
        assert envelope["value"] is None
        assert envelope["serialized"] is False
        assert envelope["error"]["type"] == "HashMismatchError"
        assert envelope["error"]["expected_hash"] == "0000"
        assert "HashMismatchError" in envelope["traceback"]

    def test_eval_error_traceback_is_compact_and_path_free(self, eval_doc):
        res = eval_code("doc.replace('Sample', 'X', paragraph='P1#0000')", connection_file=eval_doc)
        assert res.status == "error"
        assert res.traceback is not None
        assert "<docx-session eval>" in res.traceback
        assert "docx_editor/" in res.traceback  # library frames keep their relative path
        assert not re.search(r"[/~]\S+[/\\]docx_editor[/\\]", res.traceback)  # ...but no absolute prefix
        assert "site-packages" not in res.traceback
        assert len(res.traceback) < 1500  # plain traceback, not IPython's multi-kilobyte one

    def test_exec_error_traceback_is_path_free(self, eval_doc):
        res = exec_code("doc.replace('Sample', 'X', paragraph='P1#0000')", connection_file=eval_doc)
        assert res.status == "error"
        assert res.traceback is not None
        assert "HashMismatchError" in res.traceback
        assert not re.search(r"[/~]\S+[/\\]docx_editor[/\\]", res.traceback)
        assert "site-packages" not in res.traceback

    def test_eval_error_rewrites_prior_cell_frames(self, eval_doc):
        """A frame from a function defined in an earlier exec cell carries the
        ipykernel cell-file path (/tmp/ipykernel_<pid>/<n>.py) — rewritten."""
        assert exec_code("def _boom():\n    return 1 / 0", connection_file=eval_doc).status == "ok"
        res = eval_code("_boom()", connection_file=eval_doc)
        assert res.status == "error"
        assert res.error is not None
        assert res.traceback is not None
        assert res.error["type"] == "ZeroDivisionError"
        assert "<session-cell>" in res.traceback
        assert "ipykernel_" not in res.traceback


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

    def test_status_corrupt_connection_file_is_stale(self, tmp_path):
        conn = tmp_path / "kernel.json"
        conn.write_text("not json", encoding="utf-8")
        conn.with_suffix(".pid").write_text("99999999", encoding="utf-8")
        st = session_status(conn)
        assert st.running is False
        assert st.stale is True
        assert st.pid == 99999999

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
        try:
            # The kernel is executing the sleep for the whole test (10 s: long
            # enough to outlast every probe, short enough for the queued exec
            # below to complete); stop_session then shuts it down while busy.
            with _occupied(conn, "import time; time.sleep(10)"):
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
        # The hog sleeps 10 s; the budget on top covers a loaded machine's dequeue.
        res = exec_code("1 + 1", connection_file=busy_conn, timeout=10 + START_BUDGET)
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

    def test_exec_mid_exec_death_cli_before_silence_probe(self, tmp_path, capsys):
        """A timeout shorter than the silence probe still reports dead (deadline check)."""
        conn = tmp_path / "kernel.json"
        start_session(conn)
        try:
            code = "import os; os.kill(os.getpid(), 9)"
            assert main(["exec", code, "--session-file", str(conn), "--timeout", "8"]) == 4
            assert "docx-session stop" in capsys.readouterr().err
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


def test_exec_silent_code_survives_silence_probe(tmp_path):
    """>10s of iopub silence with a live kernel must not be misreported dead."""
    conn = tmp_path / "kernel.json"
    start_session(conn)
    try:
        res = exec_code("import time; time.sleep(12); 'done'", connection_file=conn, timeout=60.0)
        assert res.status == "ok"
        assert res.result == "'done'"
    finally:
        stop_session(conn)


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


def test_stop_session_honours_the_shutdown_ack(tmp_path, monkeypatch):
    """Graceful shutdown must be acknowledged, not silently dropped.

    The old code fired shutdown() then tore the socket down before it flushed, so
    every stop fell through to the SIGTERM fallback. Wall-clock cannot tell that
    apart from a loaded machine; what can is the pair "no signal was ever sent"
    and "the kernel is gone anyway" — a dropped request leaves it alive, and
    the fallback would have to signal it.
    """
    conn = tmp_path / "kernel.json"
    pid = start_session(conn)
    signalled: list[tuple[int, int]] = []
    real_kill = os.kill

    def spying_kill(target: int, sig: int) -> None:
        # _pid_alive probes with signal 0, and unrelated components may signal
        # other pids; only real signals aimed at this kernel count.
        if sig != 0 and target == pid:
            signalled.append((target, sig))
        real_kill(target, sig)

    monkeypatch.setattr(os, "kill", spying_kill)
    try:
        assert stop_session(conn, timeout=START_BUDGET) is True
        assert signalled == [], "graceful shutdown fell through to a signal"
        assert not _pid_alive(pid, reap=True), "kernel outlived an acknowledged shutdown"
    finally:
        # stop_session has unlinked the connection file, so a kernel a regression
        # leaves behind is invisible to the reaper — kill it here.
        if _pid_alive(pid, reap=True):
            real_kill(pid, signal.SIGKILL)


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


# A kernel that binds its sockets and writes the connection file but never
# serves any channel: the exact shape of every start failure in the
# ROADMAP.md #72 CI logs (ipykernel writes kernel.json before importing IPython
# and wiring its dispatchers). Ports 1-5 are unserved on loopback; ZMQ retries
# the connect silently, so the kernel_info request simply gets no reply.
_FAKE_KERNEL_SRC = """\
import json, os, sys, time
path = sys.argv[sys.argv.index("-f") + 1]
info = {
    "transport": "tcp", "ip": "127.0.0.1", "key": "", "signature_scheme": "hmac-sha256",
    "shell_port": 1, "iopub_port": 2, "stdin_port": 3, "control_port": 4, "hb_port": 5,
}
with open(path + ".tmp", "w") as f:
    json.dump(info, f)
os.replace(path + ".tmp", path)
print("fake kernel: sockets bound, never served", file=sys.stderr, flush=True)
time.sleep(60)
"""


def _assert_reported_kernel_dead(report: str) -> None:
    """The stand-in kernel's argv has no ``ipykernel_launcher``, so conftest's
    reaper cannot see it: every test must prove ``_abort`` killed it."""
    pid_match = re.search(r"Kernel pid (\d+)", report)
    assert pid_match is not None, report
    assert not _pid_alive(int(pid_match.group(1)), reap=True)


class TestStartFailureDiagnosis:
    """A start that exhausts its budget must say why, and the budget must be
    tunable for loaded machines (ROADMAP.md #72).

    The CI failures behind #72 all read "Kernel did not become ready within
    30s" — the connection file had appeared but the readiness probe never got
    an answer — with nothing to tell a slow runner from a broken one.
    """

    @pytest.fixture
    def fake_kernel(self, monkeypatch):
        monkeypatch.setattr(session_mod, "_KERNEL_COMMAND", (sys.executable, "-c", _FAKE_KERNEL_SRC))

    def test_start_timeout_explains_itself(self, tmp_path, fake_kernel):
        conn = tmp_path / "kernel.json"
        with pytest.raises(SessionError) as excinfo:
            start_session(conn, timeout=2.0)
        message = str(excinfo.value)

        assert "did not become ready within 2.0s" in message
        assert "kernel.json appeared after" in message
        assert "no kernel_info reply on the shell channel" in message
        assert "still alive at the timeout" in message
        assert "fake kernel: sockets bound, never served" in message
        assert "tcp 127.0.0.1, shell=1 control=4" in message
        assert START_TIMEOUT_ENV in message

        # Nothing left behind: no state files, and the fake is dead.
        assert not conn.exists()
        assert not conn.with_suffix(".pid").exists()
        _assert_reported_kernel_dead(message)

    def test_start_budget_comes_from_env(self, tmp_path, fake_kernel, monkeypatch):
        monkeypatch.setenv(START_TIMEOUT_ENV, "2")
        with pytest.raises(SessionError, match="did not become ready within 2.0s") as excinfo:
            start_session(tmp_path / "kernel.json")
        _assert_reported_kernel_dead(str(excinfo.value))

    def test_explicit_timeout_beats_env(self, tmp_path, fake_kernel, monkeypatch):
        monkeypatch.setenv(START_TIMEOUT_ENV, "2")
        with pytest.raises(SessionError, match="did not become ready within 1.0s") as excinfo:
            start_session(tmp_path / "kernel.json", timeout=1.0)
        _assert_reported_kernel_dead(str(excinfo.value))

    @pytest.mark.parametrize("value", ["soon", "0", "-5", "nan", "inf"])
    def test_invalid_start_budget_raises(self, tmp_path, monkeypatch, value):
        monkeypatch.setenv(START_TIMEOUT_ENV, value)
        # Were a kernel spawned anyway, this exits at once with a distinct error.
        monkeypatch.setattr(session_mod, "_KERNEL_COMMAND", (sys.executable, "-c", "raise SystemExit(99)"))
        conn = tmp_path / "kernel.json"
        with pytest.raises(SessionError, match=f"{START_TIMEOUT_ENV} must be a positive number") as excinfo:
            start_session(conn)
        assert repr(value) in str(excinfo.value)
        assert "exited during startup" not in str(excinfo.value)
        assert not conn.exists()
        assert not conn.with_suffix(".pid").exists()

    def test_blank_env_means_default(self, monkeypatch):
        monkeypatch.setenv(START_TIMEOUT_ENV, "  ")
        assert _start_timeout() == session_mod.DEFAULT_START_TIMEOUT

    def test_main_start_prints_the_diagnosis(self, tmp_path, fake_kernel, monkeypatch, capsys):
        """`docx-session start` on a loaded runner used to exit 1 with only
        "did not become ready" to show for it; the CLI must carry the report."""
        monkeypatch.setenv(START_TIMEOUT_ENV, "2")
        assert main(["start", "--session-file", str(tmp_path / "kernel.json")]) == EXIT_ERROR
        err = capsys.readouterr().err
        assert "did not become ready within 2.0s" in err
        assert "fake kernel: sockets bound, never served" in err
        assert START_TIMEOUT_ENV in err
        _assert_reported_kernel_dead(err)


class TestKernelReaping:
    """ISSUES.md #62: a kernel must not survive the test that started it.

    ``start_session`` detaches the kernel so it outlives the spawning process.
    That is correct for the CLI (the whole point is a session that persists
    across invocations) but means a test process killed mid-flight leaks a live
    kernel: unlike a failing test, a killed one never unwinds through its
    fixture teardown. Repeated over time this is what accumulated the orphaned
    kernels behind #62.
    """

    def test_kernel_orphaned_by_killed_owner_is_swept(self, tmp_path):
        conn = tmp_path / "kernel.json"
        # A stand-in for a test that starts a session and is then killed before
        # it can stop it. Runs in its own process so it can be SIGKILLed
        # without taking this test down with it.
        helper = (
            "import time\n"
            "from pathlib import Path\n"
            "from docx_editor.session import start_session\n"
            f"start_session(Path({str(conn)!r}))\n"
            "time.sleep(300)\n"
        )
        proc = subprocess.Popen(
            [sys.executable, "-c", helper],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            text=True,
        )
        try:
            # The helper's start_session honours the env budget; allow for it.
            deadline = time.monotonic() + START_BUDGET + 10
            while time.monotonic() < deadline:
                if proc.poll() is not None:
                    stderr = proc.stderr.read() if proc.stderr else ""
                    raise AssertionError(f"helper exited during startup: {stderr}")
                if conn.exists() and is_session_running(conn, timeout=2.0):
                    break
                time.sleep(0.2)
            else:
                raise AssertionError(f"kernel never became ready within {START_BUDGET + 10}s")

            # Kill the owner outright — no teardown, no stop_session.
            proc.kill()
            proc.wait(timeout=30)

            # The leak this fixture exists to catch: the kernel is detached, so
            # it is still alive with its owner gone.
            assert is_session_running(conn) is True, "expected an orphaned kernel to reproduce the leak"

            # Exactly what the session-scoped autouse fixture runs at teardown.
            assert sweep_leaked_kernels({conn}) == [conn]
            assert is_session_running(conn) is False
            assert conn.exists() is False
        finally:
            if proc.poll() is None:
                proc.kill()
                proc.wait()
            if conn.exists():
                stop_session(conn)

    def test_fixture_records_kernels_started_in_process(self, tmp_path, reap_leaked_kernels):
        """The recording half of the fixture actually captures a spawn.

        ``test_kernel_orphaned_by_killed_owner_is_swept`` starts its kernel in a
        *child* process, which the parent's ``Popen`` patch never observes, and
        then calls the sweep by hand — so without this test nothing exercises
        ``recording_init`` / ``_connection_file_from_argv``. If start_session's
        argv shape ever changed (``-f=<path>``, or launching via a module that
        binds ``from subprocess import Popen`` at import time), recording would
        silently become a no-op and no test would fail.
        """
        conn = (tmp_path / "kernel.json").resolve()
        start_session(conn)
        try:
            assert conn in reap_leaked_kernels
        finally:
            stop_session(conn)

    def test_sweep_ignores_already_stopped_sessions(self, tmp_path):
        """A kernel a test stopped itself is not re-swept (the common case)."""
        conn = tmp_path / "kernel.json"
        start_session(conn)
        assert stop_session(conn) is True
        assert sweep_leaked_kernels({conn}) == []

    def test_sweep_ignores_paths_that_never_had_a_kernel(self, tmp_path):
        assert sweep_leaked_kernels({tmp_path / "never.json"}) == []
