"""Pytest fixtures for docx_editor tests."""

import shutil
import subprocess
import tempfile
import warnings
import zipfile
from pathlib import Path
from xml.dom import minidom

import defusedxml.minidom
import pytest


def find_ref(doc, text):
    """Find the paragraph ref containing the given text."""
    for entry in doc.list_paragraphs(limit=None):
        if text in entry:
            return entry.split("|")[0]
    raise ValueError(f"Paragraph containing '{text}' not found")


def match_for(doc, text, **kwargs):
    """``find_text`` narrowed to a definite match.

    ``find_text`` returns ``SearchResult | None``, which the edit methods refuse
    (a None target must not become a silent no-op). Tests that know the text is
    present use this so the Optional is handled once, here, instead of asserting
    in every case.
    """
    match = doc.find_text(text, **kwargs)
    assert match is not None, f"no match for {text!r}"
    return match


def replace_docx_parts(src: Path, dest: Path, parts: dict[str, str | None]) -> None:
    """Copy ``src`` to ``dest``, swapping each archive part named in ``parts``.

    Keys are archive paths (e.g. ``"word/styles.xml"``); a ``None`` value
    drops the part from the copy entirely.
    """
    with (
        zipfile.ZipFile(src, "r") as z_in,
        zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z_out,
    ):
        for item in z_in.infolist():
            if item.filename in parts:
                new_content = parts[item.filename]
                if new_content is None:
                    continue
                z_out.writestr(item, new_content.encode("utf-8"))
            else:
                z_out.writestr(item, z_in.read(item.filename))


def replace_document_xml(src: Path, dest: Path, new_doc_xml: str) -> None:
    """Copy ``src`` to ``dest``, swapping ``word/document.xml`` for ``new_doc_xml``."""
    replace_docx_parts(src, dest, {"word/document.xml": new_doc_xml})


NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

# XML with an entity declaration: defusedxml refuses it (EntitiesForbidden),
# reproducing the ISSUES.md #35 parse-failure path (apache/poi XXE sample).
ENTITY_DTD_XML = '<!DOCTYPE r [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><r>&xxe;</r>'


def parse_paragraph(xml: str):
    """Parse XML string and return the first w:p element."""
    doc = defusedxml.minidom.parseString(f"<root {NS}>{xml}</root>")
    return doc.getElementsByTagName("w:p")[0]


def count_dom_walks(monkeypatch) -> list[str]:
    """Record the tag of every full-document getElementsByTagName call.

    Returns a list that grows as walks happen, so a test can pin that the
    number of full-DOM walks stays constant in the operation count. Only
    Document-level walks are counted; paragraph-local ones run on Elements.
    """
    walks: list[str] = []
    original = minidom.Document.getElementsByTagName

    def counting(self, name):
        walks.append(name)
        return original(self, name)

    monkeypatch.setattr(minidom.Document, "getElementsByTagName", counting)
    return walks


def _connection_file_from_argv(args) -> Path | None:
    """Return the ``-f <path>`` of an ipykernel_launcher command line, else None.

    ``start_session`` always spawns the kernel as
    ``[python, "-m", "ipykernel_launcher", "-f", <connection file>]``, so the
    connection file can be recovered from the argv alone — no cooperation from
    the code under test required.

    The path is resolved eagerly: cwd is correct now, but two tests chdir
    (``test_workspace.py``, ``test_document.py``), so a relative path stored
    here could resolve against a different directory at session teardown.
    """
    if isinstance(args, (str, bytes)):
        return None  # shell=True form; start_session never uses it.
    argv = [str(a) for a in args]
    if not any("ipykernel_launcher" in a for a in argv):
        return None
    try:
        return Path(argv[argv.index("-f") + 1]).resolve()
    except (ValueError, IndexError):
        return None


def sweep_leaked_kernels(connection_files) -> list[Path]:
    """Stop every kernel in ``connection_files`` that is still answering.

    ``start_session`` detaches the kernel (``start_new_session=True`` on POSIX)
    so it survives the process that spawned it. A test that fails, errors, or
    simply forgets to call ``stop_session`` therefore leaves a live kernel plus
    its connection and pid files behind; this sweep is the backstop.

    Scope, stated precisely: this runs at pytest teardown, so it covers a run
    that completes (however badly) and Ctrl-C, which unwinds finalizers. It
    does *not* cover the pytest process itself being SIGKILLed or OOM-killed —
    no in-process finalizer can. For that case the memory-capped
    ``systemd-run --user --scope`` invocation in CLAUDE.md is the real
    protection: the kernel stays inside the scope's cgroup, which systemd
    tears down with the scope.

    ``stop_session`` can shut a kernel down from its connection file alone, so
    sweeping reduces to "stop whatever still answers". Kernels a test already
    stopped are skipped: their connection file is gone, so they do not answer.

    Returns the paths actually swept, so this is directly assertable.
    """
    try:
        from docx_editor.session import DEFAULT_CONNECTION_FILE, is_session_running, stop_session
    except ImportError:
        return []  # [session] extra absent — nothing could have been started.

    swept: list[Path] = []
    for conn in connection_files:
        # Never touch the developer's own long-lived session.
        if conn == DEFAULT_CONNECTION_FILE.resolve():
            continue
        try:
            if is_session_running(conn, timeout=2.0):
                stop_session(conn, timeout=5.0)
                swept.append(conn)
        except Exception as exc:
            # A cleanup sweep must not fail the run it is cleaning up after,
            # but a kernel we could not stop is a real leak — say so loudly
            # rather than reporting a clean sweep.
            warnings.warn(f"could not reap kernel at {conn}: {exc!r}", stacklevel=2)
    return swept


@pytest.fixture(scope="session", autouse=True)
def reap_leaked_kernels():
    """Reap any ipykernel this test session started but never stopped.

    Wraps ``subprocess.Popen.__init__`` for the duration of the session,
    recording the connection file of every ``ipykernel_launcher`` spawn, then
    sweeps them all at teardown (see ``sweep_leaked_kernels`` for exactly what
    that does and does not cover).

    Two deliberate choices:

    * **Patch Popen, not start_session.** The session tests do
      ``from docx_editor.session import start_session``, binding the function
      object at import time, so patching the module attribute afterwards would
      miss those call sites. ``session.py`` always reaches Popen through the
      module, so this interception holds however the kernel was launched.
    * **Patch ``__init__``, not the class.** Rebinding ``subprocess.Popen`` to
      a function would break ``isinstance(x, subprocess.Popen)`` anywhere in
      the dependency tree (``TypeError: isinstance() arg 2 must be a type``).
      Patching the initialiser keeps the class object intact.

    Uses its own ``pytest.MonkeyPatch`` because the ``monkeypatch`` fixture is
    function-scoped and cannot be requested from a session-scoped fixture.

    Yields the set of recorded connection files; ``test_session.py`` asserts
    against it to prove the recording half actually works.
    """
    started: set[Path] = set()
    real_init = subprocess.Popen.__init__
    mp = pytest.MonkeyPatch()

    def recording_init(self, args, *popen_args, **popen_kwargs):
        # Materialise once: Popen accepts any iterable, and inspecting a bare
        # iterator here would exhaust it before the real __init__ sees it.
        if not isinstance(args, (str, bytes)):
            args = list(args)
        conn = _connection_file_from_argv(args)
        if conn is not None:
            started.add(conn)
        real_init(self, args, *popen_args, **popen_kwargs)

    mp.setattr(subprocess.Popen, "__init__", recording_init)
    try:
        yield started
    finally:
        mp.undo()
        sweep_leaked_kernels(started)


@pytest.fixture
def test_data_dir() -> Path:
    """Return the path to the test_data directory."""
    return Path(__file__).parent / "test_data"


@pytest.fixture
def simple_docx(test_data_dir) -> Path:
    """Return path to simple.docx test file."""
    return test_data_dir / "simple.docx"


@pytest.fixture
def temp_dir():
    """Create a temporary directory for test outputs."""
    temp = tempfile.mkdtemp(prefix="docx_editor_test_")
    yield Path(temp)
    shutil.rmtree(temp, ignore_errors=True)


@pytest.fixture(autouse=True)
def isolated_workspace_base(monkeypatch):
    """Isolate every test's workspace base from the real user cache.

    Points DOCX_EDITOR_WORKSPACE_DIR at a throwaway per-test directory so tests
    never write to ~/.cache/docx-editor/ and implicitly exercise the env-var
    resolution path.
    """
    base = tempfile.mkdtemp(prefix="docx_editor_ws_")
    monkeypatch.setenv("DOCX_EDITOR_WORKSPACE_DIR", base)
    yield Path(base)
    shutil.rmtree(base, ignore_errors=True)


@pytest.fixture
def temp_docx(simple_docx, temp_dir) -> Path:
    """Copy simple.docx to a temp location for testing."""
    dest = temp_dir / "test_document.docx"
    shutil.copy(simple_docx, dest)
    return dest


@pytest.fixture
def clean_workspace(temp_docx):
    """Alias for temp_docx, kept for backwards compatibility.

    Workspace isolation is handled by the autouse isolated_workspace_base
    fixture, so no manual cleanup is needed here.
    """
    return temp_docx
