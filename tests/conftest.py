"""Pytest fixtures for docx_editor tests."""

import shutil
import tempfile
import zipfile
from pathlib import Path

import defusedxml.minidom
import pytest


def find_ref(doc, text):
    """Find the paragraph ref containing the given text."""
    for entry in doc.list_paragraphs():
        if text in entry:
            return entry.split("|")[0]
    raise ValueError(f"Paragraph containing '{text}' not found")


def replace_document_xml(src: Path, dest: Path, new_doc_xml: str) -> None:
    """Copy ``src`` to ``dest``, swapping ``word/document.xml`` for ``new_doc_xml``."""
    with (
        zipfile.ZipFile(src, "r") as z_in,
        zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z_out,
    ):
        for item in z_in.infolist():
            data = new_doc_xml.encode("utf-8") if item.filename == "word/document.xml" else z_in.read(item.filename)
            z_out.writestr(item, data)


NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def parse_paragraph(xml: str):
    """Parse XML string and return the first w:p element."""
    doc = defusedxml.minidom.parseString(f"<root {NS}>{xml}</root>")
    return doc.getElementsByTagName("w:p")[0]


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
