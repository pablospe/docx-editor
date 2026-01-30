"""Pytest fixtures for docx_editor tests."""

import shutil
import tempfile
from pathlib import Path

import defusedxml.minidom
import pytest

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


@pytest.fixture
def temp_docx(simple_docx, temp_dir) -> Path:
    """Copy simple.docx to a temp location for testing."""
    dest = temp_dir / "test_document.docx"
    shutil.copy(simple_docx, dest)
    return dest


@pytest.fixture
def clean_workspace(temp_docx):
    """Ensure no workspace exists for the test document."""
    workspace_dir = temp_docx.parent / ".docx"
    if workspace_dir.exists():
        shutil.rmtree(workspace_dir)
    yield temp_docx
    # Cleanup after test
    if workspace_dir.exists():
        shutil.rmtree(workspace_dir)
