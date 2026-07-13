"""Tests for the open-document guard.

Word writes a ``~$`` owner (lock) file next to any document it has open. Saving
over that live file races Word's own writes and can corrupt the document, so
``Document.save()`` / ``Workspace.save()`` refuse unless ``force=True``. A
PermissionError from the OS (Word holding the file on Windows) maps to the same
``DocumentOpenError``.
"""

import pytest
from conftest import find_ref

from docx_editor import Document, DocumentOpenError
from docx_editor.ooxml import pack
from docx_editor.workspace import owner_file_candidates


def test_owner_file_candidates_forms(tmp_path):
    """Both the full-name and two-char-truncated ~$ forms are produced."""
    candidates = owner_file_candidates(tmp_path / "Report.docx")
    names = {c.name for c in candidates}
    assert names == {"~$Report.docx", "~$port.docx"}
    assert all(c.parent == tmp_path for c in candidates)


@pytest.mark.parametrize("stub_index", [0, 1])
def test_guard_blocks_when_document_open(clean_workspace, stub_index):
    """A ~$ owner file next to the destination makes save() raise, leaving it untouched."""
    original_bytes = clean_workspace.read_bytes()
    stub = owner_file_candidates(clean_workspace)[stub_index]
    stub.write_bytes(b"")  # Word's owner file is opaque; presence is what matters.

    doc = Document.open(clean_workspace)
    ref = find_ref(doc, "fox")
    doc.replace("fox", "cat", paragraph=ref)

    with pytest.raises(DocumentOpenError):
        doc.save()

    assert clean_workspace.read_bytes() == original_bytes  # destination untouched
    doc.close()


def test_force_overrides_guard(clean_workspace):
    """force=True saves through the guard and the edit survives a reopen."""
    stub = owner_file_candidates(clean_workspace)[0]
    stub.write_bytes(b"")

    doc = Document.open(clean_workspace)
    ref = find_ref(doc, "fox")
    doc.replace("fox", "cat", paragraph=ref)
    doc.save(force=True)
    doc.close()

    reopened = Document.open(clean_workspace)
    text = reopened.get_visible_text()
    assert "cat" in text
    assert "fox" not in text
    reopened.close()


def test_unrelated_owner_file_does_not_block(clean_workspace):
    """A ~$ stub for a *different* document in the same folder is not a false positive."""
    (clean_workspace.parent / "~$unrelated.docx").write_bytes(b"")

    doc = Document.open(clean_workspace)
    saved = doc.save()
    assert saved == clean_workspace
    doc.close()


def test_saving_to_fresh_path_ignores_source_stub(clean_workspace, temp_dir):
    """A stub on the open source must not block saving a copy to a fresh destination."""
    owner_file_candidates(clean_workspace)[0].write_bytes(b"")  # source is "open"
    fresh = temp_dir / "copy.docx"

    doc = Document.open(clean_workspace)
    saved = doc.save(fresh)
    assert saved == fresh
    assert fresh.exists()
    doc.close()


def test_permission_error_maps_to_document_open_error(clean_workspace, monkeypatch):
    """A PermissionError during packing (Word holding the file) maps to DocumentOpenError."""

    def deny(*args, **kwargs):
        raise PermissionError("Word has the file open")

    monkeypatch.setattr(pack.os, "replace", deny)

    doc = Document.open(clean_workspace)
    with pytest.raises(DocumentOpenError):
        doc.save()
    doc.close()
