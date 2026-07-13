"""Tests for the open-document guard.

Word writes a ``~$`` owner (lock) file next to any document it has open. Saving
over that live file races Word's own writes and can corrupt the document, so
``Document.save()`` / ``Workspace.save()`` refuse unless ``force=True``.

A PermissionError from the final replace maps to the same ``DocumentOpenError``,
but only on Windows — that is the one platform where a locked destination surfaces
that way. On POSIX, rename(2) over an open file always succeeds, so a denial there
means something else entirely and must not be reported as "the document is open".
"""

import os
from pathlib import Path

import pytest
from conftest import find_ref

from docx_editor import Document, DocumentOpenError
from docx_editor.ooxml import pack
from docx_editor.workspace import owner_file_candidates


class TestOwnerFileCandidates:
    """The ~$ names Word may have written for a given document."""

    def test_both_forms_for_a_normal_name(self, tmp_path):
        """Both the full-name and two-char-truncated ~$ forms are produced."""
        candidates = owner_file_candidates(tmp_path / "Report.docx")
        names = {c.name for c in candidates}
        assert names == {"~$Report.docx", "~$port.docx"}
        assert all(c.parent == tmp_path for c in candidates)

    def test_short_stem_yields_no_junk_candidate(self, tmp_path):
        """A short stem keeps the full name and produces no junk like '~$.docx'."""
        candidates = owner_file_candidates(tmp_path / "ab.docx")
        assert [c.name for c in candidates] == ["~$ab.docx"]


class TestOpenDocumentGuard:
    """save() refuses to overwrite a document that looks open in Word."""

    @pytest.mark.parametrize("stub_index", [0, 1])
    def test_guard_blocks_when_document_open(self, clean_workspace, stub_index):
        """A ~$ owner file next to the destination makes save() raise, leaving it untouched."""
        original_bytes = clean_workspace.read_bytes()
        stub = owner_file_candidates(clean_workspace)[stub_index]
        stub.write_bytes(b"")  # Word's owner file is opaque; presence is what matters.

        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "fox")
        doc.replace("fox", "cat", paragraph=ref)

        with pytest.raises(DocumentOpenError) as exc:
            doc.save()

        assert exc.value.owner_file == stub
        assert exc.value.path == Path(os.path.realpath(clean_workspace))
        assert clean_workspace.read_bytes() == original_bytes  # destination untouched
        doc.close()

    def test_guard_sees_stub_beside_a_symlinked_destination(self, clean_workspace, temp_dir):
        """Word writes its stub next to the name the user opened — the symlink.

        The save resolves the symlink to promote into the real file, so the guard has
        to check beside *both* names or it misses a genuinely open document.
        """
        link_dir = temp_dir / "links"
        link_dir.mkdir()
        link = link_dir / "link.docx"
        link.symlink_to(clean_workspace)
        (link_dir / "~$link.docx").write_bytes(b"")  # stub beside the symlink only

        doc = Document.open(clean_workspace)
        with pytest.raises(DocumentOpenError):
            doc.save(link)
        doc.close()

    def test_force_overrides_guard(self, clean_workspace):
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

    def test_unrelated_owner_file_does_not_block(self, clean_workspace):
        """A ~$ stub for a *different* document in the same folder is not a false positive."""
        (clean_workspace.parent / "~$unrelated.docx").write_bytes(b"")

        doc = Document.open(clean_workspace)
        saved = doc.save()
        assert saved == clean_workspace
        doc.close()

    def test_saving_to_fresh_path_ignores_source_stub(self, clean_workspace, temp_dir):
        """A stub on the open source must not block saving a copy to a fresh destination."""
        owner_file_candidates(clean_workspace)[0].write_bytes(b"")  # source is "open"
        fresh = temp_dir / "copy.docx"

        doc = Document.open(clean_workspace)
        saved = doc.save(fresh)
        assert saved == fresh
        assert fresh.exists()
        doc.close()


class TestPermissionErrorMapping:
    """Only a denied *replace*, and only on Windows, means "the document is open"."""

    def test_replace_denial_maps_to_document_open_error_on_windows(self, clean_workspace, monkeypatch):
        """On Windows, Word holding the destination surfaces as a PermissionError here."""
        original_bytes = clean_workspace.read_bytes()
        monkeypatch.setattr(pack.sys, "platform", "win32")

        def deny(*args, **kwargs):
            raise PermissionError("Word has the file open")

        monkeypatch.setattr(pack.os, "replace", deny)

        doc = Document.open(clean_workspace)
        with pytest.raises(DocumentOpenError):
            doc.save()

        # The failed promotion must leave the destination and its directory clean.
        assert clean_workspace.read_bytes() == original_bytes
        assert list(clean_workspace.parent.glob(f".{clean_workspace.name}.*")) == []
        doc.close()

    def test_replace_denial_stays_permission_error_on_posix(self, clean_workspace, monkeypatch):
        """On POSIX a denied rename never means "open document" — don't mislabel it.

        rename(2) over a file another process holds open always succeeds, so a denial
        means a sticky-bit directory, an immutable file, or an SELinux denial.
        """
        monkeypatch.setattr(pack.sys, "platform", "linux")

        def deny(*args, **kwargs):
            raise PermissionError("sticky-bit directory")

        monkeypatch.setattr(pack.os, "replace", deny)

        doc = Document.open(clean_workspace)
        with pytest.raises(PermissionError):
            doc.save()
        doc.close()

    def test_permission_error_before_replace_is_not_mislabeled(self, clean_workspace, monkeypatch):
        """A PermissionError from anywhere but the replace stays a PermissionError.

        Creating the temp file needs *directory* write permission, which the pre-atomic
        in-place write did not. Reporting that as "close the document in Word" would
        send the caller down a dead end — SKILL.md tells agents to stop and blame Word
        on DocumentOpenError.
        """

        def deny(*args, **kwargs):
            raise PermissionError("cannot create temp file in read-only directory")

        monkeypatch.setattr(pack.tempfile, "mkstemp", deny)

        doc = Document.open(clean_workspace)
        with pytest.raises(PermissionError):
            doc.save()
        doc.close()
