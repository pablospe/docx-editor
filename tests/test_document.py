"""Tests for the main Document class."""

import json
import re
import zipfile

import pytest
from conftest import ENTITY_DTD_XML, NS, find_ref, replace_document_xml
from defusedxml.common import EntitiesForbidden

import docx_editor
from docx_editor import Document, SearchResult
from docx_editor.exceptions import (
    DocumentOpenError,
    DocxEditError,
    InvalidDocumentError,
    WorkspaceSyncError,
)
from docx_editor.workspace import Workspace


def _simulate_process_exit(doc):
    """Drop the document's advisory lock as if its process had died.

    The adoption regressions below simulate "RUN 1 exited without close()"
    inside a single test process. A real dead session leaves a lock naming a
    dead pid, which the next open silently reclaims (covered in
    test_workspace.py); in here the pid is this very process — alive — so the
    lock must be dropped explicitly or RUN 2 sees WorkspaceLockedError instead
    of what it would actually hit after a crash.
    """
    doc._workspace._release_lock()


class TestDocumentOpen:
    """Tests for opening documents."""

    def test_open_document(self, clean_workspace):
        """Test opening a document creates workspace."""
        doc = Document.open(clean_workspace)

        assert Workspace.exists(clean_workspace)
        assert doc.source_path == clean_workspace

        doc.close()

    def test_open_with_custom_author(self, clean_workspace):
        """Test opening with custom author."""
        doc = Document.open(clean_workspace, author="Custom Author")

        assert doc.author == "Custom Author"

        doc.close()

    def test_open_force_recreate(self, clean_workspace):
        """Test force recreating workspace."""
        # Create initial workspace
        doc1 = Document.open(clean_workspace)
        doc1.close(cleanup=False)

        # Force recreate should work
        doc2 = Document.open(clean_workspace, force_recreate=True)
        doc2.close()

    def test_open_with_workspace_dir(self, clean_workspace, tmp_path):
        """Test that workspace_dir is threaded through to the Workspace."""
        doc = Document.open(clean_workspace, workspace_dir=tmp_path)

        assert doc._workspace.workspace_path.parent == tmp_path
        assert doc._workspace.workspace_path.exists()
        assert Workspace.exists(clean_workspace, workspace_dir=tmp_path)

        doc.close()

    def test_open_force_recreate_with_workspace_dir(self, clean_workspace, tmp_path):
        """Test force_recreate deletes the workspace under the given workspace_dir."""
        doc1 = Document.open(clean_workspace, workspace_dir=tmp_path)
        doc1.close(cleanup=False)
        assert Workspace.exists(clean_workspace, workspace_dir=tmp_path)

        doc2 = Document.open(clean_workspace, force_recreate=True, workspace_dir=tmp_path)
        assert doc2._workspace.workspace_path.parent == tmp_path
        doc2.close()

    def test_open_rejects_external_entities(self, simple_docx, tmp_path):
        """A document declaring XML entities must be refused, never expanded (XXE)."""
        evil = tmp_path / "entities.docx"
        replace_document_xml(
            simple_docx,
            evil,
            '<?xml version="1.0"?>\n'
            '<!DOCTYPE w:document [<!ENTITY x SYSTEM "http://example.com/x">]>\n'
            f"<w:document {NS}><w:body><w:p><w:r><w:t>&x;</w:t></w:r></w:p></w:body></w:document>",
        )
        # The refusal now surfaces as the documented InvalidDocumentError;
        # the EntitiesForbidden cause proves the entity was never expanded.
        with pytest.raises(InvalidDocumentError) as excinfo:
            Document.open(evil, workspace_dir=tmp_path / "ws")
        assert isinstance(excinfo.value.__cause__, EntitiesForbidden)

    def test_document_open_invalid_xml_raises_invalid_document_error(self, simple_docx, temp_dir):
        """ISSUES.md #35: malformed XML in a part surfaces as InvalidDocumentError."""
        bad_docx = temp_dir / "bad.docx"
        replace_document_xml(simple_docx, bad_docx, ENTITY_DTD_XML)

        with pytest.raises(DocxEditError) as excinfo:
            Document.open(bad_docx)

        assert isinstance(excinfo.value, InvalidDocumentError)



class TestDocumentSave:
    """Tests for saving documents."""

    def test_save_to_original(self, clean_workspace, temp_dir):
        """Test saving back to original path."""
        doc = Document.open(clean_workspace)

        saved_path = doc.save()

        assert saved_path == clean_workspace
        assert clean_workspace.exists()

        doc.close()

    def test_save_to_new_path(self, clean_workspace, temp_dir):
        """Test saving to a new path."""
        doc = Document.open(clean_workspace)
        new_path = temp_dir / "output.docx"

        saved_path = doc.save(new_path)

        assert saved_path == new_path
        assert new_path.exists()
        assert clean_workspace.exists()  # Original unchanged

        doc.close()

    def test_save_does_not_include_meta_json(self, clean_workspace, temp_dir):
        """Regression for issue #8: meta.json must not be packed into output.

        Word flags any non-OOXML entry in the ZIP as "unreadable content" on open.
        Reproduces the issue end-to-end via Document.open + save.
        """
        doc = Document.open(clean_workspace, author="Test")
        output_path = temp_dir / "output.docx"
        doc.save(output_path)
        doc.close()

        with zipfile.ZipFile(output_path, "r") as zf:
            names = zf.namelist()
        assert "meta.json" not in names
        # Sanity: real OOXML parts are still present.
        assert "word/document.xml" in names
        assert "[Content_Types].xml" in names


class TestStaleWorkspaceAdoption:
    """Regression tests for issue #31: a workspace left behind by a session
    that saved elsewhere (or crashed mid-save) must not be silently adopted —
    the next save would write the previous session's edits into a source
    document the user never touched."""

    def test_save_elsewhere_without_close_refuses_reopen(self, clean_workspace, temp_dir):
        """The issue's two-run repro: RUN 1 edits and saves elsewhere, RUN 2
        opens the clean source and must get an error, not RUN 1's edits."""
        # RUN 1: edit, save to a different path, exit without close().
        doc = Document.open(clean_workspace, author="Run One")
        doc.replace("fox", "cat", paragraph=find_ref(doc, "fox"))
        doc.save(temp_dir / "reviewed.docx")
        workspace_path = doc.workspace_path
        _simulate_process_exit(doc)
        del doc  # no close()

        # RUN 2: the untouched source must refuse the diverged workspace.
        with pytest.raises(WorkspaceSyncError, match="unsaved changes") as excinfo:
            Document.open(clean_workspace, author="Run Two")
        assert str(workspace_path) in str(excinfo.value)

        # The documented recovery opens clean: no leaked edits, no revisions.
        doc2 = Document.open(clean_workspace, author="Run Two", force_recreate=True)
        assert doc2.list_revisions() == []
        assert "cat" not in doc2.get_visible_text()
        doc2.close()

    def test_save_in_place_without_close_reopens_clean(self, clean_workspace):
        """An in-place save leaves source and workspace in agreement, so a
        later open (even without close()) adopts the workspace as before."""
        doc = Document.open(clean_workspace, author="Run One")
        doc.replace("fox", "cat", paragraph=find_ref(doc, "fox"))
        doc.save()
        _simulate_process_exit(doc)
        del doc  # no close()

        doc2 = Document.open(clean_workspace, author="Run Two")
        # replace() records a paired deletion + insertion.
        assert len(doc2.list_revisions()) == 2
        doc2.close()

    def test_crash_with_no_edits_still_adopts(self, clean_workspace):
        """Opening writes tracking infrastructure (people.xml, rsid) into the
        workspace; that alone must not count as unsaved changes."""
        doc = Document.open(clean_workspace, author="Run One")
        _simulate_process_exit(doc)
        del doc  # no edits, no save, no close()

        doc2 = Document.open(clean_workspace, author="Run Two")
        assert doc2.list_revisions() == []
        doc2.close()

    def test_failed_save_still_flags_workspace(self, clean_workspace):
        """Write-ahead ordering: Document.save() flushes edits into the
        workspace before packing, so a save that fails after the flush must
        already have persisted the dirty flag."""
        doc = Document.open(clean_workspace, author="Run One")
        doc.replace("fox", "cat", paragraph=find_ref(doc, "fox"))

        # Make the pack step refuse: a ~$ owner stub marks the source as open
        # in Word. The editors have already flushed by the time it is checked.
        owner_stub = clean_workspace.parent / f"~${clean_workspace.name}"
        owner_stub.write_text("stub")
        try:
            with pytest.raises(DocumentOpenError):
                doc.save()
        finally:
            owner_stub.unlink()
        _simulate_process_exit(doc)
        del doc  # no close()

        with pytest.raises(WorkspaceSyncError, match="unsaved changes"):
            Document.open(clean_workspace, author="Run Two")

    def test_relative_path_save_to_source_refreshes_mtime(self, clean_workspace, monkeypatch):
        """Regression for issue #31's adjacent bug (fixed in PR #33): saving to
        the source via a relative path must refresh the recorded mtime so the
        workspace is not reported stale on the next open."""
        doc = Document.open(clean_workspace, author="Run One")
        doc.replace("fox", "cat", paragraph=find_ref(doc, "fox"))
        monkeypatch.chdir(clean_workspace.parent)
        doc.save(clean_workspace.name)

        with open(doc.workspace_path / "meta.json") as f:
            meta = json.load(f)
        assert meta["source_mtime"] == clean_workspace.stat().st_mtime
        assert meta["dirty"] is False
        _simulate_process_exit(doc)
        del doc  # no close()

        doc2 = Document.open(clean_workspace, author="Run Two")
        # replace() records a paired deletion + insertion.
        assert len(doc2.list_revisions()) == 2
        doc2.close()


class TestDocumentClose:
    """Tests for closing documents."""

    def test_close_cleans_workspace(self, clean_workspace):
        """Test that close removes workspace."""
        doc = Document.open(clean_workspace)
        doc.close()

        assert not Workspace.exists(clean_workspace)

    def test_close_preserves_workspace(self, clean_workspace):
        """Test that close can preserve workspace."""
        doc = Document.open(clean_workspace)
        doc.close(cleanup=False)

        assert Workspace.exists(clean_workspace)

        # Manual cleanup
        Workspace.delete(clean_workspace)

    def test_operations_after_close_raise_error(self, clean_workspace):
        """Test that operations after close raise error."""
        doc = Document.open(clean_workspace)
        doc.close()

        with pytest.raises(ValueError, match="closed"):
            doc.list_revisions()


class TestDocumentContextManager:
    """Tests for using Document as context manager."""

    def test_context_manager_normal(self, clean_workspace):
        """Test using document as context manager."""
        with Document.open(clean_workspace):
            assert Workspace.exists(clean_workspace)

        # Workspace should be cleaned up
        assert not Workspace.exists(clean_workspace)

    def test_context_manager_exception(self, clean_workspace):
        """Test context manager preserves workspace on exception."""
        try:
            with Document.open(clean_workspace):
                raise RuntimeError("Test error")
        except RuntimeError:
            pass

        # Workspace should be preserved on error (cleanup=False when exc_type is not None)
        assert Workspace.exists(clean_workspace)

        # Manual cleanup
        Workspace.delete(clean_workspace)


class TestDocumentRoundTrip:
    """Tests for round-trip editing."""

    def test_edit_save_reopen(self, clean_workspace, temp_dir):
        """Test editing, saving, and reopening a document."""
        # First edit
        doc1 = Document.open(clean_workspace)
        try:
            doc1.add_comment("fox", "Test comment")
        except Exception:
            pytest.skip("Could not add comment")
        doc1.save()
        doc1.close()

        # Reopen and verify
        doc2 = Document.open(clean_workspace, force_recreate=True)
        comments = doc2.list_comments()
        assert len(comments) >= 1
        doc2.close()


class TestDocumentEdgeCases:
    """Tests for edge cases and error handling."""

    def test_close_already_closed(self, clean_workspace):
        """Test that closing an already closed document does nothing."""
        doc = Document.open(clean_workspace)
        doc.close()

        # Second close should not raise
        doc.close()

    def test_open_raises_on_sync_mismatch(self, clean_workspace):
        """Document.open must raise WorkspaceSyncError instead of silently
        deleting a workspace whose source .docx was modified out-of-band.

        Verifies both that the exception propagates AND that the workspace
        survives on disk (a buggy implementation that deleted and then
        re-raised would otherwise pass). Recovery is via the explicit
        force_recreate=True opt-in.
        """
        import os

        from docx_editor.exceptions import WorkspaceSyncError

        # Create workspace, keep it on disk so the next open hits the sync check.
        doc1 = Document.open(clean_workspace)
        doc1.close(cleanup=False)
        assert Workspace.exists(clean_workspace)

        # Mutate source bytes and bump mtime explicitly (sleep-based mtime
        # changes are unreliable on coarse-resolution filesystems).
        original_mtime = clean_workspace.stat().st_mtime
        clean_workspace.write_bytes(clean_workspace.read_bytes() + b"\x00")
        os.utime(clean_workspace, (original_mtime + 5, original_mtime + 5))

        with pytest.raises(WorkspaceSyncError):
            Document.open(clean_workspace)

        # The whole point of this PR: workspace must survive the failed open.
        assert Workspace.exists(clean_workspace), (
            "workspace must remain on disk when Document.open raises WorkspaceSyncError"
        )

        # The documented escape hatch still works: discard and re-unpack.
        doc2 = Document.open(clean_workspace, force_recreate=True)
        assert doc2.source_path == clean_workspace
        doc2.close()

    def test_operations_after_close(self, clean_workspace):
        """Test that various operations raise after close."""
        doc = Document.open(clean_workspace)
        doc.close()

        with pytest.raises(ValueError, match="closed"):
            doc.count_matches("test")

        with pytest.raises(ValueError, match="closed"):
            doc.replace("old", "new", paragraph="P1#0000")

        with pytest.raises(ValueError, match="closed"):
            doc.delete("text", paragraph="P1#0000")

        with pytest.raises(ValueError, match="closed"):
            doc.insert_after("anchor", "text", paragraph="P1#0000")

        with pytest.raises(ValueError, match="closed"):
            doc.insert_before("anchor", "text", paragraph="P1#0000")

        with pytest.raises(ValueError, match="closed"):
            doc.add_comment("anchor", "comment")

        with pytest.raises(ValueError, match="closed"):
            doc.reply_to_comment(0, "reply")

        with pytest.raises(ValueError, match="closed"):
            doc.list_comments()

        with pytest.raises(ValueError, match="closed"):
            doc.resolve_comment(0)

        with pytest.raises(ValueError, match="closed"):
            doc.delete_comment(0)

        with pytest.raises(ValueError, match="closed"):
            doc.accept_revision(0)

        with pytest.raises(ValueError, match="closed"):
            doc.reject_revision(0)

        with pytest.raises(ValueError, match="closed"):
            doc.accept_all()

        with pytest.raises(ValueError, match="closed"):
            doc.reject_all()

        with pytest.raises(ValueError, match="closed"):
            doc.save()


class TestDocumentInternalMethods:
    """Tests for internal Document methods and edge cases."""

    def test_force_recreate_with_persistent_sync_error(self, clean_workspace):
        """Test that WorkspaceSyncError propagates from Workspace.__init__ even
        when force_recreate=True (i.e. when the workspace constructor raises
        sync errors for some reason other than a pre-existing stale workspace).
        """
        from unittest.mock import patch

        from docx_editor.exceptions import WorkspaceSyncError

        # Make Workspace always raise WorkspaceSyncError
        with patch("docx_editor.document.Workspace") as mock_workspace_cls:
            mock_workspace_cls.side_effect = WorkspaceSyncError("Persistent error")
            mock_workspace_cls.delete = lambda p, workspace_dir=None: True

            with pytest.raises(WorkspaceSyncError):
                Document.open(clean_workspace, force_recreate=True)

    def test_add_relationship_for_people_missing_rels_path(self, clean_workspace):
        """Test _add_relationship_for_people returns early when rels file missing.

        This tests line 485.
        """
        doc = Document.open(clean_workspace)

        # Remove the rels file
        rels_path = doc._workspace.word_path / "_rels" / "document.xml.rels"
        if rels_path.exists():
            rels_path.unlink()

        # This should not raise, just return early
        doc._add_relationship_for_people()

        doc.close()

    def test_update_settings_missing_settings_xml(self, clean_workspace):
        """Test _update_settings returns early when settings.xml is missing.

        This tests line 512.
        """
        doc = Document.open(clean_workspace)

        # Remove settings.xml
        settings_path = doc._workspace.word_path / "settings.xml"
        if settings_path.exists():
            settings_path.unlink()

        # This should not raise, just return early
        doc._update_settings()

        doc.close()

    def test_update_settings_no_rsids_section(self, clean_workspace, temp_dir):
        """Test _update_settings creates new rsids section when none exists.

        This tests lines 528-538.
        """

        # Create a minimal settings.xml without rsids section
        minimal_settings = """<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:zoom w:percent="100"/>
</w:settings>"""

        doc = Document.open(clean_workspace)
        settings_path = doc._workspace.word_path / "settings.xml"
        settings_path.write_text(minimal_settings)

        # Call _update_settings - should create rsids section
        doc._update_settings()

        # Verify rsids section was created
        content = settings_path.read_text()
        assert "rsids" in content
        assert doc._workspace.rsid in content

        doc.close()

    def test_update_settings_no_rsids_no_compat(self, clean_workspace):
        """Test _update_settings appends rsids to root when no compat element.

        This tests lines 537-538 (the else branch).
        """
        # Create settings.xml without rsids or compat section
        minimal_settings = """<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:zoom w:percent="100"/>
</w:settings>"""

        doc = Document.open(clean_workspace)
        settings_path = doc._workspace.word_path / "settings.xml"
        settings_path.write_text(minimal_settings)

        doc._update_settings()

        content = settings_path.read_text()
        assert "rsids" in content

        doc.close()

    def test_update_settings_no_rsids_but_has_compat(self, clean_workspace):
        """Test _update_settings inserts rsids after compat when compat exists.

        This tests line 536.
        """
        # Create settings.xml with compat but without rsids section
        settings_with_compat = """<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:zoom w:percent="100"/>
    <w:compat>
        <w:compatSetting w:name="test" w:val="1"/>
    </w:compat>
</w:settings>"""

        doc = Document.open(clean_workspace)
        settings_path = doc._workspace.word_path / "settings.xml"
        settings_path.write_text(settings_with_compat)

        doc._update_settings()

        content = settings_path.read_text()
        assert "rsids" in content
        assert doc._workspace.rsid in content
        # rsids should appear after compat in the file
        compat_pos = content.find("compat")
        rsids_pos = content.find("rsids")
        assert rsids_pos > compat_pos

        doc.close()

    def test_add_author_to_people_missing_people_xml(self, clean_workspace):
        """Test _add_author_to_people returns early when people.xml missing.

        This tests line 557.
        """
        doc = Document.open(clean_workspace)

        # Remove people.xml
        people_path = doc._workspace.word_path / "people.xml"
        if people_path.exists():
            people_path.unlink()

        # Should not raise, just return early
        doc._add_author_to_people()

        doc.close()

    def test_ensure_comment_relationships_already_exists(self, clean_workspace):
        """Test _ensure_comment_relationships returns early when relationship exists.

        This tests line 596.
        """
        doc = Document.open(clean_workspace)

        # First add a comment to create comments.xml
        try:
            doc.add_comment("fox", "Test comment")
        except Exception:
            doc.close()
            pytest.skip("Could not add comment")

        # Ensure relationships are set up
        doc._ensure_comment_relationships()

        # Call again - should return early because relationship already exists
        doc._ensure_comment_relationships()

        doc.close()

    def test_ensure_comment_content_types_already_exists(self, clean_workspace):
        """Test _ensure_comment_content_types returns early when content type exists.

        This tests line 650.
        """
        doc = Document.open(clean_workspace)

        # First add a comment to create comments.xml
        try:
            doc.add_comment("fox", "Test comment")
        except Exception:
            doc.close()
            pytest.skip("Could not add comment")

        # Ensure content types are set up
        doc._ensure_comment_content_types()

        # Call again - should return early because content type already exists
        doc._ensure_comment_content_types()

        doc.close()


class TestDocumentGetVisibleText:
    """Tests for get_visible_text()."""

    def test_get_visible_text_basic(self, clean_workspace):
        """Test getting visible text from a simple document."""
        doc = Document.open(clean_workspace)
        text = doc.get_visible_text()
        # The simple.docx test fixture has some text content
        assert isinstance(text, str)
        assert len(text) > 0
        doc.close()

    def test_get_visible_text_after_insertion(self, clean_workspace):
        """Inserted text should appear in visible text."""
        doc = Document.open(clean_workspace)
        original = doc.get_visible_text()
        ref = find_ref(doc, "fox")
        doc.insert_after("fox", " INSERTED", paragraph=ref)
        updated = doc.get_visible_text()
        assert "INSERTED" in updated
        assert len(updated) > len(original)
        doc.close()

    def test_get_visible_text_after_deletion(self, clean_workspace):
        """Deleted text should NOT appear in visible text."""
        doc = Document.open(clean_workspace)
        original = doc.get_visible_text()
        assert "fox" in original
        ref = find_ref(doc, "fox")
        doc.delete("fox", paragraph=ref)
        updated = doc.get_visible_text()
        assert "fox" not in updated
        doc.close()

    def test_get_visible_text_after_replace(self, clean_workspace):
        """Replaced text should show new text, not old."""
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "fox")
        doc.replace("fox", "cat", paragraph=ref)
        text = doc.get_visible_text()
        assert "cat" in text
        assert "fox" not in text
        doc.close()

    def test_get_visible_text_after_close_raises(self, clean_workspace):
        """Should raise ValueError after document is closed."""
        doc = Document.open(clean_workspace)
        doc.close()
        with pytest.raises(ValueError, match="closed"):
            doc.get_visible_text()


class TestDocumentGetOriginalText:
    """Tests for get_original_text()."""

    def test_recovers_pre_revision_text(self, clean_workspace):
        """Original text equals the pre-edit baseline; visible is the inverse."""
        doc = Document.open(clean_workspace)
        baseline = doc.get_visible_text()
        doc.insert_after("fox", " INSERTED", paragraph=find_ref(doc, "fox"))
        doc.delete("dog", paragraph=find_ref(doc, "dog"))
        assert doc.get_original_text() == baseline
        visible = doc.get_visible_text()
        assert "INSERTED" in visible
        assert "dog" not in visible
        doc.close()

    def test_matches_visible_text_after_reject_all(self, clean_workspace):
        """Acceptance criterion: original view == post-reject_all visible text."""
        doc = Document.open(clean_workspace)
        doc.replace("fox", "cat", paragraph=find_ref(doc, "fox"))
        original = doc.get_original_text()
        doc.reject_all()
        assert doc.get_visible_text() == original
        doc.close()

    def test_no_revisions_equals_visible(self, clean_workspace):
        """Without tracked changes both views are identical."""
        doc = Document.open(clean_workspace)
        assert doc.get_original_text() == doc.get_visible_text()
        doc.close()

    def test_read_only_refs_stay_valid(self, clean_workspace):
        """Reading the original view does not invalidate paragraph refs."""
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "fox")
        doc.get_original_text()
        new_ref = doc.replace("fox", "cat", paragraph=ref)
        assert new_ref != ref
        assert "cat" in doc.get_visible_text()
        doc.close()

    def test_get_original_text_after_close_raises(self, clean_workspace):
        """Should raise ValueError after document is closed."""
        doc = Document.open(clean_workspace)
        doc.close()
        with pytest.raises(ValueError, match="closed"):
            doc.get_original_text()


class TestDocumentFindText:
    """Tests for find_text()."""

    def test_find_text_simple(self, clean_workspace):
        """Find text in a simple document."""
        doc = Document.open(clean_workspace)
        match = doc.find_text("fox")
        assert match is not None
        assert isinstance(match, SearchResult)
        assert match.text == "fox"
        assert not match.spans_revision
        assert re.fullmatch(r"P\d+#[0-9a-f]{4}", match.paragraph_ref)
        assert 0 <= match.start < match.end
        assert match.paragraph_occurrence == 0
        doc.close()

    def test_find_text_not_found(self, clean_workspace):
        """Return None when text not found."""
        doc = Document.open(clean_workspace)
        match = doc.find_text("nonexistent")
        assert match is None
        doc.close()

    def test_find_text_after_insertion(self, clean_workspace):
        """Find text that spans an insertion boundary."""
        doc = Document.open(clean_workspace)
        # insert_after with paragraph= inserts after the run containing anchor
        ref = find_ref(doc, "dog.")
        doc.insert_after("dog.", " INSERTED", paragraph=ref)
        # "dog. INSERTED" spans original run + insertion
        match = doc.find_text("dog. INSERTED")
        assert match is not None
        assert match.spans_revision
        doc.close()

    def test_find_text_ref_chains_into_edit(self, clean_workspace):
        """paragraph_ref from find_text is directly usable for a follow-up edit."""
        doc = Document.open(clean_workspace)
        match = doc.find_text("fox")
        assert match is not None
        doc.replace("fox", "cat", paragraph=match.paragraph_ref)
        assert "cat" in doc.get_visible_text()
        doc.close()

    def test_find_text_paragraph_occurrence_chains_into_edit(self, clean_workspace):
        """paragraph_occurrence pins which in-paragraph match a follow-up edit targets.

        find_text's occurrence counts document-wide; edit methods count within
        the paragraph. Without passing paragraph_occurrence through, the edit
        would silently target the paragraph's first match instead of the one
        find_text located.
        """
        doc = Document.open(clean_workspace)
        # "he" appears twice in the fox paragraph: in "The" and in "the lazy"
        match = doc.find_text("he", occurrence=1)
        assert match is not None
        assert match.paragraph_occurrence == 1
        doc.replace("he", "XX", paragraph=match.paragraph_ref, occurrence=match.paragraph_occurrence)
        text = doc.get_visible_text()
        assert "tXX lazy" in text
        assert "The quick" in text  # the paragraph's first match is untouched
        doc.close()

    def test_find_text_after_close_raises(self, clean_workspace):
        doc = Document.open(clean_workspace)
        doc.close()
        with pytest.raises(ValueError, match="closed"):
            doc.find_text("test")


class TestPublicApiSurface:
    """The text-map internals are deprecated at the top level; SearchResult is public."""

    DEPRECATED = ["TextMap", "TextMapMatch", "TextPosition", "build_text_map", "find_in_text_map"]

    @pytest.mark.parametrize("name", DEPRECATED)
    def test_deprecated_internal_warns_and_forwards(self, name):
        """Accessing an internal via the top-level package warns and returns the real object."""
        from docx_editor import xml_editor

        with pytest.warns(DeprecationWarning, match=f"docx_editor.{name} is internal"):
            obj = getattr(docx_editor, name)
        assert obj is getattr(xml_editor, name)

    def test_deprecated_internals_not_in_all(self):
        for name in self.DEPRECATED:
            assert name not in docx_editor.__all__

    def test_search_result_is_public(self):
        assert "SearchResult" in docx_editor.__all__
        assert docx_editor.SearchResult is SearchResult

    def test_unknown_attribute_raises(self):
        name = "DoesNotExist"
        with pytest.raises(AttributeError, match="no attribute 'DoesNotExist'"):
            getattr(docx_editor, name)


class TestDocumentSaveStaleness:
    def test_document_save_raises_on_external_change(self, temp_docx):
        from docx_editor.exceptions import WorkspaceSyncError

        doc = Document.open(temp_docx, author="Test")
        # Simulate an external edit: change the source file's content.
        temp_docx.write_bytes(temp_docx.read_bytes() + b"\x00")
        try:
            with pytest.raises(WorkspaceSyncError):
                doc.save()
            doc.save(force=True)  # explicit override still works
        finally:
            doc.close()
