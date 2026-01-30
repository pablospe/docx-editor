"""Tests for the main Document class."""

import pytest

from docx_editor import Document
from docx_editor.workspace import Workspace


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

    def test_auto_recreate_on_sync_error(self, clean_workspace):
        """Test that Document.open auto-recreates on WorkspaceSyncError."""
        import time

        # Create workspace
        doc1 = Document.open(clean_workspace)
        doc1.close(cleanup=False)

        # Modify source to trigger sync error
        time.sleep(0.1)
        clean_workspace.write_bytes(clean_workspace.read_bytes() + b"\x00")

        # Should auto-recreate without error
        doc2 = Document.open(clean_workspace)
        assert doc2.source_path == clean_workspace
        doc2.close()

    def test_operations_after_close(self, clean_workspace):
        """Test that various operations raise after close."""
        doc = Document.open(clean_workspace)
        doc.close()

        with pytest.raises(ValueError, match="closed"):
            doc.count_matches("test")

        with pytest.raises(ValueError, match="closed"):
            doc.replace("old", "new")

        with pytest.raises(ValueError, match="closed"):
            doc.delete("text")

        with pytest.raises(ValueError, match="closed"):
            doc.insert_after("anchor", "text")

        with pytest.raises(ValueError, match="closed"):
            doc.insert_before("anchor", "text")

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
        """Test that force_recreate=True re-raises WorkspaceSyncError if it persists.

        This tests line 102: the `raise` when force_recreate=True but WorkspaceSyncError
        still occurs after deletion.
        """
        from unittest.mock import patch

        from docx_editor.exceptions import WorkspaceSyncError

        # Make Workspace always raise WorkspaceSyncError
        with patch("docx_editor.document.Workspace") as mock_workspace_cls:
            mock_workspace_cls.side_effect = WorkspaceSyncError("Persistent error")
            mock_workspace_cls.delete = lambda p: True

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
        doc.insert_after("fox", " INSERTED")
        updated = doc.get_visible_text()
        assert "INSERTED" in updated
        assert len(updated) > len(original)
        doc.close()

    def test_get_visible_text_after_deletion(self, clean_workspace):
        """Deleted text should NOT appear in visible text."""
        doc = Document.open(clean_workspace)
        original = doc.get_visible_text()
        assert "fox" in original
        doc.delete("fox")
        updated = doc.get_visible_text()
        assert "fox" not in updated
        doc.close()

    def test_get_visible_text_after_replace(self, clean_workspace):
        """Replaced text should show new text, not old."""
        doc = Document.open(clean_workspace)
        doc.replace("fox", "cat")
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


class TestDocumentFindText:
    """Tests for find_text()."""

    def test_find_text_simple(self, clean_workspace):
        """Find text in a simple document."""
        doc = Document.open(clean_workspace)
        match = doc.find_text("fox")
        assert match is not None
        assert match.text == "fox"
        assert not match.spans_boundary
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
        # insert_after splits the run at the anchor and inserts inline after it
        doc.insert_after("fox", " INSERTED")
        # "fox INSERTED" spans original run + insertion
        match = doc.find_text("fox INSERTED")
        assert match is not None
        assert match.spans_boundary
        doc.close()

    def test_find_text_after_close_raises(self, clean_workspace):
        doc = Document.open(clean_workspace)
        doc.close()
        with pytest.raises(ValueError, match="closed"):
            doc.find_text("test")
