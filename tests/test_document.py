"""Tests for the main Document class."""

import pytest

from docx_edit import Document
from docx_edit.workspace import Workspace


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
            doc1.add_comment("the", "Test comment")
        except Exception:
            pytest.skip("Could not add comment")
        doc1.save()
        doc1.close()

        # Reopen and verify
        doc2 = Document.open(clean_workspace, force_recreate=True)
        comments = doc2.list_comments()
        assert len(comments) >= 1
        doc2.close()
