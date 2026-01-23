"""Tests for workspace management."""

import json

import pytest

from docx_editor.exceptions import (
    DocumentNotFoundError,
    InvalidDocumentError,
    WorkspaceSyncError,
)
from docx_editor.workspace import Workspace


class TestWorkspaceCreation:
    """Tests for workspace creation."""

    def test_create_workspace(self, clean_workspace):
        """Test creating a workspace for a document."""
        workspace = Workspace(clean_workspace)

        assert workspace.workspace_path.exists()
        assert workspace.word_path.exists()
        assert workspace.document_xml_path.exists()
        assert (workspace.workspace_path / "meta.json").exists()

        workspace.close()

    def test_workspace_meta_json(self, clean_workspace):
        """Test that meta.json contains correct fields."""
        workspace = Workspace(clean_workspace)

        meta_path = workspace.workspace_path / "meta.json"
        with open(meta_path) as f:
            meta = json.load(f)

        assert "source_path" in meta
        assert "source_mtime" in meta
        assert "source_size" in meta
        assert "created_at" in meta
        assert "author" in meta
        assert "rsid" in meta
        assert len(meta["rsid"]) == 8  # RSID is 8 hex chars

        workspace.close()

    def test_workspace_author_default(self, clean_workspace):
        """Test that author defaults to system user."""
        import getpass

        workspace = Workspace(clean_workspace)
        assert workspace.author == getpass.getuser()
        workspace.close()

    def test_workspace_author_custom(self, clean_workspace):
        """Test setting custom author."""
        workspace = Workspace(clean_workspace, author="Legal Team")
        assert workspace.author == "Legal Team"
        workspace.close()

    def test_document_not_found(self, temp_dir):
        """Test error when document doesn't exist."""
        with pytest.raises(DocumentNotFoundError):
            Workspace(temp_dir / "nonexistent.docx")

    def test_invalid_document_extension(self, temp_dir):
        """Test error when file is not .docx."""
        txt_file = temp_dir / "test.txt"
        txt_file.write_text("hello")

        with pytest.raises(InvalidDocumentError):
            Workspace(txt_file)


class TestWorkspacePersistence:
    """Tests for workspace loading and sync."""

    def test_workspace_exists_check(self, clean_workspace):
        """Test checking if workspace exists."""
        assert not Workspace.exists(clean_workspace)

        workspace = Workspace(clean_workspace)
        assert Workspace.exists(clean_workspace)

        workspace.close()
        assert not Workspace.exists(clean_workspace)

    def test_reopen_existing_workspace(self, clean_workspace):
        """Test reopening an existing workspace."""
        workspace1 = Workspace(clean_workspace)
        rsid1 = workspace1.rsid
        workspace1.close(cleanup=False)  # Don't delete

        # Reopen - should reuse existing workspace
        workspace2 = Workspace(clean_workspace)
        assert workspace2.rsid == rsid1

        workspace2.close()

    def test_workspace_sync_error_on_modified_source(self, clean_workspace):
        """Test error when source document is modified."""
        workspace = Workspace(clean_workspace)
        workspace.close(cleanup=False)

        # Modify the source document
        import time

        time.sleep(0.1)  # Ensure mtime changes
        clean_workspace.write_bytes(clean_workspace.read_bytes() + b"\x00")

        # Should raise sync error
        with pytest.raises(WorkspaceSyncError):
            Workspace(clean_workspace)

        # Cleanup
        Workspace.delete(clean_workspace)


class TestWorkspaceSaveClose:
    """Tests for saving and closing workspaces."""

    def test_save_to_original(self, clean_workspace):
        """Test saving back to original path."""
        workspace = Workspace(clean_workspace)

        saved_path = workspace.save()

        assert saved_path == clean_workspace
        assert clean_workspace.exists()
        # Size might change slightly due to XML processing
        assert clean_workspace.stat().st_size > 0

        workspace.close()

    def test_save_to_new_path(self, clean_workspace, temp_dir):
        """Test saving to a new path."""
        workspace = Workspace(clean_workspace)
        new_path = temp_dir / "output.docx"

        saved_path = workspace.save(new_path)

        assert saved_path == new_path
        assert new_path.exists()
        assert clean_workspace.exists()  # Original unchanged

        workspace.close()

    def test_close_with_cleanup(self, clean_workspace):
        """Test that close removes workspace."""
        workspace = Workspace(clean_workspace)
        workspace_path = workspace.workspace_path

        workspace.close(cleanup=True)

        assert not workspace_path.exists()

    def test_close_without_cleanup(self, clean_workspace):
        """Test that close can preserve workspace."""
        workspace = Workspace(clean_workspace)
        workspace_path = workspace.workspace_path

        workspace.close(cleanup=False)

        assert workspace_path.exists()

        # Manual cleanup
        Workspace.delete(clean_workspace)

    def test_delete_workspace(self, clean_workspace):
        """Test deleting workspace via class method."""
        workspace = Workspace(clean_workspace)
        workspace.close(cleanup=False)

        assert Workspace.exists(clean_workspace)
        result = Workspace.delete(clean_workspace)
        assert result is True
        assert not Workspace.exists(clean_workspace)

    def test_delete_nonexistent_workspace(self, clean_workspace):
        """Test deleting nonexistent workspace returns False."""
        result = Workspace.delete(clean_workspace)
        assert result is False
