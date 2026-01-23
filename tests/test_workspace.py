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


class TestWorkspaceEdgeCases:
    """Tests for edge cases and error handling."""

    def test_workspace_exists_no_meta_json(self, clean_workspace, temp_dir):
        """Test error when workspace exists but has no meta.json."""
        from docx_editor.exceptions import WorkspaceExistsError

        # Create workspace directory without meta.json
        workspace_dir = clean_workspace.parent / ".docx" / clean_workspace.stem
        workspace_dir.mkdir(parents=True)

        with pytest.raises(WorkspaceExistsError):
            Workspace(clean_workspace)

        # Cleanup
        import shutil

        shutil.rmtree(clean_workspace.parent / ".docx")

    def test_workspace_create_false_not_found(self, clean_workspace):
        """Test error when workspace not found with create=False."""
        from docx_editor.exceptions import WorkspaceError

        with pytest.raises(WorkspaceError, match="Workspace not found"):
            Workspace(clean_workspace, create=False)

    def test_workspace_create_false_no_meta(self, clean_workspace):
        """Test error when workspace exists but no meta.json with create=False."""
        from docx_editor.exceptions import WorkspaceError

        # Create workspace directory without meta.json
        workspace_dir = clean_workspace.parent / ".docx" / clean_workspace.stem
        workspace_dir.mkdir(parents=True)

        with pytest.raises(WorkspaceError, match="Invalid workspace"):
            Workspace(clean_workspace, create=False)

        # Cleanup
        import shutil

        shutil.rmtree(clean_workspace.parent / ".docx")

    def test_load_meta_corrupt_json(self, clean_workspace):
        """Test that corrupt meta.json returns None in _load_meta."""
        workspace = Workspace(clean_workspace)
        workspace.close(cleanup=False)

        # Corrupt the meta.json
        meta_path = workspace.workspace_path / "meta.json"
        meta_path.write_text("not valid json {{{")

        # Try to load - should raise WorkspaceExistsError because meta is None
        from docx_editor.exceptions import WorkspaceExistsError

        with pytest.raises(WorkspaceExistsError):
            Workspace(clean_workspace)

        # Cleanup
        Workspace.delete(clean_workspace)

    def test_get_xml_path(self, clean_workspace):
        """Test get_xml_path returns correct path."""
        workspace = Workspace(clean_workspace)

        xml_path = workspace.get_xml_path("word/document.xml")
        assert xml_path == workspace.workspace_path / "word/document.xml"

        workspace.close()

    def test_sync_check_in_sync(self, clean_workspace):
        """Test sync_check returns True when document is in sync."""
        workspace = Workspace(clean_workspace)

        assert workspace.sync_check() is True

        workspace.close()

    def test_sync_check_source_deleted(self, clean_workspace, temp_dir):
        """Test sync_check returns False when source is deleted."""
        workspace = Workspace(clean_workspace)

        # Delete the source file
        clean_workspace.unlink()

        assert workspace.sync_check() is False

        workspace.close(cleanup=True)

    def test_sync_check_source_modified(self, clean_workspace):
        """Test sync_check returns False when source is modified."""
        import time

        workspace = Workspace(clean_workspace)
        workspace.close(cleanup=False)

        # Modify the source
        time.sleep(0.1)
        clean_workspace.write_bytes(clean_workspace.read_bytes() + b"\x00")

        # Reopen without creating new workspace
        workspace2 = Workspace.__new__(Workspace)
        workspace2.source_path = clean_workspace.resolve()
        workspace2._author = "test"
        workspace_dir = clean_workspace.parent / ".docx"
        workspace2.workspace_path = workspace_dir / clean_workspace.stem
        workspace2.meta = workspace2._load_meta()

        assert workspace2.sync_check() is False

        # Cleanup
        Workspace.delete(clean_workspace)

    def test_close_removes_empty_parent_dir(self, clean_workspace):
        """Test that close removes empty .docx parent directory."""
        workspace = Workspace(clean_workspace)
        parent_dir = workspace.workspace_path.parent

        workspace.close(cleanup=True)

        # Parent .docx dir should be removed if empty
        assert not parent_dir.exists()

    def test_delete_removes_empty_parent_dir(self, clean_workspace):
        """Test that delete removes empty .docx parent directory."""
        workspace = Workspace(clean_workspace)
        parent_dir = workspace.workspace_path.parent
        workspace.close(cleanup=False)

        Workspace.delete(clean_workspace)

        # Parent .docx dir should be removed if empty
        assert not parent_dir.exists()

    def test_load_meta_returns_existing(self, clean_workspace):
        """Test that existing valid meta.json is returned by _load_meta.

        This tests line 97 implicitly - the early return when workspace is valid.
        """
        # Create workspace
        workspace1 = Workspace(clean_workspace)
        rsid1 = workspace1.rsid
        workspace1.close(cleanup=False)

        # Reopen - meta should be loaded and workspace reused
        workspace2 = Workspace(clean_workspace)
        assert workspace2.rsid == rsid1
        assert workspace2.meta is not None
        assert workspace2.meta.get("rsid") == rsid1

        workspace2.close()

    def test_save_fails_pack_document(self, clean_workspace):
        """Test that save raises WorkspaceError when pack_document fails.

        This tests line 191.
        """
        from unittest.mock import patch

        from docx_editor.exceptions import WorkspaceError

        workspace = Workspace(clean_workspace)

        # Mock pack_document to return False (failure)
        with patch("docx_editor.workspace.pack_document", return_value=False):
            with pytest.raises(WorkspaceError, match="Failed to pack document"):
                workspace.save()

        workspace.close()

    def test_workspace_create_false_with_valid_meta(self, clean_workspace):
        """Test loading existing workspace with create=False.

        This tests line 97.
        """
        # First create a workspace
        workspace1 = Workspace(clean_workspace)
        rsid1 = workspace1.rsid
        workspace1.close(cleanup=False)

        # Now load it with create=False
        workspace2 = Workspace(clean_workspace, create=False)
        assert workspace2.rsid == rsid1
        assert workspace2.meta is not None

        workspace2.close()
