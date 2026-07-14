"""Tests for workspace management."""

import hashlib
import json
import os
import re
import shutil
import stat
import subprocess
import sys
from pathlib import Path

import pytest
from conftest import ENTITY_DTD_XML, find_ref, replace_document_xml

from docx_editor.document import Document
from docx_editor.exceptions import (
    DocumentNotFoundError,
    InvalidDocumentError,
    WorkspaceError,
    WorkspaceLockedError,
    WorkspaceSyncError,
)
from docx_editor.workspace import Workspace, _default_cache_dir, _pid_alive


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
        assert "source_sha256" in meta
        assert "created_at" in meta
        assert "author" in meta
        assert "rsid" in meta
        assert len(meta["rsid"]) == 8  # RSID is 8 hex chars
        assert meta["dirty"] is False  # fresh workspace holds nothing unsaved

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


class TestWorkspaceParseFailureCleanup:
    """A parse failure during creation must not leave a partial workspace (ISSUES.md #35)."""

    def test_create_workspace_parse_failure_leaves_no_workspace(self, simple_docx, temp_dir, isolated_workspace_base):
        """A failed unpack leaves neither a workspace dir nor a .lock sibling."""
        bad_docx = temp_dir / "bad.docx"
        replace_document_xml(simple_docx, bad_docx, ENTITY_DTD_XML)

        with pytest.raises(InvalidDocumentError):
            Workspace(bad_docx)

        assert list(isolated_workspace_base.iterdir()) == []

    def test_open_retry_after_parse_failure_not_wedged(self, simple_docx, temp_dir):
        """After a failed open, fixing the file and reopening must succeed.

        Without cleanup the partial dir (no meta.json) makes the retry raise a
        misleading WorkspaceExistsError — the user-facing regression this guards.
        """
        doc_path = temp_dir / "retry.docx"
        replace_document_xml(simple_docx, doc_path, ENTITY_DTD_XML)

        with pytest.raises(InvalidDocumentError):
            Workspace(doc_path)

        shutil.copy(simple_docx, doc_path)
        workspace = Workspace(doc_path)
        assert workspace.document_xml_path.exists()
        workspace.close()


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
        workspace_path = Workspace._resolve_workspace_path(clean_workspace.resolve(), None)
        workspace_path.mkdir(parents=True)

        with pytest.raises(WorkspaceExistsError):
            Workspace(clean_workspace)

        # Cleanup
        import shutil

        shutil.rmtree(workspace_path)

    def test_workspace_create_false_not_found(self, clean_workspace):
        """Test error when workspace not found with create=False."""
        from docx_editor.exceptions import WorkspaceError

        with pytest.raises(WorkspaceError, match="Workspace not found"):
            Workspace(clean_workspace, create=False)

    def test_workspace_create_false_no_meta(self, clean_workspace):
        """Test error when workspace exists but no meta.json with create=False."""
        from docx_editor.exceptions import WorkspaceError

        # Create workspace directory without meta.json
        workspace_path = Workspace._resolve_workspace_path(clean_workspace.resolve(), None)
        workspace_path.mkdir(parents=True)

        with pytest.raises(WorkspaceError, match="Invalid workspace"):
            Workspace(clean_workspace, create=False)

        # Cleanup
        import shutil

        shutil.rmtree(workspace_path)

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
        workspace2.workspace_path = Workspace._resolve_workspace_path(clean_workspace.resolve(), None)
        workspace2.meta = workspace2._load_meta()

        assert workspace2.sync_check() is False

        # Cleanup
        Workspace.delete(clean_workspace)

    def test_close_keeps_shared_base_dir(self, clean_workspace):
        """close() removes only this workspace, never the shared base directory.

        The base may be a caller-supplied directory (workspace_dir= /
        DOCX_EDITOR_WORKSPACE_DIR), so the library must not delete it.
        """
        workspace = Workspace(clean_workspace)
        base_dir = workspace.workspace_path.parent

        workspace.close(cleanup=True)

        assert not workspace.workspace_path.exists()
        assert base_dir.exists()

    def test_delete_keeps_shared_base_dir(self, clean_workspace):
        """delete() removes only this workspace, never the shared base directory."""
        workspace = Workspace(clean_workspace)
        base_dir = workspace.workspace_path.parent
        workspace.close(cleanup=False)

        assert Workspace.delete(clean_workspace) is True

        assert not workspace.workspace_path.exists()
        assert base_dir.exists()

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


class TestDefaultCacheDir:
    """Tests for the _default_cache_dir platform helper."""

    def test_linux_with_xdg_cache_home(self, monkeypatch):
        monkeypatch.setattr(os, "name", "posix")
        monkeypatch.setattr(sys, "platform", "linux")
        monkeypatch.setenv("XDG_CACHE_HOME", "/custom/cache")
        assert _default_cache_dir() == Path("/custom/cache") / "docx-editor"

    def test_linux_without_xdg_cache_home(self, monkeypatch):
        monkeypatch.setattr(os, "name", "posix")
        monkeypatch.setattr(sys, "platform", "linux")
        monkeypatch.delenv("XDG_CACHE_HOME", raising=False)
        assert _default_cache_dir() == Path.home() / ".cache" / "docx-editor"

    def test_macos(self, monkeypatch):
        monkeypatch.setattr(os, "name", "posix")
        monkeypatch.setattr(sys, "platform", "darwin")
        assert _default_cache_dir() == Path.home() / "Library" / "Caches" / "docx-editor"

    # The Windows branch constructs pathlib.Path, which resolves to WindowsPath
    # only when os.name == "nt". Patching os.name on a POSIX host makes Path
    # raise UnsupportedOperation, so these run for real only on Windows.
    @pytest.mark.skipif(os.name != "nt", reason="requires a Windows host")
    def test_windows_with_localappdata(self, monkeypatch):
        monkeypatch.setenv("LOCALAPPDATA", r"C:\Users\x\AppData\Local")
        assert _default_cache_dir() == Path(r"C:\Users\x\AppData\Local") / "docx-editor"

    @pytest.mark.skipif(os.name != "nt", reason="requires a Windows host")
    def test_windows_without_localappdata(self, monkeypatch):
        monkeypatch.delenv("LOCALAPPDATA", raising=False)
        assert _default_cache_dir() == Path.home() / "AppData" / "Local" / "docx-editor"


class TestWorkspaceLocation:
    """Tests for workspace location resolution (cache dir, env var, override)."""

    @staticmethod
    def _expected_subdir(source_path):
        return hashlib.sha256(str(source_path.resolve()).encode()).hexdigest()[:16]

    def test_default_lands_in_platform_cache(self, temp_docx, monkeypatch, tmp_path):
        """With no override, workspace lands under the platform cache base."""
        monkeypatch.delenv("DOCX_EDITOR_WORKSPACE_DIR", raising=False)
        monkeypatch.setattr(os, "name", "posix")
        monkeypatch.setattr(sys, "platform", "linux")
        monkeypatch.setenv("XDG_CACHE_HOME", str(tmp_path))

        workspace = Workspace(temp_docx)
        expected = tmp_path / "docx-editor" / self._expected_subdir(temp_docx)
        assert workspace.workspace_path == expected
        assert workspace.workspace_path.exists()
        workspace.close()

    def test_env_var_used_when_no_param(self, temp_docx, tmp_path, monkeypatch):
        """DOCX_EDITOR_WORKSPACE_DIR is used as the base when no param is given."""
        monkeypatch.setenv("DOCX_EDITOR_WORKSPACE_DIR", str(tmp_path))
        workspace = Workspace(temp_docx)
        assert workspace.workspace_path == tmp_path / self._expected_subdir(temp_docx)
        workspace.close()

    def test_explicit_param_overrides_env_var(self, temp_docx, tmp_path, monkeypatch):
        """The workspace_dir param wins over the env var."""
        env_base = tmp_path / "env_base"
        param_base = tmp_path / "param_base"
        monkeypatch.setenv("DOCX_EDITOR_WORKSPACE_DIR", str(env_base))

        workspace = Workspace(temp_docx, workspace_dir=param_base)
        assert workspace.workspace_path == param_base / self._expected_subdir(temp_docx)
        # The env-var base must not have been touched.
        assert not (env_base / self._expected_subdir(temp_docx)).exists()
        workspace.close()

    def test_relative_workspace_dir_next_to_source(self, temp_docx):
        """A relative workspace_dir resolves against the source directory."""
        workspace = Workspace(temp_docx, workspace_dir=".docx")
        expected = temp_docx.parent / ".docx" / self._expected_subdir(temp_docx)
        assert workspace.workspace_path == expected
        assert workspace.workspace_path.exists()
        workspace.close()

    def test_distinct_sources_get_distinct_dirs(self, simple_docx, temp_dir, tmp_path, monkeypatch):
        """Two different source files map to different hash subdirs, same base."""
        import shutil

        monkeypatch.setenv("DOCX_EDITOR_WORKSPACE_DIR", str(tmp_path))
        doc_a = temp_dir / "a.docx"
        doc_b = temp_dir / "b.docx"
        shutil.copy(simple_docx, doc_a)
        shutil.copy(simple_docx, doc_b)

        ws_a = Workspace(doc_a)
        ws_b = Workspace(doc_b)
        assert ws_a.workspace_path != ws_b.workspace_path
        assert ws_a.workspace_path.parent == tmp_path
        assert ws_b.workspace_path.parent == tmp_path

        ws_a.close()
        ws_b.close()

    def test_sync_error_message_contains_workspace_path(self, temp_docx, tmp_path, monkeypatch):
        """WorkspaceSyncError from the cache location prints the workspace path."""
        import time

        monkeypatch.setenv("DOCX_EDITOR_WORKSPACE_DIR", str(tmp_path))
        workspace = Workspace(temp_docx)
        workspace_path = workspace.workspace_path
        workspace.close(cleanup=False)

        # Modify the source so the mtime check fails on reopen.
        time.sleep(0.1)
        temp_docx.write_bytes(temp_docx.read_bytes() + b"\x00")

        with pytest.raises(WorkspaceSyncError, match=re.escape(str(workspace_path))):
            Workspace(temp_docx)

        Workspace.delete(temp_docx)


class TestWorkspaceLocationHardening:
    """Regression tests for workspace path resolution edge cases."""

    def test_non_utf8_source_filename(self, simple_docx, temp_dir, tmp_path, monkeypatch):
        """A source path with undecodable bytes must not raise UnicodeEncodeError.

        str(Path) carries such bytes as surrogates, which strict UTF-8 .encode()
        rejects; the hash key goes through os.fsencode instead.
        """
        import shutil

        monkeypatch.setenv("DOCX_EDITOR_WORKSPACE_DIR", str(tmp_path))
        # "café.docx" in latin-1 — undecodable as UTF-8, so it round-trips as a surrogate.
        weird = temp_dir / os.fsdecode(b"caf\xe9.docx")
        shutil.copy(simple_docx, weird)

        workspace = Workspace(weird)
        assert workspace.workspace_path.exists()
        assert workspace.workspace_path.parent == tmp_path
        workspace.close()

    def test_env_var_is_tilde_expanded(self, temp_docx, tmp_path, monkeypatch):
        """A ~ in the env var is expanded, not taken as a literal directory name."""
        monkeypatch.setenv("HOME", str(tmp_path))
        monkeypatch.setenv("DOCX_EDITOR_WORKSPACE_DIR", "~/ws")

        workspace = Workspace(temp_docx)
        assert workspace.workspace_path.parent == tmp_path / "ws"
        assert "~" not in str(workspace.workspace_path)
        workspace.close()

    def test_empty_workspace_dir_is_treated_as_unset(self, temp_docx, tmp_path, monkeypatch):
        """workspace_dir="" must not rebase the workspace onto the document's own dir."""
        monkeypatch.setenv("DOCX_EDITOR_WORKSPACE_DIR", str(tmp_path))

        workspace = Workspace(temp_docx, workspace_dir="")
        assert workspace.workspace_path.parent == tmp_path
        assert workspace.workspace_path.parent != temp_docx.parent
        workspace.close()

    def test_relative_xdg_cache_home_is_ignored(self, monkeypatch, tmp_path):
        """Per the XDG spec, a relative XDG_CACHE_HOME is ignored, not honored."""
        monkeypatch.setattr(os, "name", "posix")
        monkeypatch.setattr(sys, "platform", "linux")
        monkeypatch.setenv("HOME", str(tmp_path))
        monkeypatch.setenv("XDG_CACHE_HOME", "relative-cache")

        assert _default_cache_dir() == tmp_path / ".cache" / "docx-editor"

    def test_workspace_dir_is_owner_only(self, temp_docx, tmp_path, monkeypatch):
        """The workspace holds document plaintext — it must not be group/world readable."""
        monkeypatch.setenv("DOCX_EDITOR_WORKSPACE_DIR", str(tmp_path / "base"))

        workspace = Workspace(temp_docx)
        assert stat.S_IMODE(workspace.workspace_path.stat().st_mode) == 0o700
        assert stat.S_IMODE(workspace.workspace_path.parent.stat().st_mode) == 0o700
        workspace.close()

    # NB: a skipif condition is evaluated at collection time on every platform,
    # so it must not touch os.geteuid (absent on Windows) unless already guarded
    # by the short-circuiting os.name check.
    @pytest.mark.skipif(
        os.name != "posix" or os.geteuid() == 0,
        reason="POSIX only; root ignores directory permissions",
    )
    def test_unwritable_base_raises_workspace_error(self, temp_docx, tmp_path, monkeypatch):
        """A filesystem failure surfaces as WorkspaceError, not a raw PermissionError."""
        readonly = tmp_path / "readonly"
        readonly.mkdir(mode=0o500)
        monkeypatch.setenv("DOCX_EDITOR_WORKSPACE_DIR", str(readonly / "nested"))

        with pytest.raises(WorkspaceError, match="DOCX_EDITOR_WORKSPACE_DIR"):
            Workspace(temp_docx)

    def test_adopting_another_documents_workspace_is_refused(self, simple_docx, temp_dir, tmp_path, monkeypatch):
        """A workspace whose meta names a different source must not be adopted.

        The workspace dir is keyed by a hash of the source path, so a mismatch
        means a collision or a planted directory; adopting it would let save()
        pack the other document's XML over this one's source.
        """
        import shutil

        monkeypatch.setenv("DOCX_EDITOR_WORKSPACE_DIR", str(tmp_path))
        victim = temp_dir / "victim.docx"
        shutil.copy(simple_docx, victim)

        workspace = Workspace(victim)
        workspace_path = workspace.workspace_path
        workspace.close(cleanup=False)

        # Rewrite meta.json so it claims to belong to a different document.
        meta_path = workspace_path / Workspace.META_FILE
        meta = json.loads(meta_path.read_text())
        meta["source_path"] = str(temp_dir / "other.docx")
        meta_path.write_text(json.dumps(meta))

        with pytest.raises(WorkspaceError, match="different document"):
            Workspace(victim)

        Workspace.delete(victim)


class TestSaveStaleness:
    def test_save_raises_when_source_changed_externally(self, temp_docx):
        ws = Workspace(temp_docx, author="Test")
        # Simulate an external edit: change the source file's content.
        temp_docx.write_bytes(temp_docx.read_bytes() + b"\x00")
        try:
            with pytest.raises(WorkspaceSyncError, match="changed on disk"):
                ws.save()
        finally:
            ws.close()

    def test_save_force_overwrites_changed_source(self, temp_docx):
        ws = Workspace(temp_docx, author="Test")
        temp_docx.write_bytes(temp_docx.read_bytes() + b"\x00")
        try:
            result = ws.save(force=True)
            assert result == temp_docx.resolve()
            # meta was refreshed: a follow-up save must not raise.
            ws.save()
        finally:
            ws.close()

    def test_save_to_other_destination_ignores_stale_source(self, temp_docx, temp_dir):
        ws = Workspace(temp_docx, author="Test")
        temp_docx.write_bytes(temp_docx.read_bytes() + b"\x00")
        try:
            out = ws.save(destination=temp_dir / "elsewhere.docx")
            assert out.exists()
        finally:
            ws.close()

    def test_save_recreates_deleted_source_without_force(self, temp_docx):
        """A deleted source has no external edits to lose — saving must restore it."""
        ws = Workspace(temp_docx, author="Test")
        try:
            temp_docx.unlink()
            out = ws.save()
            assert out.exists()
            assert out.stat().st_size > 0
        finally:
            ws.close()

    def test_save_returns_the_path_the_caller_passed(self, temp_docx, temp_dir, monkeypatch):
        """save() must not silently upgrade the caller's path to a resolved absolute one."""
        ws = Workspace(temp_docx, author="Test")
        monkeypatch.chdir(temp_dir)
        try:
            out = ws.save(destination=Path("relative_out.docx"))
            assert out == Path("relative_out.docx")
            assert out.exists()
        finally:
            ws.close()

    def test_reopen_is_consistent_with_save_staleness(self, temp_docx):
        """__init__ and save() must share one staleness predicate.

        Otherwise a workspace whose recorded size drifted from the source (an mtime
        match but a size mismatch) opens clean, then save() rejects it with a
        "changed on disk" error that is simply false.
        """
        ws = Workspace(temp_docx, author="Test")
        ws.meta["source_size"] = ws.meta["source_size"] + 999  # drift the size only
        ws._save_meta()
        ws.close(cleanup=False)

        with pytest.raises(WorkspaceSyncError):
            Workspace(temp_docx, author="Test")  # fails early at open, not late at save

        # And the documented recovery (drop the workspace) works.
        Workspace.delete(temp_docx)
        Workspace(temp_docx, author="Test").close()


class TestSha256Staleness:
    """Content-hash staleness detection in sync_check() (issue #22 follow-up)."""

    def test_same_size_same_mtime_content_swap_is_detected(self, temp_docx):
        """The case mtime+size provably cannot catch: same-length content swap
        with the timestamp restored."""
        ws = Workspace(temp_docx, author="Test")
        ws.close(cleanup=False)

        original = bytearray(temp_docx.read_bytes())
        stat_before = temp_docx.stat()
        original[-1] ^= 0xFF  # flip one byte, length unchanged
        temp_docx.write_bytes(bytes(original))
        os.utime(temp_docx, (stat_before.st_atime, stat_before.st_mtime))

        with pytest.raises(WorkspaceSyncError):
            Workspace(temp_docx, author="Test")

        Workspace.delete(temp_docx)

    def test_touch_with_identical_content_is_not_stale(self, temp_docx):
        """An mtime bump over identical bytes (touch, cloud-sync re-download)
        is not an external edit — the workspace adopts."""
        ws = Workspace(temp_docx, author="Test")
        ws.close(cleanup=False)

        stat = temp_docx.stat()
        os.utime(temp_docx, (stat.st_atime, stat.st_mtime + 10))

        ws2 = Workspace(temp_docx, author="Test")  # must not raise
        ws2.close()

    def test_legacy_meta_without_sha256_uses_mtime_size(self, temp_docx):
        """meta.json written before the hash existed keeps the old semantics:
        an mtime-only bump still reads as stale."""
        ws = Workspace(temp_docx, author="Test")
        ws.close(cleanup=False)

        meta_path = ws.workspace_path / "meta.json"
        meta = json.loads(meta_path.read_text())
        del meta["source_sha256"]
        meta_path.write_text(json.dumps(meta))

        stat = temp_docx.stat()
        os.utime(temp_docx, (stat.st_atime, stat.st_mtime + 10))

        with pytest.raises(WorkspaceSyncError):
            Workspace(temp_docx, author="Test")

        Workspace.delete(temp_docx)

    def test_save_to_source_refreshes_sha256(self, temp_docx):
        ws = Workspace(temp_docx, author="Test")
        try:
            ws.save()
            meta = json.loads((ws.workspace_path / "meta.json").read_text())
            assert meta["source_sha256"] == hashlib.sha256(temp_docx.read_bytes()).hexdigest()
        finally:
            ws.close()


class TestDirtyWorkspace:
    """Tests for the dirty flag guarding stale-workspace adoption (issue #31)."""

    def _meta_on_disk(self, ws):
        with open(ws.workspace_path / "meta.json") as f:
            return json.load(f)

    def test_save_elsewhere_sets_dirty_on_disk(self, temp_docx, temp_dir):
        ws = Workspace(temp_docx, author="Test")
        try:
            ws.save(destination=temp_dir / "elsewhere.docx")
            assert self._meta_on_disk(ws)["dirty"] is True
        finally:
            ws.close()

    def test_save_to_source_clears_dirty(self, temp_docx, temp_dir):
        ws = Workspace(temp_docx, author="Test")
        try:
            ws.save(destination=temp_dir / "elsewhere.docx")
            ws.save()  # back to the source: workspace no longer ahead of it
            assert self._meta_on_disk(ws)["dirty"] is False
        finally:
            ws.close()

    def test_dirty_workspace_is_refused_for_adoption(self, temp_docx, temp_dir):
        ws = Workspace(temp_docx, author="Test")
        ws.save(destination=temp_dir / "elsewhere.docx")
        ws.close(cleanup=False)

        with pytest.raises(WorkspaceSyncError, match="unsaved changes"):
            Workspace(temp_docx, author="Test")

        Workspace.delete(temp_docx)

    def test_dirty_workspace_rescue_via_create_false(self, temp_docx, temp_dir):
        """The error message's rescue hatch must keep working: create=False
        reattaches without adoption checks so the contents can be saved out."""
        ws = Workspace(temp_docx, author="Test")
        ws.save(destination=temp_dir / "elsewhere.docx")
        ws.close(cleanup=False)

        rescue = Workspace(temp_docx, create=False)
        out = rescue.save(destination=temp_dir / "rescued.docx")
        assert out.exists()

        Workspace.delete(temp_docx)

    def test_failed_in_place_save_sets_dirty(self, temp_docx):
        """save() flags the workspace write-ahead even for in-place saves, so a
        pack failure cannot leave a diverged workspace unflagged for adoption."""
        from unittest.mock import patch

        ws = Workspace(temp_docx, author="Test")
        try:
            with patch("docx_editor.workspace.pack_document", return_value=False):
                with pytest.raises(WorkspaceError, match="Failed to pack"):
                    ws.save()
            assert self._meta_on_disk(ws)["dirty"] is True
        finally:
            ws.close()

    def test_mark_dirty_is_idempotent_and_persisted(self, temp_docx):
        ws = Workspace(temp_docx, author="Test")
        try:
            ws.mark_dirty()
            assert self._meta_on_disk(ws)["dirty"] is True
            ws.mark_dirty()  # second call is a no-op, not an error
            assert self._meta_on_disk(ws)["dirty"] is True
        finally:
            ws.close()

    def test_fresh_document_open_does_not_mark_dirty(self, temp_docx):
        """Open-time tracking setup writes (people.xml, settings, rels) are
        deterministic bookkeeping — they must not flag the workspace, or every
        crashed-but-unedited session would force force_recreate at next open."""
        doc = Document.open(temp_docx, author="Test")
        try:
            assert self._meta_on_disk(doc)["dirty"] is False
        finally:
            doc.close(cleanup=False)

        # And the untouched workspace still adopts.
        Document.open(temp_docx, author="Test").close()

    def test_in_memory_edit_does_not_mark_dirty(self, temp_docx):
        """DOM edits die with the process — only disk writes need the flag."""
        doc = Document.open(temp_docx, author="Test")
        try:
            doc.replace("fox", "cat", paragraph=find_ref(doc, "fox"))
            assert self._meta_on_disk(doc)["dirty"] is False

            doc.save()  # flag set write-ahead, cleared by the in-place save
            assert self._meta_on_disk(doc)["dirty"] is False
        finally:
            doc.close()

    def test_add_comment_marks_dirty_before_save(self, temp_docx):
        """add_comment() copies comment templates into the workspace on disk —
        a post-open write, so it is flagged even though save() hasn't run."""
        doc = Document.open(temp_docx, author="Test")
        try:
            doc.add_comment("fox", "needs review")
            assert self._meta_on_disk(doc)["dirty"] is True
        finally:
            doc.close()

    def test_pre_upgrade_meta_without_dirty_key_adopts(self, temp_docx):
        """meta.json written before the flag existed must adopt exactly as before."""
        ws = Workspace(temp_docx, author="Test")
        ws.close(cleanup=False)

        meta_path = ws.workspace_path / "meta.json"
        with open(meta_path) as f:
            meta = json.load(f)
        del meta["dirty"]
        with open(meta_path, "w") as f:
            json.dump(meta, f)

        ws2 = Workspace(temp_docx, author="Test")  # must not raise
        ws2.close()


class TestWorkspaceLock:
    """Advisory per-workspace lock against concurrent opens (issue #24)."""

    def _lock_path(self, ws):
        return ws.workspace_path.with_name(ws.workspace_path.name + ".lock")

    @staticmethod
    def _lock_pid(lock):
        """PID part of the "<pid>:<token>" lock content."""
        return lock.read_text().split(":", 1)[0]

    def test_second_open_in_same_process_is_refused(self, temp_docx):
        ws = Workspace(temp_docx, author="Test")
        try:
            with pytest.raises(WorkspaceLockedError) as exc_info:
                Workspace(temp_docx, author="Test")
            assert exc_info.value.pid == os.getpid()
            assert "close()" in str(exc_info.value)
        finally:
            ws.close()

    def test_lock_holds_pid_and_close_releases_it(self, temp_docx):
        ws = Workspace(temp_docx, author="Test")
        lock = self._lock_path(ws)
        assert self._lock_pid(lock) == str(os.getpid())

        ws.close(cleanup=False)  # workspace kept — the lock must still go
        assert not lock.exists()

        Workspace(temp_docx, author="Test").close()  # and reopen adopts freely

    def test_rescue_create_false_conflicts_with_live_session(self, temp_docx):
        """The create=False rescue flow reads the workspace and writes meta, so
        it must respect a live session's lock like any other open."""
        ws = Workspace(temp_docx, author="Test")
        try:
            with pytest.raises(WorkspaceLockedError):
                Workspace(temp_docx, create=False)
        finally:
            ws.close()

    def test_failed_init_releases_lock(self, temp_docx):
        """Every __init__ failure path must release the lock, or the process
        locks itself out of its own retry/rescue."""
        ws = Workspace(temp_docx, author="Test")
        ws.close(cleanup=False)
        temp_docx.write_bytes(temp_docx.read_bytes() + b"\x00")  # now stale

        # Were the lock leaked by the first failure, the second attempt would
        # raise WorkspaceLockedError instead of the real, actionable error.
        with pytest.raises(WorkspaceSyncError):
            Workspace(temp_docx, author="Test")
        with pytest.raises(WorkspaceSyncError):
            Workspace(temp_docx, author="Test")

        Workspace.delete(temp_docx)

    @pytest.mark.skipif(os.name != "posix", reason="PID 1 is only guaranteed foreign-and-alive on POSIX")
    def test_foreign_live_pid_blocks_open(self, temp_docx):
        ws = Workspace(temp_docx, author="Test")
        ws.close(cleanup=False)
        lock = self._lock_path(ws)
        lock.write_text("1")  # init/launchd: alive, not ours, unkillable

        with pytest.raises(WorkspaceLockedError) as exc_info:
            Workspace(temp_docx, author="Test")
        assert exc_info.value.pid == 1
        assert exc_info.value.lock_path == lock

        Workspace.delete(temp_docx)

    def test_stale_lock_from_dead_process_is_reclaimed(self, temp_docx):
        proc = subprocess.Popen([sys.executable, "-c", "pass"])
        proc.wait()  # reaped: the pid is gone

        ws = Workspace(temp_docx, author="Test")
        ws.close(cleanup=False)
        self._lock_path(ws).write_text(str(proc.pid))

        ws2 = Workspace(temp_docx, author="Test")  # reclaims silently
        try:
            assert self._lock_pid(self._lock_path(ws2)) == str(os.getpid())
        finally:
            ws2.close()

    @pytest.mark.parametrize("content", ["not-a-pid", "0", "-1"])
    def test_corrupt_lock_content_is_reclaimed(self, temp_docx, content):
        """Unparseable and non-positive pids are corrupt, not holders — probing
        them would be dangerous (waitpid(-1) reaps an arbitrary child)."""
        ws = Workspace(temp_docx, author="Test")
        ws.close(cleanup=False)
        self._lock_path(ws).write_text(content)

        ws2 = Workspace(temp_docx, author="Test")
        try:
            assert self._lock_pid(self._lock_path(ws2)) == str(os.getpid())
        finally:
            ws2.close()

    def test_two_processes_conflict(self, temp_docx):
        """A second OS process is refused while the first holds the document."""
        script = (
            "import sys\n"
            "from docx_editor.workspace import Workspace\n"
            "ws = Workspace(sys.argv[1], author='Child')\n"
            "print('READY', flush=True)\n"
            "sys.stdin.read()\n"  # hold the lock until the parent is done
        )
        proc = subprocess.Popen(
            [sys.executable, "-c", script, str(temp_docx)],
            stdout=subprocess.PIPE,
            stdin=subprocess.PIPE,
            text=True,
        )
        try:
            assert proc.stdout is not None  # stdout=PIPE guarantees it
            assert proc.stdout.readline().strip() == "READY"
            with pytest.raises(WorkspaceLockedError) as exc_info:
                Workspace(temp_docx, author="Parent")
            assert exc_info.value.pid == proc.pid
        finally:
            proc.kill()
            proc.wait()

    @pytest.mark.skipif(os.name != "posix", reason="PID 1 is only guaranteed foreign-and-alive on POSIX")
    def test_force_recreate_takes_over_foreign_live_lock(self, temp_docx):
        """force_recreate is the universal escape hatch — it must break even a
        live foreign lock, and the new session ends up holding its own."""
        ws = Workspace(temp_docx, author="Test")
        ws.close(cleanup=False)
        lock = self._lock_path(ws)
        lock.write_text("1")

        doc = Document.open(temp_docx, author="Test", force_recreate=True)
        try:
            assert self._lock_pid(lock) == str(os.getpid())
        finally:
            doc.close()

    def test_superseded_session_close_keeps_new_lock(self, temp_docx):
        """After a force_recreate takeover, the superseded session's close()
        must not release the new session's lock (release is token-aware)."""
        ws_old = Workspace(temp_docx, author="Test")
        doc_new = Document.open(temp_docx, author="Test", force_recreate=True)
        try:
            ws_old.close(cleanup=False)  # superseded; must not unlink

            lock = self._lock_path(ws_old)
            assert lock.exists()
            with pytest.raises(WorkspaceLockedError):  # new session still holds
                Workspace(temp_docx, author="Test")
        finally:
            doc_new.close()

    def test_close_releases_lock_even_if_cleanup_fails(self, temp_docx, monkeypatch):
        ws = Workspace(temp_docx, author="Test")
        lock = self._lock_path(ws)

        def boom(path):
            raise OSError("simulated rmtree failure")

        monkeypatch.setattr("docx_editor.workspace.shutil.rmtree", boom)
        with pytest.raises(OSError, match="simulated"):
            ws.close()
        monkeypatch.undo()

        assert not lock.exists()
        Workspace.delete(temp_docx)

    def test_delete_removes_lock_sidecar(self, temp_docx):
        ws = Workspace(temp_docx, author="Test")
        lock = self._lock_path(ws)
        assert lock.exists()

        Workspace.delete(temp_docx)
        assert not lock.exists()

    def test_pid_alive_probe(self):
        assert _pid_alive(os.getpid()) is True

        proc = subprocess.Popen([sys.executable, "-c", "pass"])
        proc.wait()  # reaped by wait(): dead even as our own child
        assert _pid_alive(proc.pid) is False

    @pytest.mark.skipif(not Path("/proc").exists(), reason="needs /proc to observe zombie state")
    def test_zombie_child_reads_alive_unless_reaping(self):
        """Without reap=True the probe must not waitpid an arbitrary pid — that
        would steal the exit status of another component's child. The zombie
        conservatively reads as alive; only the owning caller reaps."""
        import time

        proc = subprocess.Popen([sys.executable, "-c", "pass"])
        try:
            deadline = time.monotonic() + 10
            while time.monotonic() < deadline:
                state = Path(f"/proc/{proc.pid}/stat").read_text().rsplit(")", 1)[1].split()[0]
                if state == "Z":
                    break
                time.sleep(0.05)
            else:
                pytest.fail("child never became a zombie")

            assert _pid_alive(proc.pid) is True  # unreaped: conservative
            assert _pid_alive(proc.pid, reap=True) is False  # owner may reap
        finally:
            proc.wait()  # no-op if the reap above already collected it

    def test_dropped_session_lock_is_freed_on_gc(self, temp_docx):
        """A session dropped without close() must not lock its own process out
        of the document forever — the lock names a live pid, so the staleness
        probe would never reclaim it. The GC finalizer releases it."""
        import gc

        doc = Document.open(temp_docx, author="Test")
        doc.save()  # workspace clean, so the reopen below adopts
        del doc  # dropped without close()
        gc.collect()  # refcounting already freed it on CPython; be explicit

        doc2 = Document.open(temp_docx, author="Test")  # must not raise
        doc2.close()

    def test_dropped_dirty_session_is_rescuable_in_process(self, temp_docx, temp_dir):
        """The rescue path documented by WorkspaceSyncError must work in the
        same process after the dirty session is dropped without close()."""
        import gc

        doc = Document.open(temp_docx, author="Test")
        doc.add_comment("fox", "unsaved edit")  # marks the workspace dirty
        del doc
        gc.collect()

        # The dirty refusal (not a lock error) is what the user must see...
        with pytest.raises(WorkspaceSyncError, match="unsaved changes"):
            Document.open(temp_docx, author="Test")
        # ...and its advertised rescue hatch must actually work.
        rescue = Workspace(temp_docx, create=False)
        out = rescue.save(destination=temp_dir / "rescued.docx")
        assert out.exists()
        rescue.close()

    def test_double_close_is_idempotent(self, temp_docx):
        """A second close() finds the lock already released and the workspace
        already gone — both must be silent no-ops."""
        ws = Workspace(temp_docx, author="Test")
        ws.close()
        ws.close()
        assert not self._lock_path(ws).exists()

    @pytest.mark.skipif(
        os.name != "posix" or os.geteuid() == 0,
        reason="chmod 000 only denies reads on POSIX as non-root",
    )
    def test_unreadable_lock_file_is_reclaimed(self, temp_docx):
        """A lock whose content cannot be read has no provable live holder —
        it counts as stale and is reclaimed, like corrupt content."""
        ws = Workspace(temp_docx, author="Test")
        ws.close(cleanup=False)
        lock = self._lock_path(ws)
        lock.write_text("123:abc")
        lock.chmod(0)

        ws2 = Workspace(temp_docx, author="Test")  # must not raise
        try:
            assert self._lock_pid(self._lock_path(ws2)) == str(os.getpid())
        finally:
            ws2.close()

    def test_lost_reclaim_race_raises_locked(self, temp_docx, monkeypatch):
        """If the lock content keeps changing between the stale read and the
        reclaim re-check, another process is actively re-creating it — after
        both attempts the open must give up with WorkspaceLockedError instead
        of unlinking a competitor's fresh lock or looping forever."""
        ws = Workspace(temp_docx, author="Test")
        lock = self._lock_path(ws)
        ws.close(cleanup=True)
        lock.write_text("placeholder")  # so O_EXCL keeps failing

        reads = iter(f"not-a-pid-{n}" for n in range(10))
        monkeypatch.setattr(Workspace, "_read_lock_content", lambda self: next(reads))

        try:
            with pytest.raises(WorkspaceLockedError, match="reclaiming"):
                Workspace(temp_docx, author="Test")
        finally:
            monkeypatch.undo()
            lock.unlink()

    def test_failed_token_write_removes_lock_file(self, temp_docx, monkeypatch):
        """If the token write fails right after the O_EXCL create, the file
        must be removed: an orphaned lock would name this live pid, which the
        staleness probe could never reclaim."""

        def boom(fd, *args, **kwargs):
            os.close(fd)
            raise RuntimeError("simulated fdopen failure")

        monkeypatch.setattr("docx_editor.workspace.os.fdopen", boom)
        with pytest.raises(RuntimeError, match="simulated"):
            Workspace(temp_docx, author="Test")
        monkeypatch.undo()

        # Were the lock orphaned, this open would raise WorkspaceLockedError.
        Workspace(temp_docx, author="Test").close()


class TestAtomicMetaSave:
    """_save_meta() must never leave a truncated meta.json (issue #22)."""

    def test_crash_mid_write_preserves_previous_meta(self, temp_docx, monkeypatch):
        """A crash mid-write leaves the old meta.json intact — including the
        write-ahead dirty flag, which a truncated file would destroy."""
        ws = Workspace(temp_docx, author="Test")
        try:
            ws.mark_dirty()  # dirty: true is now on disk

            def exploding_dump(obj, fp, **kwargs):
                fp.write("{ truncated garbage")
                raise RuntimeError("simulated crash mid-write")

            monkeypatch.setattr("docx_editor.workspace.json.dump", exploding_dump)
            with pytest.raises(RuntimeError, match="simulated crash"):
                ws._save_meta()
            monkeypatch.undo()

            with open(ws.workspace_path / "meta.json") as f:
                meta = json.load(f)  # still parses: the old file was never touched
            assert meta["dirty"] is True
            assert not (ws.workspace_path / "meta.json.tmp").exists()
        finally:
            ws.close()

    def test_failed_write_leaves_workspace_adoptable(self, temp_docx, monkeypatch):
        """After a crashed meta write, the next open must not see a corrupt
        workspace (the pre-fix symptom was a misleading WorkspaceExistsError)."""
        ws = Workspace(temp_docx, author="Test")

        def exploding_dump(obj, fp, **kwargs):
            fp.write("{ truncated garbage")
            raise RuntimeError("simulated crash mid-write")

        monkeypatch.setattr("docx_editor.workspace.json.dump", exploding_dump)
        with pytest.raises(RuntimeError):
            ws.meta["last_saved"] = "whenever"
            ws._save_meta()
        monkeypatch.undo()
        ws.close(cleanup=False)

        ws2 = Workspace(temp_docx, author="Test")  # adopts the intact meta
        ws2.close()
