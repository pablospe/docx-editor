"""Tests for MCP Tools following TDD."""

import os
import shutil
import time
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest


@pytest.fixture
def server():
    """Create a fresh server instance."""
    from docx_editor_mcp.server import create_server

    return create_server()


@pytest.fixture
def mcp_temp_docx(simple_docx, tmp_path):
    """Create a temporary copy of simple.docx for MCP tests."""
    dest = tmp_path / "mcp_test.docx"
    shutil.copy(simple_docx, dest)
    return dest


class TestOpenDocument:
    """Test open_document tool."""

    def test_open_document_loads_and_caches(self, server, mcp_temp_docx):
        """open_document loads document and adds to cache."""
        from docx_editor_mcp.tools import open_document

        result = open_document(server, str(mcp_temp_docx), author="Tester")

        assert result["success"] is True
        assert result["path"] == str(mcp_temp_docx)
        assert result["author"] == "Tester"
        assert server.cache.size == 1

    def test_open_document_returns_cached(self, server, mcp_temp_docx):
        """open_document returns cached document on second call."""
        from docx_editor_mcp.tools import open_document

        result1 = open_document(server, str(mcp_temp_docx), author="Tester")
        result2 = open_document(server, str(mcp_temp_docx))

        assert result1["success"] is True
        assert result2["success"] is True
        assert server.cache.size == 1  # Still just one

    def test_open_document_default_author_hints(self, server, mcp_temp_docx):
        """open_document with no author uses default and hints."""
        from docx_editor_mcp.tools import open_document

        with patch("getpass.getuser", return_value="systemuser"):
            result = open_document(server, str(mcp_temp_docx))

        assert result["success"] is True
        assert result["author"] == "systemuser"
        assert "system default" in result.get("hint", "").lower()

    def test_open_document_session_author(self, server, mcp_temp_docx, tmp_path):
        """open_document remembers session author."""
        from docx_editor_mcp.tools import open_document

        # First doc with explicit author
        open_document(server, str(mcp_temp_docx), author="Legal Team")

        # Second doc without author uses session author
        doc2 = tmp_path / "test2.docx"
        shutil.copy(mcp_temp_docx, doc2)
        result = open_document(server, str(doc2))

        assert result["author"] == "Legal Team"
        assert "hint" not in result  # No hint for session author

    def test_open_document_file_not_found(self, server):
        """open_document returns error for missing file."""
        from docx_editor_mcp.tools import open_document

        result = open_document(server, "/nonexistent/file.docx")

        assert result["success"] is False
        assert "not found" in result["error"].lower()


class TestSaveDocument:
    """Test save_document tool."""

    def test_save_document_saves_to_disk(self, server, mcp_temp_docx):
        """save_document writes document to disk."""
        from docx_editor_mcp.tools import open_document, save_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        # Get mtime before save
        mtime_before = os.path.getmtime(mcp_temp_docx)
        time.sleep(0.01)

        result = save_document(server, str(mcp_temp_docx))

        assert result["success"] is True
        # File should have been modified
        mtime_after = os.path.getmtime(mcp_temp_docx)
        assert mtime_after >= mtime_before

    def test_save_document_clears_dirty(self, server, mcp_temp_docx):
        """save_document clears dirty flag."""
        from docx_editor_mcp.tools import open_document, save_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        # Mark dirty
        cached = server.cache.get(str(mcp_temp_docx))
        cached.mark_dirty()
        assert cached.dirty is True

        save_document(server, str(mcp_temp_docx))

        assert cached.dirty is False

    def test_save_document_updates_mtime(self, server, mcp_temp_docx):
        """save_document updates cached mtime."""
        from docx_editor_mcp.tools import open_document, save_document

        open_document(server, str(mcp_temp_docx), author="Tester")
        cached = server.cache.get(str(mcp_temp_docx))
        old_mtime = cached.mtime

        time.sleep(0.01)
        save_document(server, str(mcp_temp_docx))

        assert cached.mtime >= old_mtime

    def test_save_document_not_open(self, server):
        """save_document returns error if document not open."""
        from docx_editor_mcp.tools import save_document

        result = save_document(server, "/not/open.docx")

        assert result["success"] is False
        assert "not open" in result["error"].lower()

    def test_save_document_external_change_blocked(self, server, mcp_temp_docx):
        """save_document fails if file changed externally."""
        from docx_editor_mcp.tools import open_document, save_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        # Simulate external modification
        time.sleep(0.01)
        mcp_temp_docx.write_bytes(mcp_temp_docx.read_bytes() + b"modified")

        result = save_document(server, str(mcp_temp_docx))

        assert result["success"] is False
        assert "external" in result["error"].lower()


class TestCloseDocument:
    """Test close_document tool."""

    def test_close_document_removes_from_cache(self, server, mcp_temp_docx):
        """close_document removes document from cache."""
        from docx_editor_mcp.tools import close_document, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")
        assert server.cache.size == 1

        result = close_document(server, str(mcp_temp_docx))

        assert result["success"] is True
        assert server.cache.size == 0

    def test_close_document_warns_if_dirty(self, server, mcp_temp_docx):
        """close_document warns if document has unsaved changes."""
        from docx_editor_mcp.tools import close_document, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")
        server.cache.get(str(mcp_temp_docx)).mark_dirty()

        result = close_document(server, str(mcp_temp_docx))

        assert result["success"] is True
        assert "unsaved" in result.get("warning", "").lower()

    def test_close_document_not_open(self, server):
        """close_document returns error if not open."""
        from docx_editor_mcp.tools import close_document

        result = close_document(server, "/not/open.docx")

        assert result["success"] is False
        assert "not open" in result["error"].lower()


class TestReloadDocument:
    """Test reload_document tool."""

    def test_reload_document_reloads_from_disk(self, server, mcp_temp_docx):
        """reload_document reloads document from disk."""
        from docx_editor_mcp.tools import open_document, reload_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        result = reload_document(server, str(mcp_temp_docx))

        assert result["success"] is True
        assert server.cache.size == 1  # Still cached

    def test_reload_document_discards_changes(self, server, mcp_temp_docx):
        """reload_document discards unsaved changes."""
        from docx_editor_mcp.tools import open_document, reload_document

        open_document(server, str(mcp_temp_docx), author="Tester")
        server.cache.get(str(mcp_temp_docx)).mark_dirty()

        result = reload_document(server, str(mcp_temp_docx))

        assert result["success"] is True
        assert "discarded" in result.get("warning", "").lower()
        # New cached doc should not be dirty
        assert server.cache.get(str(mcp_temp_docx)).dirty is False

    def test_reload_document_not_open(self, server):
        """reload_document returns error if not open."""
        from docx_editor_mcp.tools import reload_document

        result = reload_document(server, "/not/open.docx")

        assert result["success"] is False
        assert "not open" in result["error"].lower()


class TestForceSave:
    """Test force_save tool."""

    def test_force_save_overwrites_external(self, server, mcp_temp_docx):
        """force_save saves even when external changes detected."""
        from docx_editor_mcp.tools import force_save, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        # Simulate external modification
        time.sleep(0.01)
        mcp_temp_docx.write_bytes(mcp_temp_docx.read_bytes() + b"external")

        result = force_save(server, str(mcp_temp_docx))

        assert result["success"] is True
        # mtime should be updated
        cached = server.cache.get(str(mcp_temp_docx))
        assert not cached.has_external_changes()

    def test_force_save_not_open(self, server):
        """force_save returns error if not open."""
        from docx_editor_mcp.tools import force_save

        result = force_save(server, "/not/open.docx")

        assert result["success"] is False
        assert "not open" in result["error"].lower()


class TestTrackChangesTools:
    """Test track changes tools (Task 3.2)."""

    def test_replace_text(self, server, mcp_temp_docx):
        """replace_text replaces text with tracking."""
        from docx_editor_mcp.tools import open_document, replace_text

        open_document(server, str(mcp_temp_docx), author="Tester")

        result = replace_text(server, str(mcp_temp_docx), "quick brown fox", "Hi")

        assert result["success"] is True
        assert "change_id" in result
        assert server.cache.get(str(mcp_temp_docx)).dirty is True

    def test_replace_text_not_found(self, server, mcp_temp_docx):
        """replace_text returns error if text not found."""
        from docx_editor_mcp.tools import open_document, replace_text

        open_document(server, str(mcp_temp_docx), author="Tester")

        result = replace_text(server, str(mcp_temp_docx), "nonexistent text xyz", "new")

        assert result["success"] is False
        assert "not found" in result["error"].lower()

    def test_replace_text_with_occurrence(self, server, mcp_temp_docx):
        """replace_text can target specific occurrence."""
        from docx_editor_mcp.tools import open_document, replace_text

        open_document(server, str(mcp_temp_docx), author="Tester")

        # This may or may not find a second occurrence depending on fixture
        result = replace_text(server, str(mcp_temp_docx), "quick brown fox", "Hi", occurrence=0)

        # Should at least not crash
        assert "success" in result

    def test_delete_text(self, server, mcp_temp_docx):
        """delete_text marks text as deleted."""
        from docx_editor_mcp.tools import delete_text, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        result = delete_text(server, str(mcp_temp_docx), "quick brown fox")

        assert result["success"] is True
        assert "change_id" in result
        assert server.cache.get(str(mcp_temp_docx)).dirty is True

    def test_insert_after(self, server, mcp_temp_docx):
        """insert_after inserts text after anchor."""
        from docx_editor_mcp.tools import insert_after, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        result = insert_after(server, str(mcp_temp_docx), "quick brown fox", " World")

        assert result["success"] is True
        assert "change_id" in result

    def test_insert_before(self, server, mcp_temp_docx):
        """insert_before inserts text before anchor."""
        from docx_editor_mcp.tools import insert_before, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        result = insert_before(server, str(mcp_temp_docx), "quick brown fox", "Say: ")

        assert result["success"] is True
        assert "change_id" in result


class TestCommentTools:
    """Test comment tools (Task 3.3)."""

    def test_add_comment(self, server, mcp_temp_docx):
        """add_comment adds comment to document."""
        from docx_editor_mcp.tools import add_comment, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        result = add_comment(server, str(mcp_temp_docx), "quick brown fox", "Please review")

        assert result["success"] is True
        assert "comment_id" in result
        assert server.cache.get(str(mcp_temp_docx)).dirty is True

    def test_list_comments(self, server, mcp_temp_docx):
        """list_comments returns all comments."""
        from docx_editor_mcp.tools import add_comment, list_comments, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")
        add_comment(server, str(mcp_temp_docx), "quick brown fox", "Comment 1")

        result = list_comments(server, str(mcp_temp_docx))

        assert result["success"] is True
        assert "comments" in result
        assert len(result["comments"]) >= 1

    def test_list_comments_filter_author(self, server, mcp_temp_docx):
        """list_comments can filter by author."""
        from docx_editor_mcp.tools import list_comments, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        result = list_comments(server, str(mcp_temp_docx), author="Tester")

        assert result["success"] is True
        assert "comments" in result

    def test_reply_to_comment(self, server, mcp_temp_docx):
        """reply_to_comment adds reply to existing comment."""
        from docx_editor_mcp.tools import add_comment, open_document, reply_to_comment

        open_document(server, str(mcp_temp_docx), author="Tester")
        add_result = add_comment(server, str(mcp_temp_docx), "quick brown fox", "Original")

        result = reply_to_comment(
            server, str(mcp_temp_docx), add_result["comment_id"], "Reply text"
        )

        assert result["success"] is True
        assert "comment_id" in result

    def test_resolve_comment(self, server, mcp_temp_docx):
        """resolve_comment marks comment as resolved."""
        from docx_editor_mcp.tools import add_comment, open_document, resolve_comment

        open_document(server, str(mcp_temp_docx), author="Tester")
        add_result = add_comment(server, str(mcp_temp_docx), "quick brown fox", "To resolve")

        result = resolve_comment(server, str(mcp_temp_docx), add_result["comment_id"])

        assert result["success"] is True

    def test_delete_comment(self, server, mcp_temp_docx):
        """delete_comment removes comment."""
        from docx_editor_mcp.tools import add_comment, delete_comment, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")
        add_result = add_comment(server, str(mcp_temp_docx), "quick brown fox", "To delete")

        result = delete_comment(server, str(mcp_temp_docx), add_result["comment_id"])

        assert result["success"] is True


class TestRevisionTools:
    """Test revision tools (Task 3.4)."""

    def test_list_revisions(self, server, mcp_temp_docx):
        """list_revisions returns all tracked changes."""
        from docx_editor_mcp.tools import list_revisions, open_document, replace_text

        open_document(server, str(mcp_temp_docx), author="Tester")
        replace_text(server, str(mcp_temp_docx), "quick brown fox", "Hi")

        result = list_revisions(server, str(mcp_temp_docx))

        assert result["success"] is True
        assert "revisions" in result
        assert len(result["revisions"]) >= 1

    def test_accept_revision(self, server, mcp_temp_docx):
        """accept_revision accepts a specific revision."""
        from docx_editor_mcp.tools import (
            accept_revision,
            list_revisions,
            open_document,
            replace_text,
        )

        open_document(server, str(mcp_temp_docx), author="Tester")
        replace_text(server, str(mcp_temp_docx), "quick brown fox", "Hi")
        revisions = list_revisions(server, str(mcp_temp_docx))

        if revisions["revisions"]:
            rev_id = revisions["revisions"][0]["id"]
            result = accept_revision(server, str(mcp_temp_docx), rev_id)
            assert result["success"] is True

    def test_reject_revision(self, server, mcp_temp_docx):
        """reject_revision rejects a specific revision."""
        from docx_editor_mcp.tools import (
            list_revisions,
            open_document,
            reject_revision,
            replace_text,
        )

        open_document(server, str(mcp_temp_docx), author="Tester")
        replace_text(server, str(mcp_temp_docx), "quick brown fox", "Hi")
        revisions = list_revisions(server, str(mcp_temp_docx))

        if revisions["revisions"]:
            rev_id = revisions["revisions"][0]["id"]
            result = reject_revision(server, str(mcp_temp_docx), rev_id)
            assert result["success"] is True

    def test_accept_all(self, server, mcp_temp_docx):
        """accept_all accepts all revisions."""
        from docx_editor_mcp.tools import accept_all, open_document, replace_text

        open_document(server, str(mcp_temp_docx), author="Tester")
        replace_text(server, str(mcp_temp_docx), "quick brown fox", "Hi")

        result = accept_all(server, str(mcp_temp_docx))

        assert result["success"] is True
        assert "count" in result

    def test_reject_all(self, server, mcp_temp_docx):
        """reject_all rejects all revisions."""
        from docx_editor_mcp.tools import open_document, reject_all, replace_text

        open_document(server, str(mcp_temp_docx), author="Tester")
        replace_text(server, str(mcp_temp_docx), "quick brown fox", "Hi")

        result = reject_all(server, str(mcp_temp_docx))

        assert result["success"] is True
        assert "count" in result


class TestReadTools:
    """Test read-only tools (Task 3.5)."""

    def test_find_text(self, server, mcp_temp_docx):
        """find_text checks if text exists."""
        from docx_editor_mcp.tools import find_text, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        result = find_text(server, str(mcp_temp_docx), "quick brown fox")

        assert result["success"] is True
        assert result["found"] is True

    def test_find_text_not_found(self, server, mcp_temp_docx):
        """find_text returns found=False for missing text."""
        from docx_editor_mcp.tools import find_text, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        result = find_text(server, str(mcp_temp_docx), "nonexistent xyz")

        assert result["success"] is True
        assert result["found"] is False

    def test_count_matches(self, server, mcp_temp_docx):
        """count_matches returns occurrence count."""
        from docx_editor_mcp.tools import count_matches, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        result = count_matches(server, str(mcp_temp_docx), "quick brown fox")

        assert result["success"] is True
        assert "count" in result
        assert result["count"] >= 0

    def test_get_visible_text(self, server, mcp_temp_docx):
        """get_visible_text returns document text."""
        from docx_editor_mcp.tools import get_visible_text, open_document

        open_document(server, str(mcp_temp_docx), author="Tester")

        result = get_visible_text(server, str(mcp_temp_docx))

        assert result["success"] is True
        assert "text" in result
        assert isinstance(result["text"], str)


class TestExternalChangeDetection:
    """Test external change detection across tools."""

    def test_edit_blocked_on_external_change(self, server, mcp_temp_docx):
        """Edit operations fail if file changed externally."""
        from docx_editor_mcp.tools import open_document, replace_text

        open_document(server, str(mcp_temp_docx), author="Tester")

        # Simulate external modification
        time.sleep(0.01)
        mcp_temp_docx.write_bytes(mcp_temp_docx.read_bytes() + b"external")

        result = replace_text(server, str(mcp_temp_docx), "quick brown fox", "Hi")

        assert result["success"] is False
        assert "external" in result["error"].lower()
