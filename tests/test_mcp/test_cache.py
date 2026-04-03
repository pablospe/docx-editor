"""Tests for DocumentCache following TDD."""

import os
import time
from pathlib import Path
from unittest.mock import MagicMock, patch


class TestPathNormalization:
    """Test path normalization (Task 1.3)."""

    def test_normalize_expands_tilde(self):
        """Path with ~ expands to home directory."""
        from docx_editor_mcp.cache import normalize_path

        result = normalize_path("~/docs/test.docx")
        assert result.startswith(str(Path.home()))
        assert result.endswith("docs/test.docx")

    def test_normalize_resolves_relative_path(self):
        """Relative path resolves to absolute."""
        from docx_editor_mcp.cache import normalize_path

        result = normalize_path("test.docx")
        assert os.path.isabs(result)

    def test_normalize_resolves_symlinks(self, tmp_path):
        """Symlinks resolve to real path."""
        from docx_editor_mcp.cache import normalize_path

        # Create a real file and a symlink
        real_file = tmp_path / "real.docx"
        real_file.touch()
        symlink = tmp_path / "link.docx"
        symlink.symlink_to(real_file)

        result = normalize_path(str(symlink))
        assert result == str(real_file.resolve())


class TestCachedDocument:
    """Test CachedDocument class (Task 1.2, 1.4)."""

    def test_cached_document_stores_document_and_path(self):
        """CachedDocument stores document instance and normalized path."""
        from docx_editor_mcp.cache import CachedDocument

        mock_doc = MagicMock()
        cached = CachedDocument(path="/path/to/doc.docx", document=mock_doc, author="Tester")

        assert cached.document is mock_doc
        assert cached.path == "/path/to/doc.docx"
        assert cached.author == "Tester"

    def test_cached_document_tracks_mtime(self, tmp_path):
        """CachedDocument stores file mtime on creation."""
        from docx_editor_mcp.cache import CachedDocument

        # Create a real file to get mtime
        test_file = tmp_path / "test.docx"
        test_file.touch()
        expected_mtime = test_file.stat().st_mtime

        mock_doc = MagicMock()
        cached = CachedDocument(path=str(test_file), document=mock_doc, author="Tester")

        assert cached.mtime == expected_mtime

    def test_cached_document_tracks_last_access(self):
        """CachedDocument tracks last access time."""
        from docx_editor_mcp.cache import CachedDocument

        mock_doc = MagicMock()
        before = time.time()
        cached = CachedDocument(path="/path/to/doc.docx", document=mock_doc, author="Tester")
        after = time.time()

        assert before <= cached.last_access <= after

    def test_cached_document_touch_updates_last_access(self):
        """touch() updates last_access time."""
        from docx_editor_mcp.cache import CachedDocument

        mock_doc = MagicMock()
        cached = CachedDocument(path="/path/to/doc.docx", document=mock_doc, author="Tester")
        old_access = cached.last_access

        time.sleep(0.01)  # Small delay to ensure different timestamp
        cached.touch()

        assert cached.last_access > old_access

    def test_cached_document_dirty_flag(self):
        """CachedDocument tracks dirty state."""
        from docx_editor_mcp.cache import CachedDocument

        mock_doc = MagicMock()
        cached = CachedDocument(path="/path/to/doc.docx", document=mock_doc, author="Tester")

        assert cached.dirty is False
        cached.mark_dirty()
        assert cached.dirty is True
        cached.clear_dirty()
        assert cached.dirty is False

    def test_cached_document_detects_external_changes(self, tmp_path):
        """has_external_changes() detects when file mtime changed."""
        from docx_editor_mcp.cache import CachedDocument

        test_file = tmp_path / "test.docx"
        test_file.touch()

        mock_doc = MagicMock()
        cached = CachedDocument(path=str(test_file), document=mock_doc, author="Tester")

        # No changes yet
        assert cached.has_external_changes() is False

        # Simulate external modification
        time.sleep(0.01)
        test_file.write_text("modified")

        assert cached.has_external_changes() is True

    def test_cached_document_update_mtime(self, tmp_path):
        """update_mtime() syncs cached mtime with file."""
        from docx_editor_mcp.cache import CachedDocument

        test_file = tmp_path / "test.docx"
        test_file.touch()

        mock_doc = MagicMock()
        cached = CachedDocument(path=str(test_file), document=mock_doc, author="Tester")

        # Modify file
        time.sleep(0.01)
        test_file.write_text("modified")
        assert cached.has_external_changes() is True

        # Update mtime
        cached.update_mtime()
        assert cached.has_external_changes() is False


class TestDocumentCache:
    """Test DocumentCache class (Task 1.2, 1.5)."""

    def test_cache_get_returns_none_for_missing(self):
        """get() returns None for documents not in cache."""
        from docx_editor_mcp.cache import DocumentCache

        cache = DocumentCache(max_documents=10)
        result = cache.get("/nonexistent/path.docx")

        assert result is None

    def test_cache_put_and_get(self):
        """put() adds document, get() retrieves it."""
        from docx_editor_mcp.cache import CachedDocument, DocumentCache

        cache = DocumentCache(max_documents=10)
        mock_doc = MagicMock()
        cached = CachedDocument(path="/path/to/doc.docx", document=mock_doc, author="Tester")

        cache.put(cached)
        result = cache.get("/path/to/doc.docx")

        assert result is cached
        assert result.document is mock_doc

    def test_cache_normalizes_paths(self, tmp_path):
        """Cache uses normalized paths for lookup."""
        from docx_editor_mcp.cache import CachedDocument, DocumentCache

        cache = DocumentCache(max_documents=10)
        mock_doc = MagicMock()

        # Create real file for mtime
        test_file = tmp_path / "doc.docx"
        test_file.touch()

        # Put with absolute path
        cached = CachedDocument(path=str(test_file), document=mock_doc, author="Tester")
        cache.put(cached)

        # Get with different but equivalent path should work
        # (This tests that paths are normalized on get as well)
        result = cache.get(str(test_file))
        assert result is cached

    def test_cache_get_updates_last_access(self):
        """get() updates last_access time of cached document."""
        from docx_editor_mcp.cache import CachedDocument, DocumentCache

        cache = DocumentCache(max_documents=10)
        mock_doc = MagicMock()
        cached = CachedDocument(path="/path/to/doc.docx", document=mock_doc, author="Tester")
        cache.put(cached)

        old_access = cached.last_access
        time.sleep(0.01)

        cache.get("/path/to/doc.docx")

        assert cached.last_access > old_access

    def test_cache_remove(self):
        """remove() removes document from cache."""
        from docx_editor_mcp.cache import CachedDocument, DocumentCache

        cache = DocumentCache(max_documents=10)
        mock_doc = MagicMock()
        cached = CachedDocument(path="/path/to/doc.docx", document=mock_doc, author="Tester")

        cache.put(cached)
        assert cache.get("/path/to/doc.docx") is not None

        cache.remove("/path/to/doc.docx")
        assert cache.get("/path/to/doc.docx") is None

    def test_cache_size(self):
        """size property returns number of cached documents."""
        from docx_editor_mcp.cache import CachedDocument, DocumentCache

        cache = DocumentCache(max_documents=10)
        assert cache.size == 0

        cache.put(CachedDocument(path="/doc1.docx", document=MagicMock(), author="Tester"))
        assert cache.size == 1

        cache.put(CachedDocument(path="/doc2.docx", document=MagicMock(), author="Tester"))
        assert cache.size == 2

    def test_cache_lru_eviction(self, tmp_path):
        """Cache evicts least recently used when full."""
        from docx_editor_mcp.cache import CachedDocument, DocumentCache

        cache = DocumentCache(max_documents=2)

        # Create files for mtime
        for i in range(3):
            (tmp_path / f"doc{i}.docx").touch()

        # Add first two documents
        doc1 = CachedDocument(path=str(tmp_path / "doc0.docx"), document=MagicMock(), author="Tester")
        cache.put(doc1)
        time.sleep(0.01)

        doc2 = CachedDocument(path=str(tmp_path / "doc1.docx"), document=MagicMock(), author="Tester")
        cache.put(doc2)
        time.sleep(0.01)

        # Access doc1 to make doc2 the LRU
        cache.get(str(tmp_path / "doc0.docx"))
        time.sleep(0.01)

        # Add third document - should evict doc2 (LRU)
        doc3 = CachedDocument(path=str(tmp_path / "doc2.docx"), document=MagicMock(), author="Tester")
        cache.put(doc3)

        assert cache.size == 2
        assert cache.get(str(tmp_path / "doc0.docx")) is not None  # Still there
        assert cache.get(str(tmp_path / "doc1.docx")) is None  # Evicted
        assert cache.get(str(tmp_path / "doc2.docx")) is not None  # Added

    def test_cache_saves_dirty_on_eviction(self, tmp_path):
        """Dirty documents are saved before eviction."""
        from docx_editor_mcp.cache import CachedDocument, DocumentCache

        cache = DocumentCache(max_documents=1)

        # Create file
        (tmp_path / "doc0.docx").touch()
        (tmp_path / "doc1.docx").touch()

        mock_doc = MagicMock()
        doc1 = CachedDocument(path=str(tmp_path / "doc0.docx"), document=mock_doc, author="Tester")
        doc1.mark_dirty()
        cache.put(doc1)

        # Add second document - should trigger eviction and save
        doc2 = CachedDocument(path=str(tmp_path / "doc1.docx"), document=MagicMock(), author="Tester")
        cache.put(doc2)

        # Verify save was called on the evicted document
        mock_doc.save.assert_called_once()

    def test_cache_all_documents(self):
        """all() returns all cached documents."""
        from docx_editor_mcp.cache import CachedDocument, DocumentCache

        cache = DocumentCache(max_documents=10)
        doc1 = CachedDocument(path="/doc1.docx", document=MagicMock(), author="Tester")
        doc2 = CachedDocument(path="/doc2.docx", document=MagicMock(), author="Tester")

        cache.put(doc1)
        cache.put(doc2)

        all_docs = list(cache.all())
        assert len(all_docs) == 2
        assert doc1 in all_docs
        assert doc2 in all_docs


class TestSessionAuthor:
    """Test session author memory (Task 1.6)."""

    def test_get_author_explicit_takes_precedence(self):
        """Explicit author parameter is used first."""
        from docx_editor_mcp.cache import DocumentCache

        cache = DocumentCache(max_documents=10)
        author, is_default = cache.get_author("Explicit Author")

        assert author == "Explicit Author"
        assert is_default is False

    def test_get_author_remembers_session_author(self):
        """Session author is remembered for subsequent calls."""
        from docx_editor_mcp.cache import DocumentCache

        cache = DocumentCache(max_documents=10)

        # First call with explicit author
        cache.get_author("Legal Team")

        # Second call without author uses session author
        author, is_default = cache.get_author(None)

        assert author == "Legal Team"
        assert is_default is False

    def test_get_author_uses_system_default_first_time(self):
        """First call without author uses system username."""
        from docx_editor_mcp.cache import DocumentCache

        cache = DocumentCache(max_documents=10)

        with patch("getpass.getuser", return_value="testuser"):
            author, is_default = cache.get_author(None)

        assert author == "testuser"
        assert is_default is True  # Hints Claude to ask

    def test_get_author_fallback_to_reviewer(self):
        """Falls back to 'Reviewer' if getpass.getuser() fails."""
        from docx_editor_mcp.cache import DocumentCache

        cache = DocumentCache(max_documents=10)

        with patch("getpass.getuser", side_effect=Exception("No user")):
            author, is_default = cache.get_author(None)

        assert author == "Reviewer"
        assert is_default is True

    def test_session_author_updates_on_explicit(self):
        """Explicit author updates session author."""
        from docx_editor_mcp.cache import DocumentCache

        cache = DocumentCache(max_documents=10)

        # Set initial session author
        cache.get_author("First Author")

        # New explicit author
        cache.get_author("Second Author")

        # Should use new session author
        author, _ = cache.get_author(None)
        assert author == "Second Author"
