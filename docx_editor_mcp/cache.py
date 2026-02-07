"""Document cache for MCP server with LRU eviction and external change detection."""

import getpass
import logging
import os
import time
from dataclasses import dataclass, field
from typing import Any

logger = logging.getLogger(__name__)


def normalize_path(path: str) -> str:
    """Normalize path to absolute canonical form.

    Expands ~, resolves relative paths, and follows symlinks.
    """
    return os.path.realpath(os.path.expanduser(path))


@dataclass
class CachedDocument:
    """A cached document with metadata for cache management."""

    path: str
    document: Any  # docx_editor.Document
    author: str
    mtime: float = field(default=0.0)
    last_access: float = field(default_factory=time.time)
    dirty: bool = field(default=False)

    def __post_init__(self):
        """Normalize path and initialize mtime from file if it exists."""
        # Normalize path to ensure consistency with cache key
        self.path = normalize_path(self.path)
        if os.path.exists(self.path):
            self.mtime = os.path.getmtime(self.path)

    def touch(self) -> None:
        """Update last access time."""
        self.last_access = time.time()

    def mark_dirty(self) -> None:
        """Mark document as having unsaved changes."""
        self.dirty = True

    def clear_dirty(self) -> None:
        """Clear dirty flag after save."""
        self.dirty = False

    def has_external_changes(self) -> bool:
        """Check if file was modified externally since cached."""
        if not os.path.exists(self.path):
            return False
        current_mtime = os.path.getmtime(self.path)
        return current_mtime != self.mtime

    def update_mtime(self) -> None:
        """Update cached mtime to match current file mtime."""
        if os.path.exists(self.path):
            self.mtime = os.path.getmtime(self.path)


class DocumentCache:
    """LRU cache for Document instances with session author memory."""

    def __init__(self, max_documents: int = 10):
        """Initialize cache with maximum document count.

        Args:
            max_documents: Maximum number of documents to keep in cache.
        """
        self.max_documents = max_documents
        self._cache: dict[str, CachedDocument] = {}
        self._session_author: str | None = None

    @property
    def size(self) -> int:
        """Number of documents currently in cache."""
        return len(self._cache)

    def get(self, path: str) -> CachedDocument | None:
        """Get a cached document by path.

        Updates last_access time on hit.

        Args:
            path: Path to the document (will be normalized).

        Returns:
            CachedDocument if found, None otherwise.
        """
        normalized = normalize_path(path)
        cached = self._cache.get(normalized)
        if cached:
            cached.touch()
        return cached

    def put(self, cached_doc: CachedDocument) -> None:
        """Add a document to the cache.

        Evicts LRU document if cache is full.

        Args:
            cached_doc: CachedDocument to add.
        """
        normalized = normalize_path(cached_doc.path)

        # Evict if at capacity and this is a new document
        if normalized not in self._cache and self.size >= self.max_documents:
            self._evict_lru()

        self._cache[normalized] = cached_doc

    def remove(self, path: str) -> CachedDocument | None:
        """Remove a document from the cache.

        Args:
            path: Path to the document (will be normalized).

        Returns:
            The removed CachedDocument, or None if not found.
        """
        normalized = normalize_path(path)
        return self._cache.pop(normalized, None)

    def all(self):
        """Iterate over all cached documents."""
        yield from self._cache.values()

    def _evict_lru(self) -> None:
        """Evict the least recently used document.

        If the document is dirty, saves it first. If save fails,
        the document is NOT evicted to prevent data loss.
        """
        if not self._cache:
            return

        # Find LRU document
        lru_path = min(self._cache, key=lambda p: self._cache[p].last_access)
        lru_doc = self._cache[lru_path]

        # Save if dirty - don't evict if save fails
        if lru_doc.dirty:
            try:
                lru_doc.document.save()
            except Exception:
                logger.exception("Failed to save during eviction: %s", lru_path)
                return  # Don't evict if save failed

        # Remove from cache
        del self._cache[lru_path]

    def get_author(self, explicit_author: str | None) -> tuple[str, bool]:
        """Get author name with session memory.

        Resolution order:
        1. Explicit author parameter
        2. Previously set session author
        3. System username via getpass.getuser()
        4. "Reviewer" as fallback

        Args:
            explicit_author: Explicitly provided author, or None.

        Returns:
            Tuple of (author_name, is_default) where is_default hints
            that Claude should ask the user.
        """
        if explicit_author:
            self._session_author = explicit_author
            return explicit_author, False

        if self._session_author:
            return self._session_author, False

        # First time - use system default
        try:
            default = getpass.getuser()
        except Exception:
            default = "Reviewer"

        self._session_author = default
        return default, True
