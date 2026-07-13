"""Workspace management for docx_editor.

Manages the workspace folder that holds a document's unpacked OOXML contents.
By default the workspace lives under the platform user cache directory (see
:func:`_default_cache_dir`); the location can be overridden with the
``DOCX_EDITOR_WORKSPACE_DIR`` environment variable or an explicit
``workspace_dir`` argument.
"""

import contextlib
import getpass
import hashlib
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .exceptions import (
    DocumentNotFoundError,
    InvalidDocumentError,
    WorkspaceError,
    WorkspaceExistsError,
    WorkspaceSyncError,
)
from .ooxml import pack_document, unpack_document


def _default_cache_dir() -> Path:
    """Return the platform-appropriate user cache base for workspaces.

    - Windows: ``%LOCALAPPDATA%\\docx-editor`` (``~/AppData/Local`` fallback)
    - macOS: ``~/Library/Caches/docx-editor``
    - Linux/other: ``$XDG_CACHE_HOME/docx-editor`` (``~/.cache`` fallback)
    """
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA")
        root = Path(base) if base else Path.home() / "AppData" / "Local"
        return root / "docx-editor"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Caches" / "docx-editor"
    xdg = os.environ.get("XDG_CACHE_HOME")
    root = Path(xdg) if xdg else Path.home() / ".cache"
    return root / "docx-editor"


class Workspace:
    """Manages the per-document workspace folder.

    The workspace holds the unpacked OOXML contents of a single document. It is
    stored in a subdirectory named ``sha256(source_path)[:16]`` under a base
    directory resolved (in order) from an explicit ``workspace_dir`` argument,
    the ``DOCX_EDITOR_WORKSPACE_DIR`` environment variable, or the platform user
    cache directory. Pass ``workspace_dir=".docx"`` to keep the workspace next
    to the source file for debugging.

    Attributes:
        source_path: Path to the original .docx file
        workspace_path: Path to this document's workspace folder
        meta: Dictionary containing workspace metadata
    """

    # If you rename META_FILE, also update EXCLUDED_PATHS in ooxml/pack.py —
    # otherwise the renamed file will be packed into the .docx and Word will
    # flag it as "unreadable content" (issue #8).
    META_FILE = "meta.json"

    meta: dict[str, Any]

    def __init__(
        self,
        source_path: str | Path,
        author: str | None = None,
        create: bool = True,
        workspace_dir: str | Path | None = None,
    ):
        """Initialize workspace for a document.

        Args:
            source_path: Path to the .docx file
            author: Author name for tracked changes (defaults to system user)
            create: If True, create workspace if it doesn't exist
            workspace_dir: Base directory for the workspace. Overrides the
                DOCX_EDITOR_WORKSPACE_DIR environment variable and the platform
                cache default. A relative path resolves against the source
                document's directory (e.g. ".docx" keeps it next to the file).

        Raises:
            DocumentNotFoundError: If the source document doesn't exist
            InvalidDocumentError: If the file is not a .docx file
            WorkspaceExistsError: If workspace exists and create=True
        """
        self.source_path = Path(source_path).resolve()

        if not self.source_path.exists():
            raise DocumentNotFoundError(f"Document not found: {source_path}")

        if self.source_path.suffix.lower() != ".docx":
            raise InvalidDocumentError(f"Not a .docx file: {source_path}")

        # Determine workspace path
        self.workspace_dir = workspace_dir
        self.workspace_path = self._resolve_workspace_path(self.source_path, workspace_dir)

        # Set author (default to system user)
        self._author = author or getpass.getuser()

        if create:
            if self.workspace_path.exists():
                # Check if it's stale or matches current document
                existing_meta = self._load_meta()
                if existing_meta:
                    source_mtime = self.source_path.stat().st_mtime
                    if existing_meta.get("source_mtime") != source_mtime:
                        raise WorkspaceSyncError(
                            f"Document has changed since workspace was created. "
                            f"Delete {self.workspace_path} or use force_recreate=True"
                        )
                    # Workspace is valid, just load it
                    self.meta = existing_meta
                    return
                else:
                    raise WorkspaceExistsError(f"Workspace already exists: {self.workspace_path}")

            self._create_workspace()
        else:
            # Load existing workspace
            if not self.workspace_path.exists():
                raise WorkspaceError(f"Workspace not found: {self.workspace_path}")
            loaded_meta = self._load_meta()
            if not loaded_meta:
                raise WorkspaceError(f"Invalid workspace (no meta.json): {self.workspace_path}")
            self.meta = loaded_meta

    @classmethod
    def _resolve_workspace_path(cls, source_path: Path, workspace_dir: str | Path | None) -> Path:
        """Resolve the workspace directory for a source document.

        Precedence for the base directory:
          1. the explicit ``workspace_dir`` argument
          2. the ``DOCX_EDITOR_WORKSPACE_DIR`` environment variable
          3. the platform user cache (see :func:`_default_cache_dir`)

        A relative base resolves against ``source_path.parent`` (so
        ``workspace_dir=".docx"`` reproduces the old next-to-file layout for
        debugging); an absolute base is used as-is. The per-document
        subdirectory is always ``sha256(str(source_path))[:16]``.

        Args:
            source_path: Resolved (absolute) path to the .docx file.
            workspace_dir: Explicit base override, or None.
        """
        if workspace_dir is not None:
            base = Path(workspace_dir)
        else:
            env = os.environ.get("DOCX_EDITOR_WORKSPACE_DIR")
            base = Path(env) if env else _default_cache_dir()

        if not base.is_absolute():
            base = source_path.parent / base

        subdir = hashlib.sha256(str(source_path).encode()).hexdigest()[:16]
        return base / subdir

    @property
    def author(self) -> str:
        """Get the author name for tracked changes."""
        return self._author

    @property
    def rsid(self) -> str:
        """Get the RSID for this editing session."""
        return str(self.meta.get("rsid", ""))

    @property
    def initials(self) -> str:
        """Get author initials."""
        return str(self.meta.get("initials", self._author[0].upper() if self._author else ""))

    @property
    def word_path(self) -> Path:
        """Get the path to the word/ subfolder."""
        return self.workspace_path / "word"

    @property
    def document_xml_path(self) -> Path:
        """Get the path to word/document.xml."""
        return self.word_path / "document.xml"

    def _create_workspace(self) -> None:
        """Create the workspace by unpacking the document."""
        # Create workspace directory
        self.workspace_path.mkdir(parents=True, exist_ok=True)

        # Unpack document
        rsid = unpack_document(self.source_path, self.workspace_path)

        # Get source file info
        source_stat = self.source_path.stat()

        # Create metadata
        self.meta = {
            "source_path": str(self.source_path),
            "source_mtime": source_stat.st_mtime,
            "source_size": source_stat.st_size,
            "created_at": datetime.now(timezone.utc).isoformat(),
            "author": self._author,
            "initials": self._author[0].upper() if self._author else "",
            "rsid": rsid,
            "next_comment_id": 0,
            "next_change_id": 0,
        }

        self._save_meta()

    def _load_meta(self) -> dict[str, Any] | None:
        """Load metadata from meta.json."""
        meta_path = self.workspace_path / self.META_FILE
        if not meta_path.exists():
            return None
        try:
            with open(meta_path, encoding="utf-8") as f:
                result: dict[str, Any] = json.load(f)
                return result
        except (json.JSONDecodeError, OSError):
            return None

    def _save_meta(self) -> None:
        """Save metadata to meta.json."""
        meta_path = self.workspace_path / self.META_FILE
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(self.meta, f, indent=2)

    def save(self, destination: str | Path | None = None, validate: bool = False) -> Path:
        """Pack workspace back to a .docx file.

        Args:
            destination: Output path (defaults to original source path)
            validate: If True, validate with LibreOffice

        Returns:
            Path to the saved document

        Raises:
            WorkspaceError: If packing fails
        """
        output_path = Path(destination) if destination else self.source_path

        # Update metadata before saving
        self.meta["last_saved"] = datetime.now(timezone.utc).isoformat()
        self._save_meta()

        # Pack the document
        success = pack_document(self.workspace_path, output_path, validate=validate)

        if not success:
            raise WorkspaceError(f"Failed to pack document to {output_path}")

        # Update source_mtime if saving to original location
        if output_path == self.source_path:
            self.meta["source_mtime"] = output_path.stat().st_mtime
            self._save_meta()

        return output_path

    def close(self, cleanup: bool = True) -> None:
        """Close the workspace.

        Args:
            cleanup: If True, delete the workspace folder
        """
        if cleanup and self.workspace_path.exists():
            import shutil

            shutil.rmtree(self.workspace_path)

            # Remove the base directory if it is now empty. The base is shared
            # across documents/processes, so the check-then-rmdir can race —
            # suppress the resulting OSError (a lingering empty cache dir is fine).
            base_dir = self.workspace_path.parent
            with contextlib.suppress(OSError):
                if base_dir.exists() and not any(base_dir.iterdir()):
                    base_dir.rmdir()

    def get_xml_path(self, relative_path: str) -> Path:
        """Get the full path to an XML file in the workspace.

        Args:
            relative_path: Path relative to workspace root (e.g., "word/document.xml")

        Returns:
            Full path to the XML file
        """
        return self.workspace_path / relative_path

    def sync_check(self) -> bool:
        """Check if the workspace is in sync with the source document.

        Returns:
            True if in sync, False if source has changed
        """
        if not self.source_path.exists():
            return False

        source_stat = self.source_path.stat()
        return (
            self.meta.get("source_mtime") == source_stat.st_mtime
            and self.meta.get("source_size") == source_stat.st_size
        )

    @classmethod
    def exists(cls, source_path: str | Path, workspace_dir: str | Path | None = None) -> bool:
        """Check if a workspace exists for a document.

        Args:
            source_path: Path to the .docx file
            workspace_dir: Base directory override (see __init__)

        Returns:
            True if workspace exists
        """
        source_path = Path(source_path).resolve()
        workspace_path = cls._resolve_workspace_path(source_path, workspace_dir)
        return workspace_path.exists()

    @classmethod
    def delete(cls, source_path: str | Path, workspace_dir: str | Path | None = None) -> bool:
        """Delete workspace for a document if it exists.

        Args:
            source_path: Path to the .docx file
            workspace_dir: Base directory override (see __init__)

        Returns:
            True if workspace was deleted, False if it didn't exist
        """
        import shutil

        source_path = Path(source_path).resolve()
        workspace_path = cls._resolve_workspace_path(source_path, workspace_dir)

        if not workspace_path.exists():
            return False

        shutil.rmtree(workspace_path)

        # Remove the base directory if it is now empty (see close() for the
        # shared-base race that OSError suppression guards against).
        base_dir = workspace_path.parent
        with contextlib.suppress(OSError):
            if base_dir.exists() and not any(base_dir.iterdir()):
                base_dir.rmdir()

        return True
