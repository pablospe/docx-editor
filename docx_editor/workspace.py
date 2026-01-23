"""Workspace management for docx_editor.

Manages the .docx/ workspace folder that contains unpacked document contents.
"""

import getpass
import json
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


class Workspace:
    """Manages the .docx/ workspace folder for a document.

    The workspace is stored in a hidden .docx/ folder in the same directory
    as the source document, with a subfolder named after the document stem.

    Attributes:
        source_path: Path to the original .docx file
        workspace_path: Path to the workspace folder (.docx/{stem}/)
        meta: Dictionary containing workspace metadata
    """

    WORKSPACE_DIR = ".docx"
    META_FILE = "meta.json"

    meta: dict[str, Any]

    def __init__(
        self,
        source_path: str | Path,
        author: str | None = None,
        create: bool = True,
    ):
        """Initialize workspace for a document.

        Args:
            source_path: Path to the .docx file
            author: Author name for tracked changes (defaults to system user)
            create: If True, create workspace if it doesn't exist

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
        workspace_dir = self.source_path.parent / self.WORKSPACE_DIR
        self.workspace_path = workspace_dir / self.source_path.stem

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

            # Remove .docx parent if empty
            workspace_dir = self.workspace_path.parent
            if workspace_dir.exists() and not any(workspace_dir.iterdir()):
                workspace_dir.rmdir()

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
    def exists(cls, source_path: str | Path) -> bool:
        """Check if a workspace exists for a document.

        Args:
            source_path: Path to the .docx file

        Returns:
            True if workspace exists
        """
        source_path = Path(source_path).resolve()
        workspace_dir = source_path.parent / cls.WORKSPACE_DIR
        workspace_path = workspace_dir / source_path.stem
        return workspace_path.exists()

    @classmethod
    def delete(cls, source_path: str | Path) -> bool:
        """Delete workspace for a document if it exists.

        Args:
            source_path: Path to the .docx file

        Returns:
            True if workspace was deleted, False if it didn't exist
        """
        import shutil

        source_path = Path(source_path).resolve()
        workspace_dir = source_path.parent / cls.WORKSPACE_DIR
        workspace_path = workspace_dir / source_path.stem

        if not workspace_path.exists():
            return False

        shutil.rmtree(workspace_path)

        # Remove .docx parent if empty
        if workspace_dir.exists() and not any(workspace_dir.iterdir()):
            workspace_dir.rmdir()

        return True
