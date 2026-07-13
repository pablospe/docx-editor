"""Workspace management for docx_editor.

Manages the workspace folder that holds a document's unpacked OOXML contents.
By default the workspace lives under the platform user cache directory (see
:func:`_default_cache_dir`); the location can be overridden with the
``DOCX_EDITOR_WORKSPACE_DIR`` environment variable or an explicit
``workspace_dir`` argument.
"""

import getpass
import hashlib
import json
import os
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .exceptions import (
    DocumentNotFoundError,
    DocumentOpenError,
    InvalidDocumentError,
    WorkspaceError,
    WorkspaceExistsError,
    WorkspaceSyncError,
)
from .ooxml import pack_document, unpack_document


def _cache_root_from_env(name: str) -> Path | None:
    """Return ``$name`` as a cache root, or None if it is unusable.

    Per the XDG base directory spec, a relative value "must be ignored and a
    default equivalent [...] used instead" — the same rule is applied to
    %LOCALAPPDATA%. Without it a relative value would later be joined to the
    source document's directory, silently scattering caches next to documents.
    """
    value = os.environ.get(name, "").strip()
    if not value:
        return None
    root = Path(os.path.expanduser(value))
    return root if root.is_absolute() else None


def _home() -> Path:
    """Return the user's home directory as a WorkspaceError-domain failure."""
    try:
        return Path.home()
    except RuntimeError as exc:  # no HOME and no passwd entry (slim containers)
        raise WorkspaceError(
            "Cannot determine the home directory for the default workspace cache. "
            "Set DOCX_EDITOR_WORKSPACE_DIR to an absolute path, or pass workspace_dir=."
        ) from exc


def _default_cache_dir() -> Path:
    """Return the platform-appropriate user cache base for workspaces.

    - Windows: ``%LOCALAPPDATA%\\docx-editor`` (``~/AppData/Local`` fallback)
    - macOS: ``~/Library/Caches/docx-editor``
    - Linux/other: ``$XDG_CACHE_HOME/docx-editor`` (``~/.cache`` fallback)

    A relative ``%LOCALAPPDATA%``/``$XDG_CACHE_HOME`` is ignored (see
    :func:`_cache_root_from_env`).
    """
    if os.name == "nt":
        root = _cache_root_from_env("LOCALAPPDATA") or _home() / "AppData" / "Local"
        return root / "docx-editor"
    if sys.platform == "darwin":
        return _home() / "Library" / "Caches" / "docx-editor"
    root = _cache_root_from_env("XDG_CACHE_HOME") or _home() / ".cache"
    return root / "docx-editor"


def owner_file_candidates(path: str | Path) -> list[Path]:
    """Return the ``~$`` owner (lock) file paths Word may use for ``path``.

    Word writes a hidden ``~$`` owner file next to any document it has open. Its
    name is derived from the document's filename: Word drops the first two
    characters when the stem is longer than two characters (``Report.docx`` →
    ``~$port.docx``), and keeps the filename in full for very short stems
    (``ab.docx`` → ``~$ab.docx``).

    Both forms are returned as a deliberately conservative superset — checking a
    stub that Word would not have written costs a false positive (recoverable via
    ``force=True``), while missing one risks saving over a live document. The
    truncated form is inherently ambiguous: ``01_intro.docx`` and
    ``02_intro.docx`` share the stub ``~$_intro.docx``. That ambiguity is Word's
    own — it uses the same owner file for both, and the stub records the *user*
    who holds the lock, not the document — so it cannot be disambiguated here.
    """
    path = Path(path)
    candidates = [path.parent / f"~${path.name}"]
    # Guard on the stem so short names don't yield junk like "~$.docx".
    if len(path.stem) > 2:
        candidates.append(path.parent / f"~${path.name[2:]}")
    return candidates


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
        # Keep the name the caller actually opened. If they opened a symlink, that is
        # the name Word was told to open — and therefore the name its ~$ owner file
        # sits beside. save() needs it to find that stub. Made absolute but NOT
        # resolved: resolving would collapse the symlink and lose the very name we are
        # keeping, while leaving it relative would break if the cwd moves before save().
        given = Path(source_path)
        self._given_path = given if given.is_absolute() else Path.cwd() / given
        self.source_path = Path(source_path).resolve()

        if not self.source_path.exists():
            raise DocumentNotFoundError(f"Document not found: {source_path}")

        if self.source_path.suffix.lower() != ".docx":
            raise InvalidDocumentError(f"Not a .docx file: {source_path}")

        # Determine workspace path
        self.workspace_path = self._resolve_workspace_path(self.source_path, workspace_dir)

        # Set author (default to system user)
        self._author = author or getpass.getuser()

        if create:
            if self.workspace_path.exists():
                # Check if it's stale or matches current document
                existing_meta = self._load_meta()
                if existing_meta:
                    self._check_provenance(existing_meta)
                    self.meta = existing_meta
                    # Same staleness predicate as save(), so a workspace can never
                    # open clean here and then fail sync_check() later at save time.
                    if not self.sync_check():
                        raise WorkspaceSyncError(
                            f"Document has changed since workspace was created. "
                            f"Delete {self.workspace_path} or use force_recreate=True"
                        )
                    # Workspace is valid, just load it
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
            self._check_provenance(loaded_meta)
            self.meta = loaded_meta

    def _check_provenance(self, meta: dict[str, Any]) -> None:
        """Refuse to adopt a workspace that was unpacked from a different document.

        The workspace directory is keyed by a hash of the source path, so a
        mismatch here means either a hash collision or a directory planted by
        something else. Adopting it would let a later save() pack that other
        document's XML over this one's source file.
        """
        recorded = meta.get("source_path")
        # Compare with the same folding the workspace key uses (see
        # _resolve_workspace_path), or two spellings that map to one workspace on
        # a case-insensitive filesystem would be rejected as different documents.
        if recorded is None or os.path.normcase(recorded) != os.path.normcase(str(self.source_path)):
            raise WorkspaceError(
                f"Workspace {self.workspace_path} belongs to a different document "
                f"({recorded!r}, expected {str(self.source_path)!r}). "
                f"Delete it or use force_recreate=True."
            )

    @classmethod
    def _resolve_workspace_path(cls, source_path: Path, workspace_dir: str | Path | None) -> Path:
        """Resolve the workspace directory for a source document.

        Precedence for the base directory:
          1. the explicit ``workspace_dir`` argument
          2. the ``DOCX_EDITOR_WORKSPACE_DIR`` environment variable
          3. the platform user cache (see :func:`_default_cache_dir`)

        Both overrides are tilde-expanded, and an empty/whitespace-only value
        counts as unset. A relative base resolves against ``source_path.parent``
        (so ``workspace_dir=".docx"`` reproduces the old next-to-file layout for
        debugging); an absolute base is used as-is.

        The per-document subdirectory is ``sha256(normcase(source_path))[:16]``,
        hashed via :func:`os.fsencode` so undecodable filename bytes (which
        ``str(Path)`` carries as surrogates) do not raise UnicodeEncodeError.
        ``normcase`` folds case on Windows, so one file cannot map to two
        workspaces there; it is a no-op on POSIX.

        Args:
            source_path: Resolved (absolute) path to the .docx file.
            workspace_dir: Explicit base override, or None.
        """
        base = None
        for override in (workspace_dir, os.environ.get("DOCX_EDITOR_WORKSPACE_DIR")):
            if override is None:
                continue
            # A str override (including the env var) is stripped — stray whitespace
            # there is config noise, not part of a real path. A Path is used
            # verbatim, so an exotic path really ending in a space still works.
            text = override.strip() if isinstance(override, str) else os.fspath(override)
            if not text:  # empty/whitespace override falls through to the next level
                continue
            base = Path(os.path.expanduser(text))
            break
        if base is None:
            base = _default_cache_dir()

        if not base.is_absolute():
            base = source_path.parent / base

        key = os.fsencode(os.path.normcase(source_path))
        return base / hashlib.sha256(key).hexdigest()[:16]

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
        # Create workspace directory. The workspace holds the document's
        # plaintext in a shared, predictably-named cache base, so restrict it to
        # the owner. mkdir(parents=True, mode=...) applies the mode to the leaf
        # only, so the base is created separately to keep it 0o700 too.
        try:
            self.workspace_path.parent.mkdir(parents=True, mode=0o700, exist_ok=True)
            self.workspace_path.mkdir(mode=0o700, exist_ok=True)
        except OSError as exc:
            raise WorkspaceError(
                f"Cannot create workspace at {self.workspace_path}: {exc}. "
                f"Set DOCX_EDITOR_WORKSPACE_DIR to a writable absolute path, "
                f"or pass workspace_dir=."
            ) from exc

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

    def _check_not_open(self, output_path: Path, *, saving_in_place: bool) -> None:
        """Raise DocumentOpenError if Word appears to have the destination open.

        Word writes its ``~$`` owner file beside the name it was told to open, which
        is not necessarily the name we are about to write. So check beside every name
        this destination is known by: the path given, the path it resolves to (which
        is where the archive actually lands), and — when saving in place — the
        possibly-symlinked path the caller originally opened.

        A destination that does not exist yet is never guarded: a stale stub beside a
        name with no document behind it has nothing to protect.
        """
        if not os.path.lexists(output_path):
            return

        names = [output_path, Path(os.path.realpath(output_path))]
        if saving_in_place:
            names.append(self._given_path)

        candidates = dict.fromkeys(c for name in names for c in owner_file_candidates(name))
        owner_file = next((c for c in candidates if os.path.lexists(c)), None)
        if owner_file is None:
            return

        raise DocumentOpenError(
            f"{output_path} appears to be open in Word (found owner file "
            f"{owner_file.name}). Close the document in Word, or pass force=True if "
            f"this is a stale lock from a crashed session. Note the stub name is "
            f"ambiguous: Word derives it by dropping the first two characters, so it "
            f"may belong to a different document in the same folder.",
            path=output_path,
            owner_file=owner_file,
        )

    def save(
        self,
        destination: str | Path | None = None,
        validate: bool = False,
        force: bool = False,
    ) -> Path:
        """Pack workspace back to a .docx file.

        Args:
            destination: Output path (defaults to original source path)
            validate: If True, validate with LibreOffice
            force: If True, skip save-time safety checks (the source-changed-on-disk
                check and the open-in-Word guard)

        Returns:
            Path to the saved document

        Raises:
            WorkspaceSyncError: If saving to the original path and the source
                document changed on disk since the workspace was created
            DocumentOpenError: If the destination appears open in Word (a ``~$``
                owner file exists) and ``force`` is False, or if the OS denies the
                final replace because another program holds the destination open.
            WorkspaceError: If packing fails
        """
        # Keep the caller's path for packing and for the return value; resolution is
        # only needed to decide whether we are about to overwrite our own source.
        output_path = Path(destination) if destination else self.source_path
        overwrites_source = self._is_source(output_path)

        # A missing source has no external edits to lose, so recreating it is safe
        # and must not be blocked — only an existing, changed source is protected.
        if not force and overwrites_source and self.source_path.exists() and not self.sync_check():
            raise WorkspaceSyncError(
                f"Source document changed on disk since the workspace was created: {self.source_path}. "
                f"Saving would overwrite those changes. Use save(force=True) to overwrite anyway, "
                f"or save(destination=...) to write elsewhere."
            )

        # Refuse to overwrite a document Word currently has open. Saving into
        # Word's live file races its own writes and can corrupt the document.
        # Guard the destination only — saving a copy to a fresh path while the
        # source is open is fine. force=True skips this (and any other save-time
        # safety check) for confirmed-stale locks from a crashed session.
        if not force:
            self._check_not_open(output_path, saving_in_place=destination is None)

        # Update metadata before saving
        self.meta["last_saved"] = datetime.now(timezone.utc).isoformat()
        self._save_meta()

        # Pack the document. pack_document() maps a PermissionError from the final
        # replace to DocumentOpenError itself; any other PermissionError (e.g. a
        # non-writable directory) is a genuine filesystem error and propagates
        # unchanged rather than being mislabeled as "the document is open".
        success = pack_document(self.workspace_path, output_path, validate=validate)

        if not success:
            raise WorkspaceError(f"Failed to pack document to {output_path}")

        # Update source_mtime/source_size if saving to the original location.
        # overwrites_source uses os.path.samefile (see _is_source), so a different
        # name for the same file — including through a symlink — still refreshes the
        # workspace's stat, or the next open() would report a stale workspace. Both
        # mtime and size are updated because sync_check() compares both.
        if overwrites_source:
            saved_stat = output_path.stat()
            self.meta["source_mtime"] = saved_stat.st_mtime
            self.meta["source_size"] = saved_stat.st_size
            self._save_meta()

        return output_path

    def _is_source(self, output_path: Path) -> bool:
        """True if output_path names the same file as the workspace source.

        samefile() where possible: on case-insensitive filesystems (macOS, Windows)
        and through symlinks/hardlinks, a plain path comparison misses the match and
        would skip the staleness check entirely.
        """
        if output_path.exists() and self.source_path.exists():
            return os.path.samefile(output_path, self.source_path)
        return output_path.resolve() == self.source_path

    def close(self, cleanup: bool = True) -> None:
        """Close the workspace.

        Args:
            cleanup: If True, delete the workspace folder
        """
        # Only this document's workspace is removed. The base directory is
        # shared (and may be a caller-supplied one via workspace_dir=/
        # DOCX_EDITOR_WORKSPACE_DIR), so it is left alone even when empty —
        # _create_workspace() recreates it on demand.
        if cleanup and self.workspace_path.exists():
            shutil.rmtree(self.workspace_path)

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
        source_path = Path(source_path).resolve()
        workspace_path = cls._resolve_workspace_path(source_path, workspace_dir)

        if not workspace_path.exists():
            return False

        # As in close(), the shared base directory is left in place.
        shutil.rmtree(workspace_path)
        return True
