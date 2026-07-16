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
import secrets
import shutil
import sys
import weakref
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .exceptions import (
    DocumentNotFoundError,
    DocumentOpenError,
    InvalidDocumentError,
    WorkspaceError,
    WorkspaceExistsError,
    WorkspaceLockedError,
    WorkspaceSyncError,
)
from .ooxml import pack_document, unpack_document

# One retry after reclaiming a stale lock; losing the O_EXCL race twice means a
# live competitor, not another stale file (same pattern as _TEMP_NAME_ATTEMPTS
# in ooxml/pack.py).
_LOCK_ACQUIRE_ATTEMPTS = 2


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


def _file_sha256(path: Path) -> str:
    """Streamed sha256 hex digest of a file's content."""
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _pid_alive(pid: int, *, reap: bool = False) -> bool:
    """True if the process is still running.

    Decides whether a workspace lock (or a session pid file — see session.py)
    refers to a live process or a stale leftover. On any ambiguous probe the
    answer errs toward "alive": a false positive refuses adoption, which
    force_recreate can override; a false negative would reclaim a live
    session's lock.

    On Windows, ``os.kill(pid, sig)`` is not a probe — any signal other than
    the CTRL events calls TerminateProcess and would kill the process being
    checked — so the process is queried via the win32 API instead.

    Args:
        pid: Process ID to probe.
        reap: On POSIX, reap the pid first if it is this process's own exited
            child — otherwise a zombie keeps ``os.kill(pid, 0)`` succeeding
            forever. Only pass True when the caller owns that child (as
            session.py does for its kernel): ``os.waitpid`` on an arbitrary
            pid would steal the exit status of another component's child.
            Without reaping, a zombie child conservatively reads as alive.
    """
    if os.name == "nt":  # pragma: no cover - CI is POSIX
        import ctypes

        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        STILL_ACTIVE = 259
        ERROR_ACCESS_DENIED = 5

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)  # type: ignore[attr-defined]
        # Explicit signatures: ctypes defaults every restype to c_int, which
        # truncates a pointer-sized HANDLE on 64-bit Windows.
        kernel32.OpenProcess.restype = ctypes.c_void_p
        kernel32.OpenProcess.argtypes = (ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong)
        kernel32.GetExitCodeProcess.restype = ctypes.c_int
        kernel32.GetExitCodeProcess.argtypes = (ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong))
        kernel32.CloseHandle.restype = ctypes.c_int
        kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
        handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not handle:
            # Access denied: the process exists but belongs to someone else.
            # Any other failure (e.g. invalid parameter): no such process.
            return ctypes.get_last_error() == ERROR_ACCESS_DENIED  # type: ignore[attr-defined]
        try:
            exit_code = ctypes.c_ulong()
            if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                return True  # query failed: assume alive
            return exit_code.value == STILL_ACTIVE
        finally:
            kernel32.CloseHandle(handle)

    if reap:
        try:
            if os.waitpid(pid, os.WNOHANG)[0] == pid:
                return False
        except (ChildProcessError, OSError):
            pass  # Not our child — normal for a pid from another process.

    try:
        os.kill(pid, 0)
    except PermissionError:
        return True  # exists, just owned by another user
    except OSError:
        return False
    return True


def _release_lock_file(lock_path: Path, token: str) -> None:
    """Unlink ``lock_path`` if it still holds ``token`` (ownership-aware).

    Module-level so weakref.finalize can call it without keeping the
    Workspace alive: a session dropped without close() frees its own lock at
    garbage collection (or interpreter exit) instead of locking its process
    out of the document until restart. A SIGKILLed process still leaves the
    file behind; the liveness probe in _acquire_lock reclaims that case.
    """
    try:
        if lock_path.read_text(encoding="utf-8") == token:
            lock_path.unlink(missing_ok=True)
    except OSError:
        pass


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

    # If you rename META_FILE, also update EXCLUDED_PATHS in ooxml/pack.py
    # (both the file and its _save_meta ".tmp" twin) — otherwise the renamed
    # file will be packed into the .docx and Word will flag it as "unreadable
    # content" (issue #8).
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
            InvalidDocumentError: If the path is not a valid .docx document:
                wrong suffix, a directory, not a zip archive, a part contains
                malformed XML, or the required word/document.xml part is
                missing
            WorkspaceLockedError: If a live session already holds this
                document's workspace (see _acquire_lock). Checked first: a
                live holder masks the sync/exists errors below until it
                closes.
            WorkspaceExistsError: If workspace exists and create=True
            WorkspaceSyncError: If create=True and an existing workspace holds
                unsaved changes from a previous session, or the source document
                changed on disk since the workspace was created
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

        # One live session per workspace: acquire the advisory lock before the
        # adopt/create branch, so a concurrent open conflicts here instead of
        # silently sharing the workspace and racing last-save-wins. The lock is
        # a sidecar of the workspace dir — it cannot live inside it, because
        # the adopt-vs-create branch below keys on the dir's existence (and a
        # workspace-root file would need a pack exclusion). Both create modes
        # lock: the create=False rescue flow reads the workspace and writes
        # meta, so it must conflict with a live session too. Acquiring before
        # the dirty/staleness checks means a live holder masks those errors —
        # the correct priority: it is actively using the workspace.
        self._lock_path = self.workspace_path.with_name(self.workspace_path.name + ".lock")
        # pid + random token: the pid feeds the liveness probe; the token
        # identifies this specific instance so release stays ownership-aware
        # even when a force_recreate takeover happens within one process.
        self._lock_token = f"{os.getpid()}:{secrets.token_hex(8)}"
        self._lock_acquired = False
        self._lock_finalizer: weakref.finalize | None = None
        try:
            # The shared cache base must exist to hold the sidecar (same mkdir
            # _create_workspace performs; owner-only, see there).
            self.workspace_path.parent.mkdir(parents=True, mode=0o700, exist_ok=True)
            self._acquire_lock()
        except OSError as exc:
            # Lock *contention* raises WorkspaceLockedError (not OSError) and
            # propagates past this clause; only filesystem failures land here.
            raise WorkspaceError(
                f"Cannot create workspace at {self.workspace_path}: {exc}. "
                f"Set DOCX_EDITOR_WORKSPACE_DIR to a writable absolute path, "
                f"or pass workspace_dir=."
            ) from exc

        try:
            if create:
                if self.workspace_path.exists():
                    # Check if it's stale or matches current document
                    existing_meta = self._load_meta()
                    if existing_meta:
                        self._check_provenance(existing_meta)
                        self.meta = existing_meta
                        # Checked before staleness: when both hold, the unsaved edits
                        # are the data-loss-relevant fact to surface, not the mtime.
                        if existing_meta.get("dirty"):
                            raise WorkspaceSyncError(
                                f"Workspace {self.workspace_path} holds unsaved changes from a "
                                f"previous session. Adopting it would carry those edits into "
                                f"this session and the next save would write them into "
                                f"{self.source_path}. Use force_recreate=True (or delete the "
                                f"workspace) to discard them, or "
                                f"Workspace(source, create=False).save(destination=...) to "
                                f"rescue them first."
                            )
                        # Same staleness predicate as save(), so a workspace can never
                        # open clean here and then fail sync_check() later at save time.
                        if not self.sync_check():
                            raise WorkspaceSyncError(
                                f"Document has changed since workspace was created. "
                                f"Delete {self.workspace_path} or use force_recreate=True"
                            )
                        if self.document_xml_path.exists():
                            # Workspace is valid, just load it
                            return
                        # Clean and in-sync but missing the core part: an
                        # orphan left by a pre-0.6.1 failed open (ISSUES.md
                        # #41) or a manually mangled workspace. The checks
                        # above proved the source still matches the recorded
                        # fingerprint and no edits were flagged, so this cache
                        # can be discarded and rebuilt without losing anything
                        # the system tracks. A valid source heals; an invalid one
                        # raises InvalidDocumentError from unpack_document
                        # (single message site).
                        shutil.rmtree(self.workspace_path)
                        self._create_workspace()
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
        except BaseException:
            # A failed __init__ hands the caller no object to close(); without
            # this release the process would lock itself out of its own
            # retry/rescue (e.g. reopening after a WorkspaceSyncError).
            self._release_lock()
            raise

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

    def _read_lock_content(self) -> str | None:
        """Raw lock-file content, or None if the file is absent or unreadable."""
        try:
            return self._lock_path.read_text(encoding="utf-8")
        except OSError:
            return None

    @staticmethod
    def _parse_lock_pid(content: str | None) -> int | None:
        """PID from ``<pid>:<token>`` lock content (bare ``<pid>`` also parses).

        Non-positive values count as corrupt: probing them would be dangerous
        (``os.waitpid(-1)`` reaps an arbitrary child; ``os.kill(0/-N, 0)``
        signals whole process groups), so they read as "no holder" and the
        lock is reclaimed.
        """
        if content is None:
            return None
        try:
            pid = int(content.split(":", 1)[0])
        except ValueError:
            return None
        return pid if pid > 0 else None

    def _read_lock_pid(self) -> int | None:
        """PID recorded in the lock file, or None if absent/unreadable/corrupt."""
        return self._parse_lock_pid(self._read_lock_content())

    def _acquire_lock(self) -> None:
        """Create the sidecar lock file naming this session, exclusively.

        The file holds ``<pid>:<random token>``: the pid drives the liveness
        probe, the token makes release ownership-aware (see _release_lock).
        A lock naming a live process raises WorkspaceLockedError; a dead or
        unreadable one is stale and reclaimed. This is advisory protection
        against *accidental* concurrent opens, not a hardened mutex: the
        lock's content is re-checked right before a stale file is unlinked,
        which narrows — but cannot close — the window in which a racing
        process's fresh lock is removed instead. A lost race degrades to the
        pre-lock behavior (two sessions sharing a workspace); it cannot
        corrupt data beyond that.

        Raises:
            WorkspaceLockedError: A live process holds the lock, or the
                reclaim attempts lost the creation race.
        """
        for _ in range(_LOCK_ACQUIRE_ATTEMPTS):
            try:
                fd = os.open(self._lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
            except FileExistsError:
                stale_content = self._read_lock_content()
                holder = self._parse_lock_pid(stale_content)
                if holder is not None and _pid_alive(holder):
                    if holder == os.getpid():
                        hint = "this process already has the document open — close() the other Document/Workspace first"
                    else:
                        hint = (
                            f"close the session in process {holder} first, or use "
                            f"force_recreate=True to take the workspace over and "
                            f"discard its unsaved edits"
                        )
                    raise WorkspaceLockedError(
                        f"Workspace {self.workspace_path} is locked by a live session "
                        f"(pid {holder}, lock file {self._lock_path}): {hint}.",
                        pid=holder,
                        lock_path=self._lock_path,
                    ) from None
                # Dead process or unreadable content: stale — reclaim and
                # retry, unless another process already replaced the lock
                # since it was read (see docstring: narrowed, not airtight).
                if self._read_lock_content() == stale_content:
                    self._lock_path.unlink(missing_ok=True)
                continue
            try:
                with os.fdopen(fd, "w", encoding="utf-8") as f:
                    f.write(self._lock_token)
                # Safety net for sessions dropped without close(): release
                # the lock at garbage collection, or this process could never
                # reopen (nor rescue) the document — the lock names its own
                # live pid, so the staleness probe would never reclaim it.
                self._lock_finalizer = weakref.finalize(self, _release_lock_file, self._lock_path, self._lock_token)
            except BaseException:
                # The O_EXCL create succeeded but the token write or the
                # finalizer registration failed (disk full, MemoryError):
                # remove the file rather than orphan a lock naming this live
                # pid, which the staleness probe could never reclaim.
                self._lock_path.unlink(missing_ok=True)
                raise
            self._lock_acquired = True
            return

        # Both attempts found a fresh lock after reclaiming a stale one:
        # another process won the re-creation race.
        raise WorkspaceLockedError(
            f"Workspace {self.workspace_path} was locked by another session while "
            f"reclaiming a stale lock ({self._lock_path}). Retry, or close the "
            f"competing session.",
            pid=self._read_lock_pid(),
            lock_path=self._lock_path,
        )

    def _release_lock(self) -> None:
        """Remove the sidecar lock file if this instance still owns it.

        Ownership is verified by token, not just the _lock_acquired flag:
        after a force_recreate takeover (Workspace.delete removes even a live
        session's lock), the superseded session's close() must not unlink the
        new session's lock — and the pid alone cannot tell them apart when
        both live in one process. Release is independent of workspace cleanup:
        close(cleanup=False) keeps the workspace but must still free it for
        the next session. A session dropped without close() is covered by the
        weakref finalizer registered at acquire time; a killed process leaves
        the file behind by design, and the staleness probe in _acquire_lock
        reclaims it.
        """
        if self._lock_acquired:
            if self._lock_finalizer is not None:
                self._lock_finalizer()  # ownership-checked unlink; runs once
            else:  # pragma: no cover - acquire always registers the finalizer
                _release_lock_file(self._lock_path, self._lock_token)
            self._lock_acquired = False

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

        try:
            # Unpack document
            rsid = unpack_document(self.source_path, self.workspace_path)

            # Get source file info
            source_stat = self.source_path.stat()

            # Create metadata
            self.meta = {
                "source_path": str(self.source_path),
                "source_mtime": source_stat.st_mtime,
                "source_size": source_stat.st_size,
                "source_sha256": _file_sha256(self.source_path),
                "created_at": datetime.now(timezone.utc).isoformat(),
                "author": self._author,
                "initials": self._author[0].upper() if self._author else "",
                "rsid": rsid,
                "next_comment_id": 0,
                "next_change_id": 0,
                "dirty": False,
            }

            self._save_meta()
        except BaseException:
            # A partial workspace (no meta.json) would make the next open fail
            # with a misleading WorkspaceExistsError. This method is only
            # reached when the dir did not pre-exist, so nothing but our own
            # partial output is removed; the advisory lock is a sibling file
            # released by __init__'s handler, so rmtree composes cleanly.
            shutil.rmtree(self.workspace_path, ignore_errors=True)
            raise

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
        """Save metadata to meta.json atomically.

        Written to a temp file in the same directory, fsynced, then renamed
        over meta.json, so no crash can leave a truncated meta.json — a
        truncated file destroys the write-ahead dirty flag and makes the next
        open fail with a misleading WorkspaceExistsError (issue #22). The
        fixed temp name assumes one writer per workspace (one live session);
        EXCLUDED_PATHS in ooxml/pack.py keeps a crash-orphaned temp out of
        the packed document.
        """
        meta_path = self.workspace_path / self.META_FILE
        tmp_path = meta_path.with_name(meta_path.name + ".tmp")
        try:
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(self.meta, f, indent=2)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_path, meta_path)
        finally:
            # A failed write leaves only the intact old meta.json behind.
            tmp_path.unlink(missing_ok=True)

    def mark_dirty(self) -> None:
        """Persist that the workspace may hold content not saved to the source.

        Call this *before* mutating workspace content on disk (write-ahead), so
        that even a crash mid-mutation leaves the flag set. A dirty workspace is
        refused for adoption by a later ``Workspace(create=True)`` — otherwise a
        previous session's edits could silently carry into the new session (the
        staleness check cannot catch this case: the source itself may be
        unchanged). The flag is cleared only by a successful save() back to the
        source.

        Document.save() and Workspace.save() both call this themselves, and
        Document wires it as the write-ahead hook of every editor it
        constructs (XMLEditor's ``on_save``, CommentManager's ``on_write``),
        so editor-mediated flushes honor the contract mechanically. The
        exception is open-time tracking setup, which deliberately bypasses the
        flag — see Document._setup_tracking(). Code that writes workspace
        files directly (no editor in between) must still call this before
        touching disk, or a crash before save() leaves the divergence
        unflagged.
        """
        if not self.meta.get("dirty"):
            self.meta["dirty"] = True
            self._save_meta()

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

        The workspace is flagged as holding unsaved changes (see
        :meth:`mark_dirty`) before any check or write runs; only a successful
        save back to the source clears the flag. A failed save, or a save to a
        different destination, leaves the workspace refused for adoption by a
        later ``Workspace(create=True)``.

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

        # Write-ahead: flag the workspace before any check or write can fail,
        # so a crash from here on leaves the flag on disk. Not skipped by
        # force= — this is bookkeeping, not a safety check. A save elsewhere
        # leaves the flag set (the source never received that content); only
        # a successful save back to the source clears it, below.
        self.mark_dirty()

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

        # Update the recorded source fingerprint if saving to the original
        # location. overwrites_source uses os.path.samefile (see _is_source), so
        # a different name for the same file — including through a symlink —
        # still refreshes it, or the next open() would report a stale workspace.
        # mtime/size are refreshed alongside the hash: sync_check() needs size
        # for its cheap reject, and legacy library versions reading this meta
        # compare mtime+size only.
        if overwrites_source:
            saved_stat = output_path.stat()
            self.meta["source_mtime"] = saved_stat.st_mtime
            self.meta["source_size"] = saved_stat.st_size
            self.meta["source_sha256"] = _file_sha256(output_path)
            # The source now holds everything the workspace holds — the
            # workspace is no longer ahead of it, so adoption is safe again.
            self.meta["dirty"] = False
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
        """Close the workspace and release its advisory lock.

        The lock is released in both cleanup modes — closing is what frees
        the document for the next session (see WorkspaceLockedError).

        Args:
            cleanup: If True, delete the workspace folder
        """
        # Only this document's workspace is removed. The base directory is
        # shared (and may be a caller-supplied one via workspace_dir=/
        # DOCX_EDITOR_WORKSPACE_DIR), so it is left alone even when empty —
        # _create_workspace() recreates it on demand.
        try:
            if cleanup and self.workspace_path.exists():
                shutil.rmtree(self.workspace_path)
        finally:
            # Released in both cleanup modes, and even if rmtree fails: a
            # kept (or half-removed) workspace must still be adoptable.
            self._release_lock()

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

        Metas that record ``source_sha256`` compare content: size first (a
        cheap reject), then the hash. A touch/copy that rewrites identical
        bytes no longer counts as an external edit, and a same-size content
        swap — which mtime+size provably cannot catch — now does. Legacy metas
        without the key keep the original mtime+size comparison.

        Returns:
            True if in sync, False if source has changed
        """
        if not self.source_path.exists():
            return False

        source_stat = self.source_path.stat()
        recorded_sha256 = self.meta.get("source_sha256")
        if recorded_sha256 is not None:
            if self.meta.get("source_size") != source_stat.st_size:
                return False
            return _file_sha256(self.source_path) == recorded_sha256
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

        removed = False
        if workspace_path.exists():
            # As in close(), the shared base directory is left in place.
            shutil.rmtree(workspace_path)
            removed = True

        # The advisory lock sidecar goes with the workspace — even one held by
        # a live (possibly stuck) session — keeping force_recreate the
        # universal escape hatch.
        workspace_path.with_name(workspace_path.name + ".lock").unlink(missing_ok=True)
        return removed
