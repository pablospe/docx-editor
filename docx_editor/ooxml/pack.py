"""Pack a directory back into a .docx, .pptx, or .xlsx file."""

import contextlib
import errno
import os
import secrets
import shutil
import stat
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import BinaryIO

import defusedxml.minidom

from ..exceptions import DocumentOpenError

# Workspace-root paths that must not be packed into the output.
# meta.json is workspace bookkeeping (see Workspace.META_FILE); packing it makes
# Word flag the document as "unreadable content" on open. meta.json.tmp is
# Workspace._save_meta's atomic-write staging file — a crash can orphan it.
# Paths are workspace-root relative — a hypothetical subpart literally named
# "meta.json" is unaffected.
EXCLUDED_PATHS = {Path("meta.json"), Path("meta.json.tmp")}

# The promotion temp file is named ".<destination name>.<8 random chars>.tmp"; this
# is everything in that shape except the destination's own name.
_TEMP_NAME_OVERHEAD = len(".") + len(".") + 8 + len(".tmp")
_TEMP_NAME_ATTEMPTS = 100


def _ignore_symlinks(directory: str, names: list[str]) -> list[str]:
    """Skip symlinks so they cannot leak external host content into the archive."""
    base = Path(directory)
    return [n for n in names if (base / n).is_symlink()]


def pack_document(input_dir: str | Path, output_file: str | Path, validate: bool = False) -> bool:
    """Pack a directory into an Office file (.docx/.pptx/.xlsx).

    Args:
        input_dir: Path to unpacked Office document directory
        output_file: Path to output Office file
        validate: If True, validates with soffice (default: False)

    Returns:
        bool: True if successful, False if validation failed

    Raises:
        ValueError: If input_dir is a symlink, doesn't exist, or output_file has
            the wrong extension
        DocumentOpenError: If the OS denies the final replace (the destination is
            held open by another program, e.g. Word on Windows)
    """
    input_dir = Path(input_dir)
    output_file = Path(output_file)

    if input_dir.is_symlink():
        raise ValueError(f"{input_dir} is a symlink")
    if not input_dir.is_dir():
        raise ValueError(f"{input_dir} is not a directory")
    if output_file.suffix.lower() not in {".docx", ".pptx", ".xlsx"}:
        raise ValueError(f"{output_file} must be a .docx, .pptx, or .xlsx file")

    # Follow a symlinked destination to the file it points at. os.replace() acts on
    # the name it is given, so replacing the link itself would leave the real
    # document untouched while reporting success. The pre-atomic code opened the
    # path for writing and therefore followed the link; preserve that.
    target = Path(os.path.realpath(output_file))

    # Refuse a write-protected destination up front, before doing any work. The
    # pre-atomic code failed the moment it opened the file for writing, so a
    # validate=True run must not be able to return False first and quietly swallow
    # the refusal. _replace() checks again as a race guard.
    _assert_writable(target)

    # Work in temporary directory to avoid modifying original
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_content_dir = Path(temp_dir) / "content"
        shutil.copytree(input_dir, temp_content_dir, ignore=_ignore_symlinks)

        # Process XML files to remove pretty-printing whitespace
        for pattern in ["*.xml", "*.rels"]:
            for xml_file in temp_content_dir.rglob(pattern):
                condense_xml(xml_file)

        # Deterministic ZIP: sorted POSIX names, fixed 1980 timestamps, pinned metadata - cross-platform byte stability.
        target.parent.mkdir(parents=True, exist_ok=True)
        entries = sorted(
            (f for f in temp_content_dir.rglob("*") if f.is_file()),
            key=lambda f: f.relative_to(temp_content_dir).as_posix(),
        )

        # Atomic write: build the archive in a temp file in the destination's own
        # directory, then promote it with os.replace(). The destination is never
        # observed half-written (safe inside cloud-synced folders), and any failure
        # — a write error or a failed validation — leaves the existing destination
        # untouched. Same directory ⇒ same volume ⇒ os.replace() is a true atomic
        # rename (no cross-device error).
        tmp_file, tmp_path = _create_temp(target)
        try:
            # Write through the temp file's own fd rather than reopening it by name.
            # Reopening would fail wherever the created file is not owner-writable —
            # a directory with a restrictive default POSIX ACL masks the new file so
            # its owner cannot write it — and it would leave a window in which the name
            # could be swapped under us.
            with tmp_file:
                # Take the destination's permissions *before* writing any content, so
                # the document is never briefly more readable than the file it replaces.
                if target.exists():
                    _chmod(tmp_file, tmp_path, stat.S_IMODE(target.stat().st_mode))

                with zipfile.ZipFile(tmp_file, "w", zipfile.ZIP_DEFLATED) as zf:
                    for f in entries:
                        rel = f.relative_to(temp_content_dir)
                        if rel in EXCLUDED_PATHS:
                            continue
                        info = zipfile.ZipInfo(rel.as_posix(), date_time=(1980, 1, 1, 0, 0, 0))
                        info.compress_type = zipfile.ZIP_DEFLATED
                        info.create_system = 3  # Unix, pinned for cross-platform byte stability
                        info.external_attr = 0o644 << 16
                        # _compresslevel: stdlib's documented per-entry escape hatch (stable since 3.7);
                        # ZipFile.compresslevel is not propagated to ZipInfo entries.
                        info._compresslevel = 6  # ty: ignore[unresolved-attribute]
                        with f.open("rb") as src, zf.open(info, "w") as dst:
                            shutil.copyfileobj(src, dst)

                # Flush contents and mode to disk before promoting. os.replace() is
                # atomic with respect to other processes, but without this a crash
                # could leave the destination name pointing at an inode whose data
                # blocks were never written.
                tmp_file.flush()
                os.fsync(tmp_file.fileno())

            # Validate the temp file, never the destination, so a failure cannot
            # destroy the original (historic data-loss bug: unlink of output_file).
            # Pass the real extension: the temp file is named .tmp, and soffice's
            # export filter is chosen from the suffix. Use output_file's suffix, not
            # target's — output_file is what the extension check above accepted, and
            # a symlink may point at a name with any extension at all.
            if validate and not validate_document(tmp_path, suffix=output_file.suffix):
                return False

            _replace(tmp_path, target)
        finally:
            # The temp file carries the destination's mode, and on Windows a read-only
            # file cannot be deleted — clear that before trying, or a read-only
            # destination would leave the temp archive littering a synced folder.
            if sys.platform == "win32":  # pragma: no cover - CI is POSIX
                with contextlib.suppress(OSError):
                    os.chmod(tmp_path, stat.S_IWRITE | stat.S_IREAD)
            # Remove the temp file on every failure path (write error, validation
            # failure). On success os.replace() consumed it, so this is a no-op.
            # Suppress errors so a doomed cleanup cannot turn a completed save into
            # a reported failure (or mask an in-flight exception).
            with contextlib.suppress(OSError):
                tmp_path.unlink(missing_ok=True)

    return True


def _has_surrogates(text: str) -> bool:
    """True if ``text`` carries surrogateescape markers for undecodable bytes."""
    return any(0xDC80 <= ord(ch) <= 0xDCFF for ch in text)


def _clamp_name(name: str, budget: int) -> str:
    """Trim ``name`` to at most ``budget`` bytes, without splitting a character.

    NAME_MAX is a byte budget, not a character count, so the clamp has to happen on
    the encoding — slicing characters would still overflow for a non-ASCII name at up
    to 4 bytes each. Go through os.fsencode/os.fsdecode rather than str.encode:
    Linux filenames are bytes, and an undecodable one arrives here as surrogates that
    strict UTF-8 would reject.

    Slicing bytes can cut a multibyte character in half, and the halves come back as
    lone surrogates that macOS and Windows refuse in a filename. So back off to a
    character boundary — unless the name was *already* undecodable, in which case its
    surrogates are legitimate and must survive.
    """
    raw = os.fsencode(name)
    if len(raw) <= budget:
        return name

    was_undecodable = _has_surrogates(name)
    raw = raw[:budget]
    stem = os.fsdecode(raw)
    if not was_undecodable:
        while raw and _has_surrogates(stem):
            raw = raw[:-1]
            stem = os.fsdecode(raw)
    return stem


def _create_temp(target: Path) -> tuple[BinaryIO, Path]:
    """Create the promotion temp file next to ``target``; return it open, with its path.

    Deliberately not tempfile.mkstemp(): that forces mode 0600, which for a
    brand-new destination would then have to be corrected back to the process
    default — and the only stdlib way to read the umask (set it to 0, then restore)
    mutates global state a concurrent thread can observe. Creating the file 0666
    lets the kernel apply the umask (and any default ACL) itself, so a new
    destination lands on exactly the mode a plain open() would have produced, with
    no race. O_EXCL keeps the name ours alone.

    The file is returned already wrapped, so the raw fd is never exposed unowned.
    """
    try:
        name_max = os.pathconf(target.parent, "PC_NAME_MAX")
    except (AttributeError, OSError, ValueError):  # pragma: no cover - platform-dependent
        name_max = 255
    stem = _clamp_name(target.name, max(8, name_max - _TEMP_NAME_OVERHEAD))

    flags = os.O_RDWR | os.O_CREAT | os.O_EXCL | getattr(os, "O_BINARY", 0)
    for _ in range(_TEMP_NAME_ATTEMPTS):
        candidate = target.parent / f".{stem}.{secrets.token_hex(4)}.tmp"
        try:
            fd = os.open(candidate, flags, 0o666)
        except FileExistsError:  # pragma: no cover - requires a name collision
            continue
        try:
            return os.fdopen(fd, "w+b"), candidate
        except BaseException:  # pragma: no cover - fdopen failing on a fresh fd
            os.close(fd)
            raise
    raise OSError(f"Could not create a unique temp file next to {target}")  # pragma: no cover


def _chmod(tmp_file: BinaryIO, tmp_path: Path, mode: int) -> None:
    """Set the temp file's mode, preferring the fd so no path race is possible."""
    if hasattr(os, "fchmod"):
        os.fchmod(tmp_file.fileno(), mode)
    else:  # pragma: no cover - Windows has no fchmod
        os.chmod(tmp_path, mode)


def _assert_writable(target: Path) -> None:
    """Refuse a write-protected destination.

    rename(2) only needs write permission on the *directory*, so the atomic promotion
    would happily replace a file the user marked read-only — something the pre-atomic
    in-place write could never do, and that Windows still refuses. Keep that contract,
    and keep both platforms agreeing.
    """
    if target.exists() and not os.access(target, os.W_OK):
        raise PermissionError(errno.EACCES, os.strerror(errno.EACCES), str(target))


def _replace(tmp_path: Path, target: Path) -> None:
    """Atomically move the finished archive onto the destination, and persist it."""
    _assert_writable(target)  # re-check: the mode may have changed while we packed

    try:
        os.replace(tmp_path, target)
    except PermissionError as e:
        # On Windows a destination locked by another program (Word) surfaces here as
        # PermissionError. On POSIX it cannot: rename(2) over a file another process
        # holds open always succeeds. A denial there means something else entirely —
        # a sticky-bit directory, an immutable file, an SELinux denial — and calling
        # that "the document is open" would send the caller down a dead end.
        #
        # A read-only destination on Windows also lands here (MoveFileEx refuses to
        # replace a FILE_ATTRIBUTE_READONLY file), and a destination that does not
        # exist yet cannot be open in anything. Neither is an open document — only
        # claim "open" for an existing, otherwise-writable file.
        looks_open = sys.platform == "win32" and target.exists() and os.access(target, os.W_OK)
        if not looks_open:
            raise
        raise DocumentOpenError(
            f"{target} could not be replaced (permission denied); it is likely open "
            f"in another program. Close it and retry.",
            path=target,
        ) from e

    # Persist the rename. Directory fsync is unsupported on Windows, and unnecessary.
    with contextlib.suppress(OSError):
        dir_fd = os.open(target.parent, os.O_RDONLY)
        try:
            os.fsync(dir_fd)
        finally:
            os.close(dir_fd)


def validate_document(doc_path: Path, suffix: str | None = None) -> bool:  # pragma: no cover
    """Validate document by converting to HTML with soffice.

    Args:
        doc_path: File to validate. May be a temp file whose own extension does
            not reflect the real format.
        suffix: The real OOXML extension (".docx"/".pptx"/".xlsx") to pick the
            soffice export filter with. Defaults to doc_path's own suffix.
    """
    # Determine the correct filter based on file extension
    match (suffix or doc_path.suffix).lower():
        case ".docx":
            filter_name = "html:HTML"
        case ".pptx":
            filter_name = "html:impress_html_Export"
        case ".xlsx":
            filter_name = "html:HTML (StarCalc)"
        case _:
            filter_name = "html:HTML"

    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            result = subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to",
                    filter_name,
                    "--outdir",
                    temp_dir,
                    str(doc_path),
                ],
                capture_output=True,
                timeout=10,
                text=True,
            )
            if not (Path(temp_dir) / f"{doc_path.stem}.html").exists():
                error_msg = result.stderr.strip() or "Document validation failed"
                print(f"Validation error: {error_msg}", file=sys.stderr)
                return False
            return True
        except FileNotFoundError:
            print("Warning: soffice not found. Skipping validation.", file=sys.stderr)
            return True
        except subprocess.TimeoutExpired:
            print("Validation error: Timeout during conversion", file=sys.stderr)
            return False
        except Exception as e:
            print(f"Validation error: {e}", file=sys.stderr)
            return False


def condense_xml(xml_file: Path) -> None:
    """Strip unnecessary whitespace and remove comments from XML file."""
    with open(xml_file, encoding="utf-8") as f:
        dom = defusedxml.minidom.parse(f)

    # Process each element to remove whitespace and comments
    for element in dom.getElementsByTagName("*"):
        # Skip text-bearing OOXML elements: w:t, w:delText, w:instrText (and their
        # namespace variants). Their content is significant — including whitespace
        # fragments that minidom may split off as their own TEXT_NODE (issue #9).
        if element.tagName.endswith((":t", ":delText", ":instrText")):
            continue

        # Remove whitespace-only text nodes and comment nodes
        for child in list(element.childNodes):
            if (
                child.nodeType == child.TEXT_NODE and child.nodeValue and child.nodeValue.strip() == ""
            ) or child.nodeType == child.COMMENT_NODE:
                element.removeChild(child)

    # Write back the condensed XML
    with open(xml_file, "wb") as f:
        f.write(dom.toxml(encoding="UTF-8"))
