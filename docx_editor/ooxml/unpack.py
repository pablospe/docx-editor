"""Unpack and format XML contents of Office files (.docx, .pptx, .xlsx)."""

import random
import shutil
import stat
import zipfile
from pathlib import Path
from xml.parsers.expat import ExpatError

import defusedxml.minidom
from defusedxml.common import DefusedXmlException

from docx_editor.exceptions import DocumentNotFoundError, InvalidDocumentError


def _is_unsafe_zip_path(name: str) -> bool:
    """Reject absolute paths, colons, and '..' traversal segments."""
    if not name or name.startswith(("/", "\\")):
        return True
    # Reject any colon: covers Windows drive letters (C:\) and NTFS alternate
    # data streams (foo.xml:stream). OOXML part names never contain a colon.
    if ":" in name:
        return True
    for seg in name.replace("\\", "/").split("/"):
        # Reject literal "..", and any segment composed only of dots/spaces:
        # NTFS strips trailing spaces and dots from path components, so ".. "
        # and "..." can normalize to "..". OOXML never uses such segments.
        if seg == ".." or (seg and seg.strip(" .") == ""):
            return True
    return False


def _is_symlink_entry(info: zipfile.ZipInfo) -> bool:
    """Return True if the ZIP entry is a Unix symlink."""
    if info.create_system != 3:  # not Unix-created
        return False
    mode = (info.external_attr >> 16) & 0xFFFF
    return stat.S_ISLNK(mode)


def unpack_document(input_file: str | Path, output_dir: str | Path) -> str:
    """Unpack a .docx file to a directory with pretty-printed XML.

    On failure after extraction has started, the output directory is removed
    if this call created it; a pre-existing directory is left in place.

    Args:
        input_file: Path to the .docx file to unpack
        output_dir: Directory to extract contents to

    Returns:
        str: Suggested RSID for edit session (8-character hex string)

    Raises:
        DocumentNotFoundError: If the input file doesn't exist
        InvalidDocumentError: If the input path is a directory, the file is not
            a valid zip/docx, contains an unsafe entry path or a symlink entry,
            is missing the required word/document.xml part (rejected before
            extraction), a part contains malformed XML or XML constructs
            refused for security (DTD entity/external declarations), or the
            output directory is (or contains) a symlink or is an existing
            non-directory.
    """
    input_path = Path(input_file)
    output_path = Path(output_dir)

    if not input_path.exists():
        raise DocumentNotFoundError(f"Document not found: {input_file}")

    if input_path.is_dir():
        raise InvalidDocumentError(f"Is a directory, not a .docx file: {input_file}")

    # Reject a symlinked destination so extractall cannot write through it.
    if output_path.is_symlink():
        raise InvalidDocumentError(f"Output directory is a symlink: {output_dir}")
    if output_path.exists():
        if not output_path.is_dir():
            raise InvalidDocumentError(f"Output path is not a directory: {output_dir}")
        for entry in output_path.rglob("*"):
            if entry.is_symlink():
                raise InvalidDocumentError(f"Symlink inside output directory: {entry}")

    # Ownership of the output dir is claimed by the mkdir itself, not by a
    # separate exists() check, so a directory created concurrently by someone
    # else can never be adopted and then deleted by the failure cleanup.
    created_output_dir = False

    try:
        # Validate ZIP entries before extractall (and before mkdir) so a
        # rejection leaves the filesystem untouched.
        try:
            with zipfile.ZipFile(input_path) as zf:
                names: set[str] = set()
                for info in zf.infolist():
                    if _is_unsafe_zip_path(info.filename):
                        raise InvalidDocumentError(f"Unsafe ZIP entry path: {info.filename!r} in {input_file}")
                    if _is_symlink_entry(info):
                        raise InvalidDocumentError(f"Symlink ZIP entry: {info.filename!r} in {input_file}")
                    names.add(info.filename)
                # The one part every consumer unconditionally dereferences;
                # without it, opening would fail later with a raw
                # FileNotFoundError naming the internal cache path.
                if "word/document.xml" not in names:
                    raise InvalidDocumentError(f"Not a valid .docx: missing word/document.xml in {input_file}")
                try:
                    output_path.mkdir(parents=True)
                    created_output_dir = True
                except FileExistsError:
                    pass
                zf.extractall(output_path)
        except zipfile.BadZipFile as e:
            raise InvalidDocumentError(f"Not a valid .docx file: {input_file}") from e

        # Pretty print all XML files
        xml_files = list(output_path.rglob("*.xml")) + list(output_path.rglob("*.rels"))
        for xml_file in xml_files:
            # Parse bytes, not decoded text: expat then honors the part's XML
            # encoding declaration (a UTF-16 part is valid OOXML), and byte-level
            # garbage surfaces as ExpatError below instead of a raw
            # UnicodeDecodeError escaping the InvalidDocumentError contract.
            content = xml_file.read_bytes()
            try:
                dom = defusedxml.minidom.parseString(content)
            except (ExpatError, DefusedXmlException) as e:
                # DefusedXmlException covers EntitiesForbidden and siblings; it
                # subclasses ValueError, so catch the defused type specifically
                # to keep unrelated ValueErrors loud. `from e` preserves the
                # security exception for audit.
                raise InvalidDocumentError(
                    f"Invalid XML in {xml_file.relative_to(output_path).as_posix()} of {input_file}: {e}"
                ) from e
            xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="utf-8"))
    except BaseException:
        # A partially-extracted dir would wedge later opens; only remove it if
        # this call created it — a pre-existing dir may hold caller data.
        if created_output_dir:
            shutil.rmtree(output_path, ignore_errors=True)
        raise

    # Generate and return RSID for tracked changes
    suggested_rsid = "".join(random.choices("0123456789ABCDEF", k=8))
    return suggested_rsid
