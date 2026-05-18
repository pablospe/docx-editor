"""Unpack and format XML contents of Office files (.docx, .pptx, .xlsx)."""

import random
import stat
import zipfile
from pathlib import Path

import defusedxml.minidom

from docx_editor.exceptions import DocumentNotFoundError, InvalidDocumentError


def _is_unsafe_zip_path(name: str) -> bool:
    """Reject absolute paths, drive letters, and '..' traversal segments."""
    if not name or name.startswith(("/", "\\")):
        return True
    if len(name) >= 2 and name[1] == ":":  # Windows drive letter, e.g. C:\
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

    Args:
        input_file: Path to the .docx file to unpack
        output_dir: Directory to extract contents to

    Returns:
        str: Suggested RSID for edit session (8-character hex string)

    Raises:
        DocumentNotFoundError: If the input file doesn't exist
        InvalidDocumentError: If the file is not a valid zip/docx
    """
    input_path = Path(input_file)
    output_path = Path(output_dir)

    if not input_path.exists():
        raise DocumentNotFoundError(f"Document not found: {input_file}")

    # Reject a symlinked destination so extractall cannot write through it.
    if output_path.is_symlink():
        raise InvalidDocumentError(f"Output directory is a symlink: {output_dir}")
    if output_path.exists():
        for entry in output_path.rglob("*"):
            if entry.is_symlink():
                raise InvalidDocumentError(f"Symlink inside output directory: {entry}")

    # Validate ZIP entries before extractall (and before mkdir) so a rejection
    # leaves the filesystem untouched.
    try:
        with zipfile.ZipFile(input_path) as zf:
            for info in zf.infolist():
                if _is_unsafe_zip_path(info.filename):
                    raise InvalidDocumentError(f"Unsafe ZIP entry path: {info.filename!r} in {input_file}")
                if _is_symlink_entry(info):
                    raise InvalidDocumentError(f"Symlink ZIP entry: {info.filename!r} in {input_file}")
            output_path.mkdir(parents=True, exist_ok=True)
            zf.extractall(output_path)
    except zipfile.BadZipFile as e:
        raise InvalidDocumentError(f"Not a valid .docx file: {input_file}") from e

    # Pretty print all XML files
    xml_files = list(output_path.rglob("*.xml")) + list(output_path.rglob("*.rels"))
    for xml_file in xml_files:
        content = xml_file.read_text(encoding="utf-8")
        dom = defusedxml.minidom.parseString(content)
        xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="ascii"))

    # Generate and return RSID for tracked changes
    suggested_rsid = "".join(random.choices("0123456789ABCDEF", k=8))
    return suggested_rsid
