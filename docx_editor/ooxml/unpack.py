"""Unpack and format XML contents of Office files (.docx, .pptx, .xlsx)."""

import random
import zipfile
from pathlib import Path

import defusedxml.minidom


def unpack_document(input_file: str | Path, output_dir: str | Path) -> str:
    """Unpack a .docx file to a directory with pretty-printed XML.

    Args:
        input_file: Path to the .docx file to unpack
        output_dir: Directory to extract contents to

    Returns:
        str: Suggested RSID for edit session (8-character hex string)

    Raises:
        FileNotFoundError: If the input file doesn't exist
        zipfile.BadZipFile: If the file is not a valid zip/docx
    """
    input_path = Path(input_file)
    output_path = Path(output_dir)

    if not input_path.exists():
        raise FileNotFoundError(f"Document not found: {input_file}")

    # Extract and format
    output_path.mkdir(parents=True, exist_ok=True)
    zipfile.ZipFile(input_path).extractall(output_path)

    # Pretty print all XML files
    xml_files = list(output_path.rglob("*.xml")) + list(output_path.rglob("*.rels"))
    for xml_file in xml_files:
        content = xml_file.read_text(encoding="utf-8")
        dom = defusedxml.minidom.parseString(content)
        xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="ascii"))

    # Generate and return RSID for tracked changes
    suggested_rsid = "".join(random.choices("0123456789ABCDEF", k=8))
    return suggested_rsid
