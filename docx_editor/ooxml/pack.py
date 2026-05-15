"""Pack a directory back into a .docx, .pptx, or .xlsx file."""

import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import defusedxml.minidom

# Workspace-root paths that must not be packed into the output.
# meta.json is workspace bookkeeping (see Workspace.META_FILE); packing it makes
# Word flag the document as "unreadable content" on open. Paths are workspace-root
# relative — a hypothetical subpart literally named "meta.json" is unaffected.
EXCLUDED_PATHS = {Path("meta.json")}


def pack_document(input_dir: str | Path, output_file: str | Path, validate: bool = False) -> bool:
    """Pack a directory into an Office file (.docx/.pptx/.xlsx).

    Args:
        input_dir: Path to unpacked Office document directory
        output_file: Path to output Office file
        validate: If True, validates with soffice (default: False)

    Returns:
        bool: True if successful, False if validation failed

    Raises:
        ValueError: If input_dir doesn't exist or output_file has wrong extension
    """
    input_dir = Path(input_dir)
    output_file = Path(output_file)

    if not input_dir.is_dir():
        raise ValueError(f"{input_dir} is not a directory")
    if output_file.suffix.lower() not in {".docx", ".pptx", ".xlsx"}:
        raise ValueError(f"{output_file} must be a .docx, .pptx, or .xlsx file")

    # Work in temporary directory to avoid modifying original
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_content_dir = Path(temp_dir) / "content"
        shutil.copytree(input_dir, temp_content_dir)

        # Process XML files to remove pretty-printing whitespace
        for pattern in ["*.xml", "*.rels"]:
            for xml_file in temp_content_dir.rglob(pattern):
                condense_xml(xml_file)

        # Deterministic ZIP: sorted POSIX names, fixed 1980 timestamps, pinned metadata - cross-platform byte stability.
        output_file.parent.mkdir(parents=True, exist_ok=True)
        entries = sorted(
            (f for f in temp_content_dir.rglob("*") if f.is_file()),
            key=lambda f: f.relative_to(temp_content_dir).as_posix(),
        )
        with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in entries:
                rel = f.relative_to(temp_content_dir)
                if rel in EXCLUDED_PATHS:
                    continue
                info = zipfile.ZipInfo(rel.as_posix(), date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                info.create_system = 3  # Unix, pinned for cross-platform byte stability
                info.external_attr = 0o644 << 16
                info._compresslevel = 6  # writestr(ZipInfo) ignores ZipFile.compresslevel
                with f.open("rb") as src, zf.open(info, "w") as dst:
                    shutil.copyfileobj(src, dst)

        # Validate if requested
        if validate:  # pragma: no cover
            if not validate_document(output_file):
                output_file.unlink()  # Delete the corrupt file
                return False

    return True


def validate_document(doc_path: Path) -> bool:  # pragma: no cover
    """Validate document by converting to HTML with soffice."""
    # Determine the correct filter based on file extension
    match doc_path.suffix.lower():
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
