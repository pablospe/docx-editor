"""Tests for :meth:`Document.get_paragraph_location`.

Covers:
  * body paragraphs report no table
  * in-table paragraphs report ``TableCell`` coordinates
  * ``col`` is the logical-grid column (``w:gridSpan``-aware), not the raw
    ``<w:tc>`` count
  * nested tables increment ``depth`` and produce a depth-first doc-wide
    table ``index``
  * stale refs raise ``HashMismatchError``; out-of-range refs raise
    ``ParagraphIndexError``

The fixture is built by swapping ``word/document.xml`` inside a copy of
``simple.docx`` (already in ``tests/test_data``) so we don't need to commit
another binary and don't pull in ``python-docx``.
"""

import zipfile
from pathlib import Path

import pytest

from docx_editor import (
    Document,
    HashMismatchError,
    ParagraphIndexError,
    ParagraphLocation,
    TableCell,
)

# ---------- fixture builders -------------------------------------------------

_W_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
)


def _build_document_xml() -> str:
    """Hand-written ``word/document.xml`` with a gridSpan row and a nested table.

    Paragraph layout (P{n} = 1-based paragraph index after ``Document.open``):

      P1   body                                        "Body paragraph 1"
      P2   table 1 / row 1 / col 1                     "A"
      P3   table 1 / row 1 / col 2 (gridSpan=2)        "B (spans 2)"
      P4   table 1 / row 1 / col 4                     "C"
      P5   table 1 / row 2 / col 1                     "Row 2 col 1"
      P6   table 1 / row 2 / col 2                     "Row 2 col 2"
      P7   table 1 / row 2 / col 3                     "Row 2 col 3"
      P8   body                                        "Body paragraph 2"
      P9   table 2 / row 1 / col 1 (outer cell)        "Outer cell"
      P10  table 3 / row 1 / col 1 (nested inside P9)  "Inner cell"
      P11  body                                        "Body paragraph 3"
    """
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_W_NS}>"
        "<w:body>"
        "<w:p><w:r><w:t>Body paragraph 1</w:t></w:r></w:p>"
        # --- table 1: a row with a gridSpan=2 cell ---
        "<w:tbl>"
        "<w:tr>"
        "<w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>"
        '<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>'
        "<w:p><w:r><w:t>B (spans 2)</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>"
        "</w:tr>"
        "<w:tr>"
        "<w:tc><w:p><w:r><w:t>Row 2 col 1</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>Row 2 col 2</w:t></w:r></w:p></w:tc>"
        "<w:tc><w:p><w:r><w:t>Row 2 col 3</w:t></w:r></w:p></w:tc>"
        "</w:tr>"
        "</w:tbl>"
        "<w:p><w:r><w:t>Body paragraph 2</w:t></w:r></w:p>"
        # --- table 2: a cell that contains another (nested) table ---
        "<w:tbl>"
        "<w:tr>"
        "<w:tc>"
        "<w:p><w:r><w:t>Outer cell</w:t></w:r></w:p>"
        "<w:tbl>"
        "<w:tr>"
        "<w:tc><w:p><w:r><w:t>Inner cell</w:t></w:r></w:p></w:tc>"
        "</w:tr>"
        "</w:tbl>"
        "</w:tc>"
        "</w:tr>"
        "</w:tbl>"
        "<w:p><w:r><w:t>Body paragraph 3</w:t></w:r></w:p>"
        "</w:body>"
        "</w:document>"
    )


def _replace_document_xml(src: Path, dest: Path, new_doc_xml: str) -> None:
    """Copy ``src`` to ``dest``, swapping ``word/document.xml`` for ``new_doc_xml``."""
    with (
        zipfile.ZipFile(src, "r") as z_in,
        zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z_out,
    ):
        for item in z_in.infolist():
            data = new_doc_xml.encode("utf-8") if item.filename == "word/document.xml" else z_in.read(item.filename)
            z_out.writestr(item, data)


@pytest.fixture
def gridspan_docx(simple_docx, tmp_path) -> Path:
    """A .docx whose body contains a gridSpan row and a nested table."""
    dest = tmp_path / "gridspan.docx"
    _replace_document_xml(simple_docx, dest, _build_document_xml())
    return dest


def _ref_for_text(doc: Document, snippet: str) -> str:
    """Return the first paragraph ref whose preview contains ``snippet``."""
    for entry in doc.list_paragraphs():
        if snippet in entry:
            return entry.split("|")[0]
    raise AssertionError(f"No paragraph contains {snippet!r}")


# ---------- tests ------------------------------------------------------------


class TestParagraphLocationBody:
    def test_body_paragraph_has_no_table(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            ref = _ref_for_text(doc, "Body paragraph 1")
            loc = doc.get_paragraph_location(ref)
            assert isinstance(loc, ParagraphLocation)
            assert loc.in_table is False
            assert loc.table is None

    def test_returns_paragraph_location_dataclass(self, gridspan_docx):
        """API contract: returns ``ParagraphLocation``, not a tuple or dict."""
        with Document.open(gridspan_docx) as doc:
            ref = _ref_for_text(doc, "Row 2 col 2")
            loc = doc.get_paragraph_location(ref)
            assert isinstance(loc, ParagraphLocation)
            assert isinstance(loc.table, TableCell)


class TestParagraphLocationInTable:
    def test_reports_row_and_logical_col(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            ref = _ref_for_text(doc, "Row 2 col 2")
            loc = doc.get_paragraph_location(ref)
            assert loc.in_table is True
            assert loc.table == TableCell(index=1, row=2, col=2, depth=1)

    def test_first_row_cells_report_correct_coordinates(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            loc_a = doc.get_paragraph_location(_ref_for_text(doc, "A"))
            assert loc_a.table == TableCell(index=1, row=1, col=1, depth=1)


class TestParagraphLocationGridSpan:
    """A cell visually in column 4 must report ``col=4`` even though only 3
    ``<w:tc>`` elements precede it in the row — the second cell carries
    ``w:gridSpan=2``.
    """

    def test_logical_col_accounts_for_grid_span(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            loc_a = doc.get_paragraph_location(_ref_for_text(doc, "A"))
            loc_b = doc.get_paragraph_location(_ref_for_text(doc, "B (spans 2)"))
            loc_c = doc.get_paragraph_location(_ref_for_text(doc, "C"))
        assert loc_a.table is not None
        assert loc_b.table is not None
        assert loc_c.table is not None
        assert (loc_a.table.col, loc_b.table.col, loc_c.table.col) == (1, 2, 4)


class TestParagraphLocationNested:
    def test_outer_table_paragraph(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Outer cell"))
        # Second top-level table in the document.
        assert loc.table == TableCell(index=2, row=1, col=1, depth=1)

    def test_inner_table_paragraph(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Inner cell"))
        # depth=2 because nested inside another tbl; index=3 in doc-wide
        # depth-first order (table 1 outer, table 2 outer, table 3 nested).
        assert loc.table == TableCell(index=3, row=1, col=1, depth=2)


class TestParagraphLocationStaleRefs:
    """Mirror the error contract used by edit methods and ``add_comment``."""

    def test_hash_mismatch_raises(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            ref = _ref_for_text(doc, "Body paragraph 1")
            index_part, _hash = ref.split("#")
            bad_ref = f"{index_part}#0000"  # forged hash
            with pytest.raises(HashMismatchError) as exc:
                doc.get_paragraph_location(bad_ref)
            assert exc.value.paragraph_index == int(index_part[1:])
            assert exc.value.expected_hash == "0000"

    def test_out_of_range_index_raises(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            n = len(doc.list_paragraphs())
            bad_ref = f"P{n + 99}#0000"
            with pytest.raises(ParagraphIndexError) as exc:
                doc.get_paragraph_location(bad_ref)
            assert exc.value.index == n + 99
            assert exc.value.total_paragraphs == n

    def test_malformed_ref_raises_value_error(self, gridspan_docx):
        with (
            Document.open(gridspan_docx) as doc,
            pytest.raises(ValueError, match="Invalid paragraph reference"),
        ):
            doc.get_paragraph_location("not-a-ref")
