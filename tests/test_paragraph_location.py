"""Tests for :meth:`Document.get_paragraph_location`.

Covers:
  * body paragraphs report no table
  * in-table paragraphs report ``TableCell`` coordinates
  * ``col`` is the logical-grid column (``w:gridSpan``-aware), not the raw
    ``<w:tc>`` count
  * nested tables increment ``depth`` and produce a depth-first doc-wide
    table ``index``
  * list paragraphs report ``ListItem(num_id, ilvl)`` from their direct
    ``w:pPr/w:numPr``; non-list paragraphs report ``list=None``
  * ``style`` reports the raw ``w:pStyle`` id; ``outline_level`` comes from
    a direct ``w:outlineLvl`` or the style definition in ``word/styles.xml``
    (``w:basedOn`` chains resolved, ``w:val="9"`` = body text)
  * ``heading_path`` is the chain of nearest preceding headings, outermost
    first, excluding the paragraph itself
  * ``section`` is the 1-based section index: a direct ``w:pPr/w:sectPr``
    closes a section (the carrying paragraph belongs to the section it
    closes); the body-level ``w:sectPr`` defines the final section
  * stale refs raise ``HashMismatchError``; out-of-range refs raise
    ``ParagraphIndexError``

The fixture is built by swapping ``word/document.xml`` inside a copy of
``simple.docx`` (already in ``tests/test_data``) so we don't need to commit
another binary and don't pull in ``python-docx``.
"""

from pathlib import Path

import pytest
from conftest import replace_document_xml, replace_docx_parts

from docx_editor import (
    Document,
    HashMismatchError,
    ListItem,
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


@pytest.fixture
def gridspan_docx(simple_docx, tmp_path) -> Path:
    """A .docx whose body contains a gridSpan row and a nested table."""
    dest = tmp_path / "gridspan.docx"
    replace_document_xml(simple_docx, dest, _build_document_xml())
    return dest


def _num_pr(num_id: str, ilvl: str | None) -> str:
    """``<w:pPr><w:numPr>...`` block with the given numId and optional ilvl."""
    ilvl_xml = f'<w:ilvl w:val="{ilvl}"/>' if ilvl is not None else ""
    return f'<w:pPr><w:numPr>{ilvl_xml}<w:numId w:val="{num_id}"/></w:numPr></w:pPr>'


def _build_numbered_document_xml() -> str:
    """Hand-written ``word/document.xml`` with numbered paragraphs.

    Layout (unique text snippets for ``_ref_for_text``):

      "Plain paragraph"    body, no numPr                  → list=None
      "Numbered top"       numId=5, ilvl=0                 → ListItem(5, 0)
      "Numbered nested"    numId=5, ilvl=1                 → ListItem(5, 1)
      "Numbered deeper"    numId=5, ilvl=2                 → ListItem(5, 2)
      "Numbered no ilvl"   numId=5, no <w:ilvl>            → ListItem(5, 0)
      "Cell numbered"      table cell + numId=7, ilvl=0    → table AND list
    """
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_W_NS}>"
        "<w:body>"
        "<w:p><w:r><w:t>Plain paragraph</w:t></w:r></w:p>"
        f"<w:p>{_num_pr('5', '0')}<w:r><w:t>Numbered top</w:t></w:r></w:p>"
        f"<w:p>{_num_pr('5', '1')}<w:r><w:t>Numbered nested</w:t></w:r></w:p>"
        f"<w:p>{_num_pr('5', '2')}<w:r><w:t>Numbered deeper</w:t></w:r></w:p>"
        f"<w:p>{_num_pr('5', None)}<w:r><w:t>Numbered no ilvl</w:t></w:r></w:p>"
        # --- a one-cell table whose paragraph is itself numbered ---
        "<w:tbl><w:tr><w:tc>"
        f"<w:p>{_num_pr('7', '0')}<w:r><w:t>Cell numbered</w:t></w:r></w:p>"
        "</w:tc></w:tr></w:tbl>"
        "</w:body>"
        "</w:document>"
    )


@pytest.fixture
def numbered_docx(simple_docx, tmp_path) -> Path:
    """A .docx with numbered paragraphs at several levels, incl. one in a table."""
    dest = tmp_path / "numbered.docx"
    replace_document_xml(simple_docx, dest, _build_numbered_document_xml())
    return dest


def _p_style(style_id: str) -> str:
    """``<w:pPr><w:pStyle .../></w:pPr>`` block with the given style id."""
    return f'<w:pPr><w:pStyle w:val="{style_id}"/></w:pPr>'


def _build_styled_document_xml() -> str:
    """Hand-written ``word/document.xml`` with styled/heading paragraphs.

    Relies on the styles.xml shipped inside ``simple.docx``: ``Heading1``/
    ``Heading2`` define ``w:outlineLvl`` 0/1, and ``TOCHeading`` is based
    on ``Heading1`` but carries ``w:outlineLvl w:val="9"`` (the spec's
    "no outline" marker).

    Layout (unique text snippets for ``_ref_for_text``):

      "Preamble text"    plain, before any heading    → path ()
      "Chapter one"      pStyle=Heading1              → outline 0, path ()
      "Termination"      pStyle=Heading2              → outline 1, path ("Chapter one",)
      "Body under h2"    plain                        → path ("Chapter one", "Termination")
      "Cell under h2"    table cell paragraph         → table AND heading_path
      "Direct outline"   no pStyle, direct lvl 3      → style None, outline 3
      "Unknown style"    pStyle=NoSuchStyle           → style kept raw, outline None
      "Toc heading"      pStyle=TOCHeading (lvl 9)    → outline None despite basedOn=Heading1
      "Chapter two"      pStyle=Heading1              → path (); pops the stack
      "Body under ch2"   plain                        → path ("Chapter two",)
      "Decoy paragraph"  pStyle/outlineLvl only inside
                         a w:pPrChange revision record → style None, outline None
    """
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_W_NS}>"
        "<w:body>"
        "<w:p><w:r><w:t>Preamble text</w:t></w:r></w:p>"
        f"<w:p>{_p_style('Heading1')}<w:r><w:t>Chapter one</w:t></w:r></w:p>"
        f"<w:p>{_p_style('Heading2')}<w:r><w:t>Termination</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Body under h2</w:t></w:r></w:p>"
        "<w:tbl><w:tr><w:tc>"
        "<w:p><w:r><w:t>Cell under h2</w:t></w:r></w:p>"
        "</w:tc></w:tr></w:tbl>"
        '<w:p><w:pPr><w:outlineLvl w:val="3"/></w:pPr><w:r><w:t>Direct outline</w:t></w:r></w:p>'
        f"<w:p>{_p_style('NoSuchStyle')}<w:r><w:t>Unknown style</w:t></w:r></w:p>"
        f"<w:p>{_p_style('TOCHeading')}<w:r><w:t>Toc heading</w:t></w:r></w:p>"
        f"<w:p>{_p_style('Heading1')}<w:r><w:t>Chapter two</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Body under ch2</w:t></w:r></w:p>"
        "<w:p><w:pPr>"
        '<w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">'
        '<w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/></w:pPr>'
        "</w:pPrChange>"
        "</w:pPr><w:r><w:t>Decoy paragraph</w:t></w:r></w:p>"
        "</w:body>"
        "</w:document>"
    )


@pytest.fixture
def styled_docx(simple_docx, tmp_path) -> Path:
    """A .docx with heading-styled paragraphs (uses simple.docx's own styles.xml)."""
    dest = tmp_path / "styled.docx"
    replace_document_xml(simple_docx, dest, _build_styled_document_xml())
    return dest


def _build_sectioned_document_xml() -> str:
    """Hand-written ``word/document.xml`` with three sections.

    Layout (unique text snippets for ``_ref_for_text``):

      "Intro one"     plain                          → section 1
      "Close one"     w:pPr/w:sectPr                 → section 1 (closes it)
      "Open two"      plain                          → section 2
      "Cell in two"   table cell paragraph           → section 2
      "Close two"     w:pPr/w:sectPr                 → section 2 (closes it)
      "Tail three"    plain                          → section 3
      body-level w:sectPr at end of w:body           (final section)
    """
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_W_NS}>"
        "<w:body>"
        "<w:p><w:r><w:t>Intro one</w:t></w:r></w:p>"
        "<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>Close one</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Open two</w:t></w:r></w:p>"
        "<w:tbl><w:tr><w:tc>"
        "<w:p><w:r><w:t>Cell in two</w:t></w:r></w:p>"
        "</w:tc></w:tr></w:tbl>"
        "<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>Close two</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Tail three</w:t></w:r></w:p>"
        "<w:sectPr/>"
        "</w:body>"
        "</w:document>"
    )


@pytest.fixture
def sectioned_docx(simple_docx, tmp_path) -> Path:
    """A .docx with three sections split by paragraph-level ``w:sectPr``."""
    dest = tmp_path / "sectioned.docx"
    replace_document_xml(simple_docx, dest, _build_sectioned_document_xml())
    return dest


def _single_paragraph_body(p_pr_xml: str) -> str:
    """Full document XML with one ``Target`` paragraph carrying ``p_pr_xml``."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_W_NS}>"
        "<w:body>"
        f"<w:p>{p_pr_xml}<w:r><w:t>Target</w:t></w:r></w:p>"
        "</w:body>"
        "</w:document>"
    )


def _open_single_paragraph(simple_docx: Path, tmp_path: Path, p_pr_xml: str) -> Document:
    """Open a copy of ``simple_docx`` whose body is one ``Target`` paragraph
    carrying ``p_pr_xml`` (the ``<w:pPr>`` block under test).
    """
    dest = tmp_path / "single-paragraph.docx"
    replace_document_xml(simple_docx, dest, _single_paragraph_body(p_pr_xml))
    return Document.open(dest)


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


class TestParagraphLocationMalformedGridSpan:
    """Defensive paths in ``_direct_grid_span`` must default to span=1.

    A malformed ``w:gridSpan`` shouldn't blow up location lookup; the cell
    is treated as if it spans a single column. Verified end-to-end by
    checking the *next* cell in the row lands at ``col=2`` regardless of
    the broken first-cell metadata.
    """

    @staticmethod
    def _doc_with_first_cell_tc_pr(tc_pr_xml: str) -> str:
        """Build a doc whose first table row has two cells; the first
        carries ``tc_pr_xml`` (the ``<w:tcPr>...</w:tcPr>`` block under test)."""
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<w:document {_W_NS}>"
            "<w:body>"
            "<w:tbl><w:tr>"
            f"<w:tc>{tc_pr_xml}<w:p><w:r><w:t>First</w:t></w:r></w:p></w:tc>"
            "<w:tc><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:tc>"
            "</w:tr></w:tbl>"
            "</w:body>"
            "</w:document>"
        )

    @staticmethod
    def _open(simple_docx: Path, tmp_path: Path, body_xml: str) -> Document:
        dest = tmp_path / "malformed.docx"
        replace_document_xml(simple_docx, dest, body_xml)
        return Document.open(dest)

    def test_non_integer_grid_span_falls_back_to_one(self, simple_docx, tmp_path):
        """``<w:gridSpan w:val="abc"/>`` is unparseable; treat as span=1."""
        body = self._doc_with_first_cell_tc_pr('<w:tcPr><w:gridSpan w:val="abc"/></w:tcPr>')
        with self._open(simple_docx, tmp_path, body) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Second"))
            assert loc.table is not None
            assert loc.table.col == 2

    def test_grid_span_without_val_attribute_falls_back_to_one(self, simple_docx, tmp_path):
        """``<w:gridSpan/>`` (no ``w:val``) is treated as span=1."""
        body = self._doc_with_first_cell_tc_pr("<w:tcPr><w:gridSpan/></w:tcPr>")
        with self._open(simple_docx, tmp_path, body) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Second"))
            assert loc.table is not None
            assert loc.table.col == 2

    def test_tc_pr_without_grid_span_falls_back_to_one(self, simple_docx, tmp_path):
        """``<w:tcPr>`` carrying other properties but no ``w:gridSpan`` → span=1."""
        body = self._doc_with_first_cell_tc_pr('<w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>')
        with self._open(simple_docx, tmp_path, body) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Second"))
            assert loc.table is not None
            assert loc.table.col == 2


class TestParagraphLocationGridBefore:
    """``<w:trPr>/<w:gridBefore w:val="N"/>`` shifts the row's first cell to
    logical column ``N+1`` (ragged row). The walker must honour it.
    """

    @staticmethod
    def _doc_with_grid_before(grid_before_val: str | None) -> str:
        gb = f'<w:gridBefore w:val="{grid_before_val}"/>' if grid_before_val is not None else ""
        tr_pr = f"<w:trPr>{gb}</w:trPr>" if gb else ""
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<w:document {_W_NS}>"
            "<w:body>"
            "<w:tbl>"
            f"<w:tr>{tr_pr}"
            "<w:tc><w:p><w:r><w:t>First cell</w:t></w:r></w:p></w:tc>"
            "<w:tc><w:p><w:r><w:t>Second cell</w:t></w:r></w:p></w:tc>"
            "</w:tr>"
            "</w:tbl>"
            "</w:body>"
            "</w:document>"
        )

    def test_grid_before_shifts_first_cell(self, simple_docx, tmp_path):
        """``gridBefore=2`` → first ``<w:tc>`` is at logical col 3."""
        dest = tmp_path / "gridbefore.docx"
        replace_document_xml(simple_docx, dest, self._doc_with_grid_before("2"))
        with Document.open(dest) as doc:
            loc_a = doc.get_paragraph_location(_ref_for_text(doc, "First cell"))
            loc_b = doc.get_paragraph_location(_ref_for_text(doc, "Second cell"))
            assert loc_a.table is not None
            assert loc_b.table is not None
            assert (loc_a.table.col, loc_b.table.col) == (3, 4)

    def test_grid_before_with_non_integer_val_falls_back_to_zero(self, simple_docx, tmp_path):
        """``gridBefore w:val="abc"/`` is unparseable; treat offset as 0."""
        dest = tmp_path / "gb-bad.docx"
        replace_document_xml(simple_docx, dest, self._doc_with_grid_before("abc"))
        with Document.open(dest) as doc:
            loc_a = doc.get_paragraph_location(_ref_for_text(doc, "First cell"))
            assert loc_a.table is not None
            assert loc_a.table.col == 1

    def test_grid_before_missing_val_falls_back_to_zero(self, simple_docx, tmp_path):
        """Bare ``<w:gridBefore/>`` (no ``w:val``) → offset 0."""
        dest = tmp_path / "gb-empty.docx"
        replace_document_xml(simple_docx, dest, self._doc_with_grid_before(""))
        with Document.open(dest) as doc:
            loc_a = doc.get_paragraph_location(_ref_for_text(doc, "First cell"))
            assert loc_a.table is not None
            assert loc_a.table.col == 1


class TestParagraphLocationSdtWrappers:
    """Word templates often wrap rows and cells in
    ``<w:sdt><w:sdtContent>...</w:sdtContent></w:sdt>`` (structured document
    tags). The location walker must treat these wrappers transparently
    rather than raising ``ValueError`` on the previous "direct child" walk.
    """

    @staticmethod
    def _doc_with_sdt_row() -> str:
        """Two-row table; the second row is wrapped in an ``<w:sdt>``."""
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<w:document {_W_NS}>"
            "<w:body>"
            "<w:tbl>"
            "<w:tr><w:tc><w:p><w:r><w:t>R1</w:t></w:r></w:p></w:tc></w:tr>"
            # Row 2 wrapped in sdt/sdtContent (legal under CT_SdtRow):
            "<w:sdt><w:sdtContent>"
            "<w:tr><w:tc><w:p><w:r><w:t>R2-sdt</w:t></w:r></w:p></w:tc></w:tr>"
            "</w:sdtContent></w:sdt>"
            "</w:tbl>"
            "</w:body>"
            "</w:document>"
        )

    @staticmethod
    def _doc_with_sdt_cell() -> str:
        """Single row whose second cell is wrapped in an ``<w:sdt>``."""
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<w:document {_W_NS}>"
            "<w:body>"
            "<w:tbl><w:tr>"
            "<w:tc><w:p><w:r><w:t>C1</w:t></w:r></w:p></w:tc>"
            "<w:sdt><w:sdtContent>"
            "<w:tc><w:p><w:r><w:t>C2-sdt</w:t></w:r></w:p></w:tc>"
            "</w:sdtContent></w:sdt>"
            "<w:tc><w:p><w:r><w:t>C3</w:t></w:r></w:p></w:tc>"
            "</w:tr></w:tbl>"
            "</w:body>"
            "</w:document>"
        )

    def test_sdt_wrapped_row_does_not_raise(self, simple_docx, tmp_path):
        dest = tmp_path / "sdt-row.docx"
        replace_document_xml(simple_docx, dest, self._doc_with_sdt_row())
        with Document.open(dest) as doc:
            loc_1 = doc.get_paragraph_location(_ref_for_text(doc, "R1"))
            loc_2 = doc.get_paragraph_location(_ref_for_text(doc, "R2-sdt"))
            assert loc_1.table == TableCell(index=1, row=1, col=1, depth=1)
            # The SDT-wrapped row is still row 2 of the same table.
            assert loc_2.table == TableCell(index=1, row=2, col=1, depth=1)

    def test_sdt_wrapped_cell_does_not_raise(self, simple_docx, tmp_path):
        dest = tmp_path / "sdt-cell.docx"
        replace_document_xml(simple_docx, dest, self._doc_with_sdt_cell())
        with Document.open(dest) as doc:
            loc_1 = doc.get_paragraph_location(_ref_for_text(doc, "C1"))
            loc_2 = doc.get_paragraph_location(_ref_for_text(doc, "C2-sdt"))
            loc_3 = doc.get_paragraph_location(_ref_for_text(doc, "C3"))
            # SDT is transparent — cells keep contiguous logical columns.
            assert loc_1.table == TableCell(index=1, row=1, col=1, depth=1)
            assert loc_2.table == TableCell(index=1, row=1, col=2, depth=1)
            assert loc_3.table == TableCell(index=1, row=1, col=3, depth=1)


class TestParagraphLocationNestedRowSkip:
    """Outer-table walker must skip rows / cells that belong to nested tables.

    These exercise the ``continue`` branches in ``_row_index_in_table`` and
    ``_logical_col_in_row``. Distinct from the nested-table tests above:
    here the *outer* row count / *outer* col count must be correct in spite
    of an interposed nested table.
    """

    @staticmethod
    def _doc_with_nested_between_outer_rows() -> str:
        """Outer table with 2 rows; the first row's cell contains a nested table.

        Looking up a paragraph in row 2 of the outer table forces the walker
        to visit (and skip) the nested table's row before reaching row 2.
        """
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<w:document {_W_NS}>"
            "<w:body>"
            "<w:tbl>"
            # Row 1 — its only cell contains a nested table.
            "<w:tr><w:tc>"
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            "</w:tc></w:tr>"
            # Row 2 — the target.
            "<w:tr><w:tc><w:p><w:r><w:t>Outer R2</w:t></w:r></w:p></w:tc></w:tr>"
            "</w:tbl>"
            "</w:body>"
            "</w:document>"
        )

    @staticmethod
    def _doc_with_nested_inside_first_cell() -> str:
        """Single outer row whose first cell contains a nested table; second
        cell is the target. Forces the column walker to skip a nested ``w:tc``.
        """
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<w:document {_W_NS}>"
            "<w:body>"
            "<w:tbl><w:tr>"
            # Cell 1 contains an entire nested table.
            "<w:tc>"
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            "</w:tc>"
            # Cell 2 — target.
            "<w:tc><w:p><w:r><w:t>Outer C2</w:t></w:r></w:p></w:tc>"
            "</w:tr></w:tbl>"
            "</w:body>"
            "</w:document>"
        )

    def test_row_walker_skips_nested_rows(self, simple_docx, tmp_path):
        dest = tmp_path / "nested-rows.docx"
        replace_document_xml(simple_docx, dest, self._doc_with_nested_between_outer_rows())
        with Document.open(dest) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Outer R2"))
            # Without the continue branch, the nested-table tr would be
            # counted and "Outer R2" would land at row=3 (or raise).
            assert loc.table == TableCell(index=1, row=2, col=1, depth=1)

    def test_col_walker_skips_nested_cells(self, simple_docx, tmp_path):
        dest = tmp_path / "nested-cells.docx"
        replace_document_xml(simple_docx, dest, self._doc_with_nested_inside_first_cell())
        with Document.open(dest) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Outer C2"))
            # Without the continue branch, the nested-table tc would be
            # counted and "Outer C2" would land at col=3.
            assert loc.table == TableCell(index=1, row=1, col=2, depth=1)


class TestParagraphLocationDefensiveFallbacks:
    """The remaining defensive paths in ``_initial_grid_offset`` and
    ``_compute_paragraph_location`` — exercised on malformed-ish XML that
    minidom still parses (no schema validation).
    """

    @staticmethod
    def test_tr_pr_without_grid_before_falls_back_to_zero(simple_docx, tmp_path):
        """``<w:trPr>`` exists with other props but no ``<w:gridBefore>`` → offset 0."""
        body = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<w:document {_W_NS}>"
            "<w:body>"
            "<w:tbl>"
            "<w:tr><w:trPr><w:cantSplit/></w:trPr>"
            "<w:tc><w:p><w:r><w:t>NoGridBefore</w:t></w:r></w:p></w:tc>"
            "</w:tr>"
            "</w:tbl>"
            "</w:body>"
            "</w:document>"
        )
        dest = tmp_path / "tr-pr-no-gb.docx"
        replace_document_xml(simple_docx, dest, body)
        with Document.open(dest) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "NoGridBefore"))
            assert loc.table is not None
            assert loc.table.col == 1

    @staticmethod
    def test_tc_outside_tr_returns_table_none(simple_docx, tmp_path):
        """A ``<w:tc>`` not nested under a ``<w:tr>`` is treated as body content."""
        body = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<w:document {_W_NS}>"
            "<w:body>"
            # Orphan: tc directly under tbl, skipping the tr layer.
            "<w:tbl>"
            "<w:tc><w:p><w:r><w:t>Orphan tc</w:t></w:r></w:p></w:tc>"
            "</w:tbl>"
            "</w:body>"
            "</w:document>"
        )
        dest = tmp_path / "orphan-tc.docx"
        replace_document_xml(simple_docx, dest, body)
        with Document.open(dest) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Orphan tc"))
            # Malformed structure — falls back to body content, no raise.
            assert loc.in_table is False
            assert loc.table is None


class TestParagraphLocationLongPreview:
    """Cover the >80-char preview truncation in the ``HashMismatchError`` path."""

    @staticmethod
    def test_stale_ref_on_long_paragraph_truncates_preview(simple_docx, tmp_path):
        long_text = "x" * 200
        body = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<w:document {_W_NS}>"
            "<w:body>"
            f"<w:p><w:r><w:t>{long_text}</w:t></w:r></w:p>"
            "</w:body>"
            "</w:document>"
        )
        dest = tmp_path / "long-para.docx"
        replace_document_xml(simple_docx, dest, body)
        with Document.open(dest) as doc:
            ref = _ref_for_text(doc, "xxxxx")
            index_part, _ = ref.split("#")
            stale_ref = f"{index_part}#0000"
            with pytest.raises(HashMismatchError) as exc:
                doc.get_paragraph_location(stale_ref)
            # Preview should be truncated to 80 chars + "..."
            assert exc.value.paragraph_preview.endswith("...")
            assert len(exc.value.paragraph_preview) == 83  # 80 + "..."


class TestListParagraphLocations:
    """:meth:`Document.list_paragraph_locations` — the batch accessor.

    The defining contract: each ``(ref, loc)`` entry must be byte-identical
    to what ``get_paragraph_location(ref)`` returns per-call. The batch path
    only swaps the whole-DOM table rescan for a prebuilt index, so the two
    must never disagree.
    """

    def test_returns_one_entry_per_paragraph(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            entries = doc.list_paragraph_locations()
            assert len(entries) == len(doc.list_paragraphs())
            assert len(entries) == 11  # P1-P11 from the fixture

    def test_refs_match_list_paragraphs_tokens(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            batch_refs = [ref for ref, _ in doc.list_paragraph_locations()]
            expected_refs = [entry.split("|")[0] for entry in doc.list_paragraphs()]
            assert batch_refs == expected_refs

    def test_equivalence_with_per_call_accessor(self, gridspan_docx):
        """Every batch entry equals the per-call ``get_paragraph_location``.

        Covers body (P1/P8/P11), gridSpan row (P2-P4), plain row (P5-P7),
        and nested table (P9/P10) — the full fixture.
        """
        with Document.open(gridspan_docx) as doc:
            for ref, loc in doc.list_paragraph_locations():
                assert loc == doc.get_paragraph_location(ref)

    def test_returns_paragraph_location_dataclass(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            for ref, loc in doc.list_paragraph_locations():
                assert isinstance(ref, str)
                assert isinstance(loc, ParagraphLocation)

    def test_body_paragraphs_report_no_table(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            by_ref = dict(doc.list_paragraph_locations())
            body_ref = _ref_for_text(doc, "Body paragraph 1")
            assert by_ref[body_ref].in_table is False
            assert by_ref[body_ref].table is None

    def test_nested_table_coordinates(self, gridspan_docx):
        with Document.open(gridspan_docx) as doc:
            by_ref = dict(doc.list_paragraph_locations())
            inner_ref = _ref_for_text(doc, "Inner cell")
            assert by_ref[inner_ref].table == TableCell(index=3, row=1, col=1, depth=2)

    def test_table_free_document_all_body(self, simple_docx):
        """A document with no tables returns a non-empty all-body list."""
        with Document.open(simple_docx) as doc:
            entries = doc.list_paragraph_locations()
            assert entries  # simple.docx has at least one paragraph
            assert all(loc.table is None for _, loc in entries)
            assert all(loc.in_table is False for _, loc in entries)


class TestParagraphLocationList:
    """``location.list`` reports the raw ``w:pPr/w:numPr`` of the paragraph."""

    def test_numbered_paragraph_reports_list_item(self, numbered_docx):
        with Document.open(numbered_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Numbered top"))
            assert loc.list == ListItem(num_id=5, ilvl=0)

    def test_returns_list_item_dataclass(self, numbered_docx):
        """API contract: ``loc.list`` is a ``ListItem``, not a tuple or dict."""
        with Document.open(numbered_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Numbered top"))
            assert isinstance(loc.list, ListItem)

    def test_plain_paragraph_reports_none(self, numbered_docx):
        with Document.open(numbered_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Plain paragraph"))
            assert loc.list is None

    def test_nested_levels_report_ilvl(self, numbered_docx):
        with Document.open(numbered_docx) as doc:
            loc_1 = doc.get_paragraph_location(_ref_for_text(doc, "Numbered nested"))
            loc_2 = doc.get_paragraph_location(_ref_for_text(doc, "Numbered deeper"))
            assert loc_1.list == ListItem(num_id=5, ilvl=1)
            assert loc_2.list == ListItem(num_id=5, ilvl=2)

    def test_missing_ilvl_defaults_to_zero(self, numbered_docx):
        """``<w:numPr>`` without ``<w:ilvl>`` → level 0 (spec default)."""
        with Document.open(numbered_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Numbered no ilvl"))
            assert loc.list == ListItem(num_id=5, ilvl=0)


class TestParagraphLocationListInTable:
    """Table membership and list membership are independent axes."""

    def test_cell_paragraph_reports_both(self, numbered_docx):
        with Document.open(numbered_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Cell numbered"))
            assert loc.in_table is True
            assert loc.table == TableCell(index=1, row=1, col=1, depth=1)
            assert loc.list == ListItem(num_id=7, ilvl=0)


class TestParagraphLocationListMalformed:
    """Defensive paths in ``_extract_list_item`` — malformed or spec-edge
    ``w:numPr`` blocks must degrade to ``None`` (or ilvl=0), never raise.
    """

    def test_num_pr_without_num_id_reports_none(self, simple_docx, tmp_path):
        p_pr = '<w:pPr><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr>'
        with _open_single_paragraph(simple_docx, tmp_path, p_pr) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Target"))
            assert loc.list is None

    def test_num_id_zero_reports_none(self, simple_docx, tmp_path):
        """``numId=0`` is the spec's "numbering disabled" marker, not a list."""
        with _open_single_paragraph(simple_docx, tmp_path, _num_pr("0", "0")) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Target"))
            assert loc.list is None

    def test_non_integer_num_id_reports_none(self, simple_docx, tmp_path):
        with _open_single_paragraph(simple_docx, tmp_path, _num_pr("abc", "0")) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Target"))
            assert loc.list is None

    def test_num_id_without_val_reports_none(self, simple_docx, tmp_path):
        p_pr = "<w:pPr><w:numPr><w:numId/></w:numPr></w:pPr>"
        with _open_single_paragraph(simple_docx, tmp_path, p_pr) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Target"))
            assert loc.list is None

    def test_non_integer_ilvl_falls_back_to_zero(self, simple_docx, tmp_path):
        """Malformed ``w:ilvl`` doesn't discard the (valid) numId — level 0."""
        with _open_single_paragraph(simple_docx, tmp_path, _num_pr("5", "abc")) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Target"))
            assert loc.list == ListItem(num_id=5, ilvl=0)

    def test_stale_num_pr_inside_p_pr_change_is_ignored(self, simple_docx, tmp_path):
        """``w:pPrChange`` nests a *former* ``w:pPr`` (revision record); a
        ``w:numPr`` found only there must not be reported — direct-child
        walk regression guard against descendant search.
        """
        p_pr = (
            "<w:pPr>"
            '<w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">'
            '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr>'
            "</w:pPrChange>"
            "</w:pPr>"
        )
        with _open_single_paragraph(simple_docx, tmp_path, p_pr) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Target"))
            assert loc.list is None


class TestListParagraphLocationsListInfo:
    """Batch accessor carries the same list info as the per-call accessor."""

    def test_equivalence_with_per_call_accessor(self, numbered_docx):
        with Document.open(numbered_docx) as doc:
            entries = doc.list_paragraph_locations()
            for ref, loc in entries:
                assert loc == doc.get_paragraph_location(ref)
            assert any(loc.list is not None for _, loc in entries)


class TestParagraphStyleField:
    """``location.style`` reports the raw ``w:pPr/w:pStyle`` id."""

    def test_heading_reports_raw_style_id(self, styled_docx):
        with Document.open(styled_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Chapter one"))
            assert loc.style == "Heading1"

    def test_plain_body_reports_none(self, styled_docx):
        with Document.open(styled_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Body under h2"))
            assert loc.style is None

    def test_unknown_style_id_passes_through_raw(self, styled_docx):
        """No styles.xml lookup for the id itself — raw value, no validation."""
        with Document.open(styled_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Unknown style"))
            assert loc.style == "NoSuchStyle"
            assert loc.outline_level is None

    def test_p_style_without_val_reports_none(self, simple_docx, tmp_path):
        with _open_single_paragraph(simple_docx, tmp_path, "<w:pPr><w:pStyle/></w:pPr>") as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Target"))
            assert loc.style is None

    def test_p_style_with_empty_val_reports_none(self, simple_docx, tmp_path):
        with _open_single_paragraph(simple_docx, tmp_path, '<w:pPr><w:pStyle w:val=""/></w:pPr>') as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Target"))
            assert loc.style is None

    def test_stale_p_style_inside_p_pr_change_is_ignored(self, styled_docx):
        """A ``w:pStyle``/``w:outlineLvl`` found only inside a ``w:pPrChange``
        revision record must not be reported — direct-child walk regression
        guard against descendant search.
        """
        with Document.open(styled_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Decoy paragraph"))
            assert loc.style is None
            assert loc.outline_level is None


class TestOutlineLevel:
    """``location.outline_level`` — direct ``w:outlineLvl`` wins, else the
    style definition from ``word/styles.xml`` applies (0-based; ``None`` =
    body text).
    """

    def test_level_from_style_definition(self, styled_docx):
        with Document.open(styled_docx) as doc:
            loc_h1 = doc.get_paragraph_location(_ref_for_text(doc, "Chapter one"))
            loc_h2 = doc.get_paragraph_location(_ref_for_text(doc, "Termination"))
            assert loc_h1.outline_level == 0
            assert loc_h2.outline_level == 1

    def test_direct_outline_without_style(self, styled_docx):
        with Document.open(styled_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Direct outline"))
            assert loc.style is None
            assert loc.outline_level == 3

    def test_plain_body_reports_none(self, styled_docx):
        with Document.open(styled_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Body under h2"))
            assert loc.outline_level is None

    def test_direct_value_overrides_style_definition(self, simple_docx, tmp_path):
        """Heading1 defines level 0, but a direct ``w:outlineLvl`` wins."""
        p_pr = '<w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="4"/></w:pPr>'
        with _open_single_paragraph(simple_docx, tmp_path, p_pr) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Target"))
            assert loc.outline_level == 4

    def test_direct_body_marker_overrides_style(self, simple_docx, tmp_path):
        """A direct ``w:val="9"`` explicitly resets to body text — the
        Heading1 style definition must NOT leak back in as a fallback.
        """
        p_pr = '<w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="9"/></w:pPr>'
        with _open_single_paragraph(simple_docx, tmp_path, p_pr) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Target"))
            assert loc.outline_level is None

    @pytest.mark.parametrize("val", ["9", "12", "-1", "abc", ""])
    def test_invalid_direct_values_report_none(self, simple_docx, tmp_path, val):
        """``9`` (spec body-text marker), out-of-range, and malformed values
        all degrade to ``None``, never raise.
        """
        p_pr = f'<w:pPr><w:outlineLvl w:val="{val}"/></w:pPr>'
        with _open_single_paragraph(simple_docx, tmp_path, p_pr) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Target"))
            assert loc.outline_level is None

    def test_style_with_outline_9_reports_none(self, styled_docx):
        """``TOCHeading`` is based on ``Heading1`` but restates
        ``w:outlineLvl w:val="9"`` — the explicit reset must terminate the
        ``basedOn`` chain, not inherit level 0.
        """
        with Document.open(styled_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Toc heading"))
            assert loc.style == "TOCHeading"
            assert loc.outline_level is None


class TestHeadingPath:
    """``location.heading_path`` — nearest preceding headings, outermost first."""

    def test_preamble_before_any_heading_is_empty(self, styled_docx):
        with Document.open(styled_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Preamble text"))
            assert loc.heading_path == ()

    def test_heading_own_path_excludes_itself(self, styled_docx):
        with Document.open(styled_docx) as doc:
            loc_h1 = doc.get_paragraph_location(_ref_for_text(doc, "Chapter one"))
            loc_h2 = doc.get_paragraph_location(_ref_for_text(doc, "Termination"))
            assert loc_h1.heading_path == ()
            assert loc_h2.heading_path == ("Chapter one",)

    def test_body_under_nested_headings(self, styled_docx):
        """The issue's acceptance case: locating a paragraph "under the
        Termination heading".
        """
        with Document.open(styled_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Body under h2"))
            assert loc.heading_path == ("Chapter one", "Termination")

    def test_table_cell_carries_heading_path(self, styled_docx):
        """Table membership and heading context are independent axes."""
        with Document.open(styled_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Cell under h2"))
            assert loc.in_table is True
            assert loc.heading_path == ("Chapter one", "Termination")

    def test_direct_outline_heading_participates(self, styled_docx):
        """A heading defined only by a direct ``w:outlineLvl`` (no style)
        still nests under shallower headings and extends the path.
        """
        with Document.open(styled_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Unknown style"))
            assert loc.heading_path == ("Chapter one", "Termination", "Direct outline")

    def test_new_h1_pops_deeper_headings(self, styled_docx):
        """A later Heading1 pops every open heading at level >= 0."""
        with Document.open(styled_docx) as doc:
            loc_h1 = doc.get_paragraph_location(_ref_for_text(doc, "Chapter two"))
            loc_body = doc.get_paragraph_location(_ref_for_text(doc, "Body under ch2"))
            assert loc_h1.heading_path == ()
            assert loc_body.heading_path == ("Chapter two",)


class TestStyleOutlineBasedOnChain:
    """``w:basedOn`` chain resolution in the style outline map."""

    _STYLES_XML = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:styles {_W_NS}>"
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="heading 1"/>'
        '<w:pPr><w:outlineLvl w:val="0"/></w:pPr>'
        "</w:style>"
        # Custom style inheriting the outline level from Heading1.
        '<w:style w:type="paragraph" w:styleId="MyHead">'
        '<w:name w:val="My Head"/>'
        '<w:basedOn w:val="Heading1"/>'
        "</w:style>"
        # basedOn cycle — must terminate, not hang.
        '<w:style w:type="paragraph" w:styleId="CycleA"><w:basedOn w:val="CycleB"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="CycleB"><w:basedOn w:val="CycleA"/></w:style>'
        # Non-paragraph style carrying an outlineLvl — must be ignored.
        '<w:style w:type="character" w:styleId="CharStyle">'
        '<w:pPr><w:outlineLvl w:val="4"/></w:pPr>'
        "</w:style>"
        "</w:styles>"
    )

    _DOCUMENT_XML = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_W_NS}>"
        "<w:body>"
        f"<w:p>{_p_style('MyHead')}<w:r><w:t>Custom heading</w:t></w:r></w:p>"
        f"<w:p>{_p_style('CycleA')}<w:r><w:t>Cycle target</w:t></w:r></w:p>"
        f"<w:p>{_p_style('CharStyle')}<w:r><w:t>Char styled</w:t></w:r></w:p>"
        "</w:body>"
        "</w:document>"
    )

    @pytest.fixture
    def based_on_docx(self, simple_docx, tmp_path) -> Path:
        dest = tmp_path / "based-on.docx"
        replace_docx_parts(
            simple_docx,
            dest,
            {"word/document.xml": self._DOCUMENT_XML, "word/styles.xml": self._STYLES_XML},
        )
        return dest

    def test_based_on_chain_inherits_outline_level(self, based_on_docx):
        with Document.open(based_on_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Custom heading"))
            assert loc.style == "MyHead"
            assert loc.outline_level == 0

    def test_inherited_heading_opens_a_path(self, based_on_docx):
        """A style-inherited heading participates in ``heading_path``."""
        with Document.open(based_on_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Cycle target"))
            assert loc.heading_path == ("Custom heading",)

    def test_based_on_cycle_degrades_to_none(self, based_on_docx):
        """A ``basedOn`` cycle must terminate (visited guard), yielding no level."""
        with Document.open(based_on_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Cycle target"))
            assert loc.outline_level is None

    def test_non_paragraph_styles_are_ignored(self, based_on_docx):
        with Document.open(based_on_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Char styled"))
            assert loc.style == "CharStyle"
            assert loc.outline_level is None


class TestMissingStylesPart:
    """A document without ``word/styles.xml`` degrades gracefully."""

    @pytest.fixture
    def no_styles_docx(self, simple_docx, tmp_path) -> Path:
        dest = tmp_path / "no-styles.docx"
        replace_docx_parts(
            simple_docx,
            dest,
            {"word/document.xml": _build_styled_document_xml(), "word/styles.xml": None},
        )
        return dest

    def test_style_outline_falls_back_to_direct_only(self, no_styles_docx):
        with Document.open(no_styles_docx) as doc:
            loc_styled = doc.get_paragraph_location(_ref_for_text(doc, "Chapter one"))
            loc_direct = doc.get_paragraph_location(_ref_for_text(doc, "Direct outline"))
            # Raw style id still reported; no styles.xml → no level to inherit.
            assert loc_styled.style == "Heading1"
            assert loc_styled.outline_level is None
            # Direct w:outlineLvl needs no styles part at all.
            assert loc_direct.outline_level == 3

    def test_batch_accessor_still_works(self, no_styles_docx):
        with Document.open(no_styles_docx) as doc:
            entries = doc.list_paragraph_locations()
            assert entries
            for ref, loc in entries:
                assert loc == doc.get_paragraph_location(ref)


class TestListParagraphLocationsStyleInfo:
    """Batch accessor carries the same style/outline/heading-path info as
    the per-call accessor — including the shared style-outline map and the
    single heading pass.
    """

    def test_equivalence_with_per_call_accessor(self, styled_docx):
        with Document.open(styled_docx) as doc:
            entries = doc.list_paragraph_locations()
            for ref, loc in entries:
                assert loc == doc.get_paragraph_location(ref)
            assert any(loc.outline_level is not None for _, loc in entries)
            assert any(loc.heading_path for _, loc in entries)


class TestParagraphLocationSection:
    """``location.section`` — 1-based section index from ``w:sectPr`` walks."""

    def test_single_section_document_reports_one_everywhere(self, gridspan_docx):
        """No paragraph-level ``w:sectPr`` → body, table, and nested-table
        paragraphs all report section 1.
        """
        with Document.open(gridspan_docx) as doc:
            for _, loc in doc.list_paragraph_locations():
                assert loc.section == 1

    def test_sections_split_at_sect_pr_boundaries(self, sectioned_docx):
        expected = {
            "Intro one": 1,
            "Close one": 1,
            "Open two": 2,
            "Cell in two": 2,
            "Close two": 2,
            "Tail three": 3,
        }
        with Document.open(sectioned_docx) as doc:
            for snippet, section in expected.items():
                loc = doc.get_paragraph_location(_ref_for_text(doc, snippet))
                assert loc.section == section, snippet

    def test_carrying_paragraph_belongs_to_closed_section(self, sectioned_docx):
        """A paragraph whose ``w:pPr`` carries the ``w:sectPr`` belongs to
        the section it closes, not the next one.
        """
        with Document.open(sectioned_docx) as doc:
            assert doc.get_paragraph_location(_ref_for_text(doc, "Close one")).section == 1
            assert doc.get_paragraph_location(_ref_for_text(doc, "Close two")).section == 2

    def test_table_paragraph_gets_enclosing_section(self, sectioned_docx):
        with Document.open(sectioned_docx) as doc:
            loc = doc.get_paragraph_location(_ref_for_text(doc, "Cell in two"))
            assert loc.in_table is True
            assert loc.section == 2

    def test_stale_sect_pr_inside_p_pr_change_is_ignored(self, simple_docx, tmp_path):
        """A ``w:sectPr`` found only inside a ``w:pPrChange`` revision record
        must not close a section — direct-child walk regression guard.
        """
        body = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<w:document {_W_NS}>"
            "<w:body>"
            "<w:p><w:pPr>"
            '<w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">'
            "<w:pPr><w:sectPr/></w:pPr>"
            "</w:pPrChange>"
            "</w:pPr><w:r><w:t>Decoy paragraph</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>Following paragraph</w:t></w:r></w:p>"
            "</w:body>"
            "</w:document>"
        )
        dest = tmp_path / "sectpr-decoy.docx"
        replace_document_xml(simple_docx, dest, body)
        with Document.open(dest) as doc:
            assert doc.get_paragraph_location(_ref_for_text(doc, "Decoy paragraph")).section == 1
            assert doc.get_paragraph_location(_ref_for_text(doc, "Following paragraph")).section == 1

    def test_dataclass_defaults_to_section_one(self):
        assert ParagraphLocation(table=None).section == 1


class TestListParagraphLocationsSectionInfo:
    """Batch accessor carries the same section info as the per-call accessor."""

    def test_equivalence_with_per_call_accessor(self, sectioned_docx):
        with Document.open(sectioned_docx) as doc:
            entries = doc.list_paragraph_locations()
            for ref, loc in entries:
                assert loc == doc.get_paragraph_location(ref)
            assert [loc.section for _, loc in entries] == [1, 1, 2, 2, 2, 3]
