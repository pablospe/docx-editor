"""Moves and paragraph-property changes as first-class revisions (ISSUES.md #68).

``w:moveFrom``/``w:moveTo`` and ``w:pPrChange`` are listed by
``list_revisions()`` (types ``move_from``/``move_to``/``property_change``) and
resolved by ``accept_revision``/``reject_revision`` — so ``accept_all``,
``reject_all``, the author filter and the inferred changeset all cover them
for free. A move's four range marks are scaffolding: never listed, never
counted as unhandled, swept once no pending move content remains between them.

The fixtures mirror the two corpus files that carry these types exactly
(``benchmarks/corpus/README.md``):

- ``locore_TC-table-DnD-move.docx`` — Word's drag-and-drop of a whole table:
  a paragraph-mark ``w:moveFrom``/``w:moveTo`` marker under ``w:pPr/w:rPr`` in
  every cell, plain ``w:t`` (not ``w:delText``) inside ``w:moveFrom``, and the
  From-range's ``w:moveFromRangeEnd`` at *body* level, outside any paragraph.
- ``locore_UnknownStyleInRedline.docx`` — two ``w:pPrChange`` sharing
  ``w:id="0"``, one self-closing (LibreOffice's "previously no properties"),
  one recording a style id that ``styles.xml`` does not define, both written
  *before* ``w:rPr`` (schema order puts the record last).

``tests/test_corpus_fixtures.py`` runs the same assertions against the real
files when the corpus is built locally.
"""

import warnings
import zipfile
from pathlib import Path

import defusedxml.minidom
import pytest
from conftest import find_ref, replace_document_xml

from docx_editor import Document, UnhandledRevisionWarning
from docx_editor.track_changes import MOVE_RANGE_TAGS, RevisionManager, count_revision_elements

_W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_ANN = 'w:author="Ann" w:date="2026-01-29T16:55:00Z"'
_BOB = 'w:author="Bob" w:date="2026-01-29T16:56:00Z"'
_LN = 'w:author="László Németh" w:date="2019-10-16T12:25:00Z"'
_KG = 'w:author="Kelemen Gábor 2" w:date="2018-06-15T09:10:00Z"'

_GRID = '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'


def _document(body: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_W_NS}><w:body>{body}</w:body></w:document>"
    )


def _cell(text: str, mark_id: int, content_id: int, kind: str, range_start: tuple[int, str] | None = None) -> str:
    """One table cell of the DnD shape: paragraph-mark marker, optional range
    start, then the content wrapper around a run Word stamped with rsidDel."""
    start = f'<w:move{kind}RangeStart w:id="{range_start[0]}" w:name="{range_start[1]}" {_LN}/>' if range_start else ""
    return (
        f'<w:tc><w:p><w:pPr><w:rPr><w:move{kind} w:id="{mark_id}" {_LN}/></w:rPr></w:pPr>'
        f"{start}"
        f'<w:move{kind} w:id="{content_id}" {_LN}><w:r w:rsidDel="00AB1234"><w:t>{text}</w:t></w:r></w:move{kind}>'
        "</w:p></w:tc>"
    )


# The corpus shape, id for id: marks 0/3/5/7 and 9/12/14/16, content 2/4/6/8
# and 11/13/15/17, ranges 1 (From) and 10 (To) paired by w:name.
TABLE_MOVE = _document(
    f"<w:tbl>{_GRID}"
    "<w:tr>" + _cell("A1", 0, 2, "From", (1, "move77857376")) + _cell("B1", 3, 4, "From") + "</w:tr>"
    "<w:tr>" + _cell("A2", 5, 6, "From") + _cell("B2", 7, 8, "From") + "</w:tr>"
    "</w:tbl>"
    '<w:moveFromRangeEnd w:id="1"/>'
    "<w:p><w:r><w:t>Text</w:t></w:r></w:p>"
    f"<w:tbl>{_GRID}"
    "<w:tr>" + _cell("A1", 9, 11, "To", (10, "move77857376")) + _cell("B1", 12, 13, "To") + "</w:tr>"
    "<w:tr>" + _cell("A2", 14, 15, "To") + _cell("B2", 16, 17, "To") + "</w:tr>"
    "</w:tbl>"
    '<w:p><w:bookmarkStart w:id="18" w:name="_GoBack"/><w:bookmarkEnd w:id="18"/><w:moveToRangeEnd w:id="10"/></w:p>'
)
TABLE_MOVE_IDS = [0, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17]
TABLE_VISIBLE_MOVED = "\n\n\n\nText\nA1\nB1\nA2\nB2\n"  # moved-away cells are empty
TABLE_VISIBLE_ORIGINAL = "A1\nB1\nA2\nB2\nText\n\n\n\n\n"


def _inline_move(author: str, ids: range, name: str, text: str, *, end_offset: int = 0) -> str:
    """An inline move in the hand-authored ``w:delText`` form, two paragraphs.

    As Word writes it, a range's Start and End share one ``w:id``
    (ECMA-376 §17.13.5: the End's id "shall match" its Start's); ``end_offset``
    breaks that pairing to model a producer that does not.
    """
    s, f, s2, t = ids
    return (
        "<w:p>"
        f'<w:moveFromRangeStart w:id="{s}" {author} w:name="{name}"/>'
        f'<w:moveFrom w:id="{f}" {author}><w:r><w:delText>{text}</w:delText></w:r></w:moveFrom>'
        f'<w:moveFromRangeEnd w:id="{s + end_offset}"/>'
        "<w:r><w:t>tail</w:t></w:r>"
        "</w:p>"
        "<w:p>"
        f'<w:moveToRangeStart w:id="{s2}" {author} w:name="{name}"/>'
        f'<w:moveTo w:id="{t}" {author}><w:r><w:t>{text}</w:t></w:r></w:moveTo>'
        f'<w:moveToRangeEnd w:id="{s2 + end_offset}"/>'
        "</w:p>"
    )


INLINE_MOVE = _document(_inline_move(_ANN, range(10, 14), "move1", "relocated clause"))
TWO_AUTHORS = _document(
    _inline_move(_ANN, range(10, 14), "move1", "relocated clause")
    + _inline_move(_BOB, range(20, 24), "move2", "second clause")
)
# Ann's range marks carry mismatched ids (End 110/112, Start 10/12): they
# cannot be paired, so they fall under the "unpaired" rule.
TWO_AUTHORS_MISMATCHED = _document(
    _inline_move(_ANN, range(10, 14), "move1", "relocated clause", end_offset=100)
    + _inline_move(_BOB, range(20, 24), "move2", "second clause")
)

# Damaged files: one half without the other, or marks without their twins.
LONE_MOVE_FROM = _document(
    f'<w:p><w:moveFrom w:id="11" {_ANN}><w:r><w:delText>gone</w:delText></w:r></w:moveFrom>'
    '<w:r><w:t xml:space="preserve"> stays</w:t></w:r></w:p>'
)
LONE_MOVE_TO = _document(
    f'<w:p><w:moveTo w:id="14" {_ANN}><w:r><w:t>arrived</w:t></w:r></w:moveTo>'
    '<w:r><w:t xml:space="preserve"> here</w:t></w:r></w:p>'
)
START_WITHOUT_END = _document(
    f'<w:p><w:moveFromRangeStart w:id="10" {_ANN} w:name="m"/>'
    f'<w:moveFrom w:id="11" {_ANN}><w:r><w:delText>x</w:delText></w:r></w:moveFrom></w:p>'
    f'<w:p><w:moveToRangeStart w:id="13" {_ANN} w:name="m"/>'
    f'<w:moveTo w:id="14" {_ANN}><w:r><w:t>x</w:t></w:r></w:moveTo></w:p>'
)
END_WITHOUT_START = _document(
    f'<w:p><w:moveFrom w:id="11" {_ANN}><w:r><w:delText>x</w:delText></w:r></w:moveFrom>'
    '<w:moveFromRangeEnd w:id="12"/></w:p>'
    f'<w:p><w:moveTo w:id="14" {_ANN}><w:r><w:t>x</w:t></w:r></w:moveTo>'
    '<w:moveToRangeEnd w:id="15"/></w:p>'
)
STRAY_MARKS_ONLY = _document(
    "<w:p>"
    f'<w:moveFromRangeStart w:id="10" {_ANN} w:name="m"/><w:moveFromRangeEnd w:id="12"/>'
    f'<w:moveToRangeStart w:id="13" {_ANN} w:name="m"/><w:moveToRangeEnd w:id="15"/>'
    "<w:r><w:t>Nothing moved</w:t></w:r>"
    "</w:p>"
)


# A foreign deletion nested inside the *source* half — or inside a plain
# deletion: rejecting the host must leave the nested, still-pending w:del alone.
def _host_with_nested_del(host_tag: str) -> str:
    return _document(
        "<w:p>"
        f'<{host_tag} w:id="11" {_ANN}><w:r><w:t xml:space="preserve">kept </w:t></w:r>'
        f'<w:del w:id="20" {_BOB}><w:r w:rsidDel="00AB1234"><w:delText>gone</w:delText></w:r></w:del></{host_tag}>'
        "</w:p>"
    )


# Range marks a producer wrote badly: two From-Starts sharing an id, and an
# id-less To-Start, each still bracketing pending content.
COLLIDING_RANGE_IDS = _document(
    "<w:p>"
    f'<w:moveFromRangeStart w:id="10" {_ANN} w:name="m1"/>'
    f'<w:moveFrom w:id="11" {_ANN}><w:r><w:delText>one</w:delText></w:r></w:moveFrom>'
    f'<w:moveFromRangeStart w:id="10" {_ANN} w:name="m2"/>'
    f'<w:moveFrom w:id="12" {_ANN}><w:r><w:delText>two</w:delText></w:r></w:moveFrom>'
    '<w:moveFromRangeEnd w:id="10"/>'
    "</w:p>"
    "<w:p>"
    f'<w:moveToRangeStart {_ANN} w:name="m1"/>'
    f'<w:moveTo w:id="14" {_ANN}><w:r><w:t>one</w:t></w:r></w:moveTo>'
    '<w:moveToRangeEnd w:id="13"/>'
    f'<w:moveTo w:id="15" {_ANN}><w:r><w:t>two</w:t></w:r></w:moveTo>'
    "</w:p>"
)


# A foreign deletion nested inside the destination half of a move.
DEL_INSIDE_MOVE_TO = _document(
    "<w:p>"
    f'<w:moveToRangeStart w:id="13" {_ANN} w:name="m"/>'
    f'<w:moveTo w:id="14" {_ANN}><w:r><w:t xml:space="preserve">relocated </w:t></w:r>'
    f'<w:del w:id="20" {_BOB}><w:r><w:delText>clause</w:delText></w:r></w:del></w:moveTo>'
    '<w:moveToRangeEnd w:id="15"/>'
    "</w:p>"
)

# Paragraph-property changes. Both corpus forms carry the record *before*
# w:rPr, the LibreOffice quirk; schema order puts w:pPrChange last.
PPR_RECORDED = _document(
    '<w:p><w:pPr><w:pStyle w:val="Cmsor3"/>'
    f'<w:pPrChange w:id="5" {_KG}><w:pPr><w:pStyle w:val="UnknownStyle"/></w:pPr></w:pPrChange>'
    "<w:rPr/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>"
)
PPR_SELF_CLOSING = _document(
    '<w:p><w:pPr><w:pStyle w:val="Cmsor1"/><w:spacing w:before="240"/>'
    f'<w:pPrChange w:id="6" {_KG}/>'
    "<w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>"
)
PPR_DUPLICATE_IDS = _document(
    '<w:p><w:pPr><w:pStyle w:val="Cmsor1"/><w:spacing w:before="240"/>'
    f'<w:pPrChange w:id="0" {_KG}/><w:rPr/></w:pPr><w:r><w:t>One</w:t></w:r></w:p>'
    "<w:p><w:r><w:t>Two</w:t></w:r></w:p>"
    '<w:p><w:pPr><w:pStyle w:val="Cmsor3"/>'
    f'<w:pPrChange w:id="0" {_KG}><w:pPr><w:pStyle w:val="UnknownStyle"/></w:pPr></w:pPrChange>'
    "<w:rPr/></w:pPr><w:r><w:t>Three</w:t></w:r></w:p>"
)
PPR_NUMPR = _document(
    '<w:p><w:pPr><w:pStyle w:val="Normal"/>'
    f'<w:pPrChange w:id="7" {_ANN}><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
    '<w:jc w:val="center"/></w:pPr></w:pPrChange>'
    "</w:pPr><w:r><w:t>Was a list item</w:t></w:r></w:p>"
)
PPR_WITH_SECTPR = _document(
    '<w:p><w:pPr><w:jc w:val="right"/><w:rPr/><w:sectPr><w:type w:val="nextPage"/></w:sectPr>'
    f'<w:pPrChange w:id="8" {_ANN}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>'
    "</w:pPr><w:r><w:t>Last paragraph of a section</w:t></w:r></w:p>"
    "<w:p><w:r><w:t>Next section</w:t></w:r></w:p>"
)
# Schema-invalid: CT_PPr is the only legal parent, but the record sits
# directly under w:p, so there is no live w:pPr to restore the properties into.
PPR_MISPLACED_RECORD = _document(
    '<w:p><w:pPr><w:pStyle w:val="Cmsor1"/></w:pPr>'
    f'<w:pPrChange w:id="9" {_ANN}><w:pPr><w:pStyle w:val="UnknownStyle"/></w:pPr></w:pPrChange>'
    "<w:r><w:t>Misplaced record</w:t></w:r></w:p>"
)


@pytest.fixture
def make_docx(simple_docx, tmp_path):
    def _make(body_xml: str, name: str = "fixture") -> Path:
        dest = tmp_path / f"{name}.docx"
        replace_document_xml(simple_docx, dest, body_xml)
        return dest

    return _make


@pytest.fixture(autouse=True)
def _no_unhandled_warnings():
    """Nothing in these fixtures is unhandled: a warning here is a failure."""
    with warnings.catch_warnings():
        warnings.simplefilter("error", UnhandledRevisionWarning)
        yield


def _open(path: Path) -> Document:
    return Document.open(path, author="Tester", force_recreate=True)


def _census(doc: Document) -> dict[str, int]:
    """Revision elements still in the live DOM, by tag."""
    return count_revision_elements(doc._revision_manager.editor.dom).by_tag


def _saved_census(path: Path) -> dict[str, int]:
    """Revision elements in a saved file's word/document.xml, by tag."""
    with zipfile.ZipFile(path) as z:
        dom = defusedxml.minidom.parseString(z.read("word/document.xml"))
    return count_revision_elements(dom).by_tag


def _document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


# --------------------------------------------------------------------------
# The table drag-and-drop shape
# --------------------------------------------------------------------------


class TestTableMove:
    def test_listing_shows_both_halves_and_no_range_marks(self, make_docx):
        with _open(make_docx(TABLE_MOVE)) as doc:
            revs = doc.list_revisions()
            assert [r.id for r in revs] == TABLE_MOVE_IDS
            assert [r.type for r in revs] == ["move_from"] * 8 + ["move_to"] * 8
            # Paragraph-mark markers carry no text; content wrappers carry the cell's.
            assert [r.text for r in revs] == ["", "A1", "", "B1", "", "A2", "", "B2"] * 2
            assert {r.author for r in revs} == {"László Németh"}
            # The whole move is one inferred changeset: one call resolves it as a unit.
            assert len({r.changeset_id for r in revs}) == 1
            assert revs[0].changeset_id is not None
            # Occurrence counts in the view where the text lives: original for
            # move_from, accepted for move_to; None for the empty markers.
            by_id = {r.id: r for r in revs}
            assert by_id[2].occurrence == 0 and by_id[11].occurrence == 0
            assert by_id[0].occurrence is None and by_id[9].occurrence is None
            assert by_id[11].paragraph_ref is not None
            assert repr(by_id[2]).startswith("Revision(moveFrom 2 @P1#")
            assert repr(by_id[11]).startswith("Revision(moveTo 11 @P6#")

    def test_visible_text_excludes_moved_away_cells_before_resolution(self, make_docx):
        """Word writes plain w:t inside w:moveFrom; the accepted view must
        exclude it, or the moved text is counted twice."""
        with _open(make_docx(TABLE_MOVE)) as doc:
            assert doc.get_visible_text() == TABLE_VISIBLE_MOVED

    def test_markup_view_renders_both_halves(self, make_docx):
        with _open(make_docx(TABLE_MOVE)) as doc:
            lines = doc.get_markup_text().split("\n")
        assert lines[0] == "[moveFrom#0:László Németh][/moveFrom][moveFrom#2:László Németh]A1[/moveFrom]"
        assert lines[5] == "[moveTo#9:László Németh][/moveTo][moveTo#11:László Németh]A1[/moveTo]"

    def test_accept_all_leaves_the_table_at_its_destination_only(self, make_docx, tmp_path):
        out = tmp_path / "accepted.docx"
        with _open(make_docx(TABLE_MOVE)) as doc:
            result = doc.accept_all()
            assert result == 16
            assert result.unhandled == 0
            assert doc.get_visible_text() == TABLE_VISIBLE_MOVED
            assert _census(doc) == {}  # content wrappers, markers and range marks all gone
            assert doc.list_revisions() == []
            doc.save(out)
        assert _saved_census(out) == {}
        with _open(out) as reopened:
            assert reopened.list_revisions() == []
            assert reopened.list_unhandled_revisions() == []

    def test_reject_all_puts_the_table_back_at_its_source_only(self, make_docx, tmp_path):
        out = tmp_path / "rejected.docx"
        with _open(make_docx(TABLE_MOVE)) as doc:
            result = doc.reject_all()
            assert result == 16
            assert result.unhandled == 0
            assert doc.get_visible_text() == TABLE_VISIBLE_ORIGINAL
            assert _census(doc) == {}
            doc.save(out)
        xml = _document_xml(out)
        assert _saved_census(out) == {}
        # A rejected source half restores runs exactly as a rejected deletion does.
        assert "w:rsidDel" not in xml
        assert xml.count('<w:r w:rsidR="00AB1234"><w:t>') == 4

    def test_changeset_resolves_the_move_as_a_unit(self, make_docx):
        with _open(make_docx(TABLE_MOVE)) as doc:
            changeset_id = doc.list_revisions()[0].changeset_id
            assert changeset_id is not None
            assert doc.accept_changeset(changeset_id) == 16
            assert _census(doc) == {}
            assert doc.get_visible_text() == TABLE_VISIBLE_MOVED

    def test_range_marks_are_swept_only_when_their_content_is_gone(self, make_docx):
        with _open(make_docx(TABLE_MOVE)) as doc:
            assert doc.accept_revision(2)
            # Sibling content (ids 3..8) still lies between Start 1 and End 1.
            assert _census(doc)["w:moveFromRangeStart"] == 1
            assert _census(doc)["w:moveFromRangeEnd"] == 1
            for rev_id in (3, 4, 5, 6, 7, 8):
                assert doc.accept_revision(rev_id)
            census = _census(doc)
            assert "w:moveFromRangeStart" not in census and "w:moveFromRangeEnd" not in census
            # The To-range still has all its content; its marks stay.
            assert census["w:moveToRangeStart"] == 1 and census["w:moveToRangeEnd"] == 1
            # Marker 0 sits before Start 1 in document order, so it is not
            # "between" the pair and is still pending on its own.
            assert [r.id for r in doc.list_revisions()] == [0, 9, 11, 12, 13, 14, 15, 16, 17]

    def test_unresolved_move_survives_save_and_reopen_unchanged(self, make_docx, tmp_path):
        out = tmp_path / "carried.docx"
        with _open(make_docx(TABLE_MOVE)) as doc:
            before = [(r.id, r.type, r.text, r.author) for r in doc.list_revisions()]
            doc.save(out)
        with _open(out) as reopened:
            after = [(r.id, r.type, r.text, r.author) for r in reopened.list_revisions()]
        assert after == before
        assert _saved_census(out) == {
            "w:moveFrom": 8,
            "w:moveTo": 8,
            "w:moveFromRangeStart": 1,
            "w:moveFromRangeEnd": 1,
            "w:moveToRangeStart": 1,
            "w:moveToRangeEnd": 1,
        }


# --------------------------------------------------------------------------
# Inline moves, per-id halves, damaged files, author filter
# --------------------------------------------------------------------------


class TestInlineMove:
    def test_listing_and_markup(self, make_docx):
        with _open(make_docx(INLINE_MOVE)) as doc:
            revs = doc.list_revisions()
            assert [(r.id, r.type, r.text) for r in revs] == [
                (11, "move_from", "relocated clause"),
                (13, "move_to", "relocated clause"),
            ]
            assert revs[0].occurrence == 0  # in the original text of P1
            assert revs[1].occurrence == 0  # in the visible text of P2
            assert doc.get_visible_text() == "tail\nrelocated clause"
            assert doc.get_markup_text() == (
                "[moveFrom#11:Ann]relocated clause[/moveFrom]tail\n[moveTo#13:Ann]relocated clause[/moveTo]"
            )

    def test_accept_all(self, make_docx):
        with _open(make_docx(INLINE_MOVE)) as doc:
            assert doc.accept_all() == 2
            assert doc.get_visible_text() == "tail\nrelocated clause"
            assert _census(doc) == {}

    def test_reject_all_restores_deltext_as_text(self, make_docx):
        with _open(make_docx(INLINE_MOVE)) as doc:
            assert doc.reject_all() == 2
            assert doc.get_visible_text() == "relocated clausetail\n"
            assert _census(doc) == {}

    def test_accepting_one_half_and_rejecting_the_other_is_the_documented_footgun(self, make_docx):
        """Allowed by id, exactly as Word allows it — and with the same result."""
        with _open(make_docx(INLINE_MOVE, "double")) as doc:
            assert doc.accept_revision(13) and doc.reject_revision(11)
            assert doc.get_visible_text() == "relocated clausetail\nrelocated clause"  # duplicated
            assert _census(doc) == {}
        with _open(make_docx(INLINE_MOVE, "lost")) as doc:
            assert doc.accept_revision(11) and doc.reject_revision(13)
            assert doc.get_visible_text() == "tail\n"  # gone
            assert _census(doc) == {}

    def test_author_filter_leaves_the_other_authors_move_and_marks_intact(self, make_docx):
        with _open(make_docx(TWO_AUTHORS)) as doc:
            result = doc.accept_all(author="Ann")
            assert result == 2
            assert result.unhandled == 0
            assert [(r.id, r.author) for r in doc.list_revisions()] == [(21, "Bob"), (23, "Bob")]
            assert _census(doc) == {
                "w:moveFrom": 1,
                "w:moveTo": 1,
                "w:moveFromRangeStart": 1,
                "w:moveFromRangeEnd": 1,
                "w:moveToRangeStart": 1,
                "w:moveToRangeEnd": 1,
            }
            assert doc.accept_all(author="Bob") == 2
            assert _census(doc) == {}

    def test_marks_that_cannot_be_paired_wait_for_their_whole_family(self, make_docx):
        """The documented degrade for a producer whose Start/End ids differ:
        such marks are removed only once no move content of their family
        remains anywhere — never destructive, merely later."""
        with _open(make_docx(TWO_AUTHORS_MISMATCHED)) as doc:
            assert doc.accept_all(author="Ann") == 2
            census = _census(doc)
            assert census["w:moveFromRangeStart"] == 2 and census["w:moveFromRangeEnd"] == 2
            assert doc.accept_all(author="Bob") == 2
            assert _census(doc) == {}

    def test_deletion_nested_inside_a_move_destination(self, make_docx):
        with _open(make_docx(DEL_INSIDE_MOVE_TO)) as doc:
            by_id = {r.id: r for r in doc.list_revisions()}
            assert by_id[14].type == "move_to"
            assert by_id[14].text == "relocated clause"  # full text it moved in
            assert by_id[14].contains_ids == (20,)
            assert by_id[20].type == "deletion"
            assert by_id[20].nested_under == 14
            assert doc.accept_all() == 2
            assert doc.get_visible_text() == "relocated "
            assert _census(doc) == {}


class TestNestedDeletionInsideAHost:
    @pytest.mark.parametrize("host_tag", ["w:moveFrom", "w:del"])
    def test_rejecting_the_host_keeps_the_nested_deletion_pending(self, make_docx, host_tag):
        with _open(make_docx(_host_with_nested_del(host_tag))) as doc:
            assert doc.reject_revision(11)

            xml = doc._revision_manager.editor.dom.toxml()
            assert "<w:delText>gone</w:delText>" in xml  # not converted to w:t
            assert 'w:rsidDel="00AB1234"' in xml  # run attributes untouched
            assert [(r.id, r.type, r.text) for r in doc.list_revisions()] == [(20, "deletion", "gone")]
            assert doc.get_visible_text() == "kept "
            assert doc.reject_revision(20)
            assert doc.get_visible_text() == "kept gone"


class TestBulkResolutionSweepsOnce:
    @pytest.mark.parametrize("method", ["accept_all", "reject_all"])
    def test_accept_all_walks_the_document_a_constant_number_of_times(self, make_docx, monkeypatch, method):
        """The range-mark sweep is a full-document walk; bulk resolution must
        not pay for it once per move half (16 halves here)."""
        calls: list[int] = []
        original = RevisionManager._sweep_move_range_marks

        def counting(self):
            calls.append(1)
            original(self)

        monkeypatch.setattr(RevisionManager, "_sweep_move_range_marks", counting)
        with _open(make_docx(TABLE_MOVE)) as doc:
            assert getattr(doc, method)() == 16
            assert _census(doc) == {}
        # One deferred sweep at the end of the loop, one unconditional sweep in
        # _resolve_all_reporting.
        assert len(calls) <= 2

    def test_move_free_documents_never_sweep(self, make_docx, monkeypatch):
        """No range marks at open: the sweep — a full-document walk — is
        skipped on every path, not merely deferred."""
        calls: list[int] = []
        monkeypatch.setattr(RevisionManager, "_sweep_move_range_marks", lambda self: calls.append(1))
        with _open(make_docx(LONE_MOVE_FROM)) as doc:
            assert not doc._document_editor.holds_move_range_marks
            assert doc.accept_revision(11)
            assert doc.accept_all() == 0
        assert calls == []

    def test_sweeping_the_last_marks_switches_the_sweep_off(self, make_docx, monkeypatch):
        calls: list[int] = []
        original = RevisionManager._sweep_move_range_marks
        monkeypatch.setattr(RevisionManager, "_sweep_move_range_marks", lambda self: (calls.append(1), original(self)))
        with _open(make_docx(INLINE_MOVE)) as doc:
            editor = doc._document_editor
            assert editor.holds_move_range_marks
            assert doc.accept_revision(11)  # From half: its pair is swept
            assert editor.holds_move_range_marks  # the To pair is still there
            assert doc.accept_revision(13)
            assert not editor.holds_move_range_marks
            assert _census(doc) == {}
            n = len(calls)
            assert doc.accept_all() == 0
        assert len(calls) == n  # nothing left to sweep, so no walk

    def test_restoring_a_snapshot_recomputes_the_flag(self, make_docx):
        """A rolled-back batch replaces the DOM wholesale; the flag follows it."""
        with _open(make_docx(LONE_MOVE_FROM)) as doc:
            editor = doc._document_editor
            assert not editor.holds_move_range_marks
            editor._reload_dom_from_bytes(INLINE_MOVE.encode("utf-8"))
            assert editor.holds_move_range_marks
            editor._reload_dom_from_bytes(LONE_MOVE_FROM.encode("utf-8"))
            assert not editor.holds_move_range_marks

    def test_changeset_resolution_sweeps_once_too(self, make_docx, monkeypatch):
        calls: list[int] = []
        original = RevisionManager._sweep_move_range_marks
        monkeypatch.setattr(RevisionManager, "_sweep_move_range_marks", lambda self: (calls.append(1), original(self)))
        with _open(make_docx(TABLE_MOVE)) as doc:
            changeset_id = doc.list_revisions()[0].changeset_id
            assert changeset_id is not None
            assert doc.accept_changeset(changeset_id) == 16
            assert _census(doc) == {}
        assert len(calls) == 1


# Move content that leaves the document inside some *other* resolved
# element: the range marks around it must still be swept.
MOVE_FROM_INSIDE_DEL = _document(
    "<w:p>"
    f'<w:moveFromRangeStart w:id="10" {_ANN} w:name="m"/>'
    f'<w:del w:id="30" {_BOB}><w:moveFrom w:id="11" {_ANN}><w:r><w:delText>x</w:delText></w:r></w:moveFrom></w:del>'
    '<w:moveFromRangeEnd w:id="10"/>'
    "<w:r><w:t>tail</w:t></w:r></w:p>"
)
MOVE_TO_INSIDE_INS = _document(
    "<w:p>"
    f'<w:moveToRangeStart w:id="13" {_ANN} w:name="m"/>'
    f'<w:ins w:id="30" {_BOB}><w:moveTo w:id="14" {_ANN}><w:r><w:t>x</w:t></w:r></w:moveTo></w:ins>'
    '<w:moveToRangeEnd w:id="13"/>'
    "<w:r><w:t>tail</w:t></w:r></w:p>"
)
# A tracked split whose tail paragraph carries a moved paragraph mark: the
# rejoin drops the tail's w:pPr, marker included.
SPLIT_BEFORE_MOVED_MARK = _document(
    f'<w:p><w:pPr><w:rPr><w:ins w:id="30" {_BOB}/></w:rPr></w:pPr><w:r><w:t>head</w:t></w:r></w:p>'
    f'<w:p><w:pPr><w:rPr><w:moveFrom w:id="11" {_ANN}/></w:rPr></w:pPr>'
    f'<w:moveFromRangeStart w:id="10" {_ANN} w:name="m"/><w:moveFromRangeEnd w:id="10"/>'
    "<w:r><w:t>tail</w:t></w:r></w:p>"
)


class TestSweepAfterAnyResolution:
    @pytest.mark.parametrize(
        ("body", "resolve"),
        [
            (MOVE_FROM_INSIDE_DEL, lambda doc: doc.accept_revision(30)),
            (MOVE_TO_INSIDE_INS, lambda doc: doc.reject_revision(30)),
            (SPLIT_BEFORE_MOVED_MARK, lambda doc: doc.reject_revision(30)),
        ],
        ids=["moveFrom-inside-accepted-del", "moveTo-inside-rejected-ins", "rejoin-drops-moved-mark"],
    )
    def test_range_marks_are_swept_when_their_content_leaves_with_a_host(self, make_docx, body, resolve):
        with _open(make_docx(body)) as doc:
            assert resolve(doc)
            assert doc.list_revisions() == []
            assert doc.list_unhandled_revisions() == []
            assert _census(doc) == {}


class TestOwnEditInsideAForeignMoveTo:
    @pytest.mark.xfail(
        strict=True,
        reason="known gap: a foreign w:moveTo is not split around this session's own edits the way a "
        "foreign w:ins is, so rejecting the move carries the own edit away (see reject_revision)",
    )
    def test_own_edit_survives_rejecting_the_foreign_move(self, make_docx):
        with _open(make_docx(INLINE_MOVE)) as doc:
            result = doc.replace("relocated", "moved", paragraph=find_ref(doc, "relocated clause"))
            assert result.group_id is not None

            assert doc.reject_all(author="Ann") == 2

            assert doc.get_visible_text() == "tail\nmoved"
            assert {r.author for r in doc.list_revisions()} == {"Tester"}

    def test_the_gap_is_bounded_to_the_moved_in_text(self, make_docx):
        """What happens today, pinned so the gap cannot widen silently: the
        own edit inside the move is lost with it; edits outside survive."""
        with _open(make_docx(INLINE_MOVE)) as doc:
            doc.replace("relocated", "moved", paragraph=find_ref(doc, "relocated clause"))
            outside = doc.replace("tail", "end", paragraph=find_ref(doc, "tail"))
            assert outside.group_id is not None

            assert doc.reject_all(author="Ann") == 2

            assert doc.get_visible_text() == "relocated clauseend\n"
            assert sorted(r.type for r in doc.list_revisions() if r.author == "Tester") == ["deletion", "insertion"]
            assert _census(doc) == {"w:del": 1, "w:ins": 1}


class TestDamagedFiles:
    @pytest.mark.parametrize(
        "body,accepted,rejected",
        [
            (LONE_MOVE_FROM, " stays", "gone stays"),  # a lone source half is a deletion
            (LONE_MOVE_TO, "arrived here", " here"),  # a lone destination half is an insertion
        ],
    )
    def test_lone_half_behaves_as_what_it_structurally_is(self, make_docx, body, accepted, rejected):
        with _open(make_docx(body, "accept")) as doc:
            assert doc.accept_all() == 1
            assert doc.get_visible_text() == accepted
            assert _census(doc) == {}
        with _open(make_docx(body, "reject")) as doc:
            assert doc.reject_all() == 1
            assert doc.get_visible_text() == rejected
            assert _census(doc) == {}

    @pytest.mark.parametrize("body", [START_WITHOUT_END, END_WITHOUT_START])
    @pytest.mark.parametrize("method", ["accept_all", "reject_all"])
    def test_unpaired_marks_are_swept_once_their_family_is_empty(self, make_docx, body, method):
        with _open(make_docx(body)) as doc:
            assert getattr(doc, method)() == 2
            assert _census(doc) == {}

    def test_colliding_or_missing_range_ids_never_pair_across_ranges(self, make_docx):
        """A duplicated Start id and an id-less Start each fall under the
        unpaired rule: nothing is swept while their family has content, and
        everything goes once it is gone — in one accept_revision call."""
        with _open(make_docx(COLLIDING_RANGE_IDS)) as doc:
            assert doc.accept_revision(11)  # "one" leaves its source
            census = _census(doc)
            assert census["w:moveFromRangeStart"] == 2  # both Starts still there
            assert census["w:moveFromRangeEnd"] == 1
            assert doc.accept_revision(12)  # last From content
            assert doc.accept_revision(14) and doc.accept_revision(15)
            assert _census(doc) == {}
            assert doc.get_visible_text() == "\nonetwo"

    def test_stray_marks_with_no_content_are_swept_and_never_counted(self, make_docx):
        with _open(make_docx(STRAY_MARKS_ONLY)) as doc:
            assert doc.list_revisions() == []
            assert doc.list_unhandled_revisions() == []
            result = doc.accept_all()
            assert result == 0
            assert result.unhandled == 0
            assert _census(doc) == {}
            assert doc.get_visible_text() == "Nothing moved"

    def test_range_marks_are_neither_handled_nor_unhandled_tags(self):
        assert set(MOVE_RANGE_TAGS) == {
            "w:moveFromRangeStart",
            "w:moveFromRangeEnd",
            "w:moveToRangeStart",
            "w:moveToRangeEnd",
        }


# --------------------------------------------------------------------------
# Paragraph-property changes
# --------------------------------------------------------------------------


class TestParagraphPropertyChange:
    def test_listing(self, make_docx):
        with _open(make_docx(PPR_RECORDED)) as doc:
            (rev,) = doc.list_revisions()
            assert rev.type == "property_change"
            assert rev.id == 5
            assert rev.text == ""
            assert rev.occurrence is None
            assert rev.paragraph_ref is not None
            assert rev.author == "Kelemen Gábor 2"
            assert repr(rev).startswith("Revision(pPrChange 5 @P1#")
            assert doc.get_markup_text() == "Heading"  # nothing to render

    def test_accept_keeps_current_properties(self, make_docx, tmp_path):
        out = tmp_path / "accepted.docx"
        with _open(make_docx(PPR_RECORDED)) as doc:
            assert doc.accept_revision(5)
            assert doc.get_paragraph(1).style == "Cmsor3"
            assert _census(doc) == {}
            doc.save(out)
        assert '<w:pPr><w:pStyle w:val="Cmsor3"/><w:rPr/></w:pPr>' in _document_xml(out)

    def test_reject_restores_the_recorded_style_even_if_undefined(self, make_docx, tmp_path):
        """What the file says the paragraph had; Word falls back to Normal."""
        out = tmp_path / "rejected.docx"
        with _open(make_docx(PPR_RECORDED)) as doc:
            assert doc.reject_revision(5)
            assert doc.get_paragraph(1).style == "UnknownStyle"
            assert _census(doc) == {}
            doc.save(out)
        # Schema order holds even though the record preceded w:rPr in the input.
        assert '<w:pPr><w:pStyle w:val="UnknownStyle"/><w:rPr/></w:pPr>' in _document_xml(out)

    def test_self_closing_record_restores_no_properties(self, make_docx, tmp_path):
        out = tmp_path / "cleared.docx"
        with _open(make_docx(PPR_SELF_CLOSING)) as doc:
            assert doc.get_paragraph(1).style == "Cmsor1"
            assert doc.reject_all() == 1
            assert doc.get_paragraph(1).style is None
            doc.save(out)
        xml = _document_xml(out)
        assert "<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>" in xml  # run properties survive
        assert "w:spacing" not in xml

    def test_self_closing_record_accept_keeps_everything(self, make_docx, tmp_path):
        out = tmp_path / "kept.docx"
        with _open(make_docx(PPR_SELF_CLOSING)) as doc:
            assert doc.accept_all() == 1
            doc.save(out)
        assert (
            '<w:pPr><w:pStyle w:val="Cmsor1"/><w:spacing w:before="240"/><w:rPr><w:b/></w:rPr></w:pPr>'
            in _document_xml(out)
        )

    @pytest.mark.parametrize("method", ["accept_all", "reject_all"])
    def test_duplicate_ids_resolve_in_one_pass(self, make_docx, method):
        with _open(make_docx(PPR_DUPLICATE_IDS)) as doc:
            revs = doc.list_revisions()
            assert [r.id for r in revs] == [0, 0]
            assert [r.group_id for r in revs] == [None, None]  # ambiguous id: ungrouped
            result = getattr(doc, method)()
            assert result == 2
            assert result.unhandled == 0
            assert _census(doc) == {}
            styles = [doc.get_paragraph(i).style for i in (1, 2, 3)]
            assert styles == (["Cmsor1", None, "Cmsor3"] if method == "accept_all" else [None, None, "UnknownStyle"])

    def test_duplicate_ids_by_id_resolve_the_first_still_attached(self, make_docx):
        with _open(make_docx(PPR_DUPLICATE_IDS)) as doc:
            assert doc.accept_revision(0)
            assert doc.accept_revision(0)
            assert not doc.accept_revision(0)
            assert _census(doc) == {}

    def test_reject_restores_numbering_and_alignment(self, make_docx, tmp_path):
        out = tmp_path / "numbered.docx"
        with _open(make_docx(PPR_NUMPR)) as doc:
            assert doc.reject_revision(7)
            doc.save(out)
        assert (
            '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:jc w:val="center"/></w:pPr>'
            in _document_xml(out)
        )

    def test_reject_preserves_the_section_break(self, make_docx, tmp_path):
        out = tmp_path / "section.docx"
        with _open(make_docx(PPR_WITH_SECTPR)) as doc:
            assert doc.reject_revision(8)
            doc.save(out)
        assert (
            '<w:pPr><w:jc w:val="left"/><w:rPr/><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr>'
            in _document_xml(out)
        )

    def test_rejecting_a_record_outside_w_pPr_drops_it_and_keeps_the_live_properties(self, make_docx, tmp_path):
        """No w:pPr parent means nothing to restore into: drop the record only."""
        out = tmp_path / "misplaced.docx"
        with _open(make_docx(PPR_MISPLACED_RECORD)) as doc:
            assert doc.reject_revision(9)
            assert doc.get_paragraph(1).style == "Cmsor1"
            assert _census(doc) == {}
            doc.save(out)
        xml = _document_xml(out)
        assert "w:pPrChange" not in xml
        assert '<w:pPr><w:pStyle w:val="Cmsor1"/></w:pPr>' in xml
        assert "UnknownStyle" not in xml

    def test_resolved_document_reopens_clean(self, make_docx, tmp_path):
        out = tmp_path / "clean.docx"
        with _open(make_docx(PPR_DUPLICATE_IDS)) as doc:
            doc.reject_all()
            doc.save(out)
        assert _saved_census(out) == {}
        with _open(out) as reopened:
            assert reopened.list_revisions() == []
            assert reopened.list_unhandled_revisions() == []
            assert reopened.get_paragraph(3).style == "UnknownStyle"


# --------------------------------------------------------------------------
# Change-id allocation
# --------------------------------------------------------------------------


class TestChangeIdAllocation:
    """A new edit's ids never collide with a pending move or property change.

    Word draws every revision mark's id from one counter, and these types are
    now addressed by id: an allocation that reused one would make our own
    edit and the foreign revision the same id, so undoing ours (reject_group)
    would resolve theirs.
    """

    @pytest.mark.parametrize(
        ("body", "needle", "highest"),
        [
            (TABLE_MOVE, "Text", 17),  # moves, markers and range marks 0..17
            (PPR_RECORDED, "Heading", 5),
            (STRAY_MARKS_ONLY, "Nothing moved", 15),  # only range marks
        ],
        ids=["moves", "pPrChange", "range-marks"],
    )
    def test_new_edit_ids_start_past_every_revision_mark(self, make_docx, body, needle, highest):
        with _open(make_docx(body)) as doc:
            before = [r.id for r in doc.list_revisions()]
            doc.replace(needle, "changed", paragraph=find_ref(doc, needle))
            after = [r.id for r in doc.list_revisions()]
            new_ids = [i for i in after if i not in before]
            assert len(after) == len(set(after))  # no duplicated id
            assert new_ids and min(new_ids) > highest

    def test_undoing_our_own_edit_leaves_the_foreign_move_alone(self, make_docx):
        with _open(make_docx(TABLE_MOVE)) as doc:
            result = doc.replace("Text", "Texte", paragraph=find_ref(doc, "Text"))
            assert result.group_id is not None

            doc.reject_group(result.group_id)

            assert [r.id for r in doc.list_revisions()] == TABLE_MOVE_IDS
            assert doc.get_visible_text() == TABLE_VISIBLE_MOVED
