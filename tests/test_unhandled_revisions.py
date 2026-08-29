"""The honesty floor: revision types this library never resolves (ISSUES.md #64).

``accept_all``/``reject_all`` walk ``w:ins``/``w:del`` only, so a Word redline
whose revisions are all format changes or moves used to return ``0`` and read
as "there was nothing to accept" rather than "nothing here could be accepted".
These tests pin that the count is now accompanied by ``.unhandled`` /
``.unhandled_types``, a warning, and a ``list_unhandled_revisions()`` listing.

The corpus makes the case concrete: of 56 real-world files, ``TC-table-DnD-move``
(LibreOffice core's ooxmlexport fixtures) carries 20 move marks and
``UnknownStyleInRedline`` carries 2 ``w:pPrChange`` — both returned a silent 0.

Fixtures are hand-authored ``word/document.xml`` swapped into ``simple.docx``
(``replace_document_xml``), one per revision family, because
``tests/test_data/`` contains no foreign-type revisions at all. Several marks
deliberately omit ``w:author``/``w:date`` (``w:tblGridChange`` and the range
``*End`` marks carry only ``w:id`` in the schema), which is what exercises the
``"Unknown"`` author default and the author-filter exclusion.

Resolving any of these types is out of scope here — that is ISSUES.md #68.
"""

import warnings
from pathlib import Path

import pytest
from conftest import find_ref, replace_document_xml

from docx_editor import (
    Document,
    ResolveResult,
    UnhandledRevision,
    UnhandledRevisionWarning,
)
from docx_editor.track_changes import (
    ALL_REVISION_TAGS,
    CHANGE_RECORD_TAGS,
    HANDLED_REVISION_TAGS,
    UNHANDLED_REVISION_TAGS,
    count_revision_elements,
)

_W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_ANN = 'w:author="Ann" w:date="2026-01-29T16:55:00Z"'
_BOB = 'w:author="Bob" w:date="2026-01-29T16:56:00Z"'


def _document(body: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_W_NS}><w:body>{body}</w:body></w:document>"
    )


# Property changes. The sectPrChange sits in the body's sectPr, outside every
# <w:p> — that is the paragraph_ref=None case.
PROPERTY_CHANGES = _document(
    "<w:p>"
    f'<w:pPr><w:pPrChange w:id="1" {_ANN}><w:pPr/></w:pPrChange></w:pPr>'
    f'<w:r><w:rPr><w:rPrChange w:id="2" {_BOB}><w:rPr/></w:rPrChange></w:rPr>'
    "<w:t>Reformatted paragraph</w:t></w:r>"
    "</w:p>"
    f'<w:sectPr><w:sectPrChange w:id="3" {_ANN}><w:sectPr/></w:sectPrChange></w:sectPr>'
)

# A drag-and-drop move: the moveFrom/moveTo pair plus their four range marks.
# The *End marks carry only w:id, exactly as Word writes them.
MOVES = _document(
    "<w:p>"
    f'<w:moveFromRangeStart w:id="10" {_ANN} w:name="move1"/>'
    f'<w:moveFrom w:id="11" {_ANN}><w:r><w:delText>relocated clause</w:delText></w:r></w:moveFrom>'
    '<w:moveFromRangeEnd w:id="12"/>'
    "<w:r><w:t>tail</w:t></w:r>"
    "</w:p>"
    "<w:p>"
    f'<w:moveToRangeStart w:id="13" {_ANN} w:name="move1"/>'
    f'<w:moveTo w:id="14" {_ANN}><w:r><w:t>relocated clause</w:t></w:r></w:moveTo>'
    '<w:moveToRangeEnd w:id="15"/>'
    "</w:p>"
)

# Table-structure revisions. tblPrChange/tblGridChange/trPrChange/tcPrChange
# and the three cell marks all live outside any <w:p>.
TABLE_REVISIONS = _document(
    "<w:tbl>"
    f'<w:tblPr><w:tblPrChange w:id="20" {_ANN}><w:tblPr/></w:tblPrChange></w:tblPr>'
    '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/>'
    '<w:tblGridChange w:id="21"><w:tblGrid/></w:tblGridChange></w:tblGrid>'
    "<w:tr>"
    f'<w:tblPrEx><w:tblPrExChange w:id="27" {_ANN}><w:tblPrEx/></w:tblPrExChange></w:tblPrEx>'
    f'<w:trPr><w:trPrChange w:id="22" {_ANN}><w:trPr/></w:trPrChange></w:trPr>'
    f'<w:tc><w:tcPr><w:cellIns w:id="23" {_ANN}/>'
    f'<w:tcPrChange w:id="24" {_ANN}><w:tcPr/></w:tcPrChange></w:tcPr>'
    "<w:p><w:r><w:t>Cell one</w:t></w:r></w:p></w:tc>"
    f'<w:tc><w:tcPr><w:cellDel w:id="25" {_ANN}/><w:cellMerge w:id="26" {_BOB}/></w:tcPr>'
    "<w:p><w:r><w:t>Cell two</w:t></w:r></w:p></w:tc>"
    "</w:tr>"
    "</w:tbl>"
    "<w:p><w:r><w:t>After the table</w:t></w:r></w:p>"
)

NUMBERING_CHANGE = _document(
    f'<w:p><w:pPr><w:numberingChange w:id="30" {_ANN} w:original="1"/></w:pPr><w:r><w:t>List item</w:t></w:r></w:p>'
)

CUSTOM_XML_RANGES = _document(
    "<w:p>"
    f'<w:customXmlInsRangeStart w:id="40" {_ANN}/><w:customXmlInsRangeEnd w:id="41"/>'
    f'<w:customXmlDelRangeStart w:id="42" {_ANN}/><w:customXmlDelRangeEnd w:id="43"/>'
    f'<w:customXmlMoveFromRangeStart w:id="44" {_ANN}/><w:customXmlMoveFromRangeEnd w:id="45"/>'
    f'<w:customXmlMoveToRangeStart w:id="46" {_ANN}/><w:customXmlMoveToRangeEnd w:id="47"/>'
    "<w:r><w:t>Custom XML region</w:t></w:r>"
    "</w:p>"
)

# Two real w:ins/w:del revisions plus one foreign mark: the count must report
# the two it resolved and the one it could not.
MIXED = _document(
    "<w:p>"
    f'<w:pPr><w:pPrChange w:id="50" {_ANN}><w:pPr/></w:pPrChange></w:pPr>'
    '<w:r><w:t xml:space="preserve">This </w:t></w:r>'
    f'<w:del w:id="51" {_ANN}><w:r><w:delText xml:space="preserve">old </w:delText></w:r></w:del>'
    f'<w:ins w:id="52" {_ANN}><w:r><w:t xml:space="preserve">new </w:t></w:r></w:ins>'
    "<w:r><w:t>clause.</w:t></w:r>"
    "</w:p>"
)

# A foreign mark nested *inside* a pending insertion. Accepting the insertion
# unwraps it and the w:rPrChange survives; rejecting it deletes the whole
# subtree and the mark goes with it — the asymmetry the "counted after
# resolution" rule in ResolveResult exists to get right.
FOREIGN_INSIDE_INSERTION = _document(
    "<w:p>"
    '<w:r><w:t xml:space="preserve">Kept </w:t></w:r>'
    f'<w:ins w:id="70" {_ANN}>'
    f'<w:r><w:rPr><w:rPrChange w:id="71" {_ANN}><w:rPr/></w:rPrChange></w:rPr>'
    "<w:t>inserted and reformatted</w:t></w:r></w:ins>"
    "</w:p>"
)

# A cell marker recorded *inside* a w:tcPrChange's historical w:tcPr. The
# recorded type is CT_TcPrInner, which — alone among the recorded property
# types — still allows w:cellIns/w:cellDel/w:cellMerge. It describes state that
# is already gone, so it must not be counted as a second pending revision.
MARK_INSIDE_CHANGE_RECORD = _document(
    "<w:tbl>"
    '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
    "<w:tr><w:tc><w:tcPr>"
    f'<w:tcPrChange w:id="80" {_ANN}><w:tcPr><w:cellIns w:id="81" {_ANN}/></w:tcPr></w:tcPrChange>'
    "</w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>"
    "</w:tbl>"
    "<w:p><w:r><w:t>After the table</w:t></w:r></w:p>"
)

# A historical w:del recorded inside a paragraph-mark w:rPrChange. The recorded
# type is CT_ParaRPrOriginal, which opens with EG_ParaRPrTrackChanges, so this
# is schema-legal. The w:del describes the paragraph mark's *previous* state.
HANDLED_MARK_INSIDE_CHANGE_RECORD = _document(
    "<w:p><w:pPr><w:rPr>"
    f'<w:rPrChange w:id="90" {_ANN}><w:rPr><w:del w:id="91" {_ANN}/></w:rPr></w:rPrChange>'
    "</w:rPr></w:pPr><w:r><w:t>Paragraph</w:t></w:r></w:p>"
)

# w:ins/w:del as *structural* markers: a deleted paragraph mark and a pair of
# table-row markers. These are handled tags, so they never enter .unhandled —
# the census's ins_del_contexts is what makes them visible (ISSUES.md #68).
STRUCTURAL_INS_DEL = _document(
    "<w:p>"
    f'<w:pPr><w:rPr><w:del w:id="60" {_ANN}/></w:rPr></w:pPr>'
    "<w:r><w:t>Paragraph whose mark was deleted</w:t></w:r>"
    "</w:p>"
    "<w:p><w:r><w:t>Successor paragraph</w:t></w:r></w:p>"
    "<w:tbl>"
    '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
    f'<w:tr><w:trPr><w:ins w:id="61" {_ANN}/></w:trPr>'
    "<w:tc><w:p><w:r><w:t>Inserted row</w:t></w:r></w:p></w:tc></w:tr>"
    f'<w:tr><w:trPr><w:del w:id="62" {_ANN}/></w:trPr>'
    "<w:tc><w:p><w:r><w:t>Deleted row</w:t></w:r></w:p></w:tc></w:tr>"
    "</w:tbl>"
)

ALL_FIXTURES = {
    "property_changes": PROPERTY_CHANGES,
    "moves": MOVES,
    "table_revisions": TABLE_REVISIONS,
    "numbering_change": NUMBERING_CHANGE,
    "custom_xml_ranges": CUSTOM_XML_RANGES,
    "mixed": MIXED,
    "structural_ins_del": STRUCTURAL_INS_DEL,
    "foreign_inside_insertion": FOREIGN_INSIDE_INSERTION,
    "mark_inside_change_record": MARK_INSIDE_CHANGE_RECORD,
}


@pytest.fixture
def make_docx(simple_docx, tmp_path):
    """Build a .docx from a hand-authored document.xml and open it."""

    def _make(body_xml: str, name: str = "fixture") -> Path:
        dest = tmp_path / f"{name}.docx"
        replace_document_xml(simple_docx, dest, body_xml)
        return dest

    return _make


def _open(path: Path) -> Document:
    return Document.open(path, author="Tester", force_recreate=True)


# --------------------------------------------------------------------------
# The headline: a redline this library cannot resolve must not report success
# --------------------------------------------------------------------------


class TestForeignOnlyDocumentsAreNotSilentlyResolved:
    @pytest.mark.parametrize("method", ["accept_all", "reject_all"])
    def test_property_only_redline_reports_what_it_could_not_resolve(self, make_docx, method):
        path = make_docx(PROPERTY_CHANGES, "props")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UnhandledRevisionWarning)
            with _open(path) as doc:
                result = getattr(doc, method)()
        # The number that used to be the whole story: nothing was resolvable.
        assert result == 0
        # ...and the part that was missing: three revisions are still pending.
        assert result.unhandled == 3
        assert result.unhandled_types == {"w:pPrChange": 1, "w:rPrChange": 1, "w:sectPrChange": 1}

    @pytest.mark.parametrize("method", ["accept_all", "reject_all"])
    def test_move_only_redline_counts_every_mark_including_range_ends(self, make_docx, method):
        path = make_docx(MOVES, "moves")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UnhandledRevisionWarning)
            with _open(path) as doc:
                result = getattr(doc, method)()
        assert result == 0
        assert result.unhandled == 6
        assert result.unhandled_types == {
            "w:moveFrom": 1,
            "w:moveFromRangeStart": 1,
            "w:moveFromRangeEnd": 1,
            "w:moveTo": 1,
            "w:moveToRangeStart": 1,
            "w:moveToRangeEnd": 1,
        }

    @pytest.mark.parametrize("method", ["accept_all", "reject_all"])
    def test_table_revisions_are_counted(self, make_docx, method):
        path = make_docx(TABLE_REVISIONS, "tables")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UnhandledRevisionWarning)
            with _open(path) as doc:
                result = getattr(doc, method)()
        assert result == 0
        assert result.unhandled_types == {
            "w:tblPrChange": 1,
            "w:tblGridChange": 1,
            "w:tblPrExChange": 1,
            "w:trPrChange": 1,
            "w:tcPrChange": 1,
            "w:cellIns": 1,
            "w:cellDel": 1,
            "w:cellMerge": 1,
        }

    def test_custom_xml_ranges_are_counted(self, make_docx):
        path = make_docx(CUSTOM_XML_RANGES, "customxml")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UnhandledRevisionWarning)
            with _open(path) as doc:
                result = doc.accept_all()
        assert result == 0
        assert result.unhandled == 8

    def test_numbering_change_is_counted(self, make_docx):
        path = make_docx(NUMBERING_CHANGE, "numbering")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UnhandledRevisionWarning)
            with _open(path) as doc:
                result = doc.accept_all()
        assert result.unhandled_types == {"w:numberingChange": 1}


class TestMixedDocument:
    """Real revisions resolve; the foreign mark is reported, not absorbed."""

    def test_accept_all_separates_resolved_from_unresolved(self, make_docx):
        path = make_docx(MIXED, "mixed")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UnhandledRevisionWarning)
            with _open(path) as doc:
                result = doc.accept_all()
                assert result == 2  # the w:del and the w:ins
                assert result.unhandled == 1
                assert result.unhandled_types == {"w:pPrChange": 1}
                assert doc.list_revisions() == []

    def test_unresolved_mark_survives_save_and_reopen(self, make_docx, tmp_path):
        """The count did not over-claim: the pPrChange is still in the file."""
        path = make_docx(MIXED, "mixed_roundtrip")
        out = tmp_path / "mixed_out.docx"
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UnhandledRevisionWarning)
            with _open(path) as doc:
                doc.accept_all()
                doc.save(out)
            with _open(out) as reopened:
                assert reopened.list_revisions() == []  # nothing left to adjudicate
                rows = reopened.list_unhandled_revisions()
        assert [r.tag for r in rows] == ["w:pPrChange"]

    @pytest.mark.parametrize(
        "method,expected_types",
        [
            # Accept unwraps the w:ins, so the run — and its w:rPrChange — stay.
            ("accept_all", {"w:rPrChange": 1}),
            # Reject deletes the w:ins subtree, taking the w:rPrChange with it.
            ("reject_all", {}),
        ],
    )
    def test_foreign_mark_inside_an_insertion_is_counted_after_resolution(self, make_docx, method, expected_types):
        """The counting rule ResolveResult documents: after, not before.

        Counting on entry would report a still-pending w:rPrChange that
        reject_all() had just deleted — an over-claim in the opposite
        direction from the one the honesty floor exists to prevent.
        """
        path = make_docx(FOREIGN_INSIDE_INSERTION, f"nested_{method}")
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            with _open(path) as doc:
                assert len(doc.list_unhandled_revisions()) == 1  # present on entry
                result = getattr(doc, method)()
                assert len(doc.list_unhandled_revisions()) == len(expected_types)

        assert result == 1
        assert result.unhandled_types == expected_types
        # No warning when nothing is left pending: the claim is now true.
        n_warnings = sum(1 for w in caught if w.category is UnhandledRevisionWarning)
        assert n_warnings == (1 if expected_types else 0)

    def test_reject_all_separates_resolved_from_unresolved(self, make_docx):
        path = make_docx(MIXED, "mixed_reject")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UnhandledRevisionWarning)
            with _open(path) as doc:
                result = doc.reject_all()
        assert result == 2
        assert result.unhandled_types == {"w:pPrChange": 1}


# --------------------------------------------------------------------------
# Backward compatibility of the return value
# --------------------------------------------------------------------------


class TestResolveResultIsStillAnInt:
    @pytest.mark.parametrize("method", ["accept_all", "reject_all"])
    def test_ordinary_redline_behaves_exactly_as_before(self, temp_docx, method):
        with _open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            doc.replace("quick", "speedy", paragraph=ref)
            n_revisions = len(doc.list_revisions())
            result = getattr(doc, method)()

        assert isinstance(result, int)
        assert result == n_revisions
        assert result + 1 == n_revisions + 1
        assert f"{result}" == str(n_revisions)
        assert result.unhandled == 0
        assert result.unhandled_types == {}

    def test_clean_document_emits_no_warning(self, temp_docx):
        """warnings-as-errors + an ins/del-only document must stay silent."""
        with warnings.catch_warnings():
            warnings.simplefilter("error", UnhandledRevisionWarning)
            with _open(temp_docx) as doc:
                ref = find_ref(doc, "quick brown fox")
                doc.replace("quick", "speedy", paragraph=ref)
                assert doc.accept_all() == 2

    def test_repr_hides_the_extra_fields_when_there_is_nothing_to_report(self):
        assert repr(ResolveResult(3)) == "ResolveResult(3)"
        assert "unhandled=2" in repr(ResolveResult(3, {"w:rPrChange": 2}))

    def test_str_is_the_plain_int(self):
        assert str(ResolveResult(4, {"w:rPrChange": 2})) == "4"


class TestWarning:
    @pytest.mark.parametrize("method", ["accept_all", "reject_all"])
    def test_foreign_document_warns(self, make_docx, method):
        path = make_docx(PROPERTY_CHANGES, "warn")
        with _open(path) as doc:
            with pytest.warns(UnhandledRevisionWarning) as record:
                getattr(doc, method)()
        message = str(record[0].message)
        assert "w:pPrChange" in message
        assert "list_unhandled_revisions()" in message

    def test_warning_names_the_verb_that_fired_it(self, make_docx):
        path = make_docx(PROPERTY_CHANGES, "warn_verb")
        with _open(path) as doc:
            with pytest.warns(UnhandledRevisionWarning) as record:
                doc.reject_all()
        assert "reject_all()" in str(record[0].message)

    def test_one_warning_per_call(self, make_docx):
        path = make_docx(MOVES, "warn_once")
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            with _open(path) as doc:
                doc.accept_all()
        assert sum(1 for w in caught if w.category is UnhandledRevisionWarning) == 1


# --------------------------------------------------------------------------
# list_unhandled_revisions
# --------------------------------------------------------------------------


class TestListUnhandledRevisions:
    def test_rows_are_in_document_order_with_tags_ids_authors_dates(self, make_docx):
        path = make_docx(PROPERTY_CHANGES, "listing")
        with _open(path) as doc:
            rows = doc.list_unhandled_revisions()

        assert [r.tag for r in rows] == ["w:pPrChange", "w:rPrChange", "w:sectPrChange"]
        assert [r.id for r in rows] == [1, 2, 3]
        assert [r.author for r in rows] == ["Ann", "Bob", "Ann"]
        assert rows[0].date is not None
        assert rows[0].date.year == 2026

    def test_paragraph_ref_resolves_for_marks_inside_a_paragraph(self, make_docx):
        path = make_docx(PROPERTY_CHANGES, "refs")
        with _open(path) as doc:
            rows = doc.list_unhandled_revisions()
            paragraph_refs = doc.list_paragraphs(max_chars=0, limit=None)

        assert rows[0].paragraph_ref == paragraph_refs[0]
        assert rows[1].paragraph_ref == paragraph_refs[0]

    def test_mark_outside_any_paragraph_has_no_paragraph_ref(self, make_docx):
        """A sectPrChange lives in the body's sectPr, not in a <w:p>."""
        path = make_docx(PROPERTY_CHANGES, "outside")
        with _open(path) as doc:
            sect = [r for r in doc.list_unhandled_revisions() if r.tag == "w:sectPrChange"]
        assert len(sect) == 1
        assert sect[0].paragraph_ref is None

    def test_mark_without_w_id_reports_id_none(self, make_docx):
        """cellIns/cellDel/cellMerge are id-bearing here, but a mark can lack one."""
        body = _document(f"<w:p><w:pPr><w:pPrChange {_ANN}><w:pPr/></w:pPrChange></w:pPr></w:p>")
        path = make_docx(body, "no_id")
        with _open(path) as doc:
            rows = doc.list_unhandled_revisions()
        assert [(r.tag, r.id) for r in rows] == [("w:pPrChange", None)]

    def test_mark_without_author_reads_as_unknown(self, make_docx):
        """The range *End marks carry only w:id, exactly as Word writes them."""
        path = make_docx(MOVES, "unknown_author")
        with _open(path) as doc:
            rows = doc.list_unhandled_revisions()
        by_tag = {r.tag: r for r in rows}
        assert by_tag["w:moveFromRangeEnd"].author == "Unknown"
        assert by_tag["w:moveFromRangeEnd"].date is None
        assert by_tag["w:moveFrom"].author == "Ann"

    def test_author_filter_excludes_unattributed_marks(self, make_docx):
        path = make_docx(MOVES, "author_filter")
        with _open(path) as doc:
            everyone = doc.list_unhandled_revisions()
            ann = doc.list_unhandled_revisions(author="Ann")
            unknown = doc.list_unhandled_revisions(author="Unknown")

        assert len(everyone) == 6
        assert len(ann) == 4  # the two *RangeEnd marks carry no w:author
        assert {r.author for r in ann} == {"Ann"}
        assert [r.tag for r in unknown] == ["w:moveFromRangeEnd", "w:moveToRangeEnd"]

    def test_author_filter_on_accept_all_counts_only_that_author(self, make_docx):
        path = make_docx(TABLE_REVISIONS, "accept_filtered")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UnhandledRevisionWarning)
            with _open(path) as doc:
                bob = doc.accept_all(author="Bob")
                ann = doc.accept_all(author="Ann")
        assert bob.unhandled_types == {"w:cellMerge": 1}
        assert ann.unhandled == 6  # everything Ann authored; tblGridChange has no author

    def test_mark_recorded_inside_a_change_record_is_not_a_second_revision(self, make_docx):
        """One w:tcPrChange is one pending revision, not two.

        Its recorded w:tcPr is CT_TcPrInner, so it may legally hold a
        w:cellIns. That marker describes the cell's *previous* state, so
        counting it would overstate the number the honesty floor asks callers
        to trust before reporting a document as fully adjudicated.
        """
        path = make_docx(MARK_INSIDE_CHANGE_RECORD, "change_record")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UnhandledRevisionWarning)
            with _open(path) as doc:
                result = doc.accept_all()
                rows = doc.list_unhandled_revisions()

        assert result.unhandled_types == {"w:tcPrChange": 1}
        assert [r.tag for r in rows] == ["w:tcPrChange"]

    @pytest.mark.xfail(
        strict=True,
        reason="ISSUES.md #68: list_revisions/accept_all still walk every w:ins/w:del, "
        "including ones recorded inside a change record. Flip to passing when #68 "
        "routes the handled path through skip_change_records too.",
    )
    def test_handled_path_still_adjudicates_marks_inside_change_records(self, make_docx):
        """Known gap the honesty floor does NOT close (pre-existing).

        The unhandled path skips a change record's recorded subtree; the
        handled path does not. So a historical w:del inside a w:rPrChange is
        listed as live and accept_all() resolves it — destroying the recorded
        previous state and over-claiming the count by one.
        """
        path = make_docx(HANDLED_MARK_INSIDE_CHANGE_RECORD, "handled_in_record")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UnhandledRevisionWarning)
            with _open(path) as doc:
                assert doc.list_revisions() == []
                result = doc.accept_all()
        assert result == 0

    def test_census_still_counts_marks_inside_change_records(self):
        """The census is a raw inventory, so it deliberately differs."""
        import defusedxml.minidom

        census = count_revision_elements(defusedxml.minidom.parseString(MARK_INSIDE_CHANGE_RECORD))
        assert census.by_tag == {"w:tcPrChange": 1, "w:cellIns": 1}

    def test_clean_document_has_no_unhandled_rows(self, temp_docx):
        with _open(temp_docx) as doc:
            assert doc.list_unhandled_revisions() == []

    def test_repr_is_a_compact_one_liner(self):
        row = UnhandledRevision(tag="w:rPrChange", id=7, author="Ann", date=None, paragraph_ref="P2#a7b2")
        assert repr(row) == "UnhandledRevision(rPrChange 7 @P2#a7b2 by Ann)"
        bare = UnhandledRevision(tag="w:cellIns", id=None, author="Unknown", date=None)
        assert repr(bare) == "UnhandledRevision(cellIns by Unknown)"

    def test_rejects_use_after_close(self, make_docx):
        from docx_editor import DocumentClosedError

        path = make_docx(MOVES, "closed")
        doc = _open(path)
        doc.close()
        with pytest.raises(DocumentClosedError):
            doc.list_unhandled_revisions()


# --------------------------------------------------------------------------
# The census (the corpus harness's counting primitive)
# --------------------------------------------------------------------------


class TestRevisionCensus:
    def test_counts_every_tag_family(self):
        import defusedxml.minidom

        census = count_revision_elements(defusedxml.minidom.parseString(TABLE_REVISIONS))
        assert census.by_tag == {
            "w:tblPrChange": 1,
            "w:tblGridChange": 1,
            "w:tblPrExChange": 1,
            "w:trPrChange": 1,
            "w:cellIns": 1,
            "w:tcPrChange": 1,
            "w:cellDel": 1,
            "w:cellMerge": 1,
        }
        assert census.total == 8
        assert census.ins_del_contexts == {}  # no w:ins/w:del in this fixture

    def test_ins_del_contexts_separate_structural_markers_from_content(self):
        """The row that scopes ISSUES.md #68: which w:ins/w:del are structural."""
        import defusedxml.minidom

        census = count_revision_elements(defusedxml.minidom.parseString(STRUCTURAL_INS_DEL))
        assert census.by_tag == {"w:del": 2, "w:ins": 1}
        assert census.ins_del_contexts == {
            "w:rPr": 1,  # deleted paragraph mark (w:pPr/w:rPr/w:del)
            "w:trPr": 2,  # inserted + deleted table row
        }

    def test_content_revisions_report_their_paragraph_as_context(self):
        import defusedxml.minidom

        census = count_revision_elements(defusedxml.minidom.parseString(MIXED))
        assert census.ins_del_contexts == {"w:p": 2}

    def test_structural_markers_never_enter_the_unhandled_count(self, make_docx):
        """w:ins/w:del are handled tags even when they mark structure."""
        path = make_docx(STRUCTURAL_INS_DEL, "structural")
        with warnings.catch_warnings():
            warnings.simplefilter("error", UnhandledRevisionWarning)
            with _open(path) as doc:
                result = doc.accept_all()
        assert result.unhandled == 0

    def test_empty_document_censuses_to_nothing(self):
        import defusedxml.minidom

        census = count_revision_elements(defusedxml.minidom.parseString(_document("<w:p/>")))
        assert census.by_tag == {}
        assert census.total == 0


# --------------------------------------------------------------------------
# Coverage guard on the tag constants
# --------------------------------------------------------------------------


def test_every_unhandled_tag_has_a_fixture():
    """Adding a tag to UNHANDLED_REVISION_TAGS without a fixture fails here."""
    covered = {tag for xml in ALL_FIXTURES.values() for tag in UNHANDLED_REVISION_TAGS if f"<{tag} " in xml}
    assert covered == set(UNHANDLED_REVISION_TAGS)


def test_change_record_tags_stay_inside_the_unhandled_set():
    """CHANGE_RECORD_TAGS repeats the property-change family — pin the overlap.

    A property-change type added to UNHANDLED_REVISION_TAGS but not here would
    silently stop skip_change_records from skipping its recorded subtree.
    """
    assert set(CHANGE_RECORD_TAGS) <= set(UNHANDLED_REVISION_TAGS)
    # w:numberingChange records its previous value in an attribute, not a
    # subtree, so it is deliberately not a change record.
    assert "w:numberingChange" not in CHANGE_RECORD_TAGS
    assert set(CHANGE_RECORD_TAGS) == {t for t in UNHANDLED_REVISION_TAGS if t.endswith("Change")} - {
        "w:numberingChange"
    }


def test_tag_constants_are_disjoint_and_complete():
    assert set(HANDLED_REVISION_TAGS) & set(UNHANDLED_REVISION_TAGS) == set()
    assert set(ALL_REVISION_TAGS) == set(HANDLED_REVISION_TAGS) | set(UNHANDLED_REVISION_TAGS)
    assert len(ALL_REVISION_TAGS) == len(set(ALL_REVISION_TAGS))
