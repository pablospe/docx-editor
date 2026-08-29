"""A ``<w:tab/>`` mark is one ``"\\t"`` character in the paragraph text map (ISSUES.md #6).

Tabs are searchable and matchable, count as one character in the visible and
original text views and in ``SearchResult`` offsets, and take part in
paragraph hashes. Edits may touch a tab only at its boundary: insertions land
beside it, replace/delete targets that contain one are refused, and a rewrite
must keep every tab — the same number of them, with the text between tabs
rewritten segment by segment. Nothing writes a new tab.
"""

from pathlib import Path

import pytest
from conftest import NS, find_ref, match_for, paragraph_tokens, parse_paragraph, replace_document_xml

from docx_editor import (
    AmbiguousTextError,
    BatchOperationError,
    CommentError,
    Document,
    EditOperation,
    HashMismatchError,
    RevisionError,
    TextNotFoundError,
)
from docx_editor.track_changes import RevisionManager
from docx_editor.xml_editor import (
    DocxXMLEditor,
    build_text_map,
    compute_paragraph_hash,
    count_in_text_map,
    find_in_text_map,
)

AUTHOR = "Test Author"
OTHER = "Reviewer A"
DATE = "2024-01-01T00:00:00Z"

# <w:r><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r> — the canonical fixture
FOO_TAB_BAR = "<w:p><w:r><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r></w:p>"


@pytest.fixture
def temp_xml(tmp_path):
    """Return a function that writes a document.xml with ``body_xml`` as its body."""

    def _create_xml(body_xml: str) -> Path:
        xml = f'<?xml version="1.0" encoding="utf-8"?><w:document {NS}><w:body>{body_xml}</w:body></w:document>'
        xml_path = tmp_path / "test_doc.xml"
        xml_path.write_text(xml)
        return xml_path

    return _create_xml


def _make_manager(xml_path: Path, author: str = AUTHOR) -> RevisionManager:
    editor = DocxXMLEditor(xml_path, rsid="00000000", author=author)
    return RevisionManager(editor)


def _foreign_ins(content: str, ins_id: int = 1) -> str:
    return f'<w:ins w:id="{ins_id}" w:author="{OTHER}" w:date="{DATE}">{content}</w:ins>'


def _foreign_del(content: str, del_id: int = 2) -> str:
    return f'<w:del w:id="{del_id}" w:author="{OTHER}" w:date="{DATE}">{content}</w:del>'


def _own_ins(content: str, ins_id: int = 3) -> str:
    """A pending insertion by the editing author — edits inside it amend it in place."""
    return f'<w:ins w:id="{ins_id}" w:author="{AUTHOR}" w:date="{DATE}">{content}</w:ins>'


def _first_ref(mgr: RevisionManager) -> str:
    p = mgr.editor.dom.getElementsByTagName("w:p")[0]
    return f"P1#{compute_paragraph_hash(p)}"


def _spelled_between_markers(dom, comment_id: int) -> str:
    """The text-map characters between comment ``comment_id``'s range markers."""
    cid = str(comment_id)
    start = next(n for n in dom.getElementsByTagName("w:commentRangeStart") if n.getAttribute("w:id") == cid)
    end = next(n for n in dom.getElementsByTagName("w:commentRangeEnd") if n.getAttribute("w:id") == cid)
    inside = False
    spelled = ""
    for node in dom.getElementsByTagName("*"):
        if node is start:
            inside = True
        elif node is end:
            inside = False
        elif inside and node.parentNode.tagName == "w:r":
            if node.tagName == "w:t":
                spelled += "".join(c.data for c in node.childNodes if c.nodeType == c.TEXT_NODE)
            elif node.tagName == "w:tab":
                spelled += "\t"
    return spelled


def _docx_with_body(simple_docx: Path, tmp_path: Path, body_xml: str) -> Path:
    docx_path = tmp_path / "tabs.docx"
    replace_document_xml(
        simple_docx, docx_path, f'<?xml version="1.0"?><w:document {NS}><w:body>{body_xml}</w:body></w:document>'
    )
    return docx_path


@pytest.fixture
def tab_doc(simple_docx, tmp_path):
    """A document whose first paragraph is ``Name<tab>Value here``."""
    body = (
        "<w:p><w:r><w:t>Name</w:t><w:tab/><w:t>Value here</w:t></w:r></w:p><w:p><w:r><w:t>Other text</w:t></w:r></w:p>"
    )
    with Document.open(_docx_with_body(simple_docx, tmp_path, body), author=AUTHOR) as doc:
        yield doc


# ==================== Text map ====================


class TestTextMap:
    def test_tab_is_one_character_at_offset_zero(self):
        tm = build_text_map(parse_paragraph("<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>"))
        assert tm.text == "a\tb"
        assert [p.is_tab for p in tm.positions] == [False, True, False]
        assert tm.positions[1].offset_in_node == 0
        assert getattr(tm.positions[1].node, "tagName", None) == "w:tab"

    def test_adjacent_tabs_and_tab_in_own_run(self):
        tm = build_text_map(parse_paragraph("<w:p><w:r><w:t>a</w:t><w:tab/><w:tab/><w:t>b</w:t></w:r></w:p>"))
        assert tm.text == "a\t\tb"
        assert tm.positions[1].node is not tm.positions[2].node

        own_run = "<w:p><w:r><w:t>a</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>b</w:t></w:r></w:p>"
        assert build_text_map(parse_paragraph(own_run)).text == "a\tb"

    def test_tab_inside_insertion(self):
        p = parse_paragraph(
            f"<w:p><w:r><w:t>a</w:t></w:r>{_foreign_ins('<w:r><w:tab/><w:t>x</w:t></w:r>')}<w:r><w:t>b</w:t></w:r></w:p>"
        )
        accepted = build_text_map(p)
        assert accepted.text == "a\txb"
        assert accepted.positions[1].is_tab and accepted.positions[1].is_inside_ins
        assert build_text_map(p, view="original").text == "ab"

    def test_tab_inside_deletion(self):
        p = parse_paragraph(
            f"<w:p><w:r><w:t>a</w:t></w:r>{_foreign_del('<w:r><w:tab/><w:delText>x</w:delText></w:r>')}"
            "<w:r><w:t>b</w:t></w:r></w:p>"
        )
        assert build_text_map(p).text == "ab"
        original = build_text_map(p, view="original")
        assert original.text == "a\txb"
        assert original.positions[1].is_tab and original.positions[1].is_inside_del

    def test_tab_inside_move_source_is_not_visible(self):
        # w:moveFrom stores its text as w:delText, so only the tab could leak
        # into the accepted view — and make the moved-away region insertable.
        p = parse_paragraph(
            f'<w:p><w:moveFrom w:id="7" w:author="{OTHER}" w:date="{DATE}">'
            "<w:r><w:delText>Name</w:delText><w:tab/><w:delText>Value</w:delText></w:r></w:moveFrom>"
            "<w:r><w:t>Kept</w:t></w:r></w:p>"
        )
        assert build_text_map(p).text == "Kept"
        original = build_text_map(p, view="original")
        assert original.text == "Name\tValueKept"
        assert original.positions[4].is_tab and original.positions[4].is_inside_del

    def test_tab_stop_is_not_a_character(self):
        p = parse_paragraph(
            '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:t>ab</w:t></w:r></w:p>'
        )
        assert build_text_map(p).text == "ab"
        assert build_text_map(p, view="original").text == "ab"

    def test_tab_inside_text_box_belongs_to_the_box(self):
        p = parse_paragraph(
            "<w:p><w:r><w:t>a</w:t><w:drawing><w:txbxContent><w:p><w:r><w:tab/><w:t>box</w:t></w:r></w:p>"
            "</w:txbxContent></w:drawing><w:t>b</w:t></w:r></w:p>"
        )
        assert build_text_map(p).text == "ab"
        assert build_text_map(p, view="original").text == "ab"

    def test_find_and_count_see_the_tab(self):
        tm = build_text_map(parse_paragraph("<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>"))
        match = find_in_text_map(tm, "a\tb")
        assert match is not None and (match.start, match.end) == (0, 3)
        assert [p.is_tab for p in tm.get_nodes_for_range(0, 3)] == [False, True, False]
        assert find_in_text_map(tm, "ab") is None

        tm = build_text_map(parse_paragraph("<w:p><w:r><w:t>x</w:t><w:tab/><w:t>x x</w:t></w:r></w:p>"))
        assert count_in_text_map(tm, "x") == 3
        middle = find_in_text_map(tm, "x", occurrence=1)
        assert middle is not None and middle.start == 2


# ==================== Search, offsets, hashes ====================


class TestSearch:
    def test_find_text_matches_across_a_tab(self, tab_doc):
        m = match_for(tab_doc, "Name\tValue")
        assert (m.start, m.end, m.text) == (0, 10, "Name\tValue")

    def test_text_joined_across_the_tab_no_longer_matches(self, tab_doc):
        assert tab_doc.find_text("NameValue") is None
        ref = find_ref(tab_doc, "Name")
        with pytest.raises(TextNotFoundError) as exc:
            tab_doc.replace("NameValue", "x", paragraph=ref)
        # The preview spells the tab out — a raw tab would read like the very
        # space the caller searched with — and the message carries it too.
        assert exc.value.paragraph_preview == "Name\\tValue here"
        assert 'Current content: "Name\\tValue here"' in str(exc.value)

    def test_stale_ref_preview_spells_the_tab(self, tab_doc):
        ref = find_ref(tab_doc, "Name")
        tab_doc.replace("Value", "Worth", paragraph=ref)
        with pytest.raises(HashMismatchError) as exc:
            tab_doc.replace("Worth", "x", paragraph=ref)
        assert exc.value.paragraph_preview == "Name\\tWorth here"
        assert 'Current content: "Name\\tWorth here"' in str(exc.value)

    def test_offsets_count_the_tab_as_one_character(self, tab_doc):
        m = match_for(tab_doc, "Value")
        assert (m.start, m.end) == (5, 10)
        first_line = tab_doc.get_visible_text().splitlines()[0]
        assert first_line[m.start : m.end] == "Value"

    def test_count_and_occurrences_across_tabs(self, simple_docx, tmp_path):
        body = "<w:p><w:r><w:t>x</w:t><w:tab/><w:t>x x</w:t></w:r></w:p>"
        with Document.open(_docx_with_body(simple_docx, tmp_path, body), author=AUTHOR) as doc:
            assert doc.count_matches("x") == 3
            assert [m.start for m in doc.find_all("x")] == [0, 2, 4]
            ref = find_ref(doc, "x")
            with pytest.raises(AmbiguousTextError) as exc:
                doc.replace("x", "y", paragraph=ref)
            assert exc.value.total_occurrences == 3
            assert exc.value.paragraph_preview == "x\\tx x"
            # A SearchResult round-trips into an edit and targets the middle hit.
            doc.replace(match_for(doc, "x", occurrence=1), "y")
            assert doc.get_visible_text().splitlines()[0] == "x\ty x"

    def test_search_result_anchor_touching_a_tab(self, tab_doc):
        tab_doc.insert_after(match_for(tab_doc, "Name\t"), "X")
        assert tab_doc.get_visible_text().splitlines()[0] == "Name\tXValue here"


class TestHash:
    def test_hash_distinguishes_tab_from_no_tab(self):
        with_tab = parse_paragraph("<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>")
        without = parse_paragraph("<w:p><w:r><w:t>ab</w:t></w:r></w:p>")
        assert compute_paragraph_hash(with_tab) != compute_paragraph_hash(without)

    def test_refs_are_stable_and_resolve_through_edits(self, tab_doc):
        ref = find_ref(tab_doc, "Name")
        tab_doc.find_text("Value")
        assert find_ref(tab_doc, "Name") == ref
        assert tab_doc.get_paragraph(1).ref == ref

        new_ref = tab_doc.replace("Value", "Worth", paragraph=ref)
        assert tab_doc.get_paragraph(1).ref == str(new_ref)
        tab_doc.delete("here", paragraph=new_ref)
        assert tab_doc.get_visible_text().splitlines()[0] == "Name\tWorth "

    def test_text_views_carry_the_tab(self, tab_doc):
        assert tab_doc.get_visible_text().splitlines()[0] == "Name\tValue here"
        assert tab_doc.get_original_text().splitlines()[0] == "Name\tValue here"
        assert tab_doc.get_paragraph(1).text == "Name\tValue here"
        assert "Name\tValue here" in tab_doc.list_paragraphs()[0]


# ==================== Edit guards ====================


class TestEditGuards:
    """Replace/delete targets may not contain a tab; content inputs may not write one."""

    def test_replace_target_with_tab_is_refused_before_mutation(self, tab_doc):
        ref = find_ref(tab_doc, "Name")
        with pytest.raises(ValueError, match="ISSUES.md #6") as exc:
            tab_doc.replace("Name\tValue", "x", paragraph=ref)
        assert "either side" in str(exc.value)
        assert tab_doc.get_visible_text().splitlines()[0] == "Name\tValue here"
        assert tab_doc.list_revisions() == []

    def test_delete_target_with_tab_is_refused(self, tab_doc):
        ref = find_ref(tab_doc, "Name")
        with pytest.raises(ValueError, match="ISSUES.md #6"):
            tab_doc.delete("\t", paragraph=ref)
        with pytest.raises(ValueError, match="ISSUES.md #6"):
            tab_doc.delete("Name\t", paragraph=ref)
        assert tab_doc.list_revisions() == []

    def test_edit_operation_constructors_refuse_tab_targets(self):
        with pytest.raises(ValueError, match="ISSUES.md #6"):
            EditOperation.replace("a\tb", "x", paragraph="P1#0000")
        with pytest.raises(ValueError, match="ISSUES.md #6"):
            EditOperation.delete("a\t", paragraph="P1#0000")
        # Anchors may contain a tab.
        assert EditOperation.insert_after("a\t", "x", paragraph="P1#0000").anchor == "a\t"

    def test_batch_refuses_raw_tab_target_atomically(self, tab_doc):
        ref = find_ref(tab_doc, "Name")
        ops = [
            EditOperation.replace("Name", "Label", paragraph=ref),
            EditOperation(action="delete", paragraph=ref, text="\tValue"),
        ]
        with pytest.raises(BatchOperationError, match="ISSUES.md #6"):
            tab_doc.batch_edit(ops)
        assert tab_doc.get_visible_text().splitlines()[0] == "Name\tValue here"

    def test_content_inputs_still_reject_a_tab(self, tab_doc):
        ref = find_ref(tab_doc, "Name")
        with pytest.raises(ValueError, match="control character"):
            tab_doc.insert_after("Name", "\t", paragraph=ref)
        with pytest.raises(ValueError, match="control character"):
            tab_doc.replace("Name", "N\tame", paragraph=ref)
        with pytest.raises(ValueError, match="control character"):
            tab_doc.replace("Name", "Label", paragraph=ref, note="why\tnot")
        assert tab_doc.list_revisions() == []


# ==================== Edits beside a tab ====================


class TestAdjacentEdits:
    def test_replace_on_either_side_of_the_tab(self, temp_xml):
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        mgr.replace_text("foo", "qux")
        assert paragraph_tokens(mgr) == ["DEL(foo)", "INS(qux)", "TAB", "bar"]
        mgr.replace_text("bar", "baz")
        assert paragraph_tokens(mgr) == ["DEL(foo)", "INS(qux)", "TAB", "DEL(bar)", "INS(baz)"]
        mgr.accept_all()
        assert paragraph_tokens(mgr) == ["qux", "TAB", "baz"]

    def test_delete_beside_the_tab_then_reject(self, temp_xml):
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        mgr.suggest_deletion("foo")
        assert paragraph_tokens(mgr) == ["DEL(foo)", "TAB", "bar"]
        mgr.reject_all()
        assert paragraph_tokens(mgr) == ["foo", "TAB", "bar"]

    def test_insert_after_anchor_ending_on_the_tab(self, temp_xml):
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        mgr.insert_text_after("foo\t", "X")
        assert paragraph_tokens(mgr) == ["foo", "TAB", "INS(X)", "bar"]
        mgr.accept_all()
        assert paragraph_tokens(mgr) == ["foo", "TAB", "X", "bar"]

    def test_insert_before_anchor_starting_on_the_tab(self, temp_xml):
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        mgr.insert_text_before("\tbar", "X")
        assert paragraph_tokens(mgr) == ["foo", "INS(X)", "TAB", "bar"]
        mgr.reject_all()
        assert paragraph_tokens(mgr) == ["foo", "TAB", "bar"]

    def test_insert_after_anchor_with_interior_tab(self, temp_xml):
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        mgr.insert_text_after("foo\tbar", "X")
        assert paragraph_tokens(mgr) == ["foo", "TAB", "bar", "INS(X)"]
        mgr.insert_text_before("foo\tbar", "Y")
        assert paragraph_tokens(mgr) == ["INS(Y)", "foo", "TAB", "bar", "INS(X)"]

    def test_insert_at_tab_edge_inside_own_insertion(self, temp_xml):
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        mgr.insert_text_after("foo", "A")
        # A is our own pending insertion; the tab sits next to it in a plain run.
        assert paragraph_tokens(mgr) == ["foo", "INS(A)", "TAB", "bar"]
        mgr.insert_text_before("\tbar", "B")
        # B lands between A and the tab (after the anchor's neighbour), still one <w:ins>
        assert paragraph_tokens(mgr) == ["foo", "INS(A)", "INS(B)", "TAB", "bar"]
        assert len(mgr.editor.dom.getElementsByTagName("w:ins")) == 2

    def test_insert_at_tab_edge_inside_own_insertion_holding_the_tab(self, temp_xml):
        own_ins = (
            f'<w:ins w:id="7" w:author="{AUTHOR}" w:date="{DATE}"><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:ins>'
        )
        mgr = _make_manager(temp_xml(f"<w:p><w:r><w:t>x</w:t></w:r>{own_ins}</w:p>"))
        mgr.insert_text_after("a\t", "Y")
        assert paragraph_tokens(mgr) == ["x", "INS(a)", "INS(TAB)", "INS(Y)", "INS(b)"]
        mgr.insert_text_before("\tY", "Z")
        assert paragraph_tokens(mgr) == ["x", "INS(a)", "INS(Z)", "INS(TAB)", "INS(Y)", "INS(b)"]
        # Amending our own insertion never nests <w:ins> in <w:ins>.
        ins_elems = mgr.editor.dom.getElementsByTagName("w:ins")
        assert len(ins_elems) == 1
        assert not ins_elems[0].getElementsByTagName("w:ins")

    def test_insert_beside_tab_at_foreign_insertion_boundary(self, temp_xml):
        foreign = _foreign_ins("<w:r><w:tab/><w:t>b</w:t></w:r>")
        mgr = _make_manager(temp_xml(f"<w:p><w:r><w:t>a</w:t></w:r>{foreign}</w:p>"))
        mgr.insert_text_before("\tb", "X")
        # Tab at the foreign ins start: our ins is a plain sibling before theirs.
        assert paragraph_tokens(mgr) == ["a", "INS(X)", "INS(TAB)", "INS(b)"]
        ins_elems = mgr.editor.dom.getElementsByTagName("w:ins")
        assert [e.getAttribute("w:author") for e in ins_elems] == [AUTHOR, OTHER]

    def test_insert_beside_tab_mid_foreign_insertion(self, temp_xml):
        foreign = _foreign_ins("<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r>")
        mgr = _make_manager(temp_xml(f"<w:p>{foreign}</w:p>"))
        mgr.insert_text_after("a\t", "X")
        assert paragraph_tokens(mgr) == ["INS(a)", "INS(TAB)", "INS(X)", "INS(b)"]
        # Theirs is split into two halves around ours; every author is kept.
        authors = [e.getAttribute("w:author") for e in mgr.editor.dom.getElementsByTagName("w:ins")]
        assert authors == [OTHER, AUTHOR, OTHER]

    def test_insert_beside_tab_only_foreign_insertion(self, temp_xml):
        foreign = _foreign_ins("<w:r><w:tab/></w:r>")
        mgr = _make_manager(temp_xml(f"<w:p><w:r><w:t>a</w:t></w:r>{foreign}<w:r><w:t>b</w:t></w:r></w:p>"))
        mgr.insert_text_after("a\t", "X")
        assert paragraph_tokens(mgr) == ["a", "INS(TAB)", "INS(X)", "b"]
        mgr.insert_text_before("\tX", "Y")
        assert paragraph_tokens(mgr) == ["a", "INS(Y)", "INS(TAB)", "INS(X)", "b"]

    def test_delete_text_from_insertion_keeps_its_tab(self, temp_xml):
        foreign = _foreign_ins("<w:r><w:t>qux</w:t><w:tab/></w:r>")
        mgr = _make_manager(temp_xml(f"<w:p><w:r><w:t>a</w:t></w:r>{foreign}<w:r><w:t>b</w:t></w:r></w:p>"))
        mgr.suggest_deletion("qux")
        assert build_text_map(mgr.editor.dom.getElementsByTagName("w:p")[0]).text == "a\tb"
        assert "TAB" in " ".join(paragraph_tokens(mgr))

    def test_splice_into_a_tab_is_an_internal_error(self, temp_xml):
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        tab = mgr.editor.dom.getElementsByTagName("w:tab")[0]
        with pytest.raises(RevisionError, match="w:tab"):
            mgr._set_node_text(tab, "x")


# ==================== Paragraph splits ====================


class TestSplits:
    def _texts(self, mgr: RevisionManager) -> list[str]:
        return [build_text_map(p).text for p in mgr.editor.dom.getElementsByTagName("w:p")]

    def test_split_right_after_text_keeps_tab_in_the_tail(self, temp_xml):
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        mgr.insert_text_after("foo", "\n")
        assert self._texts(mgr) == ["foo", "\tbar"]
        mgr.reject_all()
        assert self._texts(mgr) == ["foo\tbar"]
        assert paragraph_tokens(mgr) == ["foo", "TAB", "bar"]

    def test_split_right_after_tab_keeps_tab_in_the_head(self, temp_xml):
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        mgr.insert_text_before("bar", "\n")
        assert self._texts(mgr) == ["foo\t", "bar"]
        mgr.accept_all()
        assert self._texts(mgr) == ["foo\t", "bar"]

    def test_split_before_second_text_of_a_run_keeps_the_head(self, temp_xml):
        # The run's first content child is not the split edge: the run must be
        # split rather than moved whole (the latent multi-child-run case).
        mgr = _make_manager(temp_xml("<w:p><w:r><w:t>foo</w:t><w:t>bar</w:t></w:r></w:p>"))
        mgr.insert_text_before("bar", "\n")
        assert self._texts(mgr) == ["foo", "bar"]

    def test_split_after_a_leading_break_keeps_the_break_in_the_head(self, temp_xml):
        # The same rule for any leading child: a w:br before the edge stays put.
        mgr = _make_manager(temp_xml("<w:p><w:r><w:br/><w:t>bar</w:t></w:r></w:p>"))
        mgr.insert_text_before("bar", "\n")
        assert self._texts(mgr) == ["", "bar"]
        assert paragraph_tokens(mgr) == ["BR"]

    def test_split_skips_the_run_properties_when_finding_the_first_child(self, temp_xml):
        # w:rPr precedes the content: it is not the first content child, and
        # both halves of the split run keep it.
        mgr = _make_manager(temp_xml("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r></w:p>"))
        mgr.insert_text_before("bar", "\n")
        assert self._texts(mgr) == ["foo\t", "bar"]
        assert paragraph_tokens(mgr) == ["foo", "TAB"]
        assert len(mgr.editor.dom.getElementsByTagName("w:b")) == 3  # foo run, tab run, bar run


# ==================== rewrite_paragraph ====================


class TestRewrite:
    def test_rewrite_keeps_a_tab(self, temp_xml):
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        mgr.rewrite_paragraph(_first_ref(mgr), "foo\tbaz")
        assert paragraph_tokens(mgr) == ["foo", "TAB", "DEL(bar)", "INS(baz)"]

    def test_rewrite_insert_before_the_tab(self, temp_xml):
        mgr = _make_manager(temp_xml("<w:p><w:r><w:t>a b</w:t><w:tab/><w:t>c</w:t></w:r></w:p>"))
        mgr.rewrite_paragraph(_first_ref(mgr), "a b x\tc")
        assert paragraph_tokens(mgr) == ["a b", "INS( x)", "TAB", "c"]

    def test_rewrite_insert_after_the_tab_keeps_run_siblings(self, temp_xml):
        # Regression for the hand-rolled insert path that dropped a run's other children.
        mgr = _make_manager(temp_xml("<w:p><w:r><w:t>foo</w:t><w:tab/><w:t>bar baz</w:t></w:r></w:p>"))
        mgr.rewrite_paragraph(_first_ref(mgr), "foo\tbar qux baz")
        assert paragraph_tokens(mgr) == ["foo", "TAB", "bar ", "INS(qux )", "baz"]

    @pytest.mark.parametrize(
        "new_text",
        ["foo bar", "foo\t\tbar", "foobar", ""],
        ids=["replace-tab", "add-tab", "drop-tab", "clear"],
    )
    def test_rewrite_refuses_to_add_or_remove_a_tab(self, temp_xml, new_text):
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        with pytest.raises(ValueError, match="tab marks"):
            mgr.rewrite_paragraph(_first_ref(mgr), new_text)
        assert paragraph_tokens(mgr) == ["foo", "TAB", "bar"]

    def test_rewrite_moving_words_around_the_tab_keeps_the_element(self, temp_xml):
        # Each side of the tab is diffed as its own segment, so the words on
        # either side are redlined and the <w:tab/> is never part of the diff.
        mgr = _make_manager(temp_xml("<w:p><w:r><w:t>foo</w:t><w:tab/><w:t>bar baz</w:t></w:r></w:p>"))
        mgr.rewrite_paragraph(_first_ref(mgr), "foo bar\tbaz")
        assert build_text_map(mgr.editor.dom.getElementsByTagName("w:p")[0]).text == "foo bar\tbaz"
        assert len(mgr.editor.dom.getElementsByTagName("w:tab")) == 1
        assert paragraph_tokens(mgr) == ["foo", "INS( bar)", "TAB", "DEL(bar )", "baz"]

    def test_rewrite_moves_words_across_the_tab_segment_by_segment(self, temp_xml):
        # Each tab-delimited segment is diffed on its own, so text may move
        # across a tab: the words are redlined, the <w:tab/> stays.
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        mgr.rewrite_paragraph(_first_ref(mgr), "\tfoo bar")
        assert build_text_map(mgr.editor.dom.getElementsByTagName("w:p")[0]).text == "\tfoo bar"
        assert paragraph_tokens(mgr) == ["DEL(foo)", "TAB", "INS(foo )", "bar"]

    def test_rewrite_swaps_words_around_the_tab(self, temp_xml):
        mgr = _make_manager(temp_xml(FOO_TAB_BAR))
        mgr.rewrite_paragraph(_first_ref(mgr), "bar\tfoo")
        assert build_text_map(mgr.editor.dom.getElementsByTagName("w:p")[0]).text == "bar\tfoo"
        assert len(mgr.editor.dom.getElementsByTagName("w:tab")) == 1

    def _text(self, mgr: RevisionManager) -> str:
        return build_text_map(mgr.editor.dom.getElementsByTagName("w:p")[0]).text

    def test_rewrite_appends_after_a_foreign_insertion_holding_a_tab(self, temp_xml):
        # The last character sits inside another author's insertion: our text
        # becomes a sibling <w:ins> after theirs, never spliced into it.
        body = f"<w:p><w:r><w:t>a </w:t></w:r>{_foreign_ins('<w:r><w:tab/><w:t>b</w:t></w:r>')}</w:p>"
        mgr = _make_manager(temp_xml(body))
        mgr.rewrite_paragraph(_first_ref(mgr), "a \tb X")
        assert self._text(mgr) == "a \tb X"
        assert paragraph_tokens(mgr) == ["a ", "INS(TAB)", "INS(b)", "INS( X)"]
        authors = [i.getAttribute("w:author") for i in mgr.editor.dom.getElementsByTagName("w:ins")]
        assert authors == [OTHER, AUTHOR]

    def test_rewrite_appends_after_a_tab_ending_our_own_insertion(self, temp_xml):
        # Our own pending insertion ends in a tab: the appended text lands in a
        # plain sibling run inside that insertion (an amendment, no new w:ins).
        body = f"<w:p><w:r><w:t>a</w:t></w:r>{_own_ins('<w:r><w:t>b</w:t><w:tab/></w:r>')}</w:p>"
        mgr = _make_manager(temp_xml(body))
        mgr.rewrite_paragraph(_first_ref(mgr), "ab\tX")
        assert self._text(mgr) == "ab\tX"
        assert paragraph_tokens(mgr) == ["a", "INS(b)", "INS(TAB)", "INS(X)"]
        assert len(mgr.editor.dom.getElementsByTagName("w:ins")) == 1

    def test_rewrite_inserts_inside_a_foreign_insertion_beside_a_tab(self, temp_xml):
        # Mid-insertion position in another author's w:ins: theirs is split
        # into two identity-preserving halves with our own w:ins between.
        body = f"<w:p><w:r><w:t>a</w:t><w:tab/></w:r>{_foreign_ins('<w:r><w:t>b c</w:t></w:r>')}</w:p>"
        mgr = _make_manager(temp_xml(body))
        mgr.rewrite_paragraph(_first_ref(mgr), "a\tb X c")
        assert self._text(mgr) == "a\tb X c"
        assert paragraph_tokens(mgr) == ["a", "TAB", "INS(b )", "INS(X )", "INS(c)"]
        authors = [i.getAttribute("w:author") for i in mgr.editor.dom.getElementsByTagName("w:ins")]
        assert authors == [OTHER, AUTHOR, OTHER]

    def test_rewrite_inserts_before_a_tab_inside_our_own_insertion(self, temp_xml):
        body = f"<w:p><w:r><w:t>a</w:t></w:r>{_own_ins('<w:r><w:tab/><w:t>b</w:t></w:r>')}</w:p>"
        mgr = _make_manager(temp_xml(body))
        mgr.rewrite_paragraph(_first_ref(mgr), "a X\tb")
        assert self._text(mgr) == "a X\tb"
        assert paragraph_tokens(mgr) == ["a", "INS( X)", "INS(TAB)", "INS(b)"]
        assert len(mgr.editor.dom.getElementsByTagName("w:ins")) == 1

    def test_batch_rewrite_applies_the_same_guard(self, tab_doc):
        ref = find_ref(tab_doc, "Name")
        with pytest.raises(BatchOperationError, match="tab marks"):
            tab_doc.batch_rewrite([(ref, "Name Value here")])
        assert tab_doc.get_visible_text().splitlines()[0] == "Name\tValue here"
        tab_doc.batch_rewrite([(ref, "Label\tValue here")])
        assert tab_doc.get_visible_text().splitlines()[0] == "Label\tValue here"

    def test_rewrite_idiom_on_paragraph_text(self, tab_doc):
        """``rewrite_paragraph(ref, info.text.replace(...))`` keeps working on tab-bearing text."""
        info = tab_doc.get_paragraph(1)
        tab_doc.rewrite_paragraph(info.ref, info.text.replace("Value", "Worth"))
        assert tab_doc.get_visible_text().splitlines()[0] == "Name\tWorth here"

    def test_long_paragraph_keeps_several_tabs_aligned(self, temp_xml):
        # 150 tab-delimited segments, each diffed on its own: the segment
        # splitting and hunk offsets at scale.
        cells = [f"w{i}" for i in range(150)]
        runs = "".join(f"<w:t>{c} </w:t><w:tab/>" for c in cells) + "<w:t>end</w:t>"
        mgr = _make_manager(temp_xml(f"<w:p><w:r>{runs}</w:r></w:p>"))
        old = build_text_map(mgr.editor.dom.getElementsByTagName("w:p")[0]).text
        assert old.count("\t") == 150
        new = old.replace("w10 ", "W10 ").replace("w77 ", "W77 ").replace("end", "END")
        mgr.rewrite_paragraph(_first_ref(mgr), new)
        assert build_text_map(mgr.editor.dom.getElementsByTagName("w:p")[0]).text == new
        assert len(mgr.editor.dom.getElementsByTagName("w:tab")) == 150

    def test_long_paragraph_keeps_a_tab_whose_neighbours_both_change(self, temp_xml):
        # Diffed as one token stream, a frequent "\t" was autojunk (200+ tokens)
        # and could not seed a match, so changing the words on both sides of
        # one tab left the tab inside a replace hunk; segment-wise diffing
        # cannot put a tab in a hunk at all.
        cells = [f"w{i}" for i in range(150)]
        runs = "".join(f"<w:t>{c} </w:t><w:tab/>" for c in cells) + "<w:t>end</w:t>"
        mgr = _make_manager(temp_xml(f"<w:p><w:r>{runs}</w:r></w:p>"))
        old = build_text_map(mgr.editor.dom.getElementsByTagName("w:p")[0]).text
        new = old.replace("w10 \tw11 ", "X \tY ")
        assert new.count("\t") == 150
        mgr.rewrite_paragraph(_first_ref(mgr), new)
        assert build_text_map(mgr.editor.dom.getElementsByTagName("w:p")[0]).text == new
        assert len(mgr.editor.dom.getElementsByTagName("w:tab")) == 150


# ==================== Revisions and comments ====================


class TestRevisions:
    def test_insertion_text_and_occurrence_include_the_tab(self, simple_docx, tmp_path):
        foreign = _foreign_ins("<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r>", ins_id=11)
        body = f"<w:p><w:r><w:t>a b </w:t></w:r>{foreign}</w:p>"
        with Document.open(_docx_with_body(simple_docx, tmp_path, body), author=AUTHOR) as doc:
            rev = next(r for r in doc.list_revisions() if r.id == 11)
            assert rev.text == "a\tb"
            assert rev.occurrence == 0
            assert rev.paragraph_ref is not None
            doc.add_comment(rev.text, "on the insertion", paragraph=rev.paragraph_ref, occurrence=rev.occurrence)

    def test_deletion_text_and_occurrence_include_the_tab(self, simple_docx, tmp_path):
        foreign = _foreign_del("<w:r><w:delText>a</w:delText><w:tab/><w:delText>b</w:delText></w:r>", del_id=12)
        body = f"<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b </w:t></w:r>{foreign}</w:p>"
        with Document.open(_docx_with_body(simple_docx, tmp_path, body), author=AUTHOR) as doc:
            rev = next(r for r in doc.list_revisions() if r.id == 12)
            assert rev.text == "a\tb"
            # Original text is "a\tb a\tb": the live span precedes the deleted one.
            assert rev.occurrence == 1

    def test_deletion_with_plain_wt_still_reports_the_tab(self, simple_docx, tmp_path):
        # Nonconforming producers leave w:t (not w:delText) inside w:del; the
        # fallback walk reads those and the tab between them.
        foreign = _foreign_del("<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r>", del_id=14)
        body = f"<w:p><w:r><w:t>x </w:t></w:r>{foreign}</w:p>"
        with Document.open(_docx_with_body(simple_docx, tmp_path, body), author=AUTHOR) as doc:
            rev = next(r for r in doc.list_revisions() if r.id == 14)
            assert rev.text == "a\tb"


class TestComments:
    def test_anchor_spanning_a_tab_brackets_it(self, tab_doc):
        ref = find_ref(tab_doc, "Name")
        cid = tab_doc.add_comment("Name\tValue", "spans the tab", paragraph=ref)
        # Everything between the markers spells the anchor, tab included.
        assert _spelled_between_markers(tab_doc._document_editor.dom, cid) == "Name\tValue"
        assert tab_doc.get_visible_text().splitlines()[0] == "Name\tValue here"

    @pytest.mark.parametrize("anchor", ["\tValue", "Name\t", "\t"])
    def test_anchor_on_a_tab_edge_brackets_the_tab(self, tab_doc, anchor):
        # Range markers are run-level marks: a tab edge needs no w:t split.
        ref = find_ref(tab_doc, "Name")
        cid = tab_doc.add_comment(anchor, "edge", paragraph=ref)
        assert _spelled_between_markers(tab_doc._document_editor.dom, cid) == anchor
        assert tab_doc.get_visible_text().splitlines()[0] == "Name\tValue here"

    def test_revision_starting_with_a_tab_can_be_commented(self, simple_docx, tmp_path):
        # Revision.text spells the tab and its occurrence plugs into add_comment —
        # the shape Word writes when a tabbed list line is inserted.
        foreign = _foreign_ins("<w:r><w:tab/><w:t>New item</w:t></w:r>", ins_id=13)
        body = f"<w:p><w:r><w:t>List:</w:t></w:r>{foreign}</w:p>"
        with Document.open(_docx_with_body(simple_docx, tmp_path, body), author=AUTHOR) as doc:
            rev = next(r for r in doc.list_revisions() if r.id == 13)
            assert rev.text == "\tNew item"
            cid = doc.add_comment(rev.text, "why?", paragraph=rev.paragraph_ref, occurrence=rev.occurrence)
            assert _spelled_between_markers(doc._document_editor.dom, cid) == "\tNew item"

    def test_comment_body_still_rejects_a_tab(self, tab_doc):
        ref = find_ref(tab_doc, "Name")
        with pytest.raises(CommentError, match="control character"):
            tab_doc.add_comment("Name", "bad\tbody", paragraph=ref)
