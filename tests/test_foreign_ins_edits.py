"""Tests for author-aware edits inside another author's pending insertion.

When the current author (B) edits text that sits inside a foreign author's
(A's) pending ``w:ins``, the edit must not destroy A's proposal:

- B deleting text inside A's insertion nests a ``<w:del w:author="B">``
  inside A's ``w:ins`` (Word's own behavior).
- B replacing text produces that nested deletion plus B's own *sibling*
  ``<w:ins>`` carrying the replacement — never nested inside A's.
- B inserting text mid-insertion splits A's ``w:ins`` into two siblings
  (both keeping A's author/date, fresh ids) with B's ``w:ins`` between.

Edits inside the current author's *own* insertions keep the historical
in-place behavior (no ``w:del``, direct text surgery).
"""

from pathlib import Path

import pytest
from conftest import find_ref

from docx_editor import Document
from docx_editor.track_changes import RevisionManager
from docx_editor.xml_editor import DocxXMLEditor, compute_paragraph_hash

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

AUTHOR_A = "Reviewer A"
AUTHOR_B = "Reviewer B"
DATE_A = "2024-01-01T00:00:00Z"


@pytest.fixture
def temp_xml(tmp_path):
    """Fixture that returns a function to create temp XML files."""

    def _create_xml(body_xml: str) -> Path:
        xml = f'<?xml version="1.0" encoding="utf-8"?><w:document {NS}><w:body>{body_xml}</w:body></w:document>'
        xml_path = tmp_path / "test_doc.xml"
        xml_path.write_text(xml)
        return xml_path

    return _create_xml


def _make_manager(xml_path: Path, author: str = AUTHOR_B) -> RevisionManager:
    """Create a RevisionManager editing as ``author``."""
    editor = DocxXMLEditor(xml_path, rsid="00000000", author=author)
    return RevisionManager(editor)


def _foreign_ins(content: str, ins_id: int = 1) -> str:
    """Wrap run XML in a w:ins authored by Reviewer A."""
    return f'<w:ins w:id="{ins_id}" w:author="{AUTHOR_A}" w:date="{DATE_A}">{content}</w:ins>'


def _ins_elems(manager: RevisionManager) -> list:
    return list(manager.editor.dom.getElementsByTagName("w:ins"))


def _del_elems(manager: RevisionManager) -> list:
    return list(manager.editor.dom.getElementsByTagName("w:del"))


def _visible_text(manager: RevisionManager) -> str:
    """Accepted-view text: all w:t content not inside w:del."""
    result = []
    for wt in manager.editor.dom.getElementsByTagName("w:t"):
        parent = wt.parentNode
        inside_del = False
        while parent:
            if parent.nodeType == parent.ELEMENT_NODE and parent.tagName == "w:del":
                inside_del = True
                break
            parent = parent.parentNode
        if not inside_del:
            result.append("".join(c.data for c in wt.childNodes if c.nodeType == c.TEXT_NODE))
    return "".join(result)


def _del_text(del_elem) -> str:
    parts = []
    for dt in del_elem.getElementsByTagName("w:delText"):
        parts.append("".join(c.data for c in dt.childNodes if c.nodeType == c.TEXT_NODE))
    return "".join(parts)


def _ins_visible_text(ins_elem) -> str:
    """Text of an insertion excluding any nested deletion content."""
    parts = []
    for wt in ins_elem.getElementsByTagName("w:t"):
        parent = wt.parentNode
        inside_del = False
        while parent is not ins_elem and parent is not None:
            if parent.nodeType == parent.ELEMENT_NODE and parent.tagName == "w:del":
                inside_del = True
                break
            parent = parent.parentNode
        if not inside_del:
            parts.append("".join(c.data for c in wt.childNodes if c.nodeType == c.TEXT_NODE))
    return "".join(parts)


def _assert_no_nested_ins(manager: RevisionManager) -> None:
    """No w:ins may ever nest inside another w:ins (invalid OOXML)."""
    for ins in _ins_elems(manager):
        assert len(ins.getElementsByTagName("w:ins")) == 0, f"Nested w:ins inside w:ins: {ins.toxml()}"


def _assert_no_del_in_own_ins(manager: RevisionManager, own_author: str) -> None:
    """Our own insertions are edited in place — never via nested w:del."""
    for ins in _ins_elems(manager):
        if ins.getAttribute("w:author") == own_author:
            assert len(ins.getElementsByTagName("w:del")) == 0, f"w:del nested inside our own w:ins: {ins.toxml()}"


def _nested_dels(ins_elem) -> list:
    return list(ins_elem.getElementsByTagName("w:del"))


class TestDeleteInsideForeignIns:
    """B deleting text inside A's pending insertion nests a w:del by B."""

    def test_delete_mid_single_run(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello world today</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.suggest_deletion("world")

        _assert_no_nested_ins(manager)
        ins_elems = _ins_elems(manager)
        assert len(ins_elems) == 1
        a_ins = ins_elems[0]
        # A's proposal survives untouched: element, author, and id preserved
        assert a_ins.getAttribute("w:author") == AUTHOR_A
        assert a_ins.getAttribute("w:id") == "1"
        assert a_ins.getAttribute("w:date") == DATE_A
        # B's counter-proposal nests inside it
        dels = _nested_dels(a_ins)
        assert len(dels) == 1
        assert dels[0].getAttribute("w:author") == AUTHOR_B
        assert _del_text(dels[0]) == "world"
        # Surviving insertion text keeps its order
        assert _ins_visible_text(a_ins) == "Hello  today"
        assert _visible_text(manager) == "Hello  today"

    def test_delete_at_start(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.suggest_deletion("Hello ")

        a_ins = _ins_elems(manager)[0]
        dels = _nested_dels(a_ins)
        assert len(dels) == 1
        assert _del_text(dels[0]) == "Hello "
        assert _visible_text(manager) == "world"
        # Deletion comes before the surviving text inside the ins
        children = [c for c in a_ins.childNodes if c.nodeType == c.ELEMENT_NODE]
        assert children[0].tagName == "w:del"

    def test_delete_at_end(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.suggest_deletion(" world")

        a_ins = _ins_elems(manager)[0]
        dels = _nested_dels(a_ins)
        assert len(dels) == 1
        assert _del_text(dels[0]) == " world"
        assert _visible_text(manager) == "Hello"
        children = [c for c in a_ins.childNodes if c.nodeType == c.ELEMENT_NODE]
        assert children[-1].tagName == "w:del"

    def test_delete_across_runs(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.suggest_deletion("lo wor")

        _assert_no_nested_ins(manager)
        ins_elems = _ins_elems(manager)
        assert len(ins_elems) == 1
        a_ins = ins_elems[0]
        assert a_ins.getAttribute("w:author") == AUTHOR_A
        dels = _nested_dels(a_ins)
        assert len(dels) >= 1
        for d in dels:
            assert d.getAttribute("w:author") == AUTHOR_B
        assert "".join(_del_text(d) for d in dels) == "lo wor"
        assert _visible_text(manager) == "Helld"

    def test_delete_entire_insertion_content(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.suggest_deletion("Hello")

        # A's ins survives, holding only B's nested deletion
        ins_elems = _ins_elems(manager)
        assert len(ins_elems) == 1
        a_ins = ins_elems[0]
        assert a_ins.getAttribute("w:author") == AUTHOR_A
        dels = _nested_dels(a_ins)
        assert len(dels) == 1
        assert _del_text(dels[0]) == "Hello"
        assert _visible_text(manager) == ""

    def test_delete_returns_real_change_id(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        change_id = manager.suggest_deletion("world")

        dels = _del_elems(manager)
        assert len(dels) == 1
        assert change_id == int(dels[0].getAttribute("w:id"))

    def test_delete_across_two_foreign_insertions(self, temp_xml):
        """A match spanning two adjacent foreign insertions nests a del in each."""
        body = (
            f"<w:p>{_foreign_ins('<w:r><w:t>abc </w:t></w:r>', ins_id=1)}"
            f"{_foreign_ins('<w:r><w:t>def</w:t></w:r>', ins_id=2)}</w:p>"
        )
        manager = _make_manager(temp_xml(body))

        manager.suggest_deletion("abc def")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == ""
        a_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_A]
        assert len(a_ins) == 2
        nested = [d for ins in a_ins for d in _nested_dels(ins)]
        assert len(nested) == 2
        assert all(d.getAttribute("w:author") == AUTHOR_B for d in nested)
        assert _del_text(nested[0]) == "abc "
        assert _del_text(nested[1]) == "def"


class TestReplaceInsideForeignIns:
    """B replacing text inside A's insertion: nested del + sibling B ins."""

    def test_replace_mid_splits_foreign_ins(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello world today</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        change_id = manager.replace_text("world", "earth")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "Hello earth today"

        ins_elems = _ins_elems(manager)
        a_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_A]
        b_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_B]
        # A's insertion is split into two halves around B's replacement
        assert len(a_ins) == 2
        assert len(b_ins) == 1
        # Halves keep A's identity with fresh, distinct ids
        assert a_ins[0].getAttribute("w:date") == DATE_A
        assert a_ins[1].getAttribute("w:date") == DATE_A
        assert a_ins[0].getAttribute("w:id") != a_ins[1].getAttribute("w:id")
        # The nested deletion carries B's authorship
        dels = _nested_dels(a_ins[0])
        assert len(dels) == 1
        assert dels[0].getAttribute("w:author") == AUTHOR_B
        assert _del_text(dels[0]) == "world"
        # The second half holds the split-off tail
        assert _ins_visible_text(a_ins[1]) == " today"
        # B's replacement sits between the halves, never nested
        assert _ins_visible_text(b_ins[0]) == "earth"
        assert change_id == int(b_ins[0].getAttribute("w:id"))

    def test_replace_at_end_no_split_needed(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.replace_text("world", "earth")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "Hello earth"
        ins_elems = _ins_elems(manager)
        a_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_A]
        b_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_B]
        # No trailing text — no split required
        assert len(a_ins) == 1
        assert len(b_ins) == 1
        assert _del_text(_nested_dels(a_ins[0])[0]) == "world"
        assert _ins_visible_text(b_ins[0]) == "earth"

    def test_replace_across_runs(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.replace_text("lo wor", "LO WOR")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "HelLO WORld"
        a_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_A]
        for ins in a_ins:
            assert ins.getAttribute("w:date") == DATE_A
        all_nested = [d for ins in a_ins for d in _nested_dels(ins)]
        assert "".join(_del_text(d) for d in all_nested) == "lo wor"


class TestInsertInsideForeignIns:
    """B inserting with an anchor inside A's insertion."""

    def test_insert_after_mid_splits_foreign_ins(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        change_id = manager.insert_text_after("Hello", " brave")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "Hello brave world"
        ins_elems = _ins_elems(manager)
        a_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_A]
        b_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_B]
        assert len(a_ins) == 2
        assert len(b_ins) == 1
        assert _ins_visible_text(a_ins[0]) == "Hello"
        assert _ins_visible_text(b_ins[0]) == " brave"
        assert _ins_visible_text(a_ins[1]) == " world"
        # Split halves keep A's identity, fresh distinct ids
        assert a_ins[0].getAttribute("w:date") == DATE_A
        assert a_ins[1].getAttribute("w:date") == DATE_A
        assert a_ins[0].getAttribute("w:id") != a_ins[1].getAttribute("w:id")
        assert change_id == int(b_ins[0].getAttribute("w:id"))

    def test_insert_before_mid_splits_foreign_ins(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        change_id = manager.insert_text_before("world", "brave ")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "Hello brave world"
        b_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_B]
        assert len(b_ins) == 1
        assert _ins_visible_text(b_ins[0]) == "brave "
        assert change_id == int(b_ins[0].getAttribute("w:id"))

    def test_insert_after_at_run_boundary_inside_foreign_ins(self, temp_xml):
        """Anchor ends exactly at a run boundary (not the ins end): no mid-run split."""
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello</w:t></w:r><w:r><w:t> world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.insert_text_after("Hello", "X")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "HelloX world"
        ins_elems = _ins_elems(manager)
        a_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_A]
        b_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_B]
        assert len(a_ins) == 2
        assert len(b_ins) == 1
        assert _ins_visible_text(a_ins[0]) == "Hello"
        assert _ins_visible_text(a_ins[1]) == " world"

    def test_insert_before_at_run_boundary_inside_foreign_ins(self, temp_xml):
        """Anchor starts exactly at a run boundary (not the ins start): no mid-run split."""
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.insert_text_before("world", "X")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "Hello Xworld"
        ins_elems = _ins_elems(manager)
        a_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_A]
        b_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_B]
        assert len(a_ins) == 2
        assert len(b_ins) == 1
        assert _ins_visible_text(a_ins[0]) == "Hello "
        assert _ins_visible_text(a_ins[1]) == "world"

    def test_insert_mid_run_with_non_text_child(self, temp_xml):
        """Mid-run split of a run that also carries a leading w:tab."""
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:tab/><w:t>Hello world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.insert_text_after("Hello", "X")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "HelloX world"
        a_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_A]
        assert len(a_ins) == 2
        assert len(a_ins[0].getElementsByTagName("w:tab")) == 1

    def test_insert_mid_run_tab_between_texts_stays_in_order(self, temp_xml):
        """A w:tab between two w:t nodes lands on the correct side of the split.

        Splitting A's ins after "Hello" must leave the tab (which follows the
        split point) in A's *second* ins, before "world" — not front-hoisted
        into the first (issue #20).
        """
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello</w:t><w:tab/><w:t>world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.insert_text_after("Hello", "X")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "HelloXworld"
        ins_elems = _ins_elems(manager)
        a_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_A]
        b_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_B]
        assert len(a_ins) == 2
        assert len(b_ins) == 1
        # B's ins sits between the two A halves in document order
        assert [e.getAttribute("w:author") for e in ins_elems] == [AUTHOR_A, AUTHOR_B, AUTHOR_A]
        # First A half: just "Hello", no tab
        assert _ins_visible_text(a_ins[0]) == "Hello"
        assert len(a_ins[0].getElementsByTagName("w:tab")) == 0
        # Second A half: the tab followed by "world"
        assert _ins_visible_text(a_ins[1]) == "world"
        assert len(a_ins[1].getElementsByTagName("w:tab")) == 1
        assert _ins_visible_text(b_ins[0]) == "X"

    def test_insert_mid_run_split_skips_rpr_and_empty_wt(self, temp_xml):
        """The split walk skips w:rPr, whitespace text nodes, and empty w:t
        children while keeping formatting on both rebuilt halves."""
        run_xml = "<w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t> <w:t></w:t><w:tab/><w:t>world</w:t></w:r>"
        xml_path = temp_xml(f"<w:p>{_foreign_ins(run_xml)}</w:p>")
        manager = _make_manager(xml_path)

        manager.insert_text_after("Hello", "X")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "HelloXworld"
        a_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_A]
        assert len(a_ins) == 2
        assert _ins_visible_text(a_ins[0]) == "Hello"
        assert len(a_ins[0].getElementsByTagName("w:tab")) == 0
        assert _ins_visible_text(a_ins[1]) == "world"
        assert len(a_ins[1].getElementsByTagName("w:tab")) == 1
        # Formatting survives the rebuild on both sides of the split
        for ins in a_ins:
            assert ins.getElementsByTagName("w:b")

    def test_insert_before_at_ins_start_no_split(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        change_id = manager.insert_text_before("Hello", "Say ")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "Say Hello"
        ins_elems = _ins_elems(manager)
        a_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_A]
        b_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_B]
        # Boundary anchor: plain sibling, A's ins untouched
        assert len(a_ins) == 1
        assert a_ins[0].getAttribute("w:id") == "1"
        assert len(b_ins) == 1
        assert change_id == int(b_ins[0].getAttribute("w:id"))

    def test_insert_after_at_ins_end_no_split(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        change_id = manager.insert_text_after("Hello", " world")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "Hello world"
        ins_elems = _ins_elems(manager)
        a_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_A]
        b_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_B]
        assert len(a_ins) == 1
        assert a_ins[0].getAttribute("w:id") == "1"
        assert len(b_ins) == 1
        assert change_id == int(b_ins[0].getAttribute("w:id"))

    def test_insert_anchor_across_runs_inside_foreign_ins(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        change_id = manager.insert_text_after("lo wor", "!!")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "Hello wor!!ld"
        b_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_B]
        assert len(b_ins) == 1
        assert change_id == int(b_ins[0].getAttribute("w:id"))


class TestMixedStateAcrossForeignBoundary:
    """Matches straddling regular text and a foreign insertion."""

    def test_replace_regular_into_foreign_ins(self, temp_xml):
        body = f"<w:p><w:r><w:t>Hello </w:t></w:r>{_foreign_ins('<w:r><w:t>world</w:t></w:r>')}</w:p>"
        xml_path = temp_xml(body)
        manager = _make_manager(xml_path)

        manager.replace_text("lo wor", "LO WOR")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "HelLO WORld"
        # Regular part -> top-level del; foreign part -> nested del by B
        a_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_A]
        assert len(a_ins) == 1
        nested = _nested_dels(a_ins[0])
        assert len(nested) == 1
        assert nested[0].getAttribute("w:author") == AUTHOR_B
        assert _del_text(nested[0]) == "wor"
        top_level = [d for d in _del_elems(manager) if d not in nested]
        assert "".join(_del_text(d) for d in top_level) == "lo "
        assert _ins_visible_text(a_ins[0]) == "ld"

    def test_replace_foreign_ins_into_regular(self, temp_xml):
        body = f"<w:p>{_foreign_ins('<w:r><w:t>Hello</w:t></w:r>')}<w:r><w:t> world</w:t></w:r></w:p>"
        xml_path = temp_xml(body)
        manager = _make_manager(xml_path)

        manager.replace_text("lo wor", "LO WOR")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "HelLO WORld"
        a_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_A]
        assert len(a_ins) == 1
        nested = _nested_dels(a_ins[0])
        assert len(nested) == 1
        assert _del_text(nested[0]) == "lo"
        assert _ins_visible_text(a_ins[0]) == "Hel"
        top_level = [d for d in _del_elems(manager) if d not in nested]
        assert "".join(_del_text(d) for d in top_level) == " wor"

    def test_delete_regular_into_foreign_ins(self, temp_xml):
        body = f"<w:p><w:r><w:t>Hello </w:t></w:r>{_foreign_ins('<w:r><w:t>world</w:t></w:r>')}</w:p>"
        xml_path = temp_xml(body)
        manager = _make_manager(xml_path)

        change_id = manager.suggest_deletion("lo wor")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "Helld"
        a_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_A]
        assert len(a_ins) == 1
        nested = _nested_dels(a_ins[0])
        assert len(nested) == 1
        assert _del_text(nested[0]) == "wor"
        assert change_id >= 0

    def test_mixed_replace_accept_and_reject_round_trip(self, temp_xml):
        body = f"<w:p><w:r><w:t>Hello </w:t></w:r>{_foreign_ins('<w:r><w:t>world</w:t></w:r>')}</w:p>"

        # Accepting everything lands on B's outcome
        manager = _make_manager(temp_xml(body))
        manager.replace_text("lo wor", "LO WOR")
        manager.accept_all()
        assert _visible_text(manager) == "HelLO WORld"
        assert not _ins_elems(manager)
        assert not _del_elems(manager)

        # Rejecting everything restores the pre-A original (regular text only)
        manager = _make_manager(temp_xml(body))
        manager.replace_text("lo wor", "LO WOR")
        manager.reject_all()
        assert _visible_text(manager) == "Hello "
        assert not _ins_elems(manager)
        assert not _del_elems(manager)

        # Rejecting only B leaves A's pending insertion intact
        manager = _make_manager(temp_xml(body))
        manager.replace_text("lo wor", "LO WOR")
        manager.reject_all(author=AUTHOR_B)
        assert _visible_text(manager) == "Hello world"
        a_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_A]
        assert len(a_ins) == 1
        assert _ins_visible_text(a_ins[0]) == "world"


class TestRewriteParagraphWithForeignIns:
    """rewrite_paragraph diffs that land inside a foreign insertion."""

    def _rewrite(self, manager: RevisionManager, new_text: str) -> None:
        p = manager.editor.dom.getElementsByTagName("w:p")[0]
        ref = f"P1#{compute_paragraph_hash(p)}"
        manager.rewrite_paragraph(ref, new_text)

    def test_rewrite_insert_inside_foreign_ins(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        self._rewrite(manager, "Hello brave world")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "Hello brave world"
        ins_elems = _ins_elems(manager)
        a_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_A]
        b_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_B]
        # A's content stays under A's authorship; B's words under B's
        assert "".join(_ins_visible_text(e) for e in a_ins) == "Hello world"
        assert "".join(_ins_visible_text(e) for e in b_ins) == "brave "

    def test_rewrite_delete_inside_foreign_ins(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello brave world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        self._rewrite(manager, "Hello world")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "Hello world"
        a_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_A]
        all_nested = [d for ins in a_ins for d in _nested_dels(ins)]
        assert len(all_nested) >= 1
        for d in all_nested:
            assert d.getAttribute("w:author") == AUTHOR_B
        assert "".join(_del_text(d) for d in all_nested) == "brave "

    def test_rewrite_append_at_end_of_foreign_ins(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{_foreign_ins('<w:r><w:t>Hello</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        self._rewrite(manager, "Hello world")

        _assert_no_nested_ins(manager)
        assert _visible_text(manager) == "Hello world"
        ins_elems = _ins_elems(manager)
        a_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_A]
        b_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_B]
        assert len(a_ins) == 1
        assert _ins_visible_text(a_ins[0]) == "Hello"
        assert "".join(_ins_visible_text(e) for e in b_ins) == " world"


class TestResolutionMatrix:
    """accept_all / reject_all outcomes for B's edits inside A's insertion."""

    DELETE_BODY = f"<w:p>{_foreign_ins('<w:r><w:t>Hello world</w:t></w:r>')}</w:p>"

    def _managed_delete(self, temp_xml) -> RevisionManager:
        manager = _make_manager(temp_xml(self.DELETE_BODY))
        manager.suggest_deletion("world")
        return manager

    def test_accept_all_applies_both(self, temp_xml):
        manager = self._managed_delete(temp_xml)
        manager.accept_all()
        assert _visible_text(manager) == "Hello "
        assert not _ins_elems(manager)
        assert not _del_elems(manager)

    def test_reject_all_restores_original(self, temp_xml):
        manager = self._managed_delete(temp_xml)
        manager.reject_all()
        # A's insertion (and B's edit with it) fully gone
        assert _visible_text(manager) == ""
        assert not _ins_elems(manager)
        assert not _del_elems(manager)

    def test_reject_b_keeps_a_intact(self, temp_xml):
        manager = self._managed_delete(temp_xml)
        rejected = manager.reject_all(author=AUTHOR_B)
        assert rejected == 1
        a_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_A]
        assert len(a_ins) == 1
        assert _ins_visible_text(a_ins[0]) == "Hello world"
        assert not _del_elems(manager)

    def test_accept_a_leaves_b_pending_top_level(self, temp_xml):
        manager = self._managed_delete(temp_xml)
        manager.accept_all(author=AUTHOR_A)
        # A's text becomes regular; B's deletion is now a pending top-level del
        assert _visible_text(manager) == "Hello "
        assert not _ins_elems(manager)
        dels = _del_elems(manager)
        assert len(dels) == 1
        assert dels[0].getAttribute("w:author") == AUTHOR_B
        assert _del_text(dels[0]) == "world"

    def test_accept_b_resolves_del_keeps_a_pending(self, temp_xml):
        manager = self._managed_delete(temp_xml)
        manager.accept_all(author=AUTHOR_B)
        assert _visible_text(manager) == "Hello "
        a_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_A]
        assert len(a_ins) == 1
        assert _ins_visible_text(a_ins[0]) == "Hello "
        assert not _del_elems(manager)

    def test_replace_matrix_accept_all(self, temp_xml):
        manager = _make_manager(temp_xml(self.DELETE_BODY))
        manager.replace_text("world", "earth")
        manager.accept_all()
        assert _visible_text(manager) == "Hello earth"
        assert not _ins_elems(manager)
        assert not _del_elems(manager)

    def test_replace_matrix_reject_all(self, temp_xml):
        manager = _make_manager(temp_xml(self.DELETE_BODY))
        manager.replace_text("world", "earth")
        manager.reject_all()
        assert _visible_text(manager) == ""
        assert not _ins_elems(manager)
        assert not _del_elems(manager)

    def test_replace_matrix_reject_b_only(self, temp_xml):
        manager = _make_manager(temp_xml(self.DELETE_BODY))
        manager.replace_text("world", "earth")
        manager.reject_all(author=AUTHOR_B)
        a_ins = [e for e in _ins_elems(manager) if e.getAttribute("w:author") == AUTHOR_A]
        assert "".join(_ins_visible_text(e) for e in a_ins) == "Hello world"
        assert not _del_elems(manager)
        assert _visible_text(manager) == "Hello world"

    def test_replace_matrix_accept_b_only(self, temp_xml):
        manager = _make_manager(temp_xml(self.DELETE_BODY))
        manager.replace_text("world", "earth")
        manager.accept_all(author=AUTHOR_B)
        # B's del and ins resolve; A's remaining insertion stays pending
        assert _visible_text(manager) == "Hello earth"
        assert not _del_elems(manager)
        remaining = _ins_elems(manager)
        assert remaining
        for ins in remaining:
            assert ins.getAttribute("w:author") == AUTHOR_A


class TestSameAuthorInPlaceUnchanged:
    """Author matches -> historical in-place editing, no nesting of anything."""

    def _own_ins(self, content: str) -> str:
        return f'<w:ins w:id="1" w:author="{AUTHOR_B}" w:date="{DATE_A}">{content}</w:ins>'

    def test_own_delete_stays_physical(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{self._own_ins('<w:r><w:t>Hello world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.suggest_deletion(" world")

        assert not _del_elems(manager)
        _assert_no_nested_ins(manager)
        _assert_no_del_in_own_ins(manager, AUTHOR_B)
        assert _visible_text(manager) == "Hello"

    def test_own_replace_stays_in_place(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{self._own_ins('<w:r><w:t>Hello world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.replace_text("world", "earth")

        assert not _del_elems(manager)
        _assert_no_nested_ins(manager)
        _assert_no_del_in_own_ins(manager, AUTHOR_B)
        text = _visible_text(manager)
        assert "earth" in text
        assert "world" not in text

    def test_own_insert_splices(self, temp_xml):
        xml_path = temp_xml(f"<w:p>{self._own_ins('<w:r><w:t>Hello world</w:t></w:r>')}</w:p>")
        manager = _make_manager(xml_path)

        manager.insert_text_after("Hello", " brave")

        assert not _del_elems(manager)
        _assert_no_nested_ins(manager)
        assert len(_ins_elems(manager)) == 1
        assert _visible_text(manager) == "Hello brave world"

    def test_adjacent_own_and_foreign_ins_delete(self, temp_xml):
        body = (
            f'<w:p><w:ins w:id="1" w:author="{AUTHOR_B}" w:date="{DATE_A}">'
            "<w:r><w:t>abc </w:t></w:r></w:ins>"
            f"{_foreign_ins('<w:r><w:t>def</w:t></w:r>', ins_id=2)}</w:p>"
        )
        xml_path = temp_xml(body)
        manager = _make_manager(xml_path)

        manager.suggest_deletion("abc def")

        _assert_no_nested_ins(manager)
        _assert_no_del_in_own_ins(manager, AUTHOR_B)
        assert _visible_text(manager) == ""
        # Own half physically removed; foreign half survives with nested del
        ins_elems = _ins_elems(manager)
        b_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_B]
        a_ins = [e for e in ins_elems if e.getAttribute("w:author") == AUTHOR_A]
        assert not b_ins
        assert len(a_ins) == 1
        nested = _nested_dels(a_ins[0])
        assert len(nested) == 1
        assert _del_text(nested[0]) == "def"


class TestMissingAuthorTreatedAsForeign:
    """An ins without w:author can't be attributed to us — never edit in place."""

    def test_delete_in_authorless_ins_nests(self, temp_xml):
        body = '<w:p><w:ins w:id="1" w:date="2024-01-01T00:00:00Z"><w:r><w:t>Hello world</w:t></w:r></w:ins></w:p>'
        xml_path = temp_xml(body)
        manager = _make_manager(xml_path)

        manager.suggest_deletion("world")

        ins_elems = _ins_elems(manager)
        assert len(ins_elems) == 1
        dels = _nested_dels(ins_elems[0])
        assert len(dels) == 1
        assert _del_text(dels[0]) == "world"


class TestDocumentIntegration:
    """Two-session flow through the public Document API."""

    def test_two_author_review_flow(self, temp_docx):
        # Session 1: A proposes an insertion
        doc = Document.open(temp_docx, author=AUTHOR_A)
        try:
            ref = find_ref(doc, "quick brown fox")
            doc.insert_after("fox", " PROPOSAL", paragraph=ref)
            doc.save()
        finally:
            doc.close()

        # Session 2: B edits inside A's pending insertion
        doc = Document.open(temp_docx, author=AUTHOR_B)
        try:
            ref = find_ref(doc, "PROPOSAL")
            doc.replace("PROPOSAL", "REVISED", paragraph=ref)
            doc.save()
        finally:
            doc.close()

        # Session 3: both authors' revisions are present and resolvable
        doc = Document.open(temp_docx, author="Reviewer C")
        try:
            revisions = doc.list_revisions()
            authors = {r.author for r in revisions}
            assert AUTHOR_A in authors
            assert AUTHOR_B in authors
            deletions = [r for r in revisions if r.type == "deletion"]
            assert any(r.author == AUTHOR_B and "PROPOSAL" in r.text for r in deletions)

            doc.accept_all()
            assert doc.list_revisions() == []
            full_text = "\n".join(doc.list_paragraphs(max_chars=500))
            assert "REVISED" in full_text
            assert "PROPOSAL" not in full_text
            doc.save()
        finally:
            doc.close()

    def test_reject_b_preserves_a_proposal_via_document(self, temp_docx):
        doc = Document.open(temp_docx, author=AUTHOR_A)
        try:
            ref = find_ref(doc, "quick brown fox")
            doc.insert_after("fox", " PROPOSAL", paragraph=ref)
            doc.save()
        finally:
            doc.close()

        doc = Document.open(temp_docx, author=AUTHOR_B)
        try:
            ref = find_ref(doc, "PROPOSAL")
            doc.delete("PROPOSAL", paragraph=ref)
            doc.reject_all(author=AUTHOR_B)
            # A's proposal must survive B's retracted edit
            revisions = doc.list_revisions()
            assert len(revisions) == 1
            assert revisions[0].author == AUTHOR_A
            assert "PROPOSAL" in revisions[0].text
        finally:
            doc.close()


class TestForeignInsMajorityRPrAndTrimming:
    """Site D foreign-ins replaces get majority rPr and affix trimming."""

    def test_replacement_carries_majority_rpr(self, temp_xml):
        xml_path = temp_xml(
            "<w:p>"
            + _foreign_ins(
                '<w:r><w:t xml:space="preserve">in </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bolded stuff</w:t></w:r>'
            )
            + "</w:p>"
        )
        mgr = _make_manager(xml_path)
        # find/replace share no words, so trimming cannot mask the rule
        mgr.replace_text("in bolded", "z")

        own_ins = [e for e in _ins_elems(mgr) if e.getAttribute("w:author") == AUTHOR_B]
        assert len(own_ins) == 1
        assert len(own_ins[0].getElementsByTagName("w:b")) == 1
        _assert_no_nested_ins(mgr)

    def test_trimming_inside_foreign_ins(self, temp_xml):
        xml_path = temp_xml("<w:p>" + _foreign_ins("<w:r><w:t>keep old keep</w:t></w:r>") + "</w:p>")
        mgr = _make_manager(xml_path)
        mgr.replace_text("keep old keep", "keep new keep")

        # Only the changed word is revised
        dels = _del_elems(mgr)
        assert len(dels) == 1
        assert _del_text(dels[0]) == "old"
        own_ins = [e for e in _ins_elems(mgr) if e.getAttribute("w:author") == AUTHOR_B]
        assert len(own_ins) == 1
        assert _ins_visible_text(own_ins[0]) == "new"
        assert _visible_text(mgr) == "keep new keep"
        _assert_no_nested_ins(mgr)

    def test_trim_narrows_boundary_spanning_match_into_ins(self, temp_xml):
        # Pre-trim the match spans plain text + A's insertion; post-trim only
        # "old" (inside the ins) is revised, so dispatch must follow the
        # narrowed match, nesting the del inside A's ins.
        xml_path = temp_xml(
            '<w:p><w:r><w:t xml:space="preserve">keep </w:t></w:r>'
            + _foreign_ins("<w:r><w:t>old stuff</w:t></w:r>")
            + "</w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.replace_text("keep old", "keep new")

        dels = _del_elems(mgr)
        assert len(dels) == 1
        assert _del_text(dels[0]) == "old"
        a_ins = [e for e in _ins_elems(mgr) if e.getAttribute("w:author") == AUTHOR_A]
        assert dels[0] in _nested_dels(a_ins[0])
        own_ins = [e for e in _ins_elems(mgr) if e.getAttribute("w:author") == AUTHOR_B]
        assert len(own_ins) == 1
        assert _ins_visible_text(own_ins[0]) == "new"
        assert _visible_text(mgr) == "keep new stuff"
        _assert_no_nested_ins(mgr)
