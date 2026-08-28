"""Tests for same-author in-place editing inside w:ins.

These tests verify that operations inside the current author's *own* w:ins
elements edit the insertion in place: no nested w:ins (invalid OOXML) and no
w:del inside our own w:ins (our own pending text is simply rewritten, not
counter-proposed). Edits inside *another* author's w:ins legitimately nest a
w:del — that behavior is covered by test_foreign_ins_edits.py. One fixture
here starts out multi-author (a foreign w:del already nested in our own
w:ins) to pin that amending around it leaves their deletion intact.
"""

from pathlib import Path

import pytest

from docx_editor.exceptions import TextNotFoundError
from docx_editor.track_changes import RevisionManager
from docx_editor.xml_editor import DocxXMLEditor, build_text_map

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

# Opening tag of an insertion authored by _make_manager's own author.
OWN_INS = '<w:ins w:id="1" w:author="Test Author" w:date="2024-01-01T00:00:00Z">'

# A run-level drawing whose text box holds "MID".
_BOXED_MID = "<w:drawing><w:txbxContent><w:p><w:r><w:t>MID</w:t></w:r></w:p></w:txbxContent></w:drawing>"


@pytest.fixture
def temp_xml(tmp_path):
    """Fixture that returns a function to create temp XML files."""

    def _create_xml(body_xml: str) -> Path:
        xml = f'<?xml version="1.0" encoding="utf-8"?><w:document {NS}><w:body>{body_xml}</w:body></w:document>'
        xml_path = tmp_path / "test_doc.xml"
        xml_path.write_text(xml)
        return xml_path

    return _create_xml


def _make_manager(xml_path: Path) -> RevisionManager:
    """Create a RevisionManager from an XML file path.

    Args:
        xml_path: Path to the XML file

    Returns:
        RevisionManager instance with the XML loaded
    """
    editor = DocxXMLEditor(xml_path, rsid="00000000", author="Test Author")
    return RevisionManager(editor)


def _assert_no_nested_ins(manager: RevisionManager) -> None:
    """Assert no w:ins or w:del is nested inside another w:ins.

    The fixtures this helper guards are authored by the manager's own author
    ("Test Author"), so in-place editing must never nest anything: no w:ins in
    w:ins (invalid OOXML), and no w:del in our *own* w:ins (nested w:del is
    reserved for edits inside a foreign author's insertion).

    Args:
        manager: RevisionManager to check

    Raises:
        AssertionError: If nested elements are found
    """
    dom = manager.editor.dom
    for ins in dom.getElementsByTagName("w:ins"):
        # Check no child w:ins
        nested_ins = ins.getElementsByTagName("w:ins")
        assert len(nested_ins) == 0, f"Found nested w:ins inside w:ins: {ins.toxml()}"
        # Check no child w:del
        nested_del = ins.getElementsByTagName("w:del")
        assert len(nested_del) == 0, f"Found nested w:del inside w:ins: {ins.toxml()}"


def _assert_no_empty_ins(manager: RevisionManager) -> None:
    """Assert no w:ins was left behind with all of its content removed.

    Args:
        manager: RevisionManager to check

    Raises:
        AssertionError: If a content-less w:ins is found
    """
    for ins in manager.editor.dom.getElementsByTagName("w:ins"):
        assert any(child.nodeType == child.ELEMENT_NODE for child in ins.childNodes), (
            f"Empty w:ins left behind: {ins.toxml()}"
        )


def _get_text_content(manager: RevisionManager) -> str:
    """Extract all visible text from the document.

    Args:
        manager: RevisionManager to extract text from

    Returns:
        Concatenated text content from all w:t elements not in w:del
    """
    dom = manager.editor.dom
    result = []

    # Get all w:t elements
    for wt in dom.getElementsByTagName("w:t"):
        # Check if inside w:del
        parent = wt.parentNode
        inside_del = False
        while parent:
            if (
                parent.localName == "del"
                and parent.namespaceURI == "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            ):
                inside_del = True
                break
            parent = parent.parentNode

        if not inside_del and wt.firstChild:
            result.append(wt.firstChild.data)

    return "".join(result)


class TestReplaceInsideIns:
    """Tests for replace operations inside existing w:ins elements."""

    def test_replace_inside_ins(self, temp_xml):
        """Test single-element replace inside w:ins."""
        body_xml = (
            '<w:p><w:ins w:id="1" w:author="Test Author" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello world</w:t></w:r></w:ins></w:p>"
        )
        xml_path = temp_xml(body_xml)
        manager = _make_manager(xml_path)

        # Replace "world" with "earth"
        manager.replace_text("world", "earth")

        # Assert no nested ins/del
        _assert_no_nested_ins(manager)

        # Assert "earth" appears in text
        text = _get_text_content(manager)
        assert "earth" in text


class TestDeleteInsideIns:
    """Tests for deletion operations inside existing w:ins elements."""

    def test_delete_inside_ins(self, temp_xml):
        """Test single-element delete inside w:ins."""
        body_xml = (
            '<w:p><w:ins w:id="1" w:author="Test Author" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello world</w:t></w:r></w:ins></w:p>"
        )
        xml_path = temp_xml(body_xml)
        manager = _make_manager(xml_path)

        # Delete "world"
        manager.suggest_deletion("world")

        # Assert no nested ins/del
        _assert_no_nested_ins(manager)

        # "world" should not appear in visible text
        text = _get_text_content(manager)
        assert "world" not in text


class TestInsertInsideIns:
    """Tests for insertion operations inside existing w:ins elements."""

    def test_insert_after_inside_ins(self, temp_xml):
        """Test insert_after with anchor inside w:ins."""
        body_xml = (
            '<w:p><w:ins w:id="1" w:author="Test Author" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello</w:t></w:r></w:ins></w:p>"
        )
        xml_path = temp_xml(body_xml)
        manager = _make_manager(xml_path)

        # Insert after "Hello"
        manager.insert_text_after("Hello", " world")

        # Assert no nested ins/del
        _assert_no_nested_ins(manager)

        # "world" should appear in text
        text = _get_text_content(manager)
        assert "world" in text

    def test_insert_before_inside_ins(self, temp_xml):
        """Test insert_before with anchor inside w:ins."""
        body_xml = (
            '<w:p><w:ins w:id="1" w:author="Test Author" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello</w:t></w:r></w:ins></w:p>"
        )
        xml_path = temp_xml(body_xml)
        manager = _make_manager(xml_path)

        # Insert before "Hello"
        manager.insert_text_before("Hello", "Say ")

        # Assert no nested ins/del
        _assert_no_nested_ins(manager)

        # "Say" should appear in text
        text = _get_text_content(manager)
        assert "Say" in text


class TestCrossBoundaryInsideIns:
    """Tests for cross-boundary operations inside existing w:ins elements."""

    def test_cross_boundary_replace_all_inside_ins(self, temp_xml):
        """Test replace spanning two runs both inside w:ins (site D)."""
        body_xml = (
            '<w:p><w:ins w:id="1" w:author="Test Author" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:ins></w:p>"
        )
        xml_path = temp_xml(body_xml)
        manager = _make_manager(xml_path)

        # Replace "lo wor" spanning two runs
        manager.replace_text("lo wor", "LO WOR")

        # Assert no nested ins/del
        _assert_no_nested_ins(manager)

        # "LO WOR" should appear in text
        text = _get_text_content(manager)
        assert "LO WOR" in text

    def test_cross_boundary_delete_all_inside_ins(self, temp_xml):
        """Test delete spanning two runs both inside w:ins (site F)."""
        body_xml = (
            '<w:p><w:ins w:id="1" w:author="Test Author" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:ins></w:p>"
        )
        xml_path = temp_xml(body_xml)
        manager = _make_manager(xml_path)

        # Delete "lo wor" spanning two runs
        manager.suggest_deletion("lo wor")

        # Assert no nested ins/del
        _assert_no_nested_ins(manager)

        # "lo wor" should not appear in visible text
        text = _get_text_content(manager)
        assert "lo wor" not in text

    def test_cross_boundary_insert_near_match_inside_ins(self, temp_xml):
        """Test insert near cross-boundary match inside w:ins (sites H/I)."""
        body_xml = (
            '<w:p><w:ins w:id="1" w:author="Test Author" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:ins></w:p>"
        )
        xml_path = temp_xml(body_xml)
        manager = _make_manager(xml_path)

        # Insert after "lo wor" spanning two runs
        manager.insert_text_after("lo wor", "!!")

        # Assert no nested ins/del
        _assert_no_nested_ins(manager)

        # "!!" should appear in text
        text = _get_text_content(manager)
        assert "!!" in text


class TestMixedBoundaryScenarios:
    """Tests for operations that cross w:ins boundaries."""

    def test_replace_crossing_ins_boundary_start(self, temp_xml):
        """Test replace that starts inside w:ins and ends outside."""
        body_xml = (
            '<w:p><w:ins w:id="1" w:author="Test Author" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello</w:t></w:r></w:ins><w:r><w:t> world</w:t></w:r></w:p>"
        )
        xml_path = temp_xml(body_xml)
        manager = _make_manager(xml_path)

        # Replace "lo wor" crossing boundary
        manager.replace_text("lo wor", "LO WOR")

        # Assert no nested ins/del
        _assert_no_nested_ins(manager)

        # "LO WOR" should appear
        text = _get_text_content(manager)
        assert "LO WOR" in text

    def test_replace_crossing_ins_boundary_end(self, temp_xml):
        """Test replace that starts outside w:ins and ends inside."""
        body_xml = (
            '<w:p><w:r><w:t>Hello </w:t></w:r><w:ins w:id="1" w:author="Test Author" '
            'w:date="2024-01-01T00:00:00Z"><w:r><w:t>world</w:t></w:r></w:ins></w:p>'
        )
        xml_path = temp_xml(body_xml)
        manager = _make_manager(xml_path)

        # Replace "lo wor" crossing boundary
        manager.replace_text("lo wor", "LO WOR")

        # Assert no nested ins/del
        _assert_no_nested_ins(manager)

        # "LO WOR" should appear
        text = _get_text_content(manager)
        assert "LO WOR" in text

    def test_replace_surrounding_ins(self, temp_xml):
        """Test replace that contains an entire w:ins element."""
        body_xml = (
            '<w:p><w:r><w:t>Hello </w:t></w:r><w:ins w:id="1" w:author="Test Author" '
            'w:date="2024-01-01T00:00:00Z"><w:r><w:t>big</w:t></w:r></w:ins>'
            "<w:r><w:t> world</w:t></w:r></w:p>"
        )
        xml_path = temp_xml(body_xml)
        manager = _make_manager(xml_path)

        # Replace "lo big wor" that contains the entire w:ins
        manager.replace_text("lo big wor", "LO BIG WOR")

        # Assert no nested ins/del
        _assert_no_nested_ins(manager)

        # "LO BIG WOR" should appear
        text = _get_text_content(manager)
        assert "LO BIG WOR" in text


class TestOwnInsReplaceAnchoring:
    """The replacement lands at the match position, not the insertion's start.

    Replacing text inside our own pending insertion splices the new text in
    place. The splice used to be pinned to the insertion's first child, so
    any match not starting at the insertion's very first character came out
    reordered — "Hello world" became "earthHello " (ISSUES.md #39). These
    tests assert exact visible text, which the older order-insensitive
    ``in`` assertions above never did.
    """

    def test_replace_spanning_later_runs_pending(self, temp_xml):
        """A match over runs 2-3 keeps run 1 in front of the replacement."""
        xml_path = temp_xml(
            f"<w:p>{OWN_INS}<w:r><w:t>Hello </w:t></w:r><w:r><w:t>wor</w:t></w:r><w:r><w:t>ld</w:t></w:r></w:ins></w:p>"
        )
        manager = _make_manager(xml_path)

        manager.replace_text("world", "earth")

        _assert_no_nested_ins(manager)
        assert _get_text_content(manager) == "Hello earth"

    def test_replace_spanning_later_runs_accepted(self, temp_xml):
        """Accepting the insertion keeps the spliced order."""
        xml_path = temp_xml(
            f"<w:p>{OWN_INS}<w:r><w:t>Hello </w:t></w:r><w:r><w:t>wor</w:t></w:r><w:r><w:t>ld</w:t></w:r></w:ins></w:p>"
        )
        manager = _make_manager(xml_path)

        manager.replace_text("world", "earth")
        manager.accept_all()

        assert _get_text_content(manager) == "Hello earth"

    def test_replace_spanning_later_runs_rejected(self, temp_xml):
        """Rejecting drops the whole insertion, replacement included.

        None of the insertion was ever in the original document, so a
        rejection leaves only the surrounding untracked text.
        """
        xml_path = temp_xml(
            f"<w:p><w:r><w:t>pre </w:t></w:r>{OWN_INS}<w:r><w:t>Hello </w:t></w:r>"
            "<w:r><w:t>wor</w:t></w:r><w:r><w:t>ld</w:t></w:r></w:ins>"
            "<w:r><w:t> post</w:t></w:r></w:p>"
        )
        manager = _make_manager(xml_path)

        manager.replace_text("world", "earth")
        assert _get_text_content(manager) == "pre Hello earth post"

        manager.reject_all()
        assert _get_text_content(manager) == "pre  post"

    def test_replace_at_end_of_single_run_ins(self, temp_xml):
        """A single-run insertion anchors at the match, not the run start."""
        xml_path = temp_xml(f"<w:p>{OWN_INS}<w:r><w:t>Hello world</w:t></w:r></w:ins></w:p>")
        manager = _make_manager(xml_path)

        manager.replace_text("world", "earth")

        _assert_no_nested_ins(manager)
        assert _get_text_content(manager) == "Hello earth"

    def test_replace_mid_single_run_ins(self, temp_xml):
        """Text surviving on both sides of the match keeps its order."""
        xml_path = temp_xml(f"<w:p>{OWN_INS}<w:r><w:t>Hello world today</w:t></w:r></w:ins></w:p>")
        manager = _make_manager(xml_path)

        manager.replace_text("world", "earth")

        _assert_no_nested_ins(manager)
        assert _get_text_content(manager) == "Hello earth today"

    def test_replace_mid_run_keeps_later_run_order(self, temp_xml):
        """A later run of the same insertion stays after the match's tail."""
        xml_path = temp_xml(f"<w:p>{OWN_INS}<w:r><w:t>a world b</w:t></w:r><w:r><w:t> tail</w:t></w:r></w:ins></w:p>")
        manager = _make_manager(xml_path)

        manager.replace_text("world", "earth")

        _assert_no_nested_ins(manager)
        assert _get_text_content(manager) == "a earth b tail"

    def test_replace_at_ins_start_unchanged(self, temp_xml):
        """A match at the insertion's start still replaces at the start."""
        xml_path = temp_xml(
            f"<w:p>{OWN_INS}<w:r><w:t>Hel</w:t></w:r><w:r><w:t>lo</w:t></w:r><w:r><w:t> world</w:t></w:r></w:ins></w:p>"
        )
        manager = _make_manager(xml_path)

        manager.replace_text("Hello", "Goodbye")

        _assert_no_nested_ins(manager)
        assert _get_text_content(manager) == "Goodbye world"

    def test_replace_preserves_sibling_wt_order(self, temp_xml):
        """An unmatched w:t sibling stays behind the replacement."""
        xml_path = temp_xml(f"<w:p>{OWN_INS}<w:r><w:t>ab</w:t><w:t>KEEP</w:t></w:r></w:ins></w:p>")
        manager = _make_manager(xml_path)

        manager.replace_text("b", "X")

        _assert_no_nested_ins(manager)
        assert _get_text_content(manager) == "aXKEEP"

    def test_replace_across_two_own_ins(self, temp_xml):
        """A match spanning two of our insertions leaves neither empty."""
        xml_path = temp_xml(
            f"<w:p>{OWN_INS}<w:r><w:t>Hello wor</w:t></w:r></w:ins>"
            '<w:ins w:id="2" w:author="Test Author" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>ld here</w:t></w:r></w:ins></w:p>"
        )
        manager = _make_manager(xml_path)

        manager.replace_text("world", "earth")

        _assert_no_nested_ins(manager)
        assert _get_text_content(manager) == "Hello earth here"
        _assert_no_empty_ins(manager)

    def test_replace_consuming_leading_own_ins_drops_it(self, temp_xml):
        """An insertion the match consumes whole is dropped, not left empty.

        The replacement lands in the *second* insertion (it holds the match's
        last node), so the first one is emptied by the splice.
        """
        xml_path = temp_xml(
            f"<w:p>{OWN_INS}<w:r><w:t>wor</w:t></w:r></w:ins>"
            '<w:ins w:id="2" w:author="Test Author" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>ld here</w:t></w:r></w:ins></w:p>"
        )
        manager = _make_manager(xml_path)

        manager.replace_text("world", "earth")

        _assert_no_nested_ins(manager)
        assert _get_text_content(manager) == "earth here"
        _assert_no_empty_ins(manager)
        assert manager.editor.dom.getElementsByTagName("w:ins").length == 1

    def test_replace_whole_ins_content_edits_in_place(self, temp_xml):
        """Matching an insertion entirely rewrites it, minting no new revision."""
        xml_path = temp_xml(f"<w:p>{OWN_INS}<w:r><w:t>AB</w:t></w:r><w:r><w:t>CD</w:t></w:r></w:ins></w:p>")
        manager = _make_manager(xml_path)

        change_id = manager.replace_text("ABCD", "NEW")

        assert change_id == -1
        assert _get_text_content(manager) == "NEW"
        ins_elems = manager.editor.dom.getElementsByTagName("w:ins")
        assert ins_elems.length == 1
        # The original insertion survives carrying the new text — its identity
        # is untouched because we only edited our own pending content.
        assert ins_elems[0].getAttribute("w:id") == "1"
        assert ins_elems[0].getAttribute("w:author") == "Test Author"
        assert ins_elems[0].getAttribute("w:date") == "2024-01-01T00:00:00Z"

    def test_boxed_text_is_not_part_of_the_host_paragraph(self, temp_xml):
        """The host paragraph's text runs straight past the drawing.

        Text-box content is excluded from every text map (ISSUES.md #65), so
        the host reads "abcd" — the boxed "MID" is not between "ab" and "cd".
        """
        xml_path = temp_xml(f"<w:p>{OWN_INS}<w:r><w:t>ab</w:t>{_BOXED_MID}<w:t>cd</w:t></w:r></w:ins></w:p>")
        manager = _make_manager(xml_path)

        paragraph = manager.editor.dom.getElementsByTagName("w:p")[0]
        assert build_text_map(paragraph).text == "abcd"

    @pytest.mark.parametrize("target", ["bMI", "bMIDc", "MID"])
    def test_a_match_through_the_box_no_longer_exists(self, temp_xml, target):
        """Searches that used to cross into the drawing now find nothing.

        These spans only ever existed because the boxed characters leaked into
        the host's map; with the box excluded there is no such text to edit,
        and the drawing is left untouched.
        """
        xml_path = temp_xml(f"<w:p>{OWN_INS}<w:r><w:t>ab</w:t>{_BOXED_MID}<w:t>cd</w:t></w:r></w:ins></w:p>")
        manager = _make_manager(xml_path)
        before = manager.editor.dom.getElementsByTagName("w:drawing")[0].toxml()

        with pytest.raises(TextNotFoundError):
            manager.replace_text(target, "Z")

        drawings = manager.editor.dom.getElementsByTagName("w:drawing")
        assert len(drawings) == 1
        assert drawings[0].toxml() == before

    def test_replace_beside_the_box_carries_it_through(self, temp_xml):
        """An edit in the host paragraph does not disturb the box's content.

        Rebuilding the run re-serializes the drawing, so the rsid pass stamps
        identifiers onto its elements; the boxed *text* still survives exactly
        once and never leaks into the host paragraph.
        """
        xml_path = temp_xml(f"<w:p>{OWN_INS}<w:r><w:t>ab</w:t>{_BOXED_MID}<w:t>cd</w:t></w:r></w:ins></w:p>")
        manager = _make_manager(xml_path)

        manager.replace_text("bc", "Z")

        _assert_no_nested_ins(manager)
        paragraph = manager.editor.dom.getElementsByTagName("w:p")[0]
        assert build_text_map(paragraph).text == "aZd"
        drawings = manager.editor.dom.getElementsByTagName("w:drawing")
        assert len(drawings) == 1
        assert drawings[0].toxml().count("<w:t>MID</w:t>") == 1
        assert paragraph.toxml().count("MID") == 1

    def test_consumed_own_ins_keeps_foreign_nested_del(self, temp_xml):
        """An emptied insertion still holding a foreign w:del survives.

        Another author deleted part of our insertion; our amendment then
        consumes what was left of its visible text. Dropping the now
        text-less insertion would take their tracked deletion with it, so
        the cleanup only removes an insertion with no element children at
        all.
        """
        xml_path = temp_xml(
            f"<w:p>{OWN_INS}"
            '<w:del w:id="5" w:author="Other Author" w:date="2024-01-02T00:00:00Z">'
            "<w:r><w:delText>gone</w:delText></w:r></w:del>"
            "<w:r><w:t>wor</w:t></w:r></w:ins>"
            '<w:ins w:id="2" w:author="Test Author" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>ld here</w:t></w:r></w:ins></w:p>"
        )
        manager = _make_manager(xml_path)

        manager.replace_text("world", "earth")

        assert _get_text_content(manager) == "earth here"
        dels = manager.editor.dom.getElementsByTagName("w:del")
        assert [d.getAttribute("w:author") for d in dels] == ["Other Author"]
        assert manager.editor.dom.getElementsByTagName("w:ins").length == 2
