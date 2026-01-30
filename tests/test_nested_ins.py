"""Tests for nested w:ins prevention.

These tests verify that operations inside existing w:ins elements do not create
nested w:ins or w:del elements, which are invalid in OOXML.
"""

from pathlib import Path

import pytest

from docx_editor.track_changes import RevisionManager
from docx_editor.xml_editor import DocxXMLEditor

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


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
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
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
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
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
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r><w:t>Hello</w:t></w:r></w:ins></w:p>'
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
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r><w:t>Hello</w:t></w:r></w:ins></w:p>'
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
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
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
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
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
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
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
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
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
            '<w:p><w:r><w:t>Hello </w:t></w:r><w:ins w:id="1" w:author="A" '
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
            '<w:p><w:r><w:t>Hello </w:t></w:r><w:ins w:id="1" w:author="A" '
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
