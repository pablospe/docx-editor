"""Tests for multi-w:t run safety and related fixes."""

from pathlib import Path

import pytest

from docx_editor.track_changes import RevisionManager
from docx_editor.xml_editor import DocxXMLEditor

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


@pytest.fixture
def temp_xml(tmp_path):
    def _create_xml(body_xml: str) -> Path:
        xml = f'<?xml version="1.0" encoding="utf-8"?><w:document {NS}><w:body>{body_xml}</w:body></w:document>'
        xml_path = tmp_path / "test_doc.xml"
        xml_path.write_text(xml)
        return xml_path

    return _create_xml


def _make_manager(xml_path: Path) -> RevisionManager:
    editor = DocxXMLEditor(xml_path, rsid="00000000", author="Test Author")
    return RevisionManager(editor)


def _get_text_content(manager):
    dom = manager.editor.dom
    result = []
    for wt in dom.getElementsByTagName("w:t"):
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


class TestRunDeduplication:
    """Test that replacing/deleting across multiple w:t nodes in the same run
    doesn't crash from double-removal (comment 2)."""

    def test_replace_across_multi_wt_same_run(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>Hello </w:t><w:t>world</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.replace_text("lo wor", "LO WOR")
        text = _get_text_content(mgr)
        assert "LO WOR" in text

    def test_delete_across_multi_wt_same_run(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>Hello </w:t><w:t>world</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("lo wor")
        text = _get_text_content(mgr)
        assert "lo wor" not in text


class TestSingleNodeRemovalGuard:
    """Test that removing matched text from a single w:t inside w:ins
    doesn't drop sibling w:t nodes (comment 3)."""

    def test_remove_one_run_preserves_sibling(self, temp_xml):
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello </w:t></w:r>"
            "<w:r><w:t>world</w:t></w:r>"
            "</w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("world")
        text = _get_text_content(mgr)
        assert "Hello " in text
        # w:ins element should still exist
        ins_elements = mgr.editor.dom.getElementsByTagName("w:ins")
        assert ins_elements.length > 0

    def test_remove_one_wt_preserves_sibling_in_same_run(self, temp_xml):
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello </w:t><w:t>world</w:t></w:r>"
            "</w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("world")
        text = _get_text_content(mgr)
        assert "Hello " in text
        ins_elements = mgr.editor.dom.getElementsByTagName("w:ins")
        assert ins_elements.length > 0


class TestDeleteRegularSegmentMultiWt:
    """Test that _delete_regular_segment works correctly with multi-w:t runs
    (comments 4/12)."""

    def test_delete_spanning_multi_wt_no_crash(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>abc</w:t><w:t>def</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("cd")
        text = _get_text_content(mgr)
        assert "abef" in text

    def test_delete_middle_wt_preserves_outer(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>Hello </w:t><w:t>big </w:t><w:t>world</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("big ")
        text = _get_text_content(mgr)
        assert "Hello " in text
        assert "world" in text


class TestMidNodeInsertionOrder:
    """Test that replacement text appears AFTER preserved prefix when match
    starts mid-node (comment 11)."""

    def test_replace_mid_node_preserves_prefix_order(self, temp_xml):
        xml_path = temp_xml(
            "<w:p>"
            "<w:r><w:t>Hello world</w:t></w:r>"
            '<w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t> added</w:t></w:r>"
            "</w:ins>"
            "</w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.replace_text("world added", "REPLACED")
        text = _get_text_content(mgr)
        assert text == "Hello REPLACED"
