"""Tests for specific bugs found during code review of track_changes.py."""

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


def _make_manager(xml_path) -> RevisionManager:
    editor = DocxXMLEditor(xml_path, rsid="00000000", author="Test Author")
    return RevisionManager(editor)


def _get_text_content(manager) -> str:
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


class TestSiteDPreserveInsWrapper:
    """Site D should keep replacement inside w:ins wrapper."""

    def test_replace_inside_ins_preserves_wrapper(self, temp_xml):
        # Two runs inside w:ins, replace all text -> ins_elem gets fully removed
        # Replacement should still be inside a w:ins
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>AB</w:t></w:r><w:r><w:t>CD</w:t></w:r></w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.replace_text("ABCD", "NEW")
        text = _get_text_content(mgr)
        assert "NEW" in text
        # The replacement must be inside a w:ins element
        ins_elems = mgr.editor.dom.getElementsByTagName("w:ins")
        assert ins_elems.length > 0
        ins_text = []
        for wt in ins_elems[0].getElementsByTagName("w:t"):
            if wt.firstChild:
                ins_text.append(wt.firstChild.data)
        assert "NEW" in "".join(ins_text)


class TestInsertTextMultiWtPreservesSiblings:
    """_insert_text should preserve sibling w:t nodes in multi-w:t runs."""

    def test_insert_after_in_multi_wt_run(self, temp_xml):
        # Run has two w:t children; insert after text in first w:t
        xml_path = temp_xml("<w:p><w:r><w:t>Hello</w:t><w:t> world</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.insert_text_after("Hello", " INSERTED")
        text = _get_text_content(mgr)
        assert "Hello" in text
        assert "INSERTED" in text
        assert "world" in text  # sibling w:t must be preserved


class TestRemoveFromInsertionPreservesSiblingWt:
    """_remove_from_insertion multi-node should preserve sibling w:t and set xml:space."""

    def test_multi_node_removal_preserves_sibling_wt(self, temp_xml):
        # w:ins has two runs: first run has two w:t (REMOVE + KEEP),
        # second run has one w:t (ALSO). Match "REMOVE" + "ALSO" via cross-boundary.
        # But text map concatenates as "REMOVEKEEPALSO", so we match "REMOVEKEEP"
        # which spans the first run's two w:t nodes, then the second run's w:t is safe.
        # Instead: put KEEP after the matched nodes so the text map is "REMOVEALSOKEEP".
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>REMOVE</w:t></w:r>"
            "<w:r><w:t>ALSO</w:t><w:t>KEEP</w:t></w:r>"
            "</w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("REMOVEALSO")
        text = _get_text_content(mgr)
        assert "KEEP" in text
        assert "REMOVE" not in text
        assert "ALSO" not in text.replace("KEEP", "")

    def test_truncated_nodes_get_xml_space_preserve(self, temp_xml):
        # Multi-node removal where first/last nodes are truncated
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>xxAB</w:t></w:r>"
            "<w:r><w:t>CDyy</w:t></w:r>"
            "</w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("ABCD")
        text = _get_text_content(mgr)
        assert "xx" in text
        assert "yy" in text
        # Check xml:space="preserve" on truncated nodes
        for wt in mgr.editor.dom.getElementsByTagName("w:t"):
            if wt.firstChild and wt.firstChild.data in ("xx", "yy"):
                assert wt.getAttribute("xml:space") == "preserve"


class TestReplaceSameContextPreservesMultiWt:
    """_replace_same_context non-ins path should preserve sibling w:t nodes."""

    def test_replace_preserves_unmatched_wt_siblings(self, temp_xml):
        # Two runs, first has two w:t nodes, match spans across runs
        xml_path = temp_xml("<w:p><w:r><w:t>keep</w:t><w:t>MATCH1</w:t></w:r><w:r><w:t>MATCH2</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.replace_text("MATCH1MATCH2", "NEW")
        text = _get_text_content(mgr)
        assert "keep" in text
        assert "NEW" in text
        assert "MATCH1" not in text
        assert "MATCH2" not in text
