"""Tests to improve coverage for track_changes.py uncovered lines."""

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


class TestReplaceMultiWtDelegatesToCrossBoundary:
    """Lines 143-146: replace_text delegates to cross-boundary when run has multiple w:t."""

    def test_replace_text_multi_wt_run_delegates(self, temp_xml):
        # A single run with two w:t children, where the search text is in one of them.
        # This triggers lines 143-146 in replace_text.
        xml_path = temp_xml("<w:p><w:r><w:t>Hello </w:t><w:t>world</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.replace_text("world", "EARTH")
        text = _get_text_content(mgr)
        assert "EARTH" in text
        assert "world" not in text


class TestSuggestDeletionInsideInsRemovesEntireIns:
    """Lines 254-255: suggest_deletion removes entire w:ins when sole w:t fully matched."""

    def test_delete_entire_single_wt_inside_ins(self, temp_xml):
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r><w:t>Hello</w:t></w:r></w:ins></w:p>'
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("Hello")
        # The entire w:ins should be removed
        assert mgr.editor.dom.getElementsByTagName("w:ins").length == 0
        text = _get_text_content(mgr)
        assert text == ""


class TestGetRunInfoNoParentRun:
    """Lines 294, 296: _get_run_info returns (None, '') when no w:r ancestor."""

    def test_get_run_info_no_run_parent(self, temp_xml):
        xml_path = temp_xml("<w:p><w:t>orphan</w:t></w:p>")
        mgr = _make_manager(xml_path)
        wt = mgr.editor.dom.getElementsByTagName("w:t")[0]
        run, rPr = mgr._get_run_info(wt)
        assert run is None
        assert rPr == ""


class TestBuildCrossBoundaryPartsSkipsOrphan:
    """Line 324: _build_cross_boundary_parts skips positions with no parent run."""

    def test_empty_parts_returns_minus_one(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        # _replace_same_context with empty parts returns -1 (line 385)
        from docx_editor.xml_editor import TextMapMatch

        empty_match = TextMapMatch(start=0, end=0, text="", positions=[], spans_boundary=False)
        result = mgr._replace_same_context(empty_match, "X")
        assert result == -1


class TestReplaceSameContextAllInsideIns:
    """Lines 385, 406-407: _replace_same_context when all positions inside w:ins,
    and ins_elem has no firstChild after removal."""

    def test_replace_across_runs_all_inside_ins_no_remaining(self, temp_xml):
        # Two runs inside w:ins, replace spanning both fully -> empties ins, appends
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>AB</w:t></w:r><w:r><w:t>CD</w:t></w:r></w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.replace_text("ABCD", "NEW")
        text = _get_text_content(mgr)
        assert "NEW" in text


class TestReplaceMixedStateNoDelFound:
    """Line 499: _replace_mixed_state inserts after marker when no w:del siblings."""

    def test_replace_mixed_all_ins_segment(self, temp_xml):
        # Match entirely inside w:ins with spans_boundary=True triggers _replace_mixed_state.
        # After _remove_from_insertion, no w:del exists, so insert after marker (line 499).
        xml_path = temp_xml(
            "<w:p>"
            '<w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello</w:t></w:r></w:ins>"
            "<w:r><w:t> world</w:t></w:r>"
            "</w:p>"
        )
        mgr = _make_manager(xml_path)
        # "Hello w" spans ins boundary
        mgr.replace_text("Hello w", "NEW W")
        text = _get_text_content(mgr)
        assert "NEW W" in text


class TestRemoveFromInsertionMiddleSplit:
    """Lines 582-590: _remove_from_insertion with middle split (before and after text remain)."""

    def test_remove_middle_from_single_node_ins(self, temp_xml):
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r><w:t>ABCDE</w:t></w:r></w:ins></w:p>'
        )
        mgr = _make_manager(xml_path)
        # Delete "BCD" from middle of "ABCDE" inside w:ins
        mgr.suggest_deletion("BCD")
        text = _get_text_content(mgr)
        # Should have "A" and "E" remaining
        assert "A" in text
        assert "E" in text
        assert "BCD" not in text


class TestRemoveFromInsertionSoleWtRemovesIns:
    """Line 567: _remove_from_insertion removes entire w:ins when sole w:t fully matched
    (through cross-boundary path)."""

    def test_cross_boundary_delete_entire_ins_sole_wt(self, temp_xml):
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>AB</w:t></w:r></w:ins>"
            '<w:ins w:id="2" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>CD</w:t></w:r></w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        # Delete "ABCD" spanning two w:ins elements
        mgr.suggest_deletion("ABCD")
        text = _get_text_content(mgr)
        assert text == ""


class TestRemoveFromInsertionMultiWtRemoveJustNode:
    """Lines 569-575: _remove_from_insertion removes just the w:t node when other w:t exists."""

    def test_remove_one_wt_from_multi_wt_ins_via_cross_boundary(self, temp_xml):
        # Two w:t nodes in same ins, delete one entirely via cross-boundary
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>AB</w:t></w:r>"
            "<w:r><w:t>CD</w:t></w:r>"
            "</w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("CD")
        text = _get_text_content(mgr)
        assert "AB" in text
        assert "CD" not in text
        # w:ins should still exist
        assert mgr.editor.dom.getElementsByTagName("w:ins").length > 0


class TestRemoveFromInsertionMultiNodeNoBeforeNoAfter:
    """Lines 598-600, 605-607, 611-613: Multi-node removal with no before/after text."""

    def test_multi_node_removal_no_before_no_after(self, temp_xml):
        # Three runs inside w:ins, delete spanning all three fully
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>AB</w:t></w:r>"
            "<w:r><w:t>CD</w:t></w:r>"
            "<w:r><w:t>EF</w:t></w:r>"
            "</w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("ABCDEF")
        text = _get_text_content(mgr)
        assert text == ""

    def test_multi_node_removal_with_before_and_after(self, temp_xml):
        # Three runs inside w:ins, delete middle portion spanning all three
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>xxAB</w:t></w:r>"
            "<w:r><w:t>CD</w:t></w:r>"
            "<w:r><w:t>EFyy</w:t></w:r>"
            "</w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("ABCDEF")
        text = _get_text_content(mgr)
        assert "xx" in text
        assert "yy" in text
        assert "ABCDEF" not in text


class TestGetWtNodesInAncestorNone:
    """Line 618: _get_wt_nodes_in_ancestor returns [] for None ancestor."""

    def test_none_ancestor(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>x</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        assert mgr._get_wt_nodes_in_ancestor(None) == []


class TestDeleteRegularSegmentUnmatchedWt:
    """Lines 675-678: _delete_regular_segment preserves unmatched w:t siblings."""

    def test_delete_preserves_unmatched_wt(self, temp_xml):
        # Already tested in test_multi_wt_safety.py but let's cover the
        # specific "unmatched sibling" branch in _delete_regular_segment
        xml_path = temp_xml("<w:p><w:r><w:t>prefix</w:t><w:t>MATCH</w:t><w:t>suffix</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("MATCH")
        text = _get_text_content(mgr)
        assert "prefix" in text
        assert "suffix" in text
        assert "MATCH" not in text


class TestDeleteRegularSegmentIntermediateNode:
    """Line 698: intermediate node uses entire text as matched."""

    def test_three_wt_nodes_spanning_deletion(self, temp_xml):
        # Three separate runs, delete spanning all three
        xml_path = temp_xml("<w:p><w:r><w:t>xxAB</w:t></w:r><w:r><w:t>CD</w:t></w:r><w:r><w:t>EFyy</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("ABCDEF")
        text = _get_text_content(mgr)
        assert "xx" in text
        assert "yy" in text


class TestDeleteSameContextReturnMinusOne:
    """Lines 732, 804: _delete_same_context returns -1 for empty parts."""

    def test_delete_same_context_empty_parts(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        from docx_editor.xml_editor import TextMapMatch

        empty_match = TextMapMatch(start=0, end=0, text="", positions=[], spans_boundary=False)
        result = mgr._delete_same_context(empty_match)
        assert result == -1


class TestInsertNearMatchInsideInsBeforePosition:
    """Line 955: _insert_near_match with position='before' inside w:ins."""

    def test_insert_before_cross_boundary_inside_ins(self, temp_xml):
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        # Insert before "lo wor" which spans two runs inside w:ins
        mgr.insert_text_before("lo wor", "XX")
        text = _get_text_content(mgr)
        assert "XX" in text


class TestSuggestDeletionMultiWtDelegates:
    """Lines 230-233: suggest_deletion delegates to cross-boundary for multi-w:t run."""

    def test_suggest_deletion_multi_wt_delegates(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>Hello </w:t><w:t>world</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("world")
        text = _get_text_content(mgr)
        assert "Hello " in text
        assert "world" not in text


class TestDeleteMixedState:
    """Line 933 (approximate): _delete_mixed_state processes mixed segments."""

    def test_delete_crossing_ins_boundary(self, temp_xml):
        xml_path = temp_xml(
            "<w:p>"
            "<w:r><w:t>Hello </w:t></w:r>"
            '<w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>world</w:t></w:r></w:ins>"
            "</w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("lo world")
        text = _get_text_content(mgr)
        assert "lo" not in text
        assert "world" not in text
        assert "Hel" in text


class TestReplaceSameContextInsAppendNoFirstChild:
    """Lines 405-407: _replace_same_context appends when ins_elem has no firstChild."""

    def test_replace_all_inside_ins_removes_then_appends(self, temp_xml):
        # Two runs both inside ins, replace entire content.
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>X</w:t></w:r><w:r><w:t>Y</w:t></w:r></w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.replace_text("XY", "NEW")
        text = _get_text_content(mgr)
        assert "NEW" in text
