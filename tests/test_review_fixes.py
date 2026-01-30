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


class TestOccurrenceIndexDrift:
    """Multi-w:t fallback must target the same match, not re-count with a different method."""

    def test_replace_second_occurrence_in_multi_wt_run(self, temp_xml):
        # Two paragraphs: first has "cat" in a single w:t, second has "cat" in a
        # multi-w:t run (so the fallback triggers). We ask for occurrence=1 (second).
        # Bug: _find_across_boundaries with occurrence=1 could return a different
        # match than what _get_nth_match found, because the two methods count differently.
        xml_path = temp_xml("<w:p><w:r><w:t>cat</w:t></w:r></w:p><w:p><w:r><w:t>cat</w:t><w:t> dog</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.replace_text("cat", "tiger", occurrence=1)
        text = _get_text_content(mgr)
        # First "cat" should be untouched, second replaced
        assert text.startswith("cat")
        assert "tiger" in text
        assert "dog" in text  # sibling preserved

    def test_delete_second_occurrence_in_multi_wt_run(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>cat</w:t></w:r></w:p><w:p><w:r><w:t>cat</w:t><w:t> dog</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("cat", occurrence=1)
        text = _get_text_content(mgr)
        # First "cat" untouched, second deleted, "dog" preserved
        assert text.startswith("cat")
        assert "dog" in text
        # Only one "cat" should remain (the first, untouched)
        assert text.count("cat") == 1

    def test_insert_after_second_occurrence_in_multi_wt_run(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>cat</w:t></w:r></w:p><w:p><w:r><w:t>cat</w:t><w:t> dog</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.insert_text_after("cat", "XX", occurrence=1)
        text = _get_text_content(mgr)
        # "XX" should appear right after second "cat", not at end
        assert "catXX" in text
        assert "dog" in text


class TestSiteDAttributeInjection:
    """Site D raw DOM insertion must go through editor helpers for attribute injection."""

    def test_replace_inside_ins_injects_rsid(self, temp_xml):
        # When replacing text across runs inside <w:ins>, the new run should get
        # w:rsidR injected by DocxXMLEditor's attribute injection.
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>AB</w:t></w:r><w:r><w:t>CD</w:t></w:r></w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.replace_text("ABCD", "NEW")
        # Find the new w:r containing "NEW"
        for wt in mgr.editor.dom.getElementsByTagName("w:t"):
            if wt.firstChild and wt.firstChild.data == "NEW":
                run = wt.parentNode
                assert run.hasAttribute("w:rsidR"), "New run should have w:rsidR from attribute injection"
                break
        else:
            pytest.fail("Could not find w:t with 'NEW'")

    def test_replace_inside_ins_partially_remaining_injects_rsid(self, temp_xml):
        # When ins_elem stays in DOM (partial removal), the inserted run should
        # also get attribute injection.
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>xxAB</w:t></w:r><w:r><w:t>CDyy</w:t></w:r></w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.replace_text("ABCD", "NEW")
        for wt in mgr.editor.dom.getElementsByTagName("w:t"):
            if wt.firstChild and wt.firstChild.data == "NEW":
                run = wt.parentNode
                assert run.hasAttribute("w:rsidR"), "New run should have w:rsidR from attribute injection"
                break
        else:
            pytest.fail("Could not find w:t with 'NEW'")


class TestXmlSpacePreserveOnGeneratedWt:
    """Generated <w:t> elements with leading/trailing spaces must have xml:space='preserve'."""

    def test_replace_preserves_trailing_space_in_before_text(self, temp_xml):
        # "Hello world" -> replace "world" -> before_text is "Hello " (trailing space)
        xml_path = temp_xml("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.replace_text("world", "earth")
        # Find the w:t containing "Hello "
        for wt in mgr.editor.dom.getElementsByTagName("w:t"):
            if wt.firstChild and wt.firstChild.data == "Hello ":
                assert wt.getAttribute("xml:space") == "preserve", (
                    "w:t with trailing space must have xml:space='preserve'"
                )
                break
        else:
            pytest.fail("Could not find w:t with 'Hello '")

    def test_delete_preserves_leading_space_in_after_text(self, temp_xml):
        # "Hello world" -> delete "Hello" -> after_text is " world" (leading space)
        xml_path = temp_xml("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("Hello")
        for wt in mgr.editor.dom.getElementsByTagName("w:t"):
            if wt.firstChild and wt.firstChild.data == " world":
                assert wt.getAttribute("xml:space") == "preserve", (
                    "w:t with leading space must have xml:space='preserve'"
                )
                break
        else:
            pytest.fail("Could not find w:t with ' world'")


class TestOccurrenceFallbackConsistency:
    """Initial fallback from _get_nth_match to _find_across_boundaries must find the right match."""

    def test_replace_cross_boundary_second_occurrence(self, temp_xml):
        # First paragraph: "catdog" split across two w:t (not findable by _get_nth_match)
        # Second paragraph: "catdog" also split across two w:t
        # Asking for occurrence=1 should replace the SECOND "catdog", not the first.
        xml_path = temp_xml(
            "<w:p><w:r><w:t>cat</w:t></w:r><w:r><w:t>dog</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>cat</w:t></w:r><w:r><w:t>dog</w:t></w:r></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.replace_text("catdog", "REPLACED", occurrence=1)
        text = _get_text_content(mgr)
        # First "catdog" should remain, second should be replaced
        assert text.startswith("catdog")
        assert "REPLACED" in text

    def test_delete_cross_boundary_second_occurrence(self, temp_xml):
        xml_path = temp_xml(
            "<w:p><w:r><w:t>cat</w:t></w:r><w:r><w:t>dog</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>cat</w:t></w:r><w:r><w:t>dog</w:t></w:r></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("catdog", occurrence=1)
        text = _get_text_content(mgr)
        # First "catdog" untouched
        assert text.startswith("catdog")
        # Only one "catdog" should remain
        assert text.count("catdog") == 1


class TestMixedStateMarkerPlacement:
    """_replace_mixed_state insertion must go after deletion, not before prefix text."""

    def test_replace_mixed_state_with_before_text(self, temp_xml):
        # Regular text "xxAB" followed by ins text "CD"
        # Replace "ABCD" -> before_text="xx" from regular, ins text removed
        # The replacement should appear AFTER the deletion of "AB", not before "xx"
        xml_path = temp_xml(
            "<w:p>"
            "<w:r><w:t>xxAB</w:t></w:r>"
            '<w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>CD</w:t></w:r></w:ins>"
            "</w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.replace_text("ABCD", "NEW")
        text = _get_text_content(mgr)
        assert "xx" in text
        assert "NEW" in text
        # "xx" must come BEFORE "NEW" in the visible text
        assert text.index("xx") < text.index("NEW"), f"Prefix 'xx' should appear before 'NEW', got: '{text}'"


class TestSingleNodeRemovalXmlSpaceAndGuard:
    """_remove_from_insertion single-node truncation should set xml:space and guard non-text children."""

    def test_truncate_before_sets_xml_space(self, temp_xml):
        # Delete prefix from " Hello" inside ins -> after_text=" Hello"[len(""):] needs xml:space
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r><w:t>AB cd</w:t></w:r></w:ins></w:p>'
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("AB")
        # Remaining text is " cd" which has leading space
        for wt in mgr.editor.dom.getElementsByTagName("w:t"):
            if wt.firstChild and wt.firstChild.data == " cd":
                assert wt.getAttribute("xml:space") == "preserve"
                break
        else:
            pytest.fail("Could not find w:t with ' cd'")

    def test_truncate_after_sets_xml_space(self, temp_xml):
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r><w:t>cd AB</w:t></w:r></w:ins></w:p>'
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("AB")
        for wt in mgr.editor.dom.getElementsByTagName("w:t"):
            if wt.firstChild and wt.firstChild.data == "cd ":
                assert wt.getAttribute("xml:space") == "preserve"
                break
        else:
            pytest.fail("Could not find w:t with 'cd '")

    def test_full_removal_preserves_run_with_tab(self, temp_xml):
        # Run has w:t + w:tab; removing w:t should keep run alive for the tab
        xml_path = temp_xml(
            '<w:p><w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>GONE</w:t><w:tab/></w:r>"
            "<w:r><w:t>STAY</w:t></w:r></w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("GONE")
        text = _get_text_content(mgr)
        assert "STAY" in text
        assert "GONE" not in text
        # The run with w:tab should still exist
        tabs = mgr.editor.dom.getElementsByTagName("w:tab")
        assert tabs.length > 0, "w:tab should be preserved when its sibling w:t is removed"


class TestRunRebuildPreservesNonTextChildren:
    """_replace_same_context and _delete_same_context should preserve non-text run children."""

    def test_replace_preserves_tab_in_run(self, temp_xml):
        # Run has w:tab + w:t; replacing text should keep the tab
        xml_path = temp_xml("<w:p><w:r><w:tab/><w:t>MATCH</w:t></w:r><w:r><w:t>END</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.replace_text("MATCHEND", "NEW")
        text = _get_text_content(mgr)
        assert "NEW" in text
        tabs = mgr.editor.dom.getElementsByTagName("w:tab")
        assert tabs.length > 0, "w:tab should be preserved during run rebuild"

    def test_delete_preserves_tab_in_run(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:tab/><w:t>MATCH</w:t></w:r><w:r><w:t>END</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)
        mgr.suggest_deletion("MATCHEND")
        tabs = mgr.editor.dom.getElementsByTagName("w:tab")
        assert tabs.length > 0, "w:tab should be preserved during deletion rebuild"
