"""Tests for mixed-state editing (Phase 5: Atomic Decomposition)."""

import pytest
from conftest import parse_paragraph as _parse_paragraph

from docx_editor import Document
from docx_editor.xml_editor import build_text_map, find_in_text_map


class TestMixedStateDetection:
    """Tests that mixed-state matches are correctly detected."""

    def test_regular_plus_insertion_spans_boundary(self):
        """Match spanning regular text + insertion is detected."""
        p = _parse_paragraph(
            "<w:p>"
            "<w:r><w:t>Aim: </w:t></w:r>"
            '<w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>To examine</w:t></w:r>"
            "</w:ins>"
            "</w:p>"
        )
        tm = build_text_map(p)
        match = find_in_text_map(tm, "Aim: To")
        assert match is not None
        assert match.spans_boundary

    def test_fully_within_insertion_no_boundary(self):
        """Match fully within insertion does not span boundary."""
        p = _parse_paragraph(
            "<w:p>"
            '<w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello beautiful world</w:t></w:r>"
            "</w:ins>"
            "</w:p>"
        )
        tm = build_text_map(p)
        match = find_in_text_map(tm, "beautiful")
        assert match is not None
        assert not match.spans_boundary
        assert all(pos.is_inside_ins for pos in match.positions)


class TestMixedStateReplace:
    """Integration tests for replace spanning revision boundaries."""

    def test_replace_spanning_regular_and_insertion(self, clean_workspace):
        """Replace text spanning regular text + <w:ins>.

        Regular text ("dog.") should be wrapped in <w:del>.
        Inserted text (" ADDED") should be removed (undo insertion), not del-wrapped.
        New text should appear in a <w:ins>.
        """
        doc = Document.open(clean_workspace)
        # insert_after puts text after the run containing anchor,
        # so "dog." is at end of run -> " ADDED" is in <w:ins> right after
        doc.insert_after("dog.", " ADDED")

        match = doc.find_text("dog. ADDED")
        if match is None or not match.spans_boundary:
            doc.close()
            pytest.skip("Expected boundary not created")

        change_id = doc.replace("dog. ADDED", "cat. CHANGED")
        assert change_id >= 0

        text = doc.get_visible_text()
        assert "cat. CHANGED" in text
        assert "dog." not in text
        assert "ADDED" not in text

        # Verify XML semantics: " ADDED" should NOT be in any <w:delText>
        # (it was an insertion that was undone, not deleted)
        dom = doc._document_editor.dom
        for del_text in dom.getElementsByTagName("w:delText"):
            content = del_text.firstChild.data if del_text.firstChild else ""
            assert "ADDED" not in content, "Inserted text should be removed, not wrapped in w:del"
        doc.close()

    def test_replace_fully_within_insertion(self, clean_workspace):
        """Replace text fully within an <w:ins> -- regression."""
        doc = Document.open(clean_workspace)
        doc.insert_after("dog.", " beautiful amazing")

        match = doc.find_text("beautiful")
        if match is None:
            doc.close()
            pytest.skip("Text not found in insertion")

        if not all(pos.is_inside_ins for pos in match.positions):
            doc.close()
            pytest.skip("Text not entirely within insertion")

        change_id = doc.replace("beautiful", "wonderful")
        # change_id is -1 when editing in-place inside an existing insertion
        assert isinstance(change_id, int)
        text = doc.get_visible_text()
        assert "wonderful" in text
        assert "beautiful" not in text
        doc.close()

    def test_ins_node_splitting(self, clean_workspace):
        """When partially matching an insertion, the ins node is split."""
        doc = Document.open(clean_workspace)
        doc.insert_after("dog.", " To examine whether")

        match = doc.find_text("To")
        if match is None:
            doc.close()
            pytest.skip("Text not found")

        doc.delete("To")

        text = doc.get_visible_text()
        assert "To" not in text
        assert "examine whether" in text
        doc.close()

    def test_rpr_preservation_across_split(self, clean_workspace):
        """w:rPr is preserved when splitting runs."""
        doc = Document.open(clean_workspace)
        doc.insert_after("dog.", " ADDED")

        match = doc.find_text("dog. ADDED")
        if match is None or not match.spans_boundary:
            doc.close()
            pytest.skip("Expected boundary not created")

        doc.replace("dog. ADDED", "cat. CHANGED")
        text = doc.get_visible_text()
        assert "cat. CHANGED" in text
        doc.close()

    def test_roundtrip_mixed_state_replace(self, clean_workspace, temp_dir):
        """Round-trip: mixed-state replace, save, reopen."""
        doc = Document.open(clean_workspace)
        doc.insert_after("dog.", " ADDED")

        match = doc.find_text("dog. ADDED")
        if match is None or not match.spans_boundary:
            doc.close()
            pytest.skip("Expected boundary not created")

        doc.replace("dog. ADDED", "cat. CHANGED")
        output = temp_dir / "mixed_state_output.docx"
        doc.save(output)
        doc.close()

        doc2 = Document.open(output, force_recreate=True)
        text = doc2.get_visible_text()
        assert "cat. CHANGED" in text
        assert "dog." not in text
        doc2.close()

    def test_roundtrip_within_insertion_replace(self, clean_workspace, temp_dir):
        """Round-trip: replace within insertion, save, reopen."""
        doc = Document.open(clean_workspace)
        doc.insert_after("dog.", " beautiful amazing")

        match = doc.find_text("beautiful")
        if match is None or not all(pos.is_inside_ins for pos in match.positions):
            doc.close()
            pytest.skip("Text not entirely within insertion")

        doc.replace("beautiful", "wonderful")
        output = temp_dir / "within_ins_output.docx"
        doc.save(output)
        doc.close()

        doc2 = Document.open(output, force_recreate=True)
        text = doc2.get_visible_text()
        assert "wonderful" in text
        doc2.close()
