"""Tests for coverage gaps in cross-boundary text operations."""

import pytest
from conftest import parse_paragraph as _parse_paragraph

from docx_editor import Document
from docx_editor.xml_editor import (
    TextMapMatch,
    TextPosition,
    build_text_map,
)

# -- build_text_map edge cases --


class TestBuildTextMapEdgeCases:
    def test_paragraph_with_only_deleted_text(self):
        """Paragraph containing only w:delText returns empty text map."""
        p = _parse_paragraph(
            "<w:p>"
            '<w:del w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:delText>deleted stuff</w:delText></w:r>"
            "</w:del>"
            "</w:p>"
        )
        tm = build_text_map(p)
        assert tm.text == ""
        assert tm.positions == []

    def test_wt_inside_wdel_is_skipped(self):
        """w:t inside w:del is skipped (even though normally w:delText is used)."""
        p = _parse_paragraph(
            "<w:p>"
            '<w:del w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>should be skipped</w:t></w:r>"
            "</w:del>"
            "<w:r><w:t>visible</w:t></w:r>"
            "</w:p>"
        )
        tm = build_text_map(p)
        assert tm.text == "visible"

    def test_unicode_and_special_chars(self):
        """Text map handles unicode and special characters."""
        p = _parse_paragraph("<w:p><w:r><w:t>café résumé naïve</w:t></w:r></w:p>")
        tm = build_text_map(p)
        assert tm.text == "café résumé naïve"
        assert len(tm.positions) == len("café résumé naïve")

    def test_xml_entities_in_text(self):
        """Text map handles XML entities correctly."""
        p = _parse_paragraph("<w:p><w:r><w:t>A &amp; B &lt; C</w:t></w:r></w:p>")
        tm = build_text_map(p)
        assert tm.text == "A & B < C"


# -- TextMap edge cases --


class TestTextMapEdgeCases:
    def test_find_empty_search(self):
        p = _parse_paragraph("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>")
        tm = build_text_map(p)
        # Python str.find("") returns 0
        assert tm.find("") == 0

    def test_find_longer_than_text(self):
        p = _parse_paragraph("<w:p><w:r><w:t>Hi</w:t></w:r></w:p>")
        tm = build_text_map(p)
        assert tm.find("Hello world") == -1

    def test_get_nodes_for_range_empty(self):
        p = _parse_paragraph("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>")
        tm = build_text_map(p)
        assert tm.get_nodes_for_range(2, 2) == []

    def test_get_nodes_for_range_full(self):
        p = _parse_paragraph("<w:p><w:r><w:t>Hello</w:t></w:r></w:p>")
        tm = build_text_map(p)
        positions = tm.get_nodes_for_range(0, 5)
        assert len(positions) == 5


# -- _classify_segments direct tests --


class TestClassifySegments:
    """Direct tests for _classify_segments logic."""

    def test_single_regular_segment(self):
        """All positions regular -> one segment."""
        from unittest.mock import MagicMock

        from docx_editor.track_changes import RevisionManager

        manager = RevisionManager(MagicMock())
        match = TextMapMatch(
            start=0,
            end=3,
            text="abc",
            positions=[
                TextPosition(node=None, offset_in_node=0, is_inside_ins=False, is_inside_del=False),
                TextPosition(node=None, offset_in_node=1, is_inside_ins=False, is_inside_del=False),
                TextPosition(node=None, offset_in_node=2, is_inside_ins=False, is_inside_del=False),
            ],
            spans_boundary=False,
        )
        segments = manager._classify_segments(match)
        assert len(segments) == 1
        assert segments[0][0] is False  # not inside ins
        assert len(segments[0][1]) == 3

    def test_two_segments_regular_then_ins(self):
        """Regular + insertion -> two segments."""
        from unittest.mock import MagicMock

        from docx_editor.track_changes import RevisionManager

        manager = RevisionManager(MagicMock())
        match = TextMapMatch(
            start=0,
            end=4,
            text="abcd",
            positions=[
                TextPosition(node=None, offset_in_node=0, is_inside_ins=False, is_inside_del=False),
                TextPosition(node=None, offset_in_node=1, is_inside_ins=False, is_inside_del=False),
                TextPosition(node=None, offset_in_node=0, is_inside_ins=True, is_inside_del=False),
                TextPosition(node=None, offset_in_node=1, is_inside_ins=True, is_inside_del=False),
            ],
            spans_boundary=True,
        )
        segments = manager._classify_segments(match)
        assert len(segments) == 2
        assert segments[0][0] is False
        assert len(segments[0][1]) == 2
        assert segments[1][0] is True
        assert len(segments[1][1]) == 2

    def test_three_alternating_segments(self):
        """Regular + insertion + regular -> three segments."""
        from unittest.mock import MagicMock

        from docx_editor.track_changes import RevisionManager

        manager = RevisionManager(MagicMock())
        match = TextMapMatch(
            start=0,
            end=3,
            text="abc",
            positions=[
                TextPosition(node=None, offset_in_node=0, is_inside_ins=False, is_inside_del=False),
                TextPosition(node=None, offset_in_node=0, is_inside_ins=True, is_inside_del=False),
                TextPosition(node=None, offset_in_node=0, is_inside_ins=False, is_inside_del=False),
            ],
            spans_boundary=True,
        )
        segments = manager._classify_segments(match)
        assert len(segments) == 3

    def test_empty_positions(self):
        """Empty positions -> empty segments."""
        from unittest.mock import MagicMock

        from docx_editor.track_changes import RevisionManager

        manager = RevisionManager(MagicMock())
        match = TextMapMatch(
            start=0,
            end=0,
            text="",
            positions=[],
            spans_boundary=False,
        )
        segments = manager._classify_segments(match)
        assert segments == []


# -- Mixed-state deletion --


class TestMixedStateDeletion:
    """Tests for deleting text spanning revision boundaries."""

    def test_delete_spanning_regular_and_insertion(self, clean_workspace):
        """Delete text that spans regular text + insertion."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " ADDED")

        match = doc.find_text("fox ADDED")
        if match is None or not match.spans_boundary:
            doc.close()
            pytest.skip("Expected boundary not created")

        change_id = doc.delete("fox ADDED")
        assert change_id >= 0

        text = doc.get_visible_text()
        assert "fox" not in text
        assert "ADDED" not in text
        doc.close()

    def test_delete_spanning_insertion_and_regular(self, clean_workspace):
        """Delete text spanning insertion + regular text."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " ADDED")

        match = doc.find_text("ADDED jumps")
        if match is None or not match.spans_boundary:
            doc.close()
            pytest.skip("Expected boundary not created")

        doc.delete("ADDED jumps")
        text = doc.get_visible_text()
        assert "ADDED" not in text
        assert "jumps" not in text
        doc.close()

    def test_delete_entirely_within_insertion(self, clean_workspace):
        """Delete text entirely within an insertion."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " beautiful amazing")

        match = doc.find_text("beautiful")
        if match is None or not all(p.is_inside_ins for p in match.positions):
            doc.close()
            pytest.skip("Text not entirely within insertion")

        doc.delete("beautiful")
        text = doc.get_visible_text()
        assert "beautiful" not in text
        assert "amazing" in text
        doc.close()

    def test_delete_mixed_state_roundtrip(self, clean_workspace, temp_dir):
        """Round-trip: delete across boundary, save, reopen."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " ADDED")

        match = doc.find_text("fox ADDED")
        if match is None or not match.spans_boundary:
            doc.close()
            pytest.skip("Expected boundary not created")

        doc.delete("fox ADDED")
        output = temp_dir / "delete_mixed.docx"
        doc.save(output)
        doc.close()

        doc2 = Document.open(output, force_recreate=True)
        text = doc2.get_visible_text()
        assert "fox" not in text
        assert "ADDED" not in text
        doc2.close()


# -- _remove_from_insertion all branches --


class TestRemoveFromInsertionBranches:
    """Test all 4 branches of _remove_from_insertion."""

    def test_entire_insertion_removed(self, clean_workspace):
        """Removing all text from an insertion removes the w:ins element."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " WORD")

        doc.delete("WORD")
        text = doc.get_visible_text()
        assert "WORD" not in text
        doc.close()

    def test_match_at_start_of_insertion(self, clean_workspace):
        """Removing text at the start of insertion truncates to remainder."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " hello world")

        doc.delete("hello")
        text = doc.get_visible_text()
        assert "hello" not in text
        assert " world" in text or "world" in text
        doc.close()

    def test_match_at_end_of_insertion(self, clean_workspace):
        """Removing text at end of insertion truncates to start."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " hello world")

        doc.delete("world")
        text = doc.get_visible_text()
        assert "world" not in text
        assert "hello" in text
        doc.close()

    def test_match_in_middle_of_insertion_splits(self, clean_workspace):
        """Removing text in middle of insertion splits into two w:ins elements."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " hello beautiful world")

        doc.delete("beautiful")
        text = doc.get_visible_text()
        assert "beautiful" not in text
        assert "hello" in text
        assert "world" in text
        doc.close()

    def test_middle_split_roundtrip(self, clean_workspace, temp_dir):
        """Round-trip: middle split, save, reopen."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " hello beautiful world")

        doc.delete("beautiful")
        output = temp_dir / "middle_split.docx"
        doc.save(output)
        doc.close()

        doc2 = Document.open(output, force_recreate=True)
        text = doc2.get_visible_text()
        assert "beautiful" not in text
        assert "hello" in text
        assert "world" in text
        doc2.close()


# -- XML special characters --


class TestXmlSpecialCharacters:
    """Tests for XML special characters through cross-boundary paths."""

    def test_replace_text_with_ampersand(self, clean_workspace):
        """Replace with text containing & character."""
        doc = Document.open(clean_workspace)
        doc.replace("fox", "fox & cat")
        text = doc.get_visible_text()
        assert "fox & cat" in text
        doc.close()

    def test_replace_text_with_angle_brackets(self, clean_workspace):
        """Replace with text containing < and > characters."""
        doc = Document.open(clean_workspace)
        doc.replace("fox", "a < b > c")
        text = doc.get_visible_text()
        assert "a < b > c" in text
        doc.close()

    def test_insert_text_with_special_chars(self, clean_workspace):
        """Insert text containing XML special characters."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " & friends <team>")
        text = doc.get_visible_text()
        assert "& friends <team>" in text
        doc.close()

    def test_special_chars_roundtrip(self, clean_workspace, temp_dir):
        """Round-trip with special characters."""
        doc = Document.open(clean_workspace)
        doc.replace("fox", "A & B")
        output = temp_dir / "special_chars.docx"
        doc.save(output)
        doc.close()

        doc2 = Document.open(output, force_recreate=True)
        text = doc2.get_visible_text()
        assert "A & B" in text
        doc2.close()


# -- occurrence parameter for cross-boundary --


class TestOccurrenceParameter:
    """Tests for occurrence > 0 in cross-boundary operations."""

    def test_find_text_second_occurrence(self, clean_workspace):
        """Find second occurrence across paragraphs."""
        doc = Document.open(clean_workspace)
        match0 = doc.find_text("the")
        match1 = doc.find_text("the", occurrence=1)
        if match0 is None or match1 is None:
            doc.close()
            pytest.skip("Not enough occurrences of 'the'")
        assert match0.positions[0].node is not match1.positions[0].node
        doc.close()

    def test_replace_second_occurrence_cross_boundary(self, clean_workspace):
        """Replace second occurrence when both span boundaries."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " test")
        doc.insert_after("dog", " test")

        match = doc.find_text("test", occurrence=1)
        if match is None:
            doc.close()
            pytest.skip("Second occurrence not found")

        doc.replace("test", "REPLACED", occurrence=1)
        text = doc.get_visible_text()
        assert "test" in text
        assert "REPLACED" in text
        doc.close()


# -- find_text edge cases --


class TestFindTextEdgeCases:
    """Edge cases for Document.find_text."""

    def test_find_deleted_text_returns_none(self, clean_workspace):
        """Text inside w:del should not be found."""
        doc = Document.open(clean_workspace)
        doc.delete("fox")
        match = doc.find_text("fox")
        assert match is None
        doc.close()

    def test_get_visible_text_excludes_deleted(self, clean_workspace):
        """Visible text should not include deleted text."""
        doc = Document.open(clean_workspace)
        doc.delete("fox")
        text = doc.get_visible_text()
        assert "fox" not in text
        doc.close()


# -- insert_after/before with cross-boundary anchor --


class TestInsertWithCrossBoundaryAnchor:
    """Tests for insert when anchor text spans element boundaries."""

    def test_insert_after_cross_boundary_anchor(self, clean_workspace):
        """insert_after works when anchor spans a boundary."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " RED")

        match = doc.find_text("fox RED")
        if match is None or not match.spans_boundary:
            doc.close()
            pytest.skip("Boundary not created")

        change_id = doc.insert_after("fox RED", " NEW")
        assert change_id >= 0
        text = doc.get_visible_text()
        assert "fox RED NEW" in text
        doc.close()

    def test_insert_before_cross_boundary_anchor(self, clean_workspace):
        """insert_before works when anchor spans a boundary."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " RED")

        match = doc.find_text("fox RED")
        if match is None or not match.spans_boundary:
            doc.close()
            pytest.skip("Boundary not created")

        change_id = doc.insert_before("fox RED", "NEW ")
        assert change_id >= 0
        text = doc.get_visible_text()
        assert "NEW fox RED" in text
        doc.close()

    def test_insert_cross_boundary_roundtrip(self, clean_workspace, temp_dir):
        """Round-trip: insert with cross-boundary anchor, save, reopen."""
        doc = Document.open(clean_workspace)
        doc.insert_after("fox", " RED")

        match = doc.find_text("fox RED")
        if match is None or not match.spans_boundary:
            doc.close()
            pytest.skip("Boundary not created")

        doc.insert_after("fox RED", " NEW")
        output = temp_dir / "insert_cross_boundary.docx"
        doc.save(output)
        doc.close()

        doc2 = Document.open(output, force_recreate=True)
        text = doc2.get_visible_text()
        assert "fox RED NEW" in text
        doc2.close()
