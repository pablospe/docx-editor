"""Tests for text map building (Phase 1: Core Infrastructure)."""

from conftest import parse_paragraph as _parse_paragraph

from docx_editor.xml_editor import build_text_map, find_in_text_map


class TestBuildTextMapPlainText:
    """Tests for text map with plain text (no tracked changes)."""

    def test_single_run(self):
        """Single run with plain text."""
        p = _parse_paragraph("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        tm = build_text_map(p)
        assert tm.text == "Hello world"
        assert len(tm.positions) == 11
        # All positions should reference the same node
        assert all(pos.offset_in_node == i for i, pos in enumerate(tm.positions))
        assert all(not pos.is_inside_ins for pos in tm.positions)
        assert all(not pos.is_inside_del for pos in tm.positions)

    def test_multiple_runs(self):
        """Multiple runs concatenated."""
        p = _parse_paragraph("<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p>")
        tm = build_text_map(p)
        assert tm.text == "Hello world"
        assert len(tm.positions) == 11
        # First 6 chars from first node, next 5 from second
        assert tm.positions[0].node is not tm.positions[6].node
        assert tm.positions[5].offset_in_node == 5  # space
        assert tm.positions[6].offset_in_node == 0  # 'w' in second node

    def test_empty_paragraph(self):
        """Empty paragraph returns empty text map."""
        p = _parse_paragraph("<w:p></w:p>")
        tm = build_text_map(p)
        assert tm.text == ""
        assert tm.positions == []

    def test_run_with_properties(self):
        """Run properties (w:rPr) don't affect text map."""
        p = _parse_paragraph("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r></w:p>")
        tm = build_text_map(p)
        assert tm.text == "Bold"
        assert len(tm.positions) == 4


class TestBuildTextMapTrackedChanges:
    """Tests for text map with tracked changes."""

    def test_inserted_text_included(self):
        """Text inside <w:ins> is included and marked."""
        p = _parse_paragraph(
            "<w:p>"
            "<w:r><w:t>Hello </w:t></w:r>"
            '<w:ins w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>beautiful </w:t></w:r>"
            "</w:ins>"
            "<w:r><w:t>world</w:t></w:r>"
            "</w:p>"
        )
        tm = build_text_map(p)
        assert tm.text == "Hello beautiful world"
        # Positions 0-5: "Hello " - not inside ins
        assert not tm.positions[0].is_inside_ins
        assert not tm.positions[5].is_inside_ins
        # Positions 6-15: "beautiful " - inside ins
        assert tm.positions[6].is_inside_ins
        assert tm.positions[15].is_inside_ins
        # Positions 16-20: "world" - not inside ins
        assert not tm.positions[16].is_inside_ins

    def test_deleted_text_excluded(self):
        """Text inside <w:del> with <w:delText> is excluded."""
        p = _parse_paragraph(
            "<w:p>"
            "<w:r><w:t>Hello </w:t></w:r>"
            '<w:del w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:delText>old </w:delText></w:r>"
            "</w:del>"
            "<w:r><w:t>world</w:t></w:r>"
            "</w:p>"
        )
        tm = build_text_map(p)
        assert tm.text == "Hello world"
        assert len(tm.positions) == 11

    def test_mixed_insertions_and_deletions(self):
        """Paragraph with both insertions and deletions."""
        p = _parse_paragraph(
            "<w:p>"
            "<w:r><w:t>Hello </w:t></w:r>"
            '<w:del w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:delText>old </w:delText></w:r>"
            "</w:del>"
            '<w:ins w:id="2" w:author="Test" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>new </w:t></w:r>"
            "</w:ins>"
            "<w:r><w:t>world</w:t></w:r>"
            "</w:p>"
        )
        tm = build_text_map(p)
        assert tm.text == "Hello new world"
        # "Hello " - regular
        assert not tm.positions[0].is_inside_ins
        # "new " - inside ins
        assert tm.positions[6].is_inside_ins
        assert tm.positions[9].is_inside_ins
        # "world" - regular
        assert not tm.positions[10].is_inside_ins


class TestTextMapFind:
    """Tests for TextMap.find() method."""

    def test_find_simple(self):
        """Find text in simple string."""
        p = _parse_paragraph("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        tm = build_text_map(p)
        assert tm.find("world") == 6
        assert tm.find("Hello") == 0
        assert tm.find("missing") == -1

    def test_find_across_runs(self):
        """Find text that spans multiple runs."""
        p = _parse_paragraph("<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p>")
        tm = build_text_map(p)
        assert tm.find("lo wo") == 3

    def test_find_across_insertion_boundary(self):
        """Find text spanning regular text and insertion."""
        p = _parse_paragraph(
            "<w:p>"
            "<w:r><w:t>Hello </w:t></w:r>"
            '<w:ins w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>world</w:t></w:r>"
            "</w:ins>"
            "</w:p>"
        )
        tm = build_text_map(p)
        assert tm.find("Hello world") == 0

    def test_find_with_start_offset(self):
        """Find with start offset."""
        p = _parse_paragraph("<w:p><w:r><w:t>hello hello world</w:t></w:r></w:p>")
        tm = build_text_map(p)
        assert tm.find("hello", 0) == 0
        assert tm.find("hello", 1) == 6


class TestTextMapGetNodesForRange:
    """Tests for TextMap.get_nodes_for_range()."""

    def test_single_node_range(self):
        """Range within a single node."""
        p = _parse_paragraph("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        tm = build_text_map(p)
        positions = tm.get_nodes_for_range(6, 11)  # "world"
        assert len(positions) == 5
        assert all(pos.node is tm.positions[6].node for pos in positions)

    def test_cross_node_range(self):
        """Range spanning two nodes."""
        p = _parse_paragraph("<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p>")
        tm = build_text_map(p)
        positions = tm.get_nodes_for_range(3, 8)  # "lo wo"
        assert len(positions) == 5
        # First 3 chars from first node, last 2 from second
        assert positions[0].node is not positions[3].node


class TestFindInTextMap:
    """Tests for find_in_text_map()."""

    def test_find_simple(self):
        p = _parse_paragraph("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        tm = build_text_map(p)
        match = find_in_text_map(tm, "world")
        assert match is not None
        assert match.start == 6
        assert match.end == 11
        assert match.text == "world"
        assert not match.spans_boundary

    def test_find_not_found(self):
        p = _parse_paragraph("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        tm = build_text_map(p)
        match = find_in_text_map(tm, "missing")
        assert match is None

    def test_find_nth_occurrence(self):
        p = _parse_paragraph("<w:p><w:r><w:t>hello hello hello</w:t></w:r></w:p>")
        tm = build_text_map(p)
        m0 = find_in_text_map(tm, "hello", 0)
        m1 = find_in_text_map(tm, "hello", 1)
        m2 = find_in_text_map(tm, "hello", 2)
        m3 = find_in_text_map(tm, "hello", 3)
        assert m0 is not None and m0.start == 0
        assert m1 is not None and m1.start == 6
        assert m2 is not None and m2.start == 12
        assert m3 is None

    def test_find_across_runs(self):
        p = _parse_paragraph("<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p>")
        tm = build_text_map(p)
        match = find_in_text_map(tm, "lo wo")
        assert match is not None
        assert match.start == 3
        assert not match.spans_boundary  # same context (both regular)

    def test_find_spanning_insertion_boundary(self):
        p = _parse_paragraph(
            "<w:p>"
            "<w:r><w:t>Hello </w:t></w:r>"
            '<w:ins w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>world</w:t></w:r>"
            "</w:ins>"
            "</w:p>"
        )
        tm = build_text_map(p)
        match = find_in_text_map(tm, "Hello world")
        assert match is not None
        assert match.spans_boundary  # crosses ins boundary

    def test_find_entirely_within_insertion(self):
        p = _parse_paragraph(
            "<w:p>"
            '<w:ins w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Hello beautiful world</w:t></w:r>"
            "</w:ins>"
            "</w:p>"
        )
        tm = build_text_map(p)
        match = find_in_text_map(tm, "beautiful")
        assert match is not None
        assert not match.spans_boundary  # entirely within ins
        assert all(p.is_inside_ins for p in match.positions)
