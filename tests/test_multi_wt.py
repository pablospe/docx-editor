"""Tests for multi-w:t node scenarios.

These tests verify correct behavior when a single <w:r> contains
multiple <w:t> elements, which is valid OOXML (Word produces this
with xml:space differences, tab stops, etc.).

These tests exercise the multi-node grouping logic in:
- _build_cross_boundary_parts (group by w:t node, not run)
- _delete_regular_segment (handle segment spanning multiple w:t nodes)
- _remove_from_insertion (handle insertion with multiple w:t nodes)
"""

import defusedxml.minidom
import pytest

from docx_editor.xml_editor import build_text_map, find_in_text_map

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _parse_body(xml: str):
    """Parse XML and return the w:body element."""
    full = f"<w:body {NS}>{xml}</w:body>"
    doc = defusedxml.minidom.parseString(full)
    return doc.getElementsByTagName("w:body")[0]


class TestMultiWtTextMap:
    """Verify text map correctly handles runs with multiple w:t nodes."""

    def test_two_wt_in_one_run(self):
        """A run with two w:t nodes produces correct positions."""
        body = _parse_body("<w:p><w:r><w:t>Hello </w:t><w:t>world</w:t></w:r></w:p>")
        p = body.getElementsByTagName("w:p")[0]
        tm = build_text_map(p)
        assert tm.text == "Hello world"

        # Positions for 'H' and 'w' should reference different nodes
        pos_h = tm.positions[0]
        pos_w = tm.positions[6]
        assert pos_h.node is not pos_w.node
        assert pos_h.offset_in_node == 0  # 'H' at offset 0 in first w:t
        assert pos_w.offset_in_node == 0  # 'w' at offset 0 in second w:t

    def test_find_spanning_two_wt_in_same_run(self):
        """Finding text that spans two w:t nodes in the same run."""
        body = _parse_body("<w:p><w:r><w:t>Hello </w:t><w:t>world</w:t></w:r></w:p>")
        p = body.getElementsByTagName("w:p")[0]
        tm = build_text_map(p)
        match = find_in_text_map(tm, "o w")
        assert match is not None
        assert match.text == "o w"
        # The match spans two different w:t nodes
        nodes = {id(pos.node) for pos in match.positions}
        assert len(nodes) == 2, "Match should span two different w:t nodes"


class TestMultiWtCrossBoundary:
    """Test cross-boundary operations with multi-w:t runs.

    These use the Document API with crafted XML to trigger the multi-node paths.
    """

    def _create_doc_with_multi_wt(self, tmp_path):
        """Create a minimal .docx whose document.xml has a run with two w:t nodes."""
        import shutil
        from pathlib import Path

        from docx_editor import Document

        # First create a normal document
        fixture = Path(__file__).parent / "test_data" / "simple.docx"
        test_docx = tmp_path / "multi_wt.docx"
        shutil.copy(fixture, test_docx)

        doc = Document.open(test_docx, force_recreate=True)
        return doc

    def _inject_multi_wt_xml(self, doc):
        """Replace first paragraph's XML to have a run with two w:t nodes."""
        paragraphs = doc._document_editor.dom.getElementsByTagName("w:p")
        if not paragraphs:
            return False

        p = paragraphs[0]
        # Remove existing children
        while p.firstChild:
            p.removeChild(p.firstChild)

        # Inject a run with two w:t elements
        xml = "<w:r><w:t>Hello </w:t><w:t>beautiful world</w:t></w:r>"
        ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        fragment = defusedxml.minidom.parseString(f"<root {ns}>{xml}</root>")
        for child in fragment.documentElement.childNodes:
            imported = p.ownerDocument.importNode(child, True)
            p.appendChild(imported)
        return True

    def test_delete_spanning_two_wt_in_same_run(self, tmp_path):
        """Delete text that spans two w:t nodes within the same run.

        Old code grouped by run id, so both w:t nodes would merge into
        one entry with wrong offsets. This caused incorrect text slicing.
        """
        doc = self._create_doc_with_multi_wt(tmp_path)
        if not self._inject_multi_wt_xml(doc):
            doc.close()
            pytest.skip("Could not inject XML")

        # Verify the text is correct
        text = doc.get_visible_text()
        assert "Hello beautiful world" in text

        # Delete "o beautiful w" — spans both w:t nodes in the same run
        match = doc.find_text("o beautiful w")
        assert match is not None

        doc.delete("o beautiful w")
        text = doc.get_visible_text()
        assert "o beautiful w" not in text
        # Should keep "Hell" from first w:t and "orld" from second
        assert "Hell" in text
        assert "orld" in text
        doc.close()

    def test_replace_spanning_two_wt_in_same_run(self, tmp_path):
        """Replace text that spans two w:t nodes within the same run.

        Old code stored only the first w:t node per run, so the second
        node's text was sliced using offsets from the first node.
        """
        doc = self._create_doc_with_multi_wt(tmp_path)
        if not self._inject_multi_wt_xml(doc):
            doc.close()
            pytest.skip("Could not inject XML")

        text = doc.get_visible_text()
        assert "Hello beautiful world" in text

        # Replace "o beautiful w" — spans both w:t nodes
        match = doc.find_text("o beautiful w")
        assert match is not None

        doc.replace("o beautiful w", "O REPLACED W")
        text = doc.get_visible_text()
        assert "o beautiful w" not in text
        assert "O REPLACED W" in text
        doc.close()

    def test_delete_in_multi_wt_insertion(self, tmp_path):
        """Delete text spanning two w:t nodes inside a w:ins element.

        Old _remove_from_insertion only read positions[0].node, ignoring
        positions from other w:t nodes — causing partial deletion.
        """
        doc = self._create_doc_with_multi_wt(tmp_path)

        # Inject a paragraph with an insertion containing two w:t nodes
        paragraphs = doc._document_editor.dom.getElementsByTagName("w:p")
        if not paragraphs:
            doc.close()
            pytest.skip("No paragraphs")

        p = paragraphs[0]
        while p.firstChild:
            p.removeChild(p.firstChild)

        # Regular text + insertion with two w:t nodes
        xml = (
            "<w:r><w:t>Before </w:t></w:r>"
            '<w:ins w:id="99" w:author="Test" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>hello </w:t><w:t>world</w:t></w:r>"
            "</w:ins>"
            "<w:r><w:t> after</w:t></w:r>"
        )
        ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        fragment = defusedxml.minidom.parseString(f"<root {ns}>{xml}</root>")
        for child in fragment.documentElement.childNodes:
            imported = p.ownerDocument.importNode(child, True)
            p.appendChild(imported)

        text = doc.get_visible_text()
        assert "hello world" in text

        # Delete "o wor" — spans both w:t nodes inside the insertion
        match = doc.find_text("o wor")
        assert match is not None
        assert all(pos.is_inside_ins for pos in match.positions)

        doc.delete("o wor")
        text = doc.get_visible_text()
        assert "o wor" not in text
        assert "hell" in text
        assert "ld" in text
        doc.close()
