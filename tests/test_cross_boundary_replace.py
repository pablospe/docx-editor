"""Tests for cross-boundary replace operations (Phase 4)."""

from pathlib import Path

import defusedxml.minidom

from docx_editor import Document
from docx_editor.track_changes import RevisionManager, _escape_xml
from docx_editor.xml_editor import DocxXMLEditor, build_text_map

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _make_editor_with_split_runs(tmp_path: Path, runs: list[str], rPr_xml: str = "") -> DocxXMLEditor:
    """Create a DocxXMLEditor with a paragraph containing multiple runs.

    Each string in `runs` becomes a separate <w:r><w:t>...</w:t></w:r>.
    """
    run_xml = ""
    for text in runs:
        space_attr = ' xml:space="preserve"' if text and (text[0].isspace() or text[-1].isspace()) else ""
        run_xml += f"<w:r>{rPr_xml}<w:t{space_attr}>{_escape_xml(text)}</w:t></w:r>"

    doc_xml = f"""<?xml version="1.0" encoding="utf-8"?>
<w:document {NS}>
  <w:body>
    <w:p>{run_xml}</w:p>
  </w:body>
</w:document>"""

    xml_path = tmp_path / "document.xml"
    xml_path.write_text(doc_xml, encoding="utf-8")
    return DocxXMLEditor(xml_path, rsid="00AA1234", author="Test Author")


class TestCrossBoundaryReplaceRegression:
    """Regression: single-element replace still works."""

    def test_replace_within_single_element(self, clean_workspace):
        """Replace text contained in a single w:t element."""
        doc = Document.open(clean_workspace)
        change_id = doc.replace("fox", "cat")
        assert change_id >= 0
        assert "cat" in doc.get_visible_text()
        assert "fox" not in doc.get_visible_text()
        doc.close()


class TestReplaceAcrossNodes:
    """Unit tests for _replace_across_nodes using crafted XML."""

    def test_replace_text_spanning_two_runs(self, temp_dir):
        """Replace text that spans two consecutive runs."""
        editor = _make_editor_with_split_runs(temp_dir, ["Hello wo", "rld!"])
        mgr = RevisionManager(editor)

        change_id = mgr.replace_text("wo" + "rld", "universe")
        assert change_id >= 0

        # Verify visible text
        paras = editor.dom.getElementsByTagName("w:p")
        tm = build_text_map(paras[0])
        assert "universe" in tm.text
        assert "world" not in tm.text
        # "Hello " should be preserved
        assert tm.text.startswith("Hello ")

    def test_replace_text_spanning_three_runs(self, temp_dir):
        """Replace text that spans three consecutive runs."""
        editor = _make_editor_with_split_runs(temp_dir, ["ab", "cd", "ef"])
        mgr = RevisionManager(editor)

        change_id = mgr.replace_text("bcde", "XY")
        assert change_id >= 0

        paras = editor.dom.getElementsByTagName("w:p")
        tm = build_text_map(paras[0])
        assert "aXYf" == tm.text

    def test_replace_entire_run_contents(self, temp_dir):
        """Replace text that exactly covers two full runs."""
        editor = _make_editor_with_split_runs(temp_dir, ["Hello", " World"])
        mgr = RevisionManager(editor)

        change_id = mgr.replace_text("Hello World", "Goodbye")
        assert change_id >= 0

        paras = editor.dom.getElementsByTagName("w:p")
        tm = build_text_map(paras[0])
        assert "Goodbye" == tm.text

    def test_replace_preserves_rpr(self, temp_dir):
        """Replacement insertion preserves run properties from first run."""
        rPr = "<w:rPr><w:b/></w:rPr>"
        editor = _make_editor_with_split_runs(temp_dir, ["Hel", "lo"], rPr_xml=rPr)
        mgr = RevisionManager(editor)

        mgr.replace_text("Hello", "Hi")

        # The insertion should contain w:b
        ins_elems = editor.dom.getElementsByTagName("w:ins")
        assert len(ins_elems) == 1
        rPr_elems = ins_elems[0].getElementsByTagName("w:b")
        assert len(rPr_elems) >= 1

    def test_replace_creates_del_and_ins(self, temp_dir):
        """Cross-boundary replace creates proper w:del and w:ins elements."""
        editor = _make_editor_with_split_runs(temp_dir, ["foo", "bar"])
        mgr = RevisionManager(editor)

        mgr.replace_text("foobar", "baz")

        del_elems = editor.dom.getElementsByTagName("w:del")
        ins_elems = editor.dom.getElementsByTagName("w:ins")
        assert len(del_elems) >= 1
        assert len(ins_elems) == 1

        # Verify del contains the old text
        del_texts = editor.dom.getElementsByTagName("w:delText")
        del_content = "".join(dt.firstChild.data for dt in del_texts if dt.firstChild)
        assert del_content == "foobar"

    def test_replace_saves_valid_xml(self, temp_dir):
        """After cross-boundary replace, XML can be saved and re-parsed."""
        editor = _make_editor_with_split_runs(temp_dir, ["Hello wo", "rld!"])
        mgr = RevisionManager(editor)

        mgr.replace_text("world", "universe")
        editor.save()

        # Re-parse to verify valid XML
        reparsed = defusedxml.minidom.parse(str(temp_dir / "document.xml"))
        assert reparsed is not None


class TestSuggestDeletionAcrossNodes:
    """Unit tests for cross-boundary suggest_deletion."""

    def test_delete_text_spanning_two_runs(self, temp_dir):
        """Delete text that spans two consecutive runs."""
        editor = _make_editor_with_split_runs(temp_dir, ["Hello wo", "rld!"])
        mgr = RevisionManager(editor)

        change_id = mgr.suggest_deletion("world")
        assert change_id >= 0

        paras = editor.dom.getElementsByTagName("w:p")
        tm = build_text_map(paras[0])
        assert "world" not in tm.text
        assert "Hello " in tm.text

    def test_delete_text_spanning_three_runs(self, temp_dir):
        """Delete text that spans three consecutive runs."""
        editor = _make_editor_with_split_runs(temp_dir, ["ab", "cd", "ef"])
        mgr = RevisionManager(editor)

        change_id = mgr.suggest_deletion("bcde")
        assert change_id >= 0

        paras = editor.dom.getElementsByTagName("w:p")
        tm = build_text_map(paras[0])
        assert tm.text == "af"


class TestCrossBoundaryReplaceRoundtrip:
    """Integration round-trip tests using real docx files."""

    def test_replace_across_boundary_roundtrip(self, clean_workspace, temp_dir):
        """Round-trip: create split runs, replace across boundary, save, reopen."""
        doc = Document.open(clean_workspace)

        # Manipulate the XML directly to create split runs
        editor = doc._document_editor
        paras = editor.dom.getElementsByTagName("w:p")
        for p in paras:
            tm = build_text_map(p)
            if "fox" in tm.text:
                # Find the w:t containing "fox"
                for t_node in p.getElementsByTagName("w:t"):
                    if t_node.firstChild and "fox" in t_node.firstChild.data:
                        text = t_node.firstChild.data
                        idx = text.find("fox")
                        # Split: before "fo" | "x jumps" | rest
                        before = text[: idx + 2]  # "...fo"
                        after = text[idx + 2 :]  # "x jumps..."

                        run = t_node.parentNode
                        rPr_xml = ""
                        rPr_elems = run.getElementsByTagName("w:rPr")
                        if rPr_elems:
                            rPr_xml = rPr_elems[0].toxml()

                        # Replace with two runs
                        new_xml = (
                            f"<w:r>{rPr_xml}<w:t>{_escape_xml(before)}</w:t></w:r>"
                            f"<w:r>{rPr_xml}<w:t>{_escape_xml(after)}</w:t></w:r>"
                        )
                        editor.replace_node(run, new_xml)
                        break
                break

        # Now "fox" spans two runs: "...fo" and "x jumps..."
        change_id = doc.replace("fox", "cat")
        assert change_id >= 0

        output = temp_dir / "output.docx"
        doc.save(output)
        doc.close()

        # Reopen and verify
        doc2 = Document.open(output, force_recreate=True)
        text = doc2.get_visible_text()
        assert "cat" in text
        assert "fox" not in text
        doc2.close()
