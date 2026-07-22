"""Tests for cross-boundary replace operations (Phase 4)."""

from pathlib import Path

import defusedxml.minidom
import pytest
from conftest import find_ref, replace_document_xml

from docx_editor import Document
from docx_editor.track_changes import RevisionManager, _escape_xml
from docx_editor.xml_editor import DocxXMLEditor, build_text_map

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _make_editor_with_formatted_runs(tmp_path: Path, runs: list[tuple[str, str]]) -> DocxXMLEditor:
    """Create a DocxXMLEditor with one paragraph of (text, rPr_xml) runs."""
    run_xml = ""
    for text, rPr_xml in runs:
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


def _make_editor_with_split_runs(tmp_path: Path, runs: list[str], rPr_xml: str = "") -> DocxXMLEditor:
    """Create a DocxXMLEditor with a paragraph containing multiple runs.

    Each string in `runs` becomes a separate <w:r><w:t>...</w:t></w:r>.
    """
    return _make_editor_with_formatted_runs(tmp_path, [(text, rPr_xml) for text in runs])


def _del_text(editor: DocxXMLEditor) -> str:
    """Concatenated content of every w:delText in the document."""
    return "".join(dt.firstChild.data for dt in editor.dom.getElementsByTagName("w:delText") if dt.firstChild)


def _ins_text(ins_elem) -> str:
    """Concatenated w:t content of a w:ins element."""
    return "".join(wt.firstChild.data for wt in ins_elem.getElementsByTagName("w:t") if wt.firstChild)


class TestCrossBoundaryReplaceRegression:
    """Regression: single-element replace still works."""

    def test_replace_within_single_element(self, clean_workspace):
        """Replace text contained in a single w:t element."""
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "fox")
        new_ref = doc.replace("fox", "cat", paragraph=ref)
        assert isinstance(new_ref, str)
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
        ref = find_ref(doc, "fo")
        new_ref = doc.replace("fox", "cat", paragraph=ref)
        assert isinstance(new_ref, str)

        output = temp_dir / "output.docx"
        doc.save(output)
        doc.close()

        # Reopen and verify
        doc2 = Document.open(output, force_recreate=True)
        text = doc2.get_visible_text()
        assert "cat" in text
        assert "fox" not in text
        doc2.close()


class TestReplaceTrimming:
    """Words shared by find/replace_with are trimmed before revisions are written."""

    def test_only_changed_words_become_revisions(self, temp_dir):
        editor = _make_editor_with_formatted_runs(
            temp_dir,
            [
                ("The initial term of ", ""),
                ("two (2) years", "<w:rPr><w:b/></w:rPr>"),
                (", unless terminated.", ""),
            ],
        )
        mgr = RevisionManager(editor)
        mgr.replace_text("term of two (2) years, unless", "term of three (3) years, unless")

        assert _del_text(editor) == "two (2)"
        ins_elems = editor.dom.getElementsByTagName("w:ins")
        assert len(ins_elems) == 1
        assert _ins_text(ins_elems[0]) == "three (3)"

        tm = build_text_map(editor.dom.getElementsByTagName("w:p")[0])
        assert tm.text == "The initial term of three (3) years, unless terminated."
        # Untrimmed affix words never enter a revision
        for pos in tm.positions[: len("The initial term of ")]:
            assert not pos.is_inside_ins and not pos.is_inside_del

    def test_replace_degenerating_to_pure_insert(self, temp_dir):
        editor = _make_editor_with_split_runs(temp_dir, ["The cat sat."])
        mgr = RevisionManager(editor)
        mgr.replace_text("cat sat", "cat quickly sat")

        assert len(editor.dom.getElementsByTagName("w:del")) == 0
        ins_elems = editor.dom.getElementsByTagName("w:ins")
        assert len(ins_elems) == 1
        assert _ins_text(ins_elems[0]) == "quickly "
        # Edge whitespace of the inserted fragment must be preserved
        wt = ins_elems[0].getElementsByTagName("w:t")[0]
        assert wt.getAttribute("xml:space") == "preserve"

        tm = build_text_map(editor.dom.getElementsByTagName("w:p")[0])
        assert tm.text == "The cat quickly sat."

    def test_replace_degenerating_to_pure_insert_at_start(self, temp_dir):
        editor = _make_editor_with_split_runs(temp_dir, ["sat down."])
        mgr = RevisionManager(editor)
        mgr.replace_text("sat down", "quickly sat down")

        assert len(editor.dom.getElementsByTagName("w:del")) == 0
        ins_elems = editor.dom.getElementsByTagName("w:ins")
        assert len(ins_elems) == 1
        assert _ins_text(ins_elems[0]) == "quickly "

        tm = build_text_map(editor.dom.getElementsByTagName("w:p")[0])
        assert tm.text == "quickly sat down."

    def test_replace_degenerating_to_pure_delete(self, temp_dir):
        editor = _make_editor_with_split_runs(temp_dir, ["Please delete this word now."])
        mgr = RevisionManager(editor)
        mgr.replace_text("delete this word", "delete word")

        assert len(editor.dom.getElementsByTagName("w:ins")) == 0
        assert _del_text(editor) == "this "

        tm = build_text_map(editor.dom.getElementsByTagName("w:p")[0])
        assert tm.text == "Please delete word now."

    def test_empty_replacement_writes_no_empty_ins(self, temp_dir):
        editor = _make_editor_with_split_runs(temp_dir, ["Remove here please."])
        mgr = RevisionManager(editor)
        mgr.replace_text(" here", "")

        assert len(editor.dom.getElementsByTagName("w:ins")) == 0
        assert _del_text(editor) == " here"

    def test_replace_with_identical_text_is_noop(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            revisions_before = doc.list_revisions()
            ref = find_ref(doc, "fox")
            text_before = doc.get_visible_text()

            result = doc.replace("fox", "fox", paragraph=ref)

            assert result == ref
            assert result.group_id is None
            assert result.revision_ids == ()
            assert doc.get_visible_text() == text_before
            assert doc.list_revisions() == revisions_before
        finally:
            doc.close()


class TestMajorityRPrFormatting:
    """The replacement insertion carries the majority-by-characters rPr."""

    BOLD = "<w:rPr><w:b/></w:rPr>"
    ITALIC = "<w:rPr><w:i/></w:rPr>"

    def test_majority_run_wins_over_first(self, temp_dir):
        # Plain run contributes 3 chars, bold run 17: bold must win even
        # though the match starts in the plain run.
        editor = _make_editor_with_formatted_runs(temp_dir, [("ab ", ""), ("bold body of text", self.BOLD)])
        mgr = RevisionManager(editor)
        # find/replace share no words, so trimming cannot mask the rule
        mgr.replace_text("ab bold body of text", "zz")

        ins_elems = editor.dom.getElementsByTagName("w:ins")
        assert len(ins_elems) == 1
        assert len(ins_elems[0].getElementsByTagName("w:b")) == 1

    def test_tie_breaks_to_earliest_run(self, temp_dir):
        editor = _make_editor_with_formatted_runs(temp_dir, [("abcd", self.BOLD), ("efgh", self.ITALIC)])
        mgr = RevisionManager(editor)
        mgr.replace_text("abcdefgh", "XY")

        ins_elems = editor.dom.getElementsByTagName("w:ins")
        assert len(ins_elems) == 1
        assert len(ins_elems[0].getElementsByTagName("w:b")) == 1
        assert len(ins_elems[0].getElementsByTagName("w:i")) == 0


class TestFormattingPreservedOnAccept:
    """E1/E3/E4: formatting survives accepting a boundary-spanning replace."""

    @pytest.mark.parametrize(
        ("rpr", "marker_tag"),
        [
            ("<w:rPr><w:b/></w:rPr>", "w:b"),
            ('<w:rPr><w:u w:val="single"/></w:rPr>', "w:u"),
            ("<w:rPr><w:i/></w:rPr>", "w:i"),
        ],
        ids=["bold", "underline", "italic"],
    )
    def test_formatting_survives_accept(self, simple_docx, temp_dir, rpr, marker_tag):
        doc_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<w:document {NS}><w:body><w:p>"
            "<w:r><w:t>The initial term of </w:t></w:r>"
            f"<w:r>{rpr}<w:t>two (2) years</w:t></w:r>"
            '<w:r><w:t xml:space="preserve">, unless terminated.</w:t></w:r>'
            "</w:p></w:body></w:document>"
        )
        src = temp_dir / "contract.docx"
        replace_document_xml(simple_docx, src, doc_xml)

        doc = Document.open(src, force_recreate=True)
        ref = find_ref(doc, "initial term")
        doc.replace("term of two (2) years, unless", "term of three (3) years, unless", paragraph=ref)

        # The replacement insertion carries the formatting marker
        ins_elems = doc._document_editor.dom.getElementsByTagName("w:ins")
        assert len(ins_elems) == 1
        assert len(ins_elems[0].getElementsByTagName(marker_tag)) == 1

        out = temp_dir / "contract_out.docx"
        doc.save(out)
        doc.close()

        doc2 = Document.open(out, force_recreate=True)
        try:
            doc2.accept_all()
            text = doc2.get_visible_text()
            assert "term of three (3) years, unless" in text
            assert "two (2)" not in text
            # The changed words remain inside a run carrying the marker
            dom = doc2._document_editor.dom
            marked = "".join(
                wt.firstChild.data
                for r in dom.getElementsByTagName("w:r")
                if r.getElementsByTagName(marker_tag)
                for wt in r.getElementsByTagName("w:t")
                if wt.firstChild
            )
            assert "three (3)" in marked
        finally:
            doc2.close()


class TestTrimmedReplaceRoundtrip:
    """#46 interplay: a trimmed replace's group survives save → reopen."""

    def test_group_reconstructed_and_acceptable_in_second_session(self, clean_workspace, temp_dir):
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "quick brown fox")
        result = doc.replace("quick brown fox", "quick red fox", paragraph=ref)
        assert result.group_id is not None
        assert len(result.revision_ids) == 2  # minimal pair: del "brown" + ins "red"

        out = temp_dir / "trimmed.docx"
        doc.save(out)
        doc.close()

        doc2 = Document.open(out, force_recreate=True)
        try:
            revs = [r for r in doc2.list_revisions() if r.text in ("brown", "red")]
            assert {r.type for r in revs} == {"deletion", "insertion"}
            group_id = revs[0].group_id
            assert group_id is not None
            assert all(r.group_id == group_id for r in revs)

            accepted = doc2.accept_group(group_id)
            assert accepted == 2
            text = doc2.get_visible_text()
            assert "quick red fox" in text
            assert "brown" not in text
        finally:
            doc2.close()
