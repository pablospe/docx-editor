"""Tests that tracked edits preserve run-internal child order (ISSUES.md #20).

A run like ``<w:r><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r>`` must keep its
tab between "foo" and "bar" when a tracked replace/delete/insert rebuilds the
run. Previously every non-text child (w:tab, w:br, w:drawing, field chars, …)
was hoisted into a single run emitted before all text parts.
"""

from pathlib import Path

import pytest

from docx_editor.track_changes import RevisionManager
from docx_editor.xml_editor import DocxXMLEditor, compute_paragraph_hash

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
AUTHOR = "Test Author"
DATE = "2024-01-01T00:00:00Z"


@pytest.fixture
def temp_xml(tmp_path):
    """Fixture that returns a function to create temp XML files."""

    def _create_xml(body_xml: str) -> Path:
        xml = f'<?xml version="1.0" encoding="utf-8"?><w:document {NS}><w:body>{body_xml}</w:body></w:document>'
        xml_path = tmp_path / "test_doc.xml"
        xml_path.write_text(xml)
        return xml_path

    return _create_xml


def _make_manager(xml_path: Path) -> RevisionManager:
    editor = DocxXMLEditor(xml_path, rsid="00000000", author=AUTHOR)
    return RevisionManager(editor)


def _tokens_for_run(run, wrapper: str | None = None) -> list[str]:
    """Tokenize a run's direct children in document order."""

    def wrap(token: str) -> str:
        return f"{wrapper}({token})" if wrapper else token

    tokens: list[str] = []
    for child in run.childNodes:
        if child.nodeType != child.ELEMENT_NODE:
            continue
        tag = child.tagName
        if tag == "w:rPr":
            continue
        if tag in ("w:t", "w:delText"):
            text = "".join(c.data for c in child.childNodes if c.nodeType == c.TEXT_NODE)
            if text:
                tokens.append(wrap(text))
        elif tag == "w:tab":
            tokens.append(wrap("TAB"))
        elif tag == "w:br":
            tokens.append(wrap("BR"))
        elif tag == "w:fldChar":
            tokens.append(wrap(f"FLD({child.getAttribute('w:fldCharType')})"))
        elif tag == "w:instrText":
            tokens.append(wrap("INSTR"))
        elif tag == "w:drawing":
            tokens.append(wrap("DRAWING"))
        else:  # pragma: no cover - no other run content used in these fixtures
            tokens.append(wrap(tag))
    return tokens


def paragraph_tokens(manager: RevisionManager) -> list[str]:
    """Ordered content tokens of the document's first paragraph.

    Direct runs yield their tokens bare; runs inside <w:ins>/<w:del>
    wrappers yield tokens wrapped as INS(...)/DEL(...). Exact list
    comparison makes ordering assertions precise (no substring checks).
    """
    paragraph = manager.editor.dom.getElementsByTagName("w:p")[0]
    tokens: list[str] = []
    for child in paragraph.childNodes:
        if child.nodeType != child.ELEMENT_NODE:
            continue
        if child.tagName == "w:r":
            tokens.extend(_tokens_for_run(child))
        elif child.tagName in ("w:ins", "w:del"):
            wrapper = "INS" if child.tagName == "w:ins" else "DEL"
            for run in child.childNodes:
                if run.nodeType == run.ELEMENT_NODE and run.tagName == "w:r":
                    tokens.extend(_tokens_for_run(run, wrapper))
    return tokens


MID_RUN_FIXTURES = [("<w:tab/>", "TAB"), ("<w:br/>", "BR")]


class TestReplaceMidRunOrder:
    """Tracked replace keeps a mid-run tab/br between the surrounding texts."""

    @pytest.mark.parametrize(("child_xml", "token"), MID_RUN_FIXTURES)
    def test_pending_and_accepted(self, temp_xml, child_xml, token):
        xml_path = temp_xml(f"<w:p><w:r><w:t>foo</w:t>{child_xml}<w:t>bar baz</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)

        mgr.replace_text("bar", "qux")
        assert paragraph_tokens(mgr) == ["foo", token, "DEL(bar)", "INS(qux)", " baz"]

        mgr.accept_all()
        assert paragraph_tokens(mgr) == ["foo", token, "qux", " baz"]

    @pytest.mark.parametrize(("child_xml", "token"), MID_RUN_FIXTURES)
    def test_rejected(self, temp_xml, child_xml, token):
        xml_path = temp_xml(f"<w:p><w:r><w:t>foo</w:t>{child_xml}<w:t>bar baz</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)

        mgr.replace_text("bar", "qux")
        mgr.reject_all()

        tokens = paragraph_tokens(mgr)
        assert tokens[:2] == ["foo", token]
        assert "".join(tokens[2:]) == "bar baz"


class TestDeleteMidRunOrder:
    """Tracked delete keeps a mid-run tab/br between the surrounding texts."""

    @pytest.mark.parametrize(("child_xml", "token"), MID_RUN_FIXTURES)
    def test_pending_and_accepted(self, temp_xml, child_xml, token):
        xml_path = temp_xml(f"<w:p><w:r><w:t>foo</w:t>{child_xml}<w:t>bar baz</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)

        mgr.suggest_deletion("bar")
        assert paragraph_tokens(mgr) == ["foo", token, "DEL(bar)", " baz"]

        mgr.accept_all()
        assert paragraph_tokens(mgr) == ["foo", token, " baz"]

    @pytest.mark.parametrize(("child_xml", "token"), MID_RUN_FIXTURES)
    def test_rejected(self, temp_xml, child_xml, token):
        xml_path = temp_xml(f"<w:p><w:r><w:t>foo</w:t>{child_xml}<w:t>bar baz</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)

        mgr.suggest_deletion("bar")
        mgr.reject_all()

        tokens = paragraph_tokens(mgr)
        assert tokens[:2] == ["foo", token]
        assert "".join(tokens[2:]) == "bar baz"


class TestCrossRunReplaceOrder:
    """Cross-run replace (multi-run per-run loop) keeps a trailing tab in place."""

    def test_tab_stays_between_replacement_and_trailing_text(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>alpha </w:t></w:r><w:r><w:t>beta</w:t><w:tab/><w:t>gamma</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)

        mgr.replace_text("alpha beta", "X")
        assert paragraph_tokens(mgr) == ["DEL(alpha )", "DEL(beta)", "INS(X)", "TAB", "gamma"]

        mgr.accept_all()
        assert paragraph_tokens(mgr) == ["X", "TAB", "gamma"]


class TestMixedStateDeleteOrder:
    """A delete spanning regular text and an own <w:ins> routes through
    _delete_regular_segment; the tab must stay between "foo" and the del."""

    def _fixture(self, temp_xml) -> Path:
        return temp_xml(
            "<w:p><w:r><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r>"
            f'<w:ins w:id="5" w:author="{AUTHOR}" w:date="{DATE}">'
            "<w:r><w:t>qux</w:t></w:r></w:ins></w:p>"
        )

    def test_pending_and_accepted(self, temp_xml):
        mgr = _make_manager(self._fixture(temp_xml))

        mgr.suggest_deletion("barqux")
        assert paragraph_tokens(mgr) == ["foo", "TAB", "DEL(bar)"]

        mgr.accept_all()
        assert paragraph_tokens(mgr) == ["foo", "TAB"]

    def test_rejected(self, temp_xml):
        mgr = _make_manager(self._fixture(temp_xml))

        mgr.suggest_deletion("barqux")
        mgr.reject_all()
        # The own pending "qux" insertion was consumed in place; rejecting
        # restores only the tracked deletion of "bar".
        assert paragraph_tokens(mgr) == ["foo", "TAB", "bar"]

    def test_multi_wt_regular_segment_keeps_node_order(self, temp_xml):
        """A regular segment spanning three w:t nodes (plus an empty sibling)
        wraps each node's matched text in its own <w:del>, in order."""
        xml_path = temp_xml(
            "<w:p><w:r><w:t>ab</w:t><w:t></w:t><w:t>cd</w:t><w:t>ef</w:t></w:r>"
            f'<w:ins w:id="5" w:author="{AUTHOR}" w:date="{DATE}">'
            "<w:r><w:t>gh</w:t></w:r></w:ins></w:p>"
        )
        mgr = _make_manager(xml_path)

        mgr.suggest_deletion("bcdefgh")
        assert paragraph_tokens(mgr) == ["a", "DEL(b)", "DEL(cd)", "DEL(ef)"]

        mgr.accept_all()
        assert paragraph_tokens(mgr) == ["a"]


class TestInsertAroundTab:
    """Tracked insert lands on the correct side of a mid-run tab."""

    def test_insert_after_text_precedes_tab(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)

        mgr.insert_text_after("foo", "X")
        assert paragraph_tokens(mgr) == ["foo", "INS(X)", "TAB", "bar"]

    def test_insert_before_text_follows_tab(self, temp_xml):
        xml_path = temp_xml("<w:p><w:r><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r></w:p>")
        mgr = _make_manager(xml_path)

        mgr.insert_text_before("bar", "X")
        assert paragraph_tokens(mgr) == ["foo", "TAB", "INS(X)", "bar"]


class TestEmptySiblingWt:
    """Empty sibling w:t nodes are dropped from rebuilt runs, never emitted
    as empty runs, while the surrounding order is preserved."""

    FIXTURE = "<w:p><w:r><w:t>foo</w:t><w:t></w:t><w:tab/><w:t>bar</w:t></w:r></w:p>"

    def test_replace_skips_empty_sibling(self, temp_xml):
        mgr = _make_manager(temp_xml(self.FIXTURE))
        mgr.replace_text("bar", "qux")
        assert paragraph_tokens(mgr) == ["foo", "TAB", "DEL(bar)", "INS(qux)"]

    def test_delete_skips_empty_sibling(self, temp_xml):
        mgr = _make_manager(temp_xml(self.FIXTURE))
        mgr.suggest_deletion("bar")
        assert paragraph_tokens(mgr) == ["foo", "TAB", "DEL(bar)"]

    def test_insert_skips_empty_sibling(self, temp_xml):
        mgr = _make_manager(temp_xml(self.FIXTURE))
        mgr.insert_text_after("foo", "X")
        assert paragraph_tokens(mgr) == ["foo", "INS(X)", "TAB", "bar"]


class TestFieldCharOrder:
    """Field characters (highest-risk reordering per ISSUES.md) keep their
    begin/instr/end sequence intact and in place."""

    def test_field_children_stay_in_order(self, temp_xml):
        xml_path = temp_xml(
            "<w:p><w:r><w:t>a</w:t>"
            '<w:fldChar w:fldCharType="begin"/><w:instrText>PAGE</w:instrText>'
            '<w:fldChar w:fldCharType="end"/><w:t>b</w:t></w:r></w:p>'
        )
        mgr = _make_manager(xml_path)

        mgr.replace_text("a", "A")
        assert paragraph_tokens(mgr) == ["DEL(a)", "INS(A)", "FLD(begin)", "INSTR", "FLD(end)", "b"]


class TestTextboxNoDuplication:
    """w:t nodes nested inside a drawing's text box must stay inside the
    drawing — not be re-emitted as top-level runs (text duplication)."""

    def test_replace_outside_textbox_keeps_boxed_text_once(self, temp_xml):
        xml_path = temp_xml(
            "<w:p><w:r><w:t>x</w:t>"
            "<w:drawing><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></w:drawing>"
            "<w:t>y</w:t></w:r></w:p>"
        )
        mgr = _make_manager(xml_path)

        mgr.replace_text("x", "X")

        paragraph = mgr.editor.dom.getElementsByTagName("w:p")[0]
        assert paragraph.toxml().count("boxed") == 1
        assert paragraph_tokens(mgr) == ["DEL(x)", "INS(X)", "DRAWING", "y"]
        drawings = mgr.editor.dom.getElementsByTagName("w:drawing")
        assert len(drawings) == 1
        assert "boxed" in drawings[0].toxml()


def _has_direct_rpr(run) -> bool:
    """True if ``run`` has a direct w:rPr child."""
    return any(c.nodeType == c.ELEMENT_NODE and c.tagName == "w:rPr" for c in run.childNodes)


def _top_level_runs(paragraph) -> list:
    """Direct runs of ``paragraph`` plus runs under direct w:ins/w:del wrappers."""
    runs = []
    for child in paragraph.childNodes:
        if child.nodeType != child.ELEMENT_NODE:
            continue
        if child.tagName == "w:r":
            runs.append(child)
        elif child.tagName in ("w:ins", "w:del"):
            runs.extend(r for r in child.childNodes if r.nodeType == r.ELEMENT_NODE and r.tagName == "w:r")
    return runs


class TestNestedRPrNoLeak:
    """A w:rPr nested inside a drawing's text box must not leak into rebuilt
    top-level runs when the outer run has no direct rPr (ISSUES.md #31)."""

    def test_replace_does_not_inherit_textbox_rpr(self, temp_xml):
        xml_path = temp_xml(
            "<w:p><w:r><w:t>x</w:t>"
            "<w:drawing><w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>boxed</w:t></w:r></w:p></w:txbxContent></w:drawing>"
            "<w:t>y</w:t></w:r></w:p>"
        )
        mgr = _make_manager(xml_path)

        mgr.replace_text("x", "X")

        paragraph = mgr.editor.dom.getElementsByTagName("w:p")[0]
        assert paragraph_tokens(mgr) == ["DEL(x)", "INS(X)", "DRAWING", "y"]
        # The bold prop stays inside the drawing's text box — nowhere else.
        assert paragraph.toxml().count("<w:b/>") == 1
        assert "<w:b/>" in mgr.editor.dom.getElementsByTagName("w:drawing")[0].toxml()
        for run in _top_level_runs(paragraph):
            assert not _has_direct_rpr(run)

    def test_empty_paragraph_insert_does_not_inherit_textbox_rpr(self, temp_xml):
        """Insert into a paragraph with no visible text: the created <w:ins>
        run must not adopt the drawing's nested rPr."""
        xml_path = temp_xml(
            "<w:p><w:r>"
            "<w:drawing><w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr></w:r></w:p></w:txbxContent></w:drawing>"
            "</w:r></w:p>"
        )
        mgr = _make_manager(xml_path)
        paragraph = mgr.editor.dom.getElementsByTagName("w:p")[0]
        ref = f"P1#{compute_paragraph_hash(paragraph)}"

        mgr.rewrite_paragraph(ref, "inserted")

        assert paragraph.toxml().count("<w:b/>") == 1
        ins_elems = [c for c in paragraph.childNodes if c.nodeType == c.ELEMENT_NODE and c.tagName == "w:ins"]
        assert len(ins_elems) == 1
        for run in _top_level_runs(paragraph):
            assert not _has_direct_rpr(run)
