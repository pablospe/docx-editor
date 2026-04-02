"""Tests for rewrite_paragraph() — word-level diff with tracked changes."""

import shutil
import tempfile
from pathlib import Path

import pytest

from docx_editor import Document, HashMismatchError
from docx_editor.track_changes import _tokenize_words


@pytest.fixture
def rewrite_doc():
    """Build a document with 5 known paragraphs for rewrite testing."""
    test_data = Path(__file__).parent / "test_data" / "simple.docx"
    tmp = tempfile.mkdtemp(prefix="rewrite_test_")
    dest = Path(tmp) / "test.docx"
    shutil.copy(test_data, dest)

    doc = Document.open(dest)
    doc.accept_all()
    doc.save()
    doc.close()

    # Inject paragraphs directly
    doc = Document.open(dest, force_recreate=True)
    editor = doc._document_editor
    body = editor.dom.getElementsByTagName("w:body")[0]

    # Remove existing paragraphs
    for p in list(editor.dom.getElementsByTagName("w:p")):
        if p.parentNode == body:
            body.removeChild(p)

    sect_pr = editor.dom.getElementsByTagName("w:sectPr")
    insert_before = sect_pr[0] if sect_pr else None

    paragraphs = [
        "The committee shall review and proceed with the annual budget proposal.",
        "All members must attend the quarterly meeting without exception.",
        "The report includes findings from the committee investigation.",
        "",  # empty paragraph
        "Final approval requires a majority vote by the board.",
    ]

    for text in paragraphs:
        if text:
            p_xml = f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        else:
            p_xml = "<w:p/>"
        nodes = editor._parse_fragment(p_xml)
        for node in nodes:
            if insert_before:
                body.insertBefore(node, insert_before)
            else:
                body.appendChild(node)

    editor.save()
    save_path = doc.save()
    doc.close()

    doc = Document.open(save_path, force_recreate=True)
    yield doc, Path(tmp)
    doc.close()
    shutil.rmtree(tmp, ignore_errors=True)


@pytest.fixture
def bold_doc():
    """Build a document with a bold-formatted paragraph for formatting tests."""
    test_data = Path(__file__).parent / "test_data" / "simple.docx"
    tmp = tempfile.mkdtemp(prefix="rewrite_bold_test_")
    dest = Path(tmp) / "test.docx"
    shutil.copy(test_data, dest)

    doc = Document.open(dest)
    doc.accept_all()
    doc.save()
    doc.close()

    doc = Document.open(dest, force_recreate=True)
    editor = doc._document_editor
    body = editor.dom.getElementsByTagName("w:body")[0]

    for p in list(editor.dom.getElementsByTagName("w:p")):
        if p.parentNode == body:
            body.removeChild(p)

    sect_pr = editor.dom.getElementsByTagName("w:sectPr")
    insert_before = sect_pr[0] if sect_pr else None

    # Bold paragraph
    p_xml = '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">The bold committee meets today.</w:t></w:r></w:p>'
    nodes = editor._parse_fragment(p_xml)
    for node in nodes:
        if insert_before:
            body.insertBefore(node, insert_before)
        else:
            body.appendChild(node)

    editor.save()
    save_path = doc.save()
    doc.close()

    doc = Document.open(save_path, force_recreate=True)
    yield doc, Path(tmp)
    doc.close()
    shutil.rmtree(tmp, ignore_errors=True)


class TestTokenizeWords:
    def test_basic(self):
        tokens = _tokenize_words("hello world")
        assert tokens == ["hello", " ", "world"]

    def test_multiple_spaces(self):
        tokens = _tokenize_words("a  b")
        assert tokens == ["a", "  ", "b"]

    def test_empty(self):
        tokens = _tokenize_words("")
        assert tokens == []


class TestRewriteParagraph:
    def test_word_replacement(self, rewrite_doc):
        """Replace 'committee' with 'board' and 'proceed with' with 'approve'."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        doc.rewrite_paragraph(
            ref,
            "The board shall review and approve the annual budget proposal.",
        )

        vis = doc.get_visible_text()
        assert "board" in vis
        assert "approve" in vis

        # Verify tracked changes were created
        revisions = doc.list_revisions()
        assert len(revisions) > 0

        # Check for deletions and insertions
        del_texts = [r.text for r in revisions if r.type == "deletion"]
        ins_texts = [r.text for r in revisions if r.type == "insertion"]
        assert any("committee" in t for t in del_texts)
        assert any("board" in t for t in ins_texts)

    def test_addition_only(self, rewrite_doc):
        """Add a word — should produce insertions only."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[1].split("|")[0]

        doc.rewrite_paragraph(
            ref,
            "All members must always attend the quarterly meeting without exception.",
        )

        vis = doc.get_visible_text()
        assert "always" in vis

        revisions = doc.list_revisions()
        # Should have insertion(s) but no deletions
        ins_revs = [r for r in revisions if r.type == "insertion"]
        del_revs = [r for r in revisions if r.type == "deletion"]
        assert len(ins_revs) > 0
        assert len(del_revs) == 0

    def test_deletion_only(self, rewrite_doc):
        """Remove a word — should produce deletions only."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[1].split("|")[0]

        doc.rewrite_paragraph(
            ref,
            "All members must attend the meeting without exception.",
        )

        vis = doc.get_visible_text()
        assert "quarterly" not in vis

        revisions = doc.list_revisions()
        del_revs = [r for r in revisions if r.type == "deletion"]
        ins_revs = [r for r in revisions if r.type == "insertion"]
        assert len(del_revs) > 0
        assert len(ins_revs) == 0

    def test_noop_rewrite(self, rewrite_doc):
        """Same text produces no revisions."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        original_text = "The committee shall review and proceed with the annual budget proposal."
        doc.rewrite_paragraph(ref, original_text)

        revisions = doc.list_revisions()
        assert len(revisions) == 0

    def test_hash_mismatch(self, rewrite_doc):
        """Stale hash raises HashMismatchError."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        # Mutate the paragraph
        doc.rewrite_paragraph(ref, "Changed text here.")

        # Now use old ref (stale hash)
        with pytest.raises(HashMismatchError):
            doc.rewrite_paragraph(ref, "Another change.")

    def test_rewrite_empty_paragraph(self, rewrite_doc):
        """Empty paragraph to text — all inserted."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        # P4 is the empty paragraph (index 3)
        ref = refs[3].split("|")[0]

        doc.rewrite_paragraph(ref, "Brand new text.")

        vis = doc.get_visible_text()
        assert "Brand new text." in vis

        revisions = doc.list_revisions()
        ins_revs = [r for r in revisions if r.type == "insertion"]
        assert len(ins_revs) > 0

    def test_rewrite_to_empty(self, rewrite_doc):
        """Text to empty string — all deleted."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[2].split("|")[0]

        original = "The report includes findings from the committee investigation."
        doc.rewrite_paragraph(ref, "")

        vis = doc.get_visible_text()
        assert original not in vis

        revisions = doc.list_revisions()
        del_revs = [r for r in revisions if r.type == "deletion"]
        assert len(del_revs) > 0

    def test_formatting_preserved(self, bold_doc):
        """Bold formatting is inherited by inserted text."""
        doc, _ = bold_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        doc.rewrite_paragraph(ref, "The bold board meets today.")

        # Check that inserted text has bold rPr
        revisions = doc.list_revisions()
        ins_revs = [r for r in revisions if r.type == "insertion"]
        assert len(ins_revs) > 0

        # Verify the w:ins element contains w:b in rPr
        editor = doc._document_editor
        for ins_elem in editor.dom.getElementsByTagName("w:ins"):
            rPr_elems = ins_elem.getElementsByTagName("w:rPr")
            if rPr_elems:
                b_elems = rPr_elems[0].getElementsByTagName("w:b")
                assert len(b_elems) > 0, "Inserted text should inherit bold formatting"
