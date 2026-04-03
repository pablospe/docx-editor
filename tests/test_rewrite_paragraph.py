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
            assert len(rPr_elems) > 0, "Inserted run should have rPr"
            b_elems = rPr_elems[0].getElementsByTagName("w:b")
            assert len(b_elems) > 0, "Inserted text should inherit bold formatting"


class TestRewriteDuplicateText:
    """Tests for paragraphs containing duplicate text (Bug #1 fix)."""

    @pytest.fixture
    def dup_doc(self):
        """Build a document with a paragraph containing repeated words."""
        test_data = Path(__file__).parent / "test_data" / "simple.docx"
        tmp = tempfile.mkdtemp(prefix="rewrite_dup_test_")
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

        # Paragraph with "the" appearing 3 times
        text = "The cat and the dog and the bird"
        p_xml = f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
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

    def test_replace_second_occurrence(self, dup_doc):
        """Replacing only the second 'the' should not touch the first or third."""
        doc, _ = dup_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        # Change "the dog" to "the hound" — second "the" stays, "dog" becomes "hound"
        doc.rewrite_paragraph(ref, "The cat and the hound and the bird")

        vis = doc.get_visible_text()
        assert "The cat and the hound and the bird" in vis

        revisions = doc.list_revisions()
        del_texts = [r.text for r in revisions if r.type == "deletion"]
        ins_texts = [r.text for r in revisions if r.type == "insertion"]
        assert any("dog" in t for t in del_texts)
        assert any("hound" in t for t in ins_texts)
        # "the" should NOT appear in deletions or insertions
        assert not any(t.strip() == "the" for t in del_texts)
        assert not any(t.strip() == "the" for t in ins_texts)

    def test_replace_last_occurrence(self, dup_doc):
        """Replacing only the last 'the' should not touch earlier occurrences."""
        doc, _ = dup_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        doc.rewrite_paragraph(ref, "The cat and the dog and a bird")

        vis = doc.get_visible_text()
        assert "The cat and the dog and a bird" in vis

        revisions = doc.list_revisions()
        del_texts = [r.text for r in revisions if r.type == "deletion"]
        ins_texts = [r.text for r in revisions if r.type == "insertion"]
        assert any("the" in t.lower() for t in del_texts)
        assert any(t.strip() == "a" for t in ins_texts)


def _make_doc_with_paragraphs(paragraphs: list[str]):
    """Helper: build a document with specific paragraph texts. Returns (doc, tmp_dir)."""
    test_data = Path(__file__).parent / "test_data" / "simple.docx"
    tmp = tempfile.mkdtemp(prefix="rewrite_stress_")
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
    return doc, Path(tmp)


class TestRewriteStructuralEdits:
    """Stress tests for structural edits where rewrite_paragraph is the right tool."""

    def test_passive_to_active_voice(self):
        """Passive→active voice conversion restructures clauses."""
        doc, tmp = _make_doc_with_paragraphs(
            ["The proposal was reviewed by the committee and was approved by the board."]
        )
        try:
            refs = doc.list_paragraphs()
            ref = refs[0].split("|")[0]

            doc.rewrite_paragraph(
                ref, "The committee reviewed the proposal and the board approved it."
            )

            vis = doc.get_visible_text()
            assert "The committee reviewed the proposal and the board approved it." in vis

            revisions = doc.list_revisions()
            assert len(revisions) > 0
            del_texts = [r.text for r in revisions if r.type == "deletion"]
            ins_texts = [r.text for r in revisions if r.type == "insertion"]
            assert len(del_texts) > 0
            assert len(ins_texts) > 0
        finally:
            doc.close()
            shutil.rmtree(tmp, ignore_errors=True)

    def test_reorder_list_items(self):
        """Reordering items in a list produces correct tracked changes."""
        doc, tmp = _make_doc_with_paragraphs(
            ["Deliverables include the final report, executive summary, and presentation slides."]
        )
        try:
            refs = doc.list_paragraphs()
            ref = refs[0].split("|")[0]

            doc.rewrite_paragraph(
                ref,
                "Deliverables include the presentation slides, final report, and executive summary.",
            )

            vis = doc.get_visible_text()
            assert "presentation slides, final report, and executive summary" in vis

            revisions = doc.list_revisions()
            assert len(revisions) > 0
        finally:
            doc.close()
            shutil.rmtree(tmp, ignore_errors=True)

    def test_full_sentence_rephrasing(self):
        """Complete rephrasing produces tracked changes for restructured sentence."""
        doc, tmp = _make_doc_with_paragraphs(
            ["The committee recommends that the project timeline be extended by three months to allow for additional stakeholder consultation and review."]
        )
        try:
            refs = doc.list_paragraphs()
            ref = refs[0].split("|")[0]

            doc.rewrite_paragraph(
                ref,
                "The board has approved a three-month extension for further stakeholder review.",
            )

            vis = doc.get_visible_text()
            assert "The board has approved a three-month extension" in vis

            revisions = doc.list_revisions()
            del_texts = [r.text for r in revisions if r.type == "deletion"]
            ins_texts = [r.text for r in revisions if r.type == "insertion"]
            assert len(del_texts) > 0
            assert len(ins_texts) > 0
        finally:
            doc.close()
            shutil.rmtree(tmp, ignore_errors=True)

    def test_prose_tightening(self):
        """Removing filler words (multiple independent deletions)."""
        doc, tmp = _make_doc_with_paragraphs(
            ["We need to ensure that all stakeholders are informed and that all risks are mitigated."]
        )
        try:
            refs = doc.list_paragraphs()
            ref = refs[0].split("|")[0]

            doc.rewrite_paragraph(
                ref,
                "We need to ensure all stakeholders are informed and all risks mitigated.",
            )

            vis = doc.get_visible_text()
            assert "ensure all stakeholders" in vis
            assert "and all risks mitigated" in vis

            revisions = doc.list_revisions()
            del_texts = [r.text for r in revisions if r.type == "deletion"]
            # "that" should be deleted (twice) and "are " before "mitigated"
            assert len(del_texts) >= 2
        finally:
            doc.close()
            shutil.rmtree(tmp, ignore_errors=True)

    def test_duplicate_phrase_different_changes(self):
        """Different changes to different occurrences of the same phrase."""
        doc, tmp = _make_doc_with_paragraphs(
            ["The team will meet on Monday, the team will present on Wednesday, and the team will review on Friday."]
        )
        try:
            refs = doc.list_paragraphs()
            ref = refs[0].split("|")[0]

            doc.rewrite_paragraph(
                ref,
                "The team will meet on Monday, management will present on Wednesday, and everyone will review on Friday.",
            )

            vis = doc.get_visible_text()
            assert "management will present" in vis
            assert "everyone will review" in vis
            # First "The team will meet" unchanged
            assert "The team will meet on Monday" in vis

            revisions = doc.list_revisions()
            del_texts = [r.text for r in revisions if r.type == "deletion"]
            ins_texts = [r.text for r in revisions if r.type == "insertion"]
            assert any("the team" in t.lower() for t in del_texts)
            assert any("management" in t for t in ins_texts)
            assert any("everyone" in t for t in ins_texts)
        finally:
            doc.close()
            shutil.rmtree(tmp, ignore_errors=True)

    def test_add_clause_and_change_value(self):
        """Insert a clause in the middle while also changing a value elsewhere."""
        doc, tmp = _make_doc_with_paragraphs(
            ["The contract expires on December 31st."]
        )
        try:
            refs = doc.list_paragraphs()
            ref = refs[0].split("|")[0]

            doc.rewrite_paragraph(
                ref,
                "The contract, unless renewed by either party, expires on January 15th.",
            )

            vis = doc.get_visible_text()
            assert "unless renewed by either party" in vis
            assert "January 15th" in vis

            revisions = doc.list_revisions()
            del_texts = [r.text for r in revisions if r.type == "deletion"]
            ins_texts = [r.text for r in revisions if r.type == "insertion"]
            assert any("December" in t for t in del_texts)
            assert any("January" in t for t in ins_texts)
            assert any("unless renewed" in t for t in ins_texts)
        finally:
            doc.close()
            shutil.rmtree(tmp, ignore_errors=True)

    def test_duplicate_manager_different_replacements(self):
        """Two occurrences of 'manager' replaced with different words."""
        doc, tmp = _make_doc_with_paragraphs(
            ["The manager will review the report and the manager will approve the budget."]
        )
        try:
            refs = doc.list_paragraphs()
            ref = refs[0].split("|")[0]

            doc.rewrite_paragraph(
                ref,
                "The director will review the findings and the supervisor will approve the budget.",
            )

            vis = doc.get_visible_text()
            assert "director will review" in vis
            assert "supervisor will approve" in vis
            assert "findings" in vis

            revisions = doc.list_revisions()
            del_texts = [r.text for r in revisions if r.type == "deletion"]
            ins_texts = [r.text for r in revisions if r.type == "insertion"]
            assert any("manager" in t for t in del_texts)
            assert any("director" in t for t in ins_texts)
            assert any("supervisor" in t for t in ins_texts)
            assert any("report" in t for t in del_texts)
            assert any("findings" in t for t in ins_texts)
        finally:
            doc.close()
            shutil.rmtree(tmp, ignore_errors=True)


class TestEditResultRef:
    """Tests for edit methods returning new paragraph refs."""

    def test_replace_returns_ref(self, rewrite_doc):
        """replace() returns EditResult with .ref for follow-up edits."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        result = doc.replace("committee", "board", paragraph=ref)

        # .ref is a fresh paragraph reference
        assert result.ref.startswith("P1#")
        assert result.ref != ref  # hash changed

        # Can use .ref for a follow-up edit without list_paragraphs()
        result2 = doc.replace("board", "council", paragraph=result.ref)
        assert result2.ref.startswith("P1#")
        assert result2.ref != result.ref  # hash changed again

        vis = doc.get_visible_text()
        assert "council" in vis

    def test_delete_returns_ref(self, rewrite_doc):
        """delete() returns EditResult with .ref."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        result = doc.delete("committee", paragraph=ref)
        assert result.ref.startswith("P1#")
        assert result.ref != ref

    def test_insert_after_returns_ref(self, rewrite_doc):
        """insert_after() returns EditResult with .ref."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        result = doc.insert_after("committee", " meeting", paragraph=ref)
        assert result.ref.startswith("P1#")
        assert result.ref != ref

    def test_insert_before_returns_ref(self, rewrite_doc):
        """insert_before() returns EditResult with .ref."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        result = doc.insert_before("committee", "special ", paragraph=ref)
        assert result.ref.startswith("P1#")
        assert result.ref != ref

    def test_rewrite_returns_ref(self, rewrite_doc):
        """rewrite_paragraph() returns new ref string."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        new_ref = doc.rewrite_paragraph(ref, "Completely new text.")
        assert new_ref.startswith("P1#")
        assert new_ref != ref

        # Can chain with another rewrite
        new_ref2 = doc.rewrite_paragraph(new_ref, "Even newer text.")
        assert new_ref2.startswith("P1#")
        assert new_ref2 != new_ref

        vis = doc.get_visible_text()
        assert "Even newer text." in vis

    def test_batch_rewrite_returns_refs(self, rewrite_doc):
        """batch_rewrite() returns list of new refs in input order."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()

        new_refs = doc.batch_rewrite([
            (refs[0].split("|")[0], "First changed."),
            (refs[2].split("|")[0], "Third changed."),
        ])

        assert len(new_refs) == 2
        assert new_refs[0].startswith("P1#")
        assert new_refs[1].startswith("P3#")

    def test_batch_edit_returns_refs(self, rewrite_doc):
        """batch_edit() returns list of EditResults with .ref."""
        from docx_editor import EditOperation

        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()

        results = doc.batch_edit([
            EditOperation(action="replace", find="committee", replace_with="board", paragraph=refs[0].split("|")[0]),
            EditOperation(action="replace", find="quarterly", replace_with="monthly", paragraph=refs[1].split("|")[0]),
        ])

        assert len(results) == 2
        assert results[0].ref.startswith("P1#")
        assert results[1].ref.startswith("P2#")

    def test_edit_result_backwards_compatible(self, rewrite_doc):
        """EditResult works as int for accept/reject operations."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        result = doc.insert_after("committee", " NEW", paragraph=ref)

        # Acts as int
        assert isinstance(result, int)
        assert int(result) >= 0

        # Can be used for accept/reject
        accepted = doc.accept_revision(result)
        assert accepted is True

    def test_sequential_edits_without_list_paragraphs(self, rewrite_doc):
        """Chain multiple edits using .ref without ever calling list_paragraphs() again."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        # Chain 3 edits on the same paragraph, using .ref each time
        r1 = doc.replace("committee", "board", paragraph=ref)
        r2 = doc.replace("shall", "must", paragraph=r1.ref)
        r3 = doc.replace("annual", "quarterly", paragraph=r2.ref)

        vis = doc.get_visible_text()
        assert "board" in vis
        assert "must" in vis
        assert "quarterly" in vis


class TestRewriteWithTrackedChanges:
    def test_rewrite_after_prior_edit(self, rewrite_doc):
        """Rewriting a paragraph that already has tracked changes works."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref1 = refs[0].split("|")[0]

        # Make a surgical edit first (creates tracked insertion/deletion)
        doc.replace("committee", "board", paragraph=ref1)

        # Now rewrite the paragraph using fresh hash
        refs2 = doc.list_paragraphs()
        ref2 = refs2[0].split("|")[0]
        doc.rewrite_paragraph(ref2, "The board shall review and approve the annual budget proposal.")

        vis = doc.get_visible_text()
        assert "approve" in vis
        assert "proceed with" not in vis

    def test_save_and_reopen(self, rewrite_doc):
        """Document can be saved and reopened after rewrite."""
        doc, tmp = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        doc.rewrite_paragraph(ref, "Completely rewritten paragraph text.")
        save_path = doc.save()
        doc.close()

        # Reopen and verify
        doc2 = Document.open(save_path, force_recreate=True)
        vis = doc2.get_visible_text()
        assert "Completely rewritten paragraph text." in vis

        # Verify revisions are preserved
        revisions = doc2.list_revisions()
        assert len(revisions) > 0
        doc2.close()

    def test_rewrite_after_batch(self, rewrite_doc):
        """Individual rewrite works after batch rewrite."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()

        doc.batch_rewrite([
            (refs[0].split("|")[0], "First changed."),
            (refs[2].split("|")[0], "Third changed."),
        ])

        # Now do another rewrite on a different paragraph
        refs2 = doc.list_paragraphs()
        ref = refs2[4].split("|")[0]
        doc.rewrite_paragraph(ref, "Fifth also changed.")

        vis = doc.get_visible_text()
        assert "First changed." in vis
        assert "Third changed." in vis
        assert "Fifth also changed." in vis


class TestBatchRewrite:
    def test_batch_multiple_paragraphs(self, rewrite_doc):
        """Batch rewrite of multiple paragraphs succeeds."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()

        doc.batch_rewrite([
            (refs[0].split("|")[0], "The board shall review and approve."),
            (refs[2].split("|")[0], "The report is complete."),
            (refs[4].split("|")[0], "Final decision pending."),
        ])

        vis = doc.get_visible_text()
        assert "The board shall review and approve." in vis
        assert "The report is complete." in vis
        assert "Final decision pending." in vis

    def test_batch_rejected_on_stale_hash(self, rewrite_doc):
        """Stale hash in batch rejects entire batch (no changes applied)."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()

        with pytest.raises(HashMismatchError):
            doc.batch_rewrite([
                (refs[0].split("|")[0], "Good text"),
                ("P3#0000", "Bad hash"),
            ])

        # Verify no changes were applied (P1 unchanged)
        vis = doc.get_visible_text()
        assert "The committee shall review" in vis

    def test_batch_rejected_on_duplicate(self, rewrite_doc):
        """Duplicate paragraph in batch raises ValueError."""
        doc, _ = rewrite_doc
        refs = doc.list_paragraphs()
        ref = refs[0].split("|")[0]

        with pytest.raises(ValueError, match="duplicate"):
            doc.batch_rewrite([
                (ref, "First"),
                (ref, "Second"),
            ])

    def test_batch_empty(self, rewrite_doc):
        """Empty batch succeeds with no changes."""
        doc, _ = rewrite_doc
        doc.batch_rewrite([])
        assert len(doc.list_revisions()) == 0
