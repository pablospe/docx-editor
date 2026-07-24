"""Tests for tracked paragraph splits on ``\\n`` (ISSUES.md #58+61).

A ``\\n`` in edit text means a *tracked paragraph split* at that point: the
first paragraph's paragraph mark is flagged as an inserted revision, the tail
runs move to a new paragraph as unchanged content. Accepting keeps the split;
rejecting removes the mark and rejoins. One ``\\n``-containing operation is ONE
revision group covering the deletion, the replacement runs in both paragraphs,
and the inserted mark; ``reject_group`` reverts the whole split atomically.
"""

import re
import zipfile

import pytest
from conftest import find_ref, replace_document_xml

from docx_editor import Document, EditOperation
from docx_editor.xml_editor import ParagraphRef


@pytest.fixture
def doc(temp_docx):
    """An open Document over the simple.docx copy, closed after the test."""
    document = Document.open(temp_docx)
    yield document
    document.close()


def _lines(doc: Document) -> list[str]:
    return doc.get_visible_text().splitlines()


class TestReplaceSplitsParagraph:
    def test_replace_with_newline_splits_into_two_paragraphs(self, doc):
        ref = find_ref(doc, "quick brown fox")
        before = doc.paragraph_count()

        result = doc.replace("lazy dog.", "lazy dog.\nA new paragraph.", paragraph=ref)

        assert doc.paragraph_count() == before + 1
        lines = _lines(doc)
        # The fox paragraph keeps its (deleted-then-reinserted) text; the new
        # segment lands on its own line immediately after.
        fox_idx = next(i for i, line in enumerate(lines) if "lazy dog." in line)
        assert lines[fox_idx + 1] == "A new paragraph."
        # No literal newline leaked into a run.
        assert "\n" not in "".join(lines)
        assert result.group_id is not None

    def test_multi_newline_makes_multiple_splits_one_group(self, doc):
        ref = find_ref(doc, "quick brown fox")
        before = doc.paragraph_count()

        result = doc.replace("fox", "fox\nSECOND\nTHIRD", paragraph=ref)

        assert doc.paragraph_count() == before + 2
        lines = _lines(doc)
        i = next(idx for idx, line in enumerate(lines) if line.endswith("fox"))
        assert lines[i + 1] == "SECOND"
        assert lines[i + 2].startswith("THIRD")
        # One operation = one group across every resulting paragraph.
        groups = {rev.group_id for rev in doc.list_revisions()}
        assert groups == {result.group_id}
        # Two inserted paragraph marks (empty-text insertions), one per split.
        marks = [rev for rev in doc.list_revisions() if rev.type == "insertion" and rev.text == ""]
        assert len(marks) == 2

    def test_accept_group_keeps_the_split(self, doc):
        ref = find_ref(doc, "quick brown fox")
        before = doc.paragraph_count()
        result = doc.replace("dog.", "dog.\nKept line.", paragraph=ref)

        doc.accept_group(result.group_id)

        assert doc.paragraph_count() == before + 1
        assert "Kept line." in _lines(doc)
        # Nothing tracked remains after accepting.
        assert doc.list_revisions() == []


class TestRejectRejoins:
    def test_reject_group_rejoins_exactly(self, doc):
        ref = find_ref(doc, "quick brown fox")
        original = doc.get_visible_text()
        before = doc.paragraph_count()

        result = doc.replace("dog.", "dog.\nExtra sentence.", paragraph=ref)
        assert doc.paragraph_count() == before + 1

        doc.reject_group(result.group_id)

        assert doc.paragraph_count() == before
        assert doc.get_visible_text() == original
        assert doc.list_revisions() == []

    def test_reject_group_rejoins_multi_split(self, doc):
        ref = find_ref(doc, "quick brown fox")
        original = doc.get_visible_text()
        before = doc.paragraph_count()

        result = doc.replace("fox", "fox\nB\nC\nD", paragraph=ref)
        assert doc.paragraph_count() == before + 3

        doc.reject_group(result.group_id)

        assert doc.paragraph_count() == before
        assert doc.get_visible_text() == original
        assert doc.list_revisions() == []

    def test_reject_all_rejoins(self, doc):
        ref = find_ref(doc, "quick brown fox")
        original = doc.get_visible_text()
        before = doc.paragraph_count()

        doc.replace("dog.", "dog.\nMore.", paragraph=ref)
        doc.reject_all()

        assert doc.paragraph_count() == before
        assert doc.get_visible_text() == original

    def test_reject_single_mark_revision_rejoins(self, doc):
        ref = find_ref(doc, "quick brown fox")
        before = doc.paragraph_count()
        doc.replace("dog.", "dog.\nTail.", paragraph=ref)
        # The paragraph-mark insertion is the empty-text insertion.
        mark = next(r for r in doc.list_revisions() if r.type == "insertion" and r.text == "")

        doc.reject_revision(mark.id)

        assert doc.paragraph_count() == before


class TestInsertAndRewriteSplit:
    def test_insert_after_with_newline_splits(self, doc):
        ref = find_ref(doc, "quick brown fox")
        before = doc.paragraph_count()

        doc.insert_after("fox", "\nInserted paragraph.", paragraph=ref)

        assert doc.paragraph_count() == before + 1
        lines = _lines(doc)
        i = next(idx for idx, line in enumerate(lines) if line.endswith("fox"))
        # The break falls right after "fox"; the inserted text opens the new
        # paragraph, and the original tail follows it.
        assert lines[i + 1].startswith("Inserted paragraph.")
        assert lines[i + 1].endswith("jumps over the lazy dog.")

    def test_insert_before_with_newline_splits(self, doc):
        ref = find_ref(doc, "quick brown fox")
        before = doc.paragraph_count()

        doc.insert_before("jumps", "END.\n", paragraph=ref)

        assert doc.paragraph_count() == before + 1
        lines = _lines(doc)
        i = next(idx for idx, line in enumerate(lines) if line.endswith("END."))
        assert lines[i + 1].startswith("jumps")

    def test_rewrite_paragraph_with_newline_splits(self, doc):
        ref = find_ref(doc, "quick brown fox")
        before = doc.paragraph_count()

        doc.rewrite_paragraph(ref, "First half.\nSecond half.")

        assert doc.paragraph_count() == before + 1
        lines = _lines(doc)
        assert "First half." in lines
        assert "Second half." in lines


class TestSplitParagraphSugar:
    def test_split_paragraph_before_anchor(self, doc):
        ref = find_ref(doc, "quick brown fox")
        before = doc.paragraph_count()

        doc.split_paragraph(ref, before="jumps")

        assert doc.paragraph_count() == before + 1
        lines = _lines(doc)
        i = next(idx for idx, line in enumerate(lines) if line.startswith("jumps"))
        assert lines[i - 1].endswith("fox ")
        assert lines[i] == "jumps over the lazy dog."

    def test_split_paragraph_reject_rejoins(self, doc):
        ref = find_ref(doc, "quick brown fox")
        original = doc.get_visible_text()
        before = doc.paragraph_count()

        result = doc.split_paragraph(ref, before="jumps")
        doc.reject_group(result.group_id)

        assert doc.paragraph_count() == before
        assert doc.get_visible_text() == original

    def test_split_paragraph_requires_before_keyword(self, doc):
        ref = find_ref(doc, "quick brown fox")
        with pytest.raises(TypeError):
            doc.split_paragraph(ref, "jumps")  # `before` is keyword-only


class TestEditResultRefs:
    def test_non_split_edit_has_single_ref(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.replace("quick", "speedy", paragraph=ref)
        assert result.refs == (str(result),)

    def test_split_refs_cover_both_paragraphs(self, doc):
        from docx_editor.xml_editor import ParagraphRef

        ref = find_ref(doc, "quick brown fox")
        result = doc.replace("dog.", "dog.\nNew line.", paragraph=ref)

        assert len(result.refs) == 2
        assert result.refs[0] == str(result)
        r1, r2 = result.refs
        assert ParagraphRef.parse(r2).index == ParagraphRef.parse(r1).index + 1
        # Both refs are usable for follow-up edits.
        assert "New line." in doc.get_paragraph(ParagraphRef.parse(r2).index).text

    def test_multi_split_has_three_refs(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.replace("fox", "fox\nB\nC", paragraph=ref)
        assert len(result.refs) == 3

    def test_split_paragraph_refs(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.split_paragraph(ref, before="jumps")
        assert len(result.refs) == 2


class TestBatchSplit:
    def test_batch_with_split_op_is_one_changeset(self, doc):
        fox = find_ref(doc, "quick brown fox")
        sample = find_ref(doc, "sample document")
        sample_index = ParagraphRef.parse(sample).index

        results = doc.batch_edit([
            EditOperation.replace("fox", "fox\nSPLIT", paragraph=fox),
            EditOperation.replace("sample", "example", paragraph=sample),
        ])

        # The whole call is one changeset; each op keeps its own group.
        changesets = {rev.changeset_id for rev in doc.list_revisions()}
        assert len(changesets) == 1
        assert results[0].group_id != results[1].group_id
        # The split op reports both resulting paragraphs.
        assert len(results[0].refs) == 2
        # The later op's paragraph shifted down by the split; its result ref
        # tracks the LIVE position, not the stale original index.
        assert ParagraphRef.parse(results[1]).index == sample_index + 1
        assert "example" in doc.get_paragraph(ParagraphRef.parse(results[1]).index).text

    def test_reject_changeset_reverts_batch_split(self, doc):
        original = doc.get_visible_text()
        before = doc.paragraph_count()
        fox = find_ref(doc, "quick brown fox")
        sample = find_ref(doc, "sample document")

        results = doc.batch_edit([
            EditOperation.replace("fox", "fox\nSPLIT", paragraph=fox),
            EditOperation.replace("sample", "example", paragraph=sample),
        ])
        doc.reject_changeset(results[0].changeset_id)

        assert doc.paragraph_count() == before
        assert doc.get_visible_text() == original


class TestReopenRejoin:
    def test_reopened_split_is_one_group_and_rejoins(self, temp_docx):
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            original = doc.get_visible_text()
            doc.replace("dog.", "dog.\nAppended.", paragraph=ref)
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            gids = {r.group_id for r in revisions}
            assert len(gids) == 1  # reconstructed as ONE inferred group across two paragraphs
            assert all(r.group_source == "inferred" for r in revisions)
            gid = gids.pop()
            assert gid is not None
            doc.reject_group(gid)
            assert doc.get_visible_text() == original

    def test_word_stripped_split_still_reconstructs_and_rejoins(self, temp_docx, tmp_path):
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            original = doc.get_visible_text()
            doc.replace("dog.", "dog.\nAppended.", paragraph=ref)
            doc.save()

        # Simulate Word normalizing on its own save: drop our extension
        # attributes (rsids, dateUtc, paraId/textId). The rejoin rule must key
        # only on the durable author+date+mark signal.
        with zipfile.ZipFile(temp_docx) as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")
        stripped = re.sub(r'\s+(w:rsid\w*|w16du:dateUtc|w14:paraId|w14:textId)="[^"]*"', "", doc_xml)
        assert "w16du:dateUtc" not in stripped
        dest = tmp_path / "stripped.docx"
        replace_document_xml(temp_docx, dest, stripped)

        with Document.open(dest) as doc:
            revisions = doc.list_revisions()
            assert len({r.group_id for r in revisions}) == 1
            gid = revisions[0].group_id
            assert gid is not None
            doc.reject_group(gid)
            assert doc.get_visible_text() == original
