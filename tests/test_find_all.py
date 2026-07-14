"""Tests for Document.find_all() — one-call enumeration of every match."""

import shutil
import tempfile
from pathlib import Path

import pytest

from docx_editor import (
    Document,
    HashMismatchError,
    ParagraphIndexError,
    SearchResult,
)


@pytest.fixture
def multi_para_doc():
    """Build a document with 10 paragraphs, each containing repeated phrases."""
    test_data = Path(__file__).parent / "test_data" / "simple.docx"
    tmp = tempfile.mkdtemp(prefix="find_all_test_")
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

    for i in range(1, 11):
        text = f"[P{i:02d}] The committee shall review item {i}. The report includes findings from the committee."
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


class TestFindAllDocumentWide:
    def test_enumerates_every_match(self, multi_para_doc):
        """Result count matches count_matches; results are SearchResults."""
        doc, _ = multi_para_doc
        results = doc.find_all("committee")
        assert len(results) == doc.count_matches("committee")
        assert len(results) == 20  # 2 per paragraph x 10
        assert all(isinstance(r, SearchResult) for r in results)
        assert all(r.text == "committee" for r in results)

    def test_refs_are_hash_anchored_and_current(self, multi_para_doc):
        """Each result's paragraph_ref equals the list_paragraphs() ref."""
        doc, _ = multi_para_doc
        refs = [r.split("|")[0] for r in doc.list_paragraphs()]
        results = doc.find_all("committee")
        assert {r.paragraph_ref for r in results} == set(refs)

    def test_paragraph_occurrence_resets_per_paragraph(self, multi_para_doc):
        """paragraph_occurrence counts within each paragraph, not document-wide."""
        doc, _ = multi_para_doc
        results = doc.find_all("committee")
        by_ref: dict[str, list[int]] = {}
        for r in results:
            by_ref.setdefault(r.paragraph_ref, []).append(r.paragraph_occurrence)
        assert all(occs == [0, 1] for occs in by_ref.values())

    def test_offsets_are_paragraph_local(self, multi_para_doc):
        """start/end are offsets in the containing paragraph's visible text."""
        doc, _ = multi_para_doc
        results = doc.find_all("[P05]")
        assert len(results) == 1
        assert results[0].start == 0
        assert results[0].end == len("[P05]")

    def test_results_in_document_order(self, multi_para_doc):
        doc, _ = multi_para_doc
        results = doc.find_all("item")
        markers = [int(r.paragraph_ref.split("#")[0][1:]) for r in results]
        assert markers == sorted(markers)

    def test_result_chains_into_edit(self, multi_para_doc):
        """paragraph_ref + paragraph_occurrence target exactly that match."""
        doc, _ = multi_para_doc
        target = [r for r in doc.find_all("committee") if r.paragraph_ref.startswith("P3#")][1]
        assert target.paragraph_occurrence == 1

        doc.replace(
            target.text,
            "CHANGED",
            paragraph=target.paragraph_ref,
            occurrence=target.paragraph_occurrence,
        )
        doc.accept_all()
        vis = doc.get_visible_text()
        # Exactly the targeted (second) match in P3 changed; the first is untouched.
        assert "[P03] The committee shall review item 3. The report includes findings from the CHANGED." in vis
        assert vis.count("CHANGED") == 1

    def test_no_matches_returns_empty_list(self, multi_para_doc):
        """No-match is not an error for an enumeration API."""
        doc, _ = multi_para_doc
        assert doc.find_all("__absent_everywhere__") == []

    def test_empty_text_raises_value_error(self, multi_para_doc):
        doc, _ = multi_para_doc
        with pytest.raises(ValueError, match="non-empty"):
            doc.find_all("")

    def test_match_spanning_revision_sets_flag(self, multi_para_doc):
        """A match crossing a w:ins boundary reports spans_revision=True."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        doc.insert_after("review", "XX", paragraph=ref)

        results = doc.find_all("reviewXX")
        assert len(results) == 1
        assert results[0].spans_revision is True

        plain = doc.find_all("[P01]")
        assert plain[0].spans_revision is False


class TestFindAllScoped:
    def test_scoped_returns_only_that_paragraphs_hits(self, multi_para_doc):
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[2].split("|")[0]
        results = doc.find_all("committee", paragraph=ref)
        assert len(results) == 2
        assert all(r.paragraph_ref == ref for r in results)
        assert [r.paragraph_occurrence for r in results] == [0, 1]

    def test_scoped_no_matches_returns_empty_list(self, multi_para_doc):
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        assert doc.find_all("item 5", paragraph=ref) == []

    def test_stale_hash_raises(self, multi_para_doc):
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        doc.replace("item 1", "MUTATED", paragraph=ref)
        with pytest.raises(HashMismatchError):
            doc.find_all("committee", paragraph=ref)

    def test_malformed_ref_raises_value_error(self, multi_para_doc):
        doc, _ = multi_para_doc
        with pytest.raises(ValueError, match="Invalid paragraph reference"):
            doc.find_all("committee", paragraph="not-a-ref")

    def test_out_of_range_index_raises(self, multi_para_doc):
        doc, _ = multi_para_doc
        with pytest.raises(ParagraphIndexError):
            doc.find_all("committee", paragraph="P999#0000")
