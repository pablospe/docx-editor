"""Baseline + improvement comparison for LLM-facing error quality.

Captures the information content of docx-editor errors so that a reader can
see, side-by-side, what an LLM agent learns from each error class.

For every LLM-facing failure mode this file documents:
- **Before**: what an agent got from the bare string / stdlib exception.
- **After**: the structured attributes added by this change and the new
  message contents.

Each test asserts ONLY the "after" state (the current behaviour). The
"before" description lives in the docstring so that reviewers comparing
commits can see what was lifted into the type system without re-running
the old code.

This file is the artefact that proves the change's claim: diagnostics
move from string-parsing into structured fields.
"""

import shutil
import tempfile
from pathlib import Path

import pytest

from docx_editor import (
    AmbiguousTextError,
    BatchOperationError,
    Document,
    EditOperation,
    ParagraphIndexError,
    TextNotFoundError,
)


@pytest.fixture
def doc_path():
    test_data = Path(__file__).parent / "test_data" / "simple.docx"
    tmp = tempfile.mkdtemp(prefix="error_quality_")
    dest = Path(tmp) / "test.docx"
    shutil.copy(test_data, dest)
    yield dest
    shutil.rmtree(tmp, ignore_errors=True)


def _build_doc_with_paragraphs(doc_path: Path, texts: list[str]) -> Document:
    """Replace the fixture document's body with one paragraph per entry."""
    doc = Document.open(doc_path, force_recreate=True)
    editor = doc._document_editor
    body = editor.dom.getElementsByTagName("w:body")[0]
    for p in list(editor.dom.getElementsByTagName("w:p")):
        if p.parentNode == body:
            body.removeChild(p)
    sect_pr = editor.dom.getElementsByTagName("w:sectPr")
    insert_before = sect_pr[0] if sect_pr else None
    for text in texts:
        xml = f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        for node in editor._parse_fragment(xml):
            if insert_before:
                body.insertBefore(node, insert_before)
            else:
                body.appendChild(node)
    editor.save()
    saved = doc.save()
    doc.close()
    return Document.open(saved, force_recreate=True)


class TestTextNotFoundErrorQuality:
    """
    Before: `raise TextNotFoundError(f"Text not found in paragraph P{idx}: '{find}'")`
        An agent saw only a string. To retry it had to:
          1. Parse the paragraph index out of the message.
          2. Re-read the paragraph from the document to see current content.
          3. Guess at occurrence counts for nth-match failures.

    After: structured fields (`search_text`, `paragraph_ref`,
    `paragraph_preview`, `occurrence`, `total_occurrences`) let the
    agent self-correct without a second round-trip.
    """

    def test_scoped_search_exposes_paragraph_state(self, doc_path):
        """Scoped miss carries ref + preview — agent can diff without re-reading."""
        doc = Document.open(doc_path)
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            with pytest.raises(TextNotFoundError) as exc:
                doc.replace("__definitely_not_here__", "x", paragraph=ref)

            err = exc.value
            assert err.search_text == "__definitely_not_here__"
            assert err.paragraph_ref == ref
            assert err.paragraph_preview is not None
            assert err.paragraph_preview  # non-empty preview of current text
            assert len(err.paragraph_preview) <= 83  # 80 + "..." cap
            msg = str(err)
            assert "__definitely_not_here__" in msg
            assert ref in msg
            assert err.paragraph_preview.removesuffix("...") in msg
        finally:
            doc.close()

    def test_unscoped_search_leaves_scope_fields_none(self, doc_path):
        """Unscoped miss has no paragraph context — honest None, not a lie."""
        doc = Document.open(doc_path)
        try:
            with pytest.raises(TextNotFoundError) as exc:
                doc.add_comment("__absent_from_entire_doc__", "c")

            err = exc.value
            assert err.search_text == "__absent_from_entire_doc__"
            assert err.paragraph_ref is None
            assert err.paragraph_preview is None
            assert "__absent_from_entire_doc__" in str(err)
        finally:
            doc.close()

    def test_scoped_miss_without_preview_omits_content_line(self):
        """paragraph_ref set, no preview → message names the paragraph, skips content."""
        err = TextNotFoundError("needle", paragraph_ref="P3#a7b2")
        msg = str(err)
        assert "needle" in msg
        assert "P3#a7b2" in msg
        assert "Current content" not in msg
        assert err.paragraph_ref == "P3#a7b2"
        assert err.paragraph_preview is None

    def test_occurrence_miss_with_paragraph_ref_names_scope(self):
        """nth-match miss with paragraph_ref appends the ref so the agent knows where."""
        err = TextNotFoundError(
            "needle",
            paragraph_ref="P2#cafe",
            occurrence=3,
            total_occurrences=1,
        )
        msg = str(err)
        assert "Only 1 occurrence(s) of 'needle'" in msg
        assert "occurrence=3 requested" in msg
        assert "Paragraph: P2#cafe" in msg


class TestParagraphIndexErrorQuality:
    """
    Before: stdlib `IndexError` with a free-form message. `except IndexError`
    also caught list indexing bugs inside the library; structured recovery
    required the agent to regex the message.

    After: `ParagraphIndexError(DocxEditError)` with `.index` and
    `.total_paragraphs` instance attributes. The agent can compute a
    valid index programmatically and retry.
    """

    def test_out_of_range_surfaces_structured_index(self, doc_path):
        doc = Document.open(doc_path)
        try:
            n = len(doc.list_paragraphs())
            bad_ref = f"P{n + 99}#0000"
            with pytest.raises(ParagraphIndexError) as exc:
                doc.replace("anything", "else", paragraph=bad_ref)

            err = exc.value
            assert err.index == n + 99
            assert err.total_paragraphs == n
            msg = str(err)
            assert str(n + 99) in msg
            assert str(n) in msg
        finally:
            doc.close()

    def test_empty_document_uses_distinct_phrasing(self):
        """total_paragraphs=0 gets its own message — no misleading 'valid: P1-P0'."""
        err = ParagraphIndexError(index=1, total_paragraphs=0)
        msg = str(err)
        assert "no paragraphs" in msg
        assert "P1-P0" not in msg
        assert err.index == 1
        assert err.total_paragraphs == 0


class TestBatchOperationErrorQuality:
    """
    Before: `raise ValueError(f"Operation {i}: ...")` — the index was only
    in the message string. Worse, `_apply_single_edit` raised bare
    `ValueError("replace requires 'find' and 'replace_with'")` with NO
    index at all. On a 10-op batch, the agent could not tell which op
    failed without re-running operations one by one.

    After: `BatchOperationError(DocxEditError)` with `.operation_index`
    and `.reason`. Every validation path — pre-dispatch and inside
    `_apply_single_edit` — is wrapped so the index is always present.
    """

    def _build_batch_doc(self, doc_path: Path) -> Document:
        return _build_doc_with_paragraphs(doc_path, [f"Paragraph {i} content." for i in range(1, 4)])

    def test_predispatch_failure_carries_operation_index(self, doc_path):
        """Structural validation inside batch_edit surfaces the index."""
        doc = self._build_batch_doc(doc_path)
        try:
            refs = doc.list_paragraphs()
            ops = [
                EditOperation(action="replace", find="Paragraph 1", replace_with="x", paragraph=refs[0].split("|")[0]),
                EditOperation(action="replace", find="a", replace_with="b", paragraph=""),
            ]
            with pytest.raises(BatchOperationError) as exc:
                doc.batch_edit(ops)

            err = exc.value
            assert err.operation_index == 1
            assert "paragraph" in err.reason.lower()
            assert "1" in str(err)
        finally:
            doc.close()

    def test_inner_missing_field_failure_carries_operation_index(self, doc_path):
        """A `ValueError` inside _apply_single_edit is wrapped with the index."""
        doc = self._build_batch_doc(doc_path)
        try:
            refs = doc.list_paragraphs()
            ops = [
                EditOperation(
                    action="replace",
                    find=None,  # missing required field, fires inside _apply_single_edit
                    replace_with="x",
                    paragraph=refs[0].split("|")[0],
                ),
            ]
            with pytest.raises(BatchOperationError) as exc:
                doc.batch_edit(ops)

            err = exc.value
            assert err.operation_index == 0
            assert "replace requires" in err.reason
        finally:
            doc.close()


class TestAmbiguousTextErrorQuality:
    """
    Before: every edit method defaulted to occurrence=0, so an ambiguous
    target was silently resolved to the first match — the agent got no
    signal that it may have edited the wrong text.

    After: an omitted occurrence requires the target to be unique in the
    search scope. `AmbiguousTextError` carries `total_occurrences` and
    names both recovery paths (explicit `occurrence=` or `find_all()`);
    explicit `occurrence=0` keeps the old first-match behavior per call.
    """

    AMBIGUOUS = "alpha beta alpha gamma alpha."
    UNIQUE = "one needle only."

    def _build_doc(self, doc_path: Path) -> Document:
        return _build_doc_with_paragraphs(doc_path, [self.AMBIGUOUS, self.UNIQUE])

    def _edit_calls(self, doc: Document, paragraph: str):
        """Every scoped edit method, called with an ambiguous target and no occurrence."""
        return [
            lambda: doc.replace("alpha", "x", paragraph=paragraph),
            lambda: doc.delete("alpha", paragraph=paragraph),
            lambda: doc.insert_after("alpha", "x", paragraph=paragraph),
            lambda: doc.insert_before("alpha", "x", paragraph=paragraph),
            lambda: doc.add_comment("alpha", "note", paragraph=paragraph),
        ]

    def test_scoped_ambiguous_target_raises_with_totals(self, doc_path):
        """All five scoped methods raise AmbiguousTextError with the count."""
        doc = self._build_doc(doc_path)
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            for call in self._edit_calls(doc, ref):
                with pytest.raises(AmbiguousTextError) as exc:
                    call()
                err = exc.value
                assert err.search_text == "alpha"
                assert err.total_occurrences == 3
                assert err.paragraph_ref == ref
                assert err.paragraph_preview is not None
                msg = str(err)
                assert "matches 3 times" in msg
                assert "occurrence=" in msg
                assert "find_all()" in msg
        finally:
            doc.close()

    def test_docwide_ambiguous_anchor_raises(self, doc_path):
        """Document-wide add_comment with a repeated anchor raises too."""
        doc = self._build_doc(doc_path)
        try:
            with pytest.raises(AmbiguousTextError) as exc:
                doc.add_comment("alpha", "note")
            err = exc.value
            assert err.total_occurrences == 3
            assert err.paragraph_ref is None
            assert err.paragraph_preview is None
            assert "in the document" in str(err)
        finally:
            doc.close()

    def test_explicit_occurrence_zero_edits_first_silently(self, doc_path):
        """occurrence=0 is the per-call opt-out: old first-match behavior."""
        doc = self._build_doc(doc_path)
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            doc.replace("alpha", "FIRST", paragraph=ref, occurrence=0)
            doc.accept_all()
            assert "FIRST beta alpha gamma alpha." in doc.get_visible_text()
        finally:
            doc.close()

    def test_unique_target_with_omitted_occurrence_works(self, doc_path):
        """A unique target never trips the ambiguity check."""
        doc = self._build_doc(doc_path)
        try:
            ref = doc.list_paragraphs()[1].split("|")[0]
            doc.replace("needle", "thread", paragraph=ref)
            doc.accept_all()
            assert "one thread only." in doc.get_visible_text()
        finally:
            doc.close()

    def test_ambiguity_surfaces_in_dry_run_error(self, doc_path):
        """batch_edit(dry_run=True) reports the same ambiguity message."""
        doc = self._build_doc(doc_path)
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            ops = [EditOperation.replace("alpha", "x", paragraph=ref)]
            results = doc.batch_edit(ops, dry_run=True)
            assert results[0].valid is False
            error = results[0].error
            assert error is not None
            assert "matches 3 times" in error
        finally:
            doc.close()


class TestOccurrenceOutOfRangeQuality:
    """
    Before: a scoped edit with an out-of-range occurrence raised
    `TextNotFoundError` built with only ref + preview — the message said
    "Text not found" while the preview visibly contained the text.

    After: the scoped paths count total matches, so the error reports
    "Only N occurrence(s) ... occurrence=X requested" with both fields
    populated — never a self-contradicting "not found".
    """

    TEXT = "term one, term two."

    def test_scoped_out_of_range_names_totals_not_notfound(self, doc_path):
        doc = _build_doc_with_paragraphs(doc_path, [self.TEXT])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            with pytest.raises(TextNotFoundError) as exc:
                doc.replace("term", "x", paragraph=ref, occurrence=9)
            err = exc.value
            assert err.occurrence == 9
            assert err.total_occurrences == 2
            msg = str(err)
            assert "Only 2 occurrence(s)" in msg
            assert "occurrence=9" in msg
            assert "Text not found" not in msg
        finally:
            doc.close()

    def test_dry_run_error_gets_the_same_fix(self, doc_path):
        """EditValidationResult.error also reports totals, not 'not found'."""
        doc = _build_doc_with_paragraphs(doc_path, [self.TEXT])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            ops = [EditOperation.replace("term", "x", paragraph=ref, occurrence=9)]
            results = doc.batch_edit(ops, dry_run=True)
            assert results[0].valid is False
            error = results[0].error
            assert error is not None
            assert "Only 2 occurrence(s)" in error
            assert "occurrence=9" in error
            assert "Text not found" not in error
        finally:
            doc.close()


class TestSharedBaseClass:
    """All structured errors inherit from `DocxEditError`.

    This means `except DocxEditError` is a correct catch-all for LLM
    consumers — no need to enumerate the leaf classes.
    """

    def test_all_structured_errors_share_docx_edit_error_base(self):
        from docx_editor import DocxEditError, HashMismatchError

        assert issubclass(TextNotFoundError, DocxEditError)
        assert issubclass(AmbiguousTextError, DocxEditError)
        assert issubclass(HashMismatchError, DocxEditError)
        assert issubclass(ParagraphIndexError, DocxEditError)
        assert issubclass(BatchOperationError, DocxEditError)

    def test_ambiguous_is_not_a_textnotfound_subclass(self):
        """The text WAS found — code catching TextNotFoundError to mean
        'pick different text' must not swallow ambiguity."""
        assert not issubclass(AmbiguousTextError, TextNotFoundError)

    def test_top_level_import_smoke(self):
        """Every structured error is importable from the top-level package."""
        from docx_editor import (  # noqa: F401
            AmbiguousTextError,
            BatchOperationError,
            HashMismatchError,
            ParagraphIndexError,
            TextNotFoundError,
        )
