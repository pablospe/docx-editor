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
from collections.abc import Callable
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

    def test_non_editoperation_element_rejected_with_index(self, doc_path):
        """A dict element is a batch-level rule violation, not a raw AttributeError."""
        doc = self._build_batch_doc(doc_path)
        try:
            before = doc.get_visible_text()
            ops = [{"action": "replace", "find": "a", "replace_with": "b"}]
            with pytest.raises(BatchOperationError) as exc:
                doc.batch_edit(ops)  # type: ignore[arg-type]

            err = exc.value
            assert err.operation_index == 0
            assert err.original is None
            assert "expected EditOperation, got dict" in err.reason
            assert "EditOperation.replace()" in err.reason  # names the recovery path
            assert doc.get_visible_text() == before
        finally:
            doc.close()

    def test_non_editoperation_mixed_batch_names_offending_index(self, doc_path):
        """Atomicity: the valid op at index 0 is not applied when index 1 is rejected."""
        doc = self._build_batch_doc(doc_path)
        try:
            before = doc.get_visible_text()
            refs = doc.list_paragraphs()
            valid = EditOperation.replace("Paragraph 1", "x", paragraph=refs[0].split("|")[0])
            with pytest.raises(BatchOperationError) as exc:
                doc.batch_edit([valid, "not an op"])  # type: ignore[list-item]

            assert exc.value.operation_index == 1
            assert "expected EditOperation, got str" in exc.value.reason
            assert doc.get_visible_text() == before
        finally:
            doc.close()

    def test_dry_run_reports_non_editoperation_without_raising(self, doc_path):
        """validate_batch never raises — a bad element type comes back as an invalid row."""
        doc = self._build_batch_doc(doc_path)
        try:
            refs = doc.list_paragraphs()
            valid = EditOperation.replace("Paragraph 1", "x", paragraph=refs[0].split("|")[0])
            results = doc.batch_edit([valid, {"action": "delete"}], dry_run=True)  # type: ignore[list-item]

            assert len(results) == 2
            assert results[0].valid
            bad = results[1]
            assert bad.index == 1
            assert not bad.valid
            assert bad.paragraph is None
            assert bad.error is not None
            assert "expected EditOperation, got dict" in bad.error
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

    def _edit_calls(self, doc: Document, paragraph: str) -> list[Callable[[], object]]:
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


class TestNegativeOccurrenceRejected:
    """
    Before: a negative occurrence slipped past the direct (non-batch) edit
    methods into ``find_in_text_map``, where ``range(occurrence + 1)`` never
    runs and the scoped paths crashed with ``UnboundLocalError`` (the
    document-wide paths produced a nonsensical "occurrence=-1 requested").

    After: every locate path rejects a negative occurrence up front with the
    same ValueError the EditOperation constructors raise.
    """

    def test_all_paths_raise_value_error(self, doc_path):
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            calls = [
                lambda: doc.replace("needle", "x", paragraph=ref, occurrence=-1),
                lambda: doc.add_comment("needle", "note", paragraph=ref, occurrence=-1),
                lambda: doc.add_comment("needle", "note", occurrence=-1),
                lambda: doc._revision_manager.replace_text("needle", "x", occurrence=-1),
            ]
            for call in calls:
                with pytest.raises(ValueError, match="occurrence must be >= 0"):
                    call()
        finally:
            doc.close()


class TestEmptySearchTextRejected:
    """
    Before: an empty search string slipped into ``find_in_text_map``, where
    ``str.find("", start)`` matches at every position. With an explicit
    ``occurrence`` the zero-width match carried no DOM positions, so the
    apply helpers created zero revisions and the edit silently vanished;
    without one, the caller got a meaningless AmbiguousTextError
    ("'' matches N times").

    After: both shared locate paths reject empty (and None) search text up
    front with the same ValueError, before any change is made.
    (``add_comment`` never had this bug — it pre-rejects empty anchors with
    its own CommentError; pinned below so the asymmetry stays deliberate.)
    """

    def test_empty_find_with_occurrence_no_longer_a_silent_no_op(self, doc_path):
        """The original regression: explicit occurrence=0 made the edit vanish."""
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            with pytest.raises(ValueError, match="search text must be a non-empty string"):
                doc.replace("", "TEXT", paragraph=ref, occurrence=0)
            assert doc.list_revisions() == []  # nothing applied, nothing vanished
        finally:
            doc.close()

    def test_all_paths_raise_value_error(self, doc_path):
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            rm = doc._revision_manager
            calls = [
                # scoped, no occurrence (was: meaningless AmbiguousTextError)
                lambda: doc.replace("", "x", paragraph=ref),
                lambda: doc.delete("", paragraph=ref),
                lambda: doc.insert_after("", "x", paragraph=ref),
                lambda: doc.insert_before("", "x", paragraph=ref),
                # scoped, explicit occurrence (was: silent no-op)
                lambda: doc.delete("", paragraph=ref, occurrence=0),
                # document-wide RM paths, both failure modes
                lambda: rm.replace_text("", "x"),
                lambda: rm.replace_text("", "x", occurrence=0),
            ]
            for call in calls:
                with pytest.raises(ValueError, match="search text must be a non-empty string"):
                    call()
            assert doc.list_revisions() == []
        finally:
            doc.close()

    def test_none_search_text_names_the_value(self, doc_path):
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            with pytest.raises(ValueError, match="search text must be a non-empty string, got None"):
                doc.replace(None, "x", paragraph=ref)  # type: ignore[arg-type]
        finally:
            doc.close()

    def test_add_comment_still_rejects_bad_anchors_its_own_way(self, doc_path):
        """add_comment pre-rejects empty AND non-string anchors with CommentError,
        not ValueError (its documented type) — and never a raw TypeError."""
        from docx_editor import CommentError

        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            with pytest.raises(CommentError, match="anchor_text must be a non-empty string, got ''"):
                doc.add_comment("", "note", paragraph=ref)
            with pytest.raises(CommentError, match="anchor_text must be a non-empty string, got 123"):
                doc.add_comment(123, "note", paragraph=ref)  # type: ignore[arg-type]
            with pytest.raises(CommentError, match="anchor_text must be a non-empty string, got b'needle'"):
                doc.add_comment(b"needle", "note")  # type: ignore[arg-type]
        finally:
            doc.close()


class TestNonStringParagraphRefRejected:
    """
    Before: ``Document.replace(..., paragraph=None)`` didn't error at all —
    None slipped through the keyword-only ``paragraph: str`` signature into
    the RevisionManager, silently selecting its document-wide search branch.

    After: the four Document edit methods reject non-string paragraph refs
    up front. (RevisionManager keeps ``paragraph=None`` as its intended
    document-wide mode.)
    """

    def test_all_edit_methods_reject_none(self, doc_path):
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            calls = [
                lambda: doc.replace("needle", "x", paragraph=None),  # type: ignore[arg-type]
                lambda: doc.delete("needle", paragraph=None),  # type: ignore[arg-type]
                lambda: doc.insert_after("needle", "x", paragraph=None),  # type: ignore[arg-type]
                lambda: doc.insert_before("needle", "x", paragraph=None),  # type: ignore[arg-type]
            ]
            for call in calls:
                with pytest.raises(ValueError, match="'paragraph' must be a paragraph ref string"):
                    call()
            assert doc.list_revisions() == []  # the doc-wide fallback never fired
        finally:
            doc.close()


class TestNonIntOccurrenceRejected:
    """
    Before: ``occurrence`` was the one input the type-hardening pass skipped.
    ``occurrence="0"`` hit ``"0" < 0`` → raw TypeError on every path
    (constructors, direct API, batch apply, dry-run, add_comment);
    ``occurrence=1.5`` passed the ``< 0`` check and TypeErrored deeper in
    the search; ``occurrence=True`` was silently misread as occurrence 1.

    After: a shared guard rejects non-int (including bool) occurrences with
    the same ValueError at every boundary, before any comparison or search.
    """

    def test_direct_api_rejects_str_occurrence(self, doc_path):
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            with pytest.raises(ValueError, match="occurrence must be a non-negative integer, got '0'"):
                doc.replace("needle", "x", paragraph=ref, occurrence="0")  # type: ignore[arg-type]
            assert doc.list_revisions() == []
        finally:
            doc.close()

    def test_constructor_rejects_float_and_bool(self):
        with pytest.raises(
            ValueError, match=r"EditOperation\.replace\(\): occurrence must be a non-negative integer, got 1\.5"
        ):
            EditOperation.replace("a", "b", paragraph="P2#f3c1", occurrence=1.5)  # type: ignore[arg-type]
        with pytest.raises(
            ValueError, match=r"EditOperation\.delete\(\): occurrence must be a non-negative integer, got True"
        ):
            EditOperation.delete("a", paragraph="P2#f3c1", occurrence=True)  # type: ignore[arg-type]

    def test_batch_wraps_str_occurrence_not_raw_typeerror(self, doc_path):
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            before = doc.get_visible_text()
            ref = doc.list_paragraphs()[0].split("|")[0]
            op = EditOperation(action="replace", paragraph=ref, find="needle", replace_with="x", occurrence="0")  # type: ignore[arg-type]
            with pytest.raises(BatchOperationError) as exc:
                doc.batch_edit([op])

            assert exc.value.operation_index == 0
            assert "occurrence must be a non-negative integer" in exc.value.reason
            assert doc.get_visible_text() == before
        finally:
            doc.close()

    def test_dry_run_reports_str_occurrence_as_invalid(self, doc_path):
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            op = EditOperation(action="replace", paragraph=ref, find="needle", replace_with="x", occurrence="0")  # type: ignore[arg-type]
            results = doc.batch_edit([op], dry_run=True)

            assert len(results) == 1
            assert not results[0].valid
            assert results[0].error is not None
            assert "occurrence must be a non-negative integer" in results[0].error
        finally:
            doc.close()

    def test_add_comment_rejects_str_occurrence(self, doc_path):
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            with pytest.raises(ValueError, match="occurrence must be a non-negative integer, got '0'"):
                doc.add_comment("needle", "note", paragraph=ref, occurrence="0")  # type: ignore[arg-type]
        finally:
            doc.close()


class TestNonStringSearchAndPayloadRejected:
    """
    Before: a truthy non-string search target (``find=123``) slipped past
    the falsiness checks into the text-map search and surfaced as a raw
    ``TypeError`` from ``str.find`` — escaping both documented batch
    contracts (``batch_edit`` raises only BatchOperationError;
    ``validate_batch`` never raises). A non-string payload
    (``replace_with=123``) was worse: dry-run reported the op as *valid*,
    then apply crashed with a raw ``AttributeError``.

    After: search/anchor text must be a non-empty ``str`` and payloads must
    be ``str`` at every boundary — direct API, typed constructors, batch
    apply (wrapped with the op index), and dry-run (invalid row, no raise).
    """

    def test_direct_api_rejects_non_string_search_text(self, doc_path):
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            with pytest.raises(ValueError, match="search text must be a non-empty string, got 123"):
                doc.replace(123, "x", paragraph=ref)  # type: ignore[arg-type]
            with pytest.raises(ValueError, match="search text must be a non-empty string, got b'needle'"):
                doc.delete(b"needle", paragraph=ref)  # type: ignore[arg-type]
            assert doc.list_revisions() == []
        finally:
            doc.close()

    def test_direct_api_rejects_non_string_payload(self, doc_path):
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            with pytest.raises(ValueError, match="'replace_with' must be a string"):
                doc.replace("needle", 123, paragraph=ref)  # type: ignore[arg-type]
            with pytest.raises(ValueError, match="'text' must be a string"):
                doc.insert_after("needle", 123, paragraph=ref)  # type: ignore[arg-type]
            with pytest.raises(ValueError, match="'text' must be a string"):
                doc.insert_before("needle", 123, paragraph=ref)  # type: ignore[arg-type]
            assert doc.list_revisions() == []
        finally:
            doc.close()

    def test_batch_wraps_non_string_target_not_raw_typeerror(self, doc_path):
        """Hand-built op (bypassing the typed constructors) with an int find."""
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            before = doc.get_visible_text()
            ref = doc.list_paragraphs()[0].split("|")[0]
            op = EditOperation(action="replace", paragraph=ref, find=123, replace_with="x")  # type: ignore[arg-type]
            with pytest.raises(BatchOperationError) as exc:
                doc.batch_edit([op])

            assert exc.value.operation_index == 0
            assert isinstance(exc.value.original, ValueError)
            assert "search text must be a non-empty string" in exc.value.reason
            assert doc.get_visible_text() == before
        finally:
            doc.close()

    def test_batch_wraps_non_string_payload_not_raw_attributeerror(self, doc_path):
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            before = doc.get_visible_text()
            ref = doc.list_paragraphs()[0].split("|")[0]
            op = EditOperation(action="replace", paragraph=ref, find="needle", replace_with=123)  # type: ignore[arg-type]
            with pytest.raises(BatchOperationError) as exc:
                doc.batch_edit([op])

            assert exc.value.operation_index == 0
            assert "a string 'replace_with'" in exc.value.reason
            assert doc.get_visible_text() == before
        finally:
            doc.close()

    def test_dry_run_reports_non_string_target_as_invalid(self, doc_path):
        """validate_batch's never-raises contract holds for an int find."""
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            op = EditOperation(action="replace", paragraph=ref, find=123, replace_with="x")  # type: ignore[arg-type]
            results = doc.batch_edit([op], dry_run=True)

            assert len(results) == 1
            assert not results[0].valid
            assert results[0].error is not None
            assert "search text must be a non-empty string" in results[0].error
        finally:
            doc.close()

    def test_dry_run_reports_non_string_payload_as_invalid(self, doc_path):
        """Before the fix this came back valid=True, then crashed at apply."""
        doc = _build_doc_with_paragraphs(doc_path, ["one needle only."])
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            op = EditOperation(action="replace", paragraph=ref, find="needle", replace_with=123)  # type: ignore[arg-type]
            results = doc.batch_edit([op], dry_run=True)

            assert len(results) == 1
            assert not results[0].valid
            assert results[0].error is not None
            assert "a string 'replace_with'" in results[0].error
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
