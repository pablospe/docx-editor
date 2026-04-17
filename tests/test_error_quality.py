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
            assert err.paragraph_preview.rstrip(".") in msg
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

    def _build_batch_doc(self, doc_path):
        doc = Document.open(doc_path, force_recreate=True)
        editor = doc._document_editor
        body = editor.dom.getElementsByTagName("w:body")[0]
        for p in list(editor.dom.getElementsByTagName("w:p")):
            if p.parentNode == body:
                body.removeChild(p)
        sect_pr = editor.dom.getElementsByTagName("w:sectPr")
        insert_before = sect_pr[0] if sect_pr else None
        for i in range(1, 4):
            xml = f'<w:p><w:r><w:t xml:space="preserve">Paragraph {i} content.</w:t></w:r></w:p>'
            for node in editor._parse_fragment(xml):
                if insert_before:
                    body.insertBefore(node, insert_before)
                else:
                    body.appendChild(node)
        editor.save()
        saved = doc.save()
        doc.close()
        return Document.open(saved, force_recreate=True)

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


class TestSharedBaseClass:
    """All four structured errors inherit from `DocxEditError`.

    This means `except DocxEditError` is a correct catch-all for LLM
    consumers — no need to enumerate the leaf classes.
    """

    def test_all_structured_errors_share_docx_edit_error_base(self):
        from docx_editor import DocxEditError, HashMismatchError

        assert issubclass(TextNotFoundError, DocxEditError)
        assert issubclass(HashMismatchError, DocxEditError)
        assert issubclass(ParagraphIndexError, DocxEditError)
        assert issubclass(BatchOperationError, DocxEditError)

    def test_top_level_import_smoke(self):
        """Every structured error is importable from the top-level package."""
        from docx_editor import (  # noqa: F401
            BatchOperationError,
            HashMismatchError,
            ParagraphIndexError,
            TextNotFoundError,
        )
