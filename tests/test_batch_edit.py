"""Tests for batch_edit() with reverse-order application."""

import shutil
import tempfile
from pathlib import Path

import pytest

from docx_editor import (
    AmbiguousTextError,
    BatchOperationError,
    Document,
    EditOperation,
    EditValidationResult,
    HashMismatchError,
    TextNotFoundError,
)


@pytest.fixture
def multi_para_doc():
    """Build a document with 10 paragraphs, each containing repeated phrases."""
    test_data = Path(__file__).parent / "test_data" / "simple.docx"
    tmp = tempfile.mkdtemp(prefix="batch_test_")
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


class TestBatchEdit:
    def test_batch_multiple_paragraphs(self, multi_para_doc):
        """Batch of edits to different paragraphs all succeed."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        ops = [
            EditOperation(
                action="replace",
                find="item 3",
                replace_with="ITEM_THREE",
                paragraph=refs[2].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="item 7",
                replace_with="ITEM_SEVEN",
                paragraph=refs[6].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="item 9",
                replace_with="ITEM_NINE",
                paragraph=refs[8].split("|")[0],
            ),
        ]

        result = doc.batch_edit(ops)
        assert len(result) == 3

        vis = doc.get_visible_text()
        assert "ITEM_THREE" in vis
        assert "ITEM_SEVEN" in vis
        assert "ITEM_NINE" in vis

    def test_batch_rejected_on_stale_hash(self, multi_para_doc):
        """Entire batch rejected if any hash is stale."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        # Edit P5 to make its hash stale
        p5_ref = refs[4].split("|")[0]
        doc.replace("item 5", "CHANGED", paragraph=p5_ref)

        # Now try a batch using OLD refs (P5 hash is stale)
        ops = [
            EditOperation(
                action="replace",
                find="item 3",
                replace_with="EDIT_3",
                paragraph=refs[2].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="CHANGED",
                replace_with="EDIT_5",
                paragraph=p5_ref,  # STALE hash
            ),
            EditOperation(
                action="replace",
                find="item 8",
                replace_with="EDIT_8",
                paragraph=refs[7].split("|")[0],
            ),
        ]

        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        assert exc.value.operation_index == 1
        assert isinstance(exc.value.original, HashMismatchError)

        # Verify NO edits were applied
        vis = doc.get_visible_text()
        assert "EDIT_3" not in vis
        assert "EDIT_5" not in vis
        assert "EDIT_8" not in vis

    def test_batch_reverse_order_preserves_hashes(self, multi_para_doc):
        """Edits applied in reverse order — P10 edit doesn't invalidate P2 hash."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        # Take ONE snapshot, edit both ends
        ops = [
            EditOperation(
                action="replace",
                find="item 2",
                replace_with="SECOND",
                paragraph=refs[1].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="item 10",
                replace_with="TENTH",
                paragraph=refs[9].split("|")[0],
            ),
        ]

        result = doc.batch_edit(ops)
        assert len(result) == 2

        vis = doc.get_visible_text()
        assert "SECOND" in vis
        assert "TENTH" in vis

    def test_single_snapshot_suffices(self, multi_para_doc):
        """One list_paragraphs() call works for all 10 edits."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        ops = []
        for i in range(10):
            ops.append(
                EditOperation(
                    action="replace",
                    find=f"item {i + 1}",
                    replace_with=f"BATCH_{i + 1}",
                    paragraph=refs[i].split("|")[0],
                )
            )

        result = doc.batch_edit(ops)
        assert len(result) == 10

        vis = doc.get_visible_text()
        for i in range(10):
            assert f"BATCH_{i + 1}" in vis

    def test_duplicate_paragraph_targets(self, multi_para_doc):
        """Two edits to the same paragraph both apply."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        p3_ref = refs[2].split("|")[0]
        ops = [
            EditOperation(
                action="replace",
                find="item 3",
                replace_with="FIRST_EDIT",
                paragraph=p3_ref,
            ),
            EditOperation(
                action="replace",
                find="committee",
                replace_with="BOARD",
                paragraph=p3_ref,
                occurrence=0,
            ),
        ]

        result = doc.batch_edit(ops)
        assert len(result) == 2

        vis = doc.get_visible_text()
        assert "FIRST_EDIT" in vis
        assert "BOARD" in vis

    def test_empty_batch(self, multi_para_doc):
        """Empty batch returns empty list."""
        doc, _ = multi_para_doc
        result = doc.batch_edit([])
        assert result == []

    def test_batch_mixed_actions(self, multi_para_doc):
        """Batch with replace, delete, and insert actions."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        ops = [
            EditOperation(
                action="replace",
                find="item 2",
                replace_with="REPLACED",
                paragraph=refs[1].split("|")[0],
            ),
            EditOperation(
                action="delete",
                text="item 5",
                paragraph=refs[4].split("|")[0],
            ),
            EditOperation(
                action="insert_after",
                anchor="item 8",
                text=" [APPENDED]",
                paragraph=refs[7].split("|")[0],
            ),
        ]

        result = doc.batch_edit(ops)
        assert len(result) == 3

        vis = doc.get_visible_text()
        assert "REPLACED" in vis
        # Verify "item 5" was deleted (but [P05] marker remains in the paragraph)
        assert "[P05]" in vis
        p05_line = [line for line in vis.split("\n") if "[P05]" in line][0]
        assert "item 5" not in p05_line
        assert "[APPENDED]" in vis

    def test_batch_missing_paragraph_raises(self, multi_para_doc):
        """Batch with missing paragraph field raises BatchOperationError."""
        doc, _ = multi_para_doc

        ops = [
            EditOperation(action="replace", find="a", replace_with="b", paragraph=""),
        ]

        with pytest.raises(BatchOperationError, match="paragraph reference is required"):
            doc.batch_edit(ops)

    def test_batch_replace_missing_find(self, multi_para_doc):
        """Replace without find raises BatchOperationError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="replace", find=None, replace_with="x", paragraph=ref)]
        with pytest.raises(BatchOperationError, match="replace requires"):
            doc.batch_edit(ops)

    def test_batch_replace_text_not_found(self, multi_para_doc):
        """Replace with non-existent text raises BatchOperationError wrapping TextNotFoundError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="replace", find="NONEXISTENT", replace_with="x", paragraph=ref)]
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        assert exc.value.operation_index == 0
        assert isinstance(exc.value.original, TextNotFoundError)

    def test_batch_delete_missing_text(self, multi_para_doc):
        """Delete without text raises BatchOperationError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="delete", text=None, paragraph=ref)]
        with pytest.raises(BatchOperationError, match="delete requires"):
            doc.batch_edit(ops)

    def test_batch_delete_text_not_found(self, multi_para_doc):
        """Delete with non-existent text raises BatchOperationError wrapping TextNotFoundError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="delete", text="NONEXISTENT", paragraph=ref)]
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        assert exc.value.operation_index == 0
        assert isinstance(exc.value.original, TextNotFoundError)

    def test_batch_insert_missing_anchor(self, multi_para_doc):
        """Insert without anchor raises BatchOperationError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="insert_after", anchor=None, text="x", paragraph=ref)]
        with pytest.raises(BatchOperationError, match="insert_after requires"):
            doc.batch_edit(ops)

    def test_batch_insert_anchor_not_found(self, multi_para_doc):
        """Insert with non-existent anchor raises BatchOperationError wrapping TextNotFoundError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="insert_after", anchor="NONEXISTENT", text="x", paragraph=ref)]
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        assert exc.value.operation_index == 0
        assert isinstance(exc.value.original, TextNotFoundError)

    def test_batch_unknown_action(self, multi_para_doc):
        """Unknown action raises BatchOperationError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="unknown", paragraph=ref)]  # type: ignore[arg-type]
        with pytest.raises(BatchOperationError, match="Unknown action"):
            doc.batch_edit(ops)


class TestBatchEditDryRun:
    """dry_run=True validates every op without mutating the document."""

    def test_all_valid(self, multi_para_doc):
        """All-matching ops report valid=True, in input order, doc unchanged."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        before = doc.get_visible_text()

        ops = [
            EditOperation(
                action="replace",
                find="item 3",
                replace_with="ITEM_THREE",
                paragraph=refs[2].split("|")[0],
            ),
            EditOperation(
                action="delete",
                text="item 5",
                paragraph=refs[4].split("|")[0],
            ),
            EditOperation(
                action="insert_after",
                anchor="item 8",
                text=" [APPENDED]",
                paragraph=refs[7].split("|")[0],
            ),
        ]

        results = doc.batch_edit(ops, dry_run=True)

        assert len(results) == len(ops)
        assert all(isinstance(r, EditValidationResult) for r in results)
        assert all(r.valid for r in results)
        assert all(r.error is None for r in results)
        assert [r.index for r in results] == [0, 1, 2]
        assert [r.paragraph for r in results] == [op.paragraph for op in ops]

        # No edits applied.
        assert doc.get_visible_text() == before

    def test_stale_hash(self, multi_para_doc):
        """A stale-hash op reports invalid with a hash-mismatch error; doc unchanged."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        # Make P5's hash stale by editing it.
        p5_ref = refs[4].split("|")[0]
        doc.replace("item 5", "CHANGED", paragraph=p5_ref)
        before = doc.get_visible_text()

        ops = [
            EditOperation(
                action="replace",
                find="CHANGED",
                replace_with="EDIT_5",
                paragraph=p5_ref,  # STALE hash
            ),
        ]

        results = doc.batch_edit(ops, dry_run=True)

        assert len(results) == 1
        assert results[0].valid is False
        assert "hash" in results[0].error.lower()
        assert doc.get_visible_text() == before

    def test_missing_text(self, multi_para_doc):
        """A valid ref but non-existent find/text reports invalid; doc unchanged."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        before = doc.get_visible_text()

        ops = [
            EditOperation(
                action="replace",
                find="NONEXISTENT",
                replace_with="x",
                paragraph=ref,
            ),
        ]

        results = doc.batch_edit(ops, dry_run=True)

        assert results[0].valid is False
        assert "not found" in results[0].error.lower()
        assert doc.get_visible_text() == before

    def test_mixed_valid_invalid(self, multi_para_doc):
        """Valid + stale-hash + missing-text ops report per-op; order preserved; doc unchanged."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        # Make P5 stale.
        p5_ref = refs[4].split("|")[0]
        doc.replace("item 5", "CHANGED", paragraph=p5_ref)
        before = doc.get_visible_text()

        ops = [
            EditOperation(
                action="replace",
                find="item 3",
                replace_with="ITEM_THREE",
                paragraph=refs[2].split("|")[0],  # valid
            ),
            EditOperation(
                action="replace",
                find="CHANGED",
                replace_with="x",
                paragraph=p5_ref,  # stale hash
            ),
            EditOperation(
                action="delete",
                text="NONEXISTENT",
                paragraph=refs[7].split("|")[0],  # missing text
            ),
        ]

        results = doc.batch_edit(ops, dry_run=True)

        assert [r.index for r in results] == [0, 1, 2]
        assert results[0].valid is True
        assert results[0].error is None
        assert results[1].valid is False
        assert "hash" in results[1].error.lower()
        assert results[2].valid is False
        assert "not found" in results[2].error.lower()

        assert doc.get_visible_text() == before

    def test_empty_batch(self, multi_para_doc):
        """Empty dry-run batch returns an empty list."""
        doc, _ = multi_para_doc
        assert doc.batch_edit([], dry_run=True) == []

    def test_not_found_error_includes_preview(self, multi_para_doc):
        """A not-found error carries the paragraph preview for easier debugging."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]

        ops = [
            EditOperation(action="replace", find="NONEXISTENT", replace_with="x", paragraph=ref),
        ]

        results = doc.batch_edit(ops, dry_run=True)

        assert results[0].valid is False
        # The paragraph's real text ("item 1") should appear in the error preview.
        assert "item 1" in results[0].error

    def test_bad_occurrence_reports_instead_of_raising(self, multi_para_doc):
        """A malformed op (negative occurrence) is reported, never raised; doc unchanged."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        before = doc.get_visible_text()

        ops = [
            EditOperation(
                action="replace",
                find="item 1",
                replace_with="x",
                paragraph=ref,
                occurrence=-1,  # malformed
            ),
        ]

        # Must not raise despite the malformed occurrence.
        results = doc.batch_edit(ops, dry_run=True)

        assert results[0].valid is False
        assert "occurrence" in results[0].error
        assert doc.get_visible_text() == before

    def test_missing_paragraph_ref(self, multi_para_doc):
        """An op without a paragraph ref is invalid, with paragraph=None on the result."""
        doc, _ = multi_para_doc

        ops = [EditOperation(action="replace", find="item 1", replace_with="x", paragraph="")]

        results = doc.batch_edit(ops, dry_run=True)

        assert results[0].valid is False
        assert results[0].paragraph == ""
        assert "paragraph" in results[0].error.lower()

    def test_malformed_paragraph_ref(self, multi_para_doc):
        """An unparseable paragraph ref is reported invalid, never raised."""
        doc, _ = multi_para_doc

        ops = [EditOperation(action="replace", find="x", replace_with="y", paragraph="not-a-ref")]

        results = doc.batch_edit(ops, dry_run=True)

        assert results[0].valid is False
        assert results[0].error

    def test_missing_action_arguments(self, multi_para_doc):
        """Each action reports its missing-argument error without raising."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        r0 = refs[0].split("|")[0]
        r1 = refs[1].split("|")[0]
        r2 = refs[2].split("|")[0]

        ops = [
            # replace without find/replace_with
            EditOperation(action="replace", paragraph=r0),
            # delete without text
            EditOperation(action="delete", paragraph=r1),
            # insert_after without anchor/text
            EditOperation(action="insert_after", paragraph=r2),
        ]

        results = doc.batch_edit(ops, dry_run=True)

        assert all(r.valid is False for r in results)
        assert "replace requires" in results[0].error
        assert "delete requires" in results[1].error
        assert "insert_after requires" in results[2].error

    def test_unknown_action(self, multi_para_doc):
        """An unrecognized action is reported invalid, never raised."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]

        ops = [EditOperation(action="frobnicate", paragraph=ref)]  # type: ignore[arg-type]

        results = doc.batch_edit(ops, dry_run=True)

        assert results[0].valid is False
        assert "Unknown action" in results[0].error


class TestStructuredBatchOperationError:
    def test_predispatch_missing_paragraph_has_operation_index(self, multi_para_doc):
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        ops = [
            EditOperation(
                action="replace",
                find="item 1",
                replace_with="x",
                paragraph=refs[0].split("|")[0],
            ),
            EditOperation(action="replace", find="a", replace_with="b", paragraph=""),
        ]
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        err = exc.value
        assert err.operation_index == 1
        assert "paragraph" in err.reason.lower()
        assert "1" in str(err)

    def test_inner_missing_field_has_operation_index(self, multi_para_doc):
        """`ValueError` raised inside `_apply_single_edit` is wrapped with the op index."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        ops = [
            EditOperation(
                action="replace",
                find=None,
                replace_with="x",
                paragraph=refs[0].split("|")[0],
            ),
        ]
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        err = exc.value
        assert err.operation_index == 0
        assert "replace requires" in err.reason

    def test_inner_unknown_action_has_operation_index(self, multi_para_doc):
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        ops = [
            EditOperation(
                action="replace",
                find="item 1",
                replace_with="x",
                paragraph=refs[0].split("|")[0],
            ),
            EditOperation(action="bogus", paragraph=refs[1].split("|")[0]),  # type: ignore[arg-type]
        ]
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        err = exc.value
        assert err.operation_index == 1
        assert "Unknown action" in err.reason

    def test_inner_negative_occurrence_has_operation_index(self, multi_para_doc):
        """A negative occurrence fails cleanly on the apply path (no internal error)."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        before = doc.get_visible_text()
        ops = [
            EditOperation(
                action="replace",
                find="item 1",
                replace_with="x",
                paragraph=refs[0].split("|")[0],
                occurrence=-1,
            ),
        ]
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        err = exc.value
        assert err.operation_index == 0
        assert "occurrence" in err.reason
        # Rejected before any mutation.
        assert doc.get_visible_text() == before

    def test_non_batch_value_error_unchanged(self):
        """Malformed paragraph refs outside batch paths still raise plain ValueError."""
        from docx_editor import ParagraphRef

        with pytest.raises(ValueError, match="Invalid paragraph reference"):
            ParagraphRef.parse("not-a-valid-ref")

    def test_batch_operation_error_is_docx_edit_error(self, multi_para_doc):
        from docx_editor import DocxEditError

        doc, _ = multi_para_doc
        ops = [EditOperation(action="replace", find="a", replace_with="b", paragraph="")]
        with pytest.raises(DocxEditError):
            doc.batch_edit(ops)


class TestSingleExceptionContract:
    """batch_edit raises ONLY BatchOperationError for per-op failures — both
    phases — so one recovery path (`ops.pop(e.operation_index)`) always works.
    The underlying typed exception stays reachable via `original`/`__cause__`."""

    def test_malformed_ref_wrapped_with_index(self, multi_para_doc):
        """Validation-phase ValueError from ParagraphRef.parse is wrapped."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        ops = [
            EditOperation(action="replace", find="item 1", replace_with="x", paragraph=refs[0].split("|")[0]),
            EditOperation(action="replace", find="item 2", replace_with="x", paragraph="garbage-ref"),
        ]
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        assert exc.value.operation_index == 1
        assert isinstance(exc.value.original, ValueError)
        assert exc.value.__cause__ is exc.value.original
        assert "Invalid paragraph reference" in exc.value.reason

    def test_ambiguous_op_wrapped_and_rolled_back(self, multi_para_doc):
        """An ambiguous target inside a batch wraps AmbiguousTextError; the
        document is unchanged afterwards."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        before = doc.get_visible_text()
        ops = [
            EditOperation(action="replace", find="item 1", replace_with="ONE", paragraph=refs[0].split("|")[0]),
            # "committee" appears twice in every paragraph — ambiguous without occurrence
            EditOperation(action="replace", find="committee", replace_with="x", paragraph=refs[1].split("|")[0]),
        ]
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        assert exc.value.operation_index == 1
        assert isinstance(exc.value.original, AmbiguousTextError)
        assert exc.value.original.total_occurrences == 2
        assert doc.get_visible_text() == before

    def test_stale_hash_original_carries_usable_actual_hash(self, multi_para_doc):
        """`e.original.actual_hash` re-targets the stale op without list_paragraphs()."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        stale_ref = refs[4].split("|")[0]
        doc.replace("item 5", "CHANGED", paragraph=stale_ref)

        ops = [EditOperation(action="replace", find="CHANGED", replace_with="FIXED", paragraph=stale_ref)]
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        original = exc.value.original
        assert isinstance(original, HashMismatchError)

        # Re-target using the structured fields alone.
        fresh_ref = f"P{original.paragraph_index}#{original.actual_hash}"
        ops[exc.value.operation_index] = EditOperation(
            action="replace", find="CHANGED", replace_with="FIXED", paragraph=fresh_ref
        )
        doc.batch_edit(ops)
        assert "FIXED" in doc.get_visible_text()

    def test_documented_recovery_loop(self, multi_para_doc):
        """The SKILL.md recovery pattern: pop the failing op, retry, succeed."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        ops = [
            EditOperation(action="replace", find="item 2", replace_with="TWO", paragraph=refs[1].split("|")[0]),
            EditOperation(action="replace", find="MISSING", replace_with="x", paragraph=refs[3].split("|")[0]),
            EditOperation(action="replace", find="item 6", replace_with="SIX", paragraph=refs[5].split("|")[0]),
        ]

        while ops:
            try:
                doc.batch_edit(ops)
                break
            except BatchOperationError as e:
                ops.pop(e.operation_index)

        vis = doc.get_visible_text()
        assert "TWO" in vis
        assert "SIX" in vis
        assert "MISSING" not in vis


class TestSameParagraphBatch:
    """The documented pattern for several matches in ONE paragraph: batch the
    ops in DESCENDING occurrence order — an edit never shifts the matches
    before it. Ascending order breaks because each edit shifts the occurrence
    numbering of the matches after it in the accepted text view."""

    def test_descending_occurrence_order_hits_every_intended_match(self, multi_para_doc):
        doc, _ = multi_para_doc
        results = [r for r in doc.find_all("committee") if r.paragraph_ref.startswith("P3#")]
        assert [r.paragraph_occurrence for r in results] == [0, 1]

        ops = [
            EditOperation.replace(
                r.text, f"EDIT{r.paragraph_occurrence}", paragraph=r.paragraph_ref, occurrence=r.paragraph_occurrence
            )
            for r in sorted(results, key=lambda r: r.paragraph_occurrence, reverse=True)
        ]
        doc.batch_edit(ops)
        doc.accept_all()
        assert (
            "[P03] The EDIT0 shall review item 3. The report includes findings from the EDIT1."
            in doc.get_visible_text()
        )

    def test_ascending_occurrence_order_fails_atomically(self, multi_para_doc):
        """find_all order (ascending) is NOT batchable within a paragraph: the
        first edit renumbers the rest, so a later op goes out of range and the
        atomic contract rolls everything back."""
        doc, _ = multi_para_doc
        before = doc.get_visible_text()
        results = [r for r in doc.find_all("committee") if r.paragraph_ref.startswith("P3#")]
        ops = [
            EditOperation.replace(r.text, "X", paragraph=r.paragraph_ref, occurrence=r.paragraph_occurrence)
            for r in results
        ]
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        assert isinstance(exc.value.original, TextNotFoundError)
        assert doc.get_visible_text() == before


class TestBatchRewriteSingleExceptionContract:
    """batch_rewrite honors the same single-exception contract."""

    def test_stale_hash_wrapped_with_index(self, multi_para_doc):
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        before = doc.get_visible_text()
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_rewrite([
                (refs[0].split("|")[0], "New first paragraph."),
                ("P5#0000", "Stale hash."),
            ])
        assert exc.value.operation_index == 1
        assert isinstance(exc.value.original, HashMismatchError)
        assert doc.get_visible_text() == before

    def test_malformed_ref_wrapped_with_index(self, multi_para_doc):
        doc, _ = multi_para_doc
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_rewrite([("bogus", "text")])
        assert exc.value.operation_index == 0
        assert isinstance(exc.value.original, ValueError)


class TestBatchEditRollback:
    """Mid-batch failure must leave the document untouched (atomic contract)."""

    def test_rollback_on_text_not_found_mid_batch(self, multi_para_doc):
        """First op applies, second op's find text doesn't exist — DOM must roll back."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        before_text = doc.get_visible_text()
        before_revisions = len(doc.list_revisions())

        ops = [
            EditOperation(
                action="replace",
                find="item 3",
                replace_with="ITEM_THREE",
                paragraph=refs[2].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="DOES_NOT_EXIST_IN_P7",
                replace_with="x",
                paragraph=refs[6].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="item 9",
                replace_with="ITEM_NINE",
                paragraph=refs[8].split("|")[0],
            ),
        ]

        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        assert exc.value.operation_index == 1
        assert isinstance(exc.value.original, TextNotFoundError)

        # Reverse-paragraph order means P9 applied before P7 failed; full-text
        # equality proves P9's mutation was rolled back.
        assert doc.get_visible_text() == before_text
        assert len(doc.list_revisions()) == before_revisions

    def test_rollback_on_structural_error_mid_batch(self, multi_para_doc):
        """First op is valid; second op has replace_with=None — DOM must roll back."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        before_text = doc.get_visible_text()
        before_revisions = len(doc.list_revisions())

        ops = [
            EditOperation(
                action="replace",
                find="item 2",
                replace_with="SECOND",
                paragraph=refs[1].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="item 4",
                replace_with=None,
                paragraph=refs[3].split("|")[0],
            ),
        ]

        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        assert exc.value.operation_index == 1

        assert doc.get_visible_text() == before_text
        assert len(doc.list_revisions()) == before_revisions

    def test_rollback_preserves_hash_anchored_refs(self, multi_para_doc):
        """After a failed batch, pre-batch paragraph refs still resolve cleanly."""
        doc, _ = multi_para_doc
        refs_before = doc.list_paragraphs()

        ops = [
            EditOperation(
                action="replace",
                find="item 2",
                replace_with="WILL_ROLLBACK",
                paragraph=refs_before[1].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="MISSING",
                replace_with="x",
                paragraph=refs_before[5].split("|")[0],
            ),
        ]

        with pytest.raises(BatchOperationError):
            doc.batch_edit(ops)

        # The pre-batch P2 ref must still resolve — proves hashes did not drift.
        doc.replace("item 2", "POST_ROLLBACK", paragraph=refs_before[1].split("|")[0])
        assert "POST_ROLLBACK" in doc.get_visible_text()
        assert "WILL_ROLLBACK" not in doc.get_visible_text()

    def test_rollback_preserves_parse_position_metadata(self, multi_para_doc):
        """After rollback, parse_position attributes on elements must survive,
        so XMLEditor.get_node(line_number=...) keeps working."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        editor = doc._document_editor

        # Sanity: parse_position is set on every element before any batch.
        paragraphs_before = editor.dom.getElementsByTagName("w:p")
        assert paragraphs_before
        assert all(hasattr(p, "parse_position") for p in paragraphs_before)

        ops = [
            EditOperation(
                action="replace",
                find="item 2",
                replace_with="x",
                paragraph=refs[1].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="MISSING",
                replace_with="x",
                paragraph=refs[5].split("|")[0],
            ),
        ]

        with pytest.raises(BatchOperationError):
            doc.batch_edit(ops)

        # After rollback, parse_position must still be present on every element
        # — proves the line-tracking parser was used for the restore.
        paragraphs_after = editor.dom.getElementsByTagName("w:p")
        assert paragraphs_after
        assert all(hasattr(p, "parse_position") for p in paragraphs_after)

    def test_rollback_failure_surfaces_original_error(self, multi_para_doc, monkeypatch):
        """If the rollback re-parse itself fails, the original edit error must
        still propagate — not the rollback failure."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        editor = doc._document_editor
        before_text = doc.get_visible_text()

        def boom(xml_bytes):
            raise RuntimeError("simulated rollback failure")

        monkeypatch.setattr(editor, "_reload_dom_from_bytes", boom)

        ops = [
            EditOperation(
                action="replace",
                find="item 2",
                replace_with="x",
                paragraph=refs[1].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="MISSING",
                replace_with="x",
                paragraph=refs[5].split("|")[0],
            ),
        ]

        # Original edit error wins over the rollback failure.
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)
        assert isinstance(exc.value.original, TextNotFoundError)

        # Reverse-paragraph order means the failing P6 op runs first and
        # raises before any mutation, so visible text is still unchanged
        # even though the rollback re-parse itself was swallowed.
        assert doc.get_visible_text() == before_text


class TestEditOperationConstructors:
    """Typed constructors validate at construction time, mirroring apply-time rules."""

    def test_replace_builds_op(self):
        op = EditOperation.replace("old", "new", paragraph="P2#f3c1", occurrence=1)
        assert op == EditOperation(action="replace", paragraph="P2#f3c1", find="old", replace_with="new", occurrence=1)

    def test_delete_builds_op(self):
        op = EditOperation.delete("gone", paragraph="P2#f3c1")
        assert op == EditOperation(action="delete", paragraph="P2#f3c1", text="gone")

    def test_insert_after_builds_op(self):
        op = EditOperation.insert_after("anchor", " tail", paragraph="P2#f3c1")
        assert op == EditOperation(action="insert_after", paragraph="P2#f3c1", anchor="anchor", text=" tail")

    def test_insert_before_builds_op(self):
        op = EditOperation.insert_before("anchor", "head ", paragraph="P2#f3c1")
        assert op == EditOperation(action="insert_before", paragraph="P2#f3c1", anchor="anchor", text="head ")

    def test_typed_ops_apply_end_to_end(self, multi_para_doc):
        """A batch built exclusively from typed constructors applies cleanly."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        ops = [
            EditOperation.replace("item 3", "ITEM_THREE", paragraph=refs[2].split("|")[0]),
            EditOperation.delete("The report includes findings", paragraph=refs[4].split("|")[0]),
            EditOperation.insert_after("item 7", " (amended)", paragraph=refs[6].split("|")[0]),
            EditOperation.insert_before("The committee", "NOTE: ", paragraph=refs[8].split("|")[0]),
        ]
        doc.batch_edit(ops)
        # Inserts land relative to the run containing the anchor (each fixture
        # paragraph is a single run), so assert per-paragraph outcomes.
        paras = doc.list_paragraphs(max_chars=200)
        assert "ITEM_THREE" in paras[2]
        assert "The report includes findings" not in paras[4]
        assert "(amended)" in paras[6]
        assert "NOTE: " in paras[8]

    def test_constructor_rules_match_apply_rules(self, multi_para_doc):
        """Drift guard: whatever the constructors accept must pass apply-time validation.

        If ``_resolve_action_target`` ever becomes stricter than the typed
        constructors (or vice versa), this dry-run disagrees and fails.
        """
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()
        ops = [
            # Boundary inputs the constructors deliberately allow:
            EditOperation.replace("item 1", "", paragraph=refs[0].split("|")[0]),
            EditOperation.delete("item 2", paragraph=refs[1].split("|")[0]),
            EditOperation.insert_after("item 3", "", paragraph=refs[2].split("|")[0]),
            EditOperation.insert_before("item 4", "x", paragraph=refs[3].split("|")[0]),
        ]
        results = doc.batch_edit(ops, dry_run=True)
        assert all(r.valid for r in results), [r.error for r in results]

    def test_malformed_paragraph_ref_rejected(self):
        with pytest.raises(ValueError, match="Invalid paragraph reference 'P3'"):
            EditOperation.replace("a", "b", paragraph="P3")

    def test_empty_paragraph_ref_rejected(self):
        with pytest.raises(ValueError, match="Invalid paragraph reference"):
            EditOperation.delete("a", paragraph="")

    def test_none_paragraph_ref_rejected(self):
        """paragraph=None gets the field-specific ValueError, not a raw regex TypeError."""
        with pytest.raises(ValueError, match="Invalid paragraph reference None"):
            EditOperation.replace("a", "b", paragraph=None)  # type: ignore[arg-type]

    def test_delete_none_paragraph_ref_rejected(self):
        with pytest.raises(ValueError, match="Invalid paragraph reference None"):
            EditOperation.delete("x", paragraph=None)  # type: ignore[arg-type]

    def test_non_string_paragraph_ref_names_the_type(self):
        with pytest.raises(ValueError, match="expected a string like 'P3#a7b2', got int"):
            EditOperation.insert_after("a", "x", paragraph=3)  # type: ignore[arg-type]

    def test_non_string_search_text_rejected(self):
        """A truthy non-string target fails at construction, not deep in the search."""
        with pytest.raises(ValueError, match=r"'find' must be a non-empty string"):
            EditOperation.replace(123, "x", paragraph="P2#f3c1")  # type: ignore[arg-type]
        with pytest.raises(ValueError, match=r"'text' must be a non-empty string"):
            EditOperation.delete(123, paragraph="P2#f3c1")  # type: ignore[arg-type]
        with pytest.raises(ValueError, match=r"'anchor' must be a non-empty string"):
            EditOperation.insert_after(123, "x", paragraph="P2#f3c1")  # type: ignore[arg-type]

    def test_non_string_payload_rejected(self):
        with pytest.raises(ValueError, match=r"'replace_with' must be a string"):
            EditOperation.replace("a", 123, paragraph="P2#f3c1")  # type: ignore[arg-type]
        with pytest.raises(ValueError, match=r"'text' must be a string"):
            EditOperation.insert_before("a", 123, paragraph="P2#f3c1")  # type: ignore[arg-type]

    def test_negative_occurrence_rejected(self):
        with pytest.raises(ValueError, match=r"EditOperation\.replace\(\): occurrence must be >= 0, got -1"):
            EditOperation.replace("a", "b", paragraph="P2#f3c1", occurrence=-1)

    def test_replace_empty_find_rejected(self):
        with pytest.raises(ValueError, match=r"EditOperation\.replace\(\): 'find' must be a non-empty string"):
            EditOperation.replace("", "b", paragraph="P2#f3c1")

    def test_delete_empty_text_rejected(self):
        with pytest.raises(ValueError, match=r"EditOperation\.delete\(\): 'text' must be a non-empty string"):
            EditOperation.delete("", paragraph="P2#f3c1")

    def test_insert_after_empty_anchor_rejected(self):
        with pytest.raises(ValueError, match=r"EditOperation\.insert_after\(\): 'anchor' must be a non-empty"):
            EditOperation.insert_after("", "x", paragraph="P2#f3c1")

    def test_insert_before_empty_anchor_rejected(self):
        with pytest.raises(ValueError, match=r"EditOperation\.insert_before\(\): 'anchor' must be a non-empty"):
            EditOperation.insert_before("", "x", paragraph="P2#f3c1")

    def test_replace_with_empty_string_accepted(self):
        """Replacing with nothing is a valid tracked deletion — parity with apply-time rules."""
        op = EditOperation.replace("old", "", paragraph="P2#f3c1")
        assert op.replace_with == ""

    def test_replace_with_none_rejected(self):
        with pytest.raises(ValueError, match=r"EditOperation\.replace\(\): 'replace_with' must be a string"):
            EditOperation.replace("old", None, paragraph="P2#f3c1")  # type: ignore[arg-type]

    def test_insert_empty_text_accepted(self):
        """Empty insert text is allowed at apply time, so the constructor allows it too."""
        op = EditOperation.insert_after("anchor", "", paragraph="P2#f3c1")
        assert op.text == ""

    def test_insert_none_text_rejected(self):
        with pytest.raises(ValueError, match=r"EditOperation\.insert_before\(\): 'text' must be a string"):
            EditOperation.insert_before("anchor", None, paragraph="P2#f3c1")  # type: ignore[arg-type]
