## 1. Structured TextNotFoundError (additive)

- [x] 1.1 Write failing tests in `tests/test_paragraph_hash.py`: scoped `TextNotFoundError` carries `search_text`, `paragraph_ref`, `paragraph_preview`; message includes current paragraph content; unscoped call leaves `paragraph_ref` and `paragraph_preview` as `None`. Also test occurrence-based failure carries `occurrence` and `total_occurrences`.
- [x] 1.2 Update `docx_editor/exceptions.py`: add `__init__` to `TextNotFoundError` accepting `search_text` (positional) and keyword-only `paragraph_ref`, `paragraph_preview`, `occurrence`, `total_occurrences`; compose message adaptively based on which fields are set.
- [x] 1.3 Update `docx_editor/track_changes.py`: thread paragraph ref and current paragraph text into every `raise TextNotFoundError(...)` site. Unscoped paths pass `paragraph_ref=None`. At the `_get_nth_match` raise site (line ~423), pass `occurrence=` and `total_occurrences=` so the LLM can extract counts without parsing the message string.
- [x] 1.3a Update `docx_editor/comments.py` line 110: change the `raise TextNotFoundError("full message")` to use the new signature (`TextNotFoundError(search_text, paragraph_ref=..., paragraph_preview=...)`).
- [x] 1.4 Run `uv run pytest tests/test_paragraph_hash.py -v` — new tests pass.
- [x] 1.5 Run `uv run pytest` — full suite passes (existing `pytest.raises(TextNotFoundError)` still matches).

## 2. ParagraphIndexError (breaking)

- [x] 2.1 Write failing tests in `tests/test_paragraph_hash.py`: out-of-range ref raises `ParagraphIndexError` with `index`, `total_paragraphs`; message states valid range.
- [x] 2.2 Add `ParagraphIndexError(DocxEditError)` to `docx_editor/exceptions.py` with `__init__(index, total_paragraphs)` and range-stating message.
- [x] 2.3 Update `docx_editor/track_changes.py::_resolve_paragraph`: replace `raise IndexError(...)` with `raise ParagraphIndexError(ref.index, len(paragraphs))`.
- [x] 2.4 Grep `tests/` for `pytest.raises(IndexError)` in paragraph-related tests; update to `pytest.raises(ParagraphIndexError)`.
- [x] 2.5 Run `uv run pytest` — all pass.

## 3. BatchOperationError (breaking, scoped)

- [x] 3.1 Write failing tests in `tests/test_batch_edit.py`: invalid batch op raises `BatchOperationError` with `operation_index` and `reason`.
- [x] 3.2 Add `BatchOperationError(DocxEditError)` to `docx_editor/exceptions.py` with `__init__(operation_index, reason)` and message `f"Operation {operation_index}: {reason}"`.
- [x] 3.3 Update `docx_editor/track_changes.py` batch validation: replace `raise ValueError(f"Operation {i}: ...")` with `raise BatchOperationError(i, ...)`. **Scope-check**: leave non-batch `ValueError` sites untouched.
- [x] 3.4 In `batch_edit`, wrap each `_apply_single_edit(op)` call in `try/except ValueError as e` and re-raise as `BatchOperationError(i, str(e))`. This catches the 4 `ValueError` sites inside `_apply_single_edit` (lines 151, 159, 167, 175 — missing-field checks like `"replace requires 'find' and 'replace_with'"`) that fire exclusively from batch context but don't carry `operation_index`.
- [x] 3.5 Grep `tests/test_batch_edit.py` for `pytest.raises(ValueError)` on batch validation; switch to `pytest.raises(BatchOperationError)`. Leave non-batch `ValueError` tests alone.
- [x] 3.6 Run `uv run pytest` — all pass.

## 4. Public API and documentation

- [x] 4.1 Update `docx_editor/__init__.py`: add `ParagraphIndexError` and `BatchOperationError` to imports and `__all__`.
- [x] 4.2 Add an import smoke test: `from docx_editor import HashMismatchError, TextNotFoundError, ParagraphIndexError, BatchOperationError`.
- [x] 4.3 Update `skills/docx/SKILL.md` error-handling section: add a table listing each of the four structured errors, their instance attributes, and the recovery action.
- [x] 4.4 Run `uv run pytest` and `uv run ruff check . && uv run ruff format --check .` — clean.

## 5. Validation

- [x] 5.1 Run `openspec validate add-structured-error-types --strict --no-interactive` — passes.
- [x] 5.2 Run `uv run pytest` one more time from a clean checkout — green.
