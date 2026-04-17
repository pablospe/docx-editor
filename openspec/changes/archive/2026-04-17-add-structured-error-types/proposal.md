## Why

LLM agents using docx-editor cannot recover from most tool errors in-loop because the errors are opaque strings. `HashMismatchError` already shows the working pattern — it carries `paragraph_index`, `expected_hash`, `actual_hash`, and `paragraph_preview` so the caller can retry without re-reading the document. The remaining LLM-facing errors (`TextNotFoundError`, stdlib `IndexError` for bad paragraph indices, stdlib `ValueError` for batch validation) force an external feedback loop or a full document re-read. ARIA Phase 0B' experiments confirmed that pushing diagnostics into error types delivers L3-grade recovery at L2 cost.

## What Changes

- Add structured fields to `TextNotFoundError`: `search_text`, `paragraph_ref`, `paragraph_preview`. Message embeds current paragraph content when scoped.
- Introduce `ParagraphIndexError(DocxEditError)` with fields `index`, `total_paragraphs`. **BREAKING**: replaces stdlib `IndexError` raised from `_resolve_paragraph()`.
- Introduce `BatchOperationError(DocxEditError)` with fields `operation_index`, `reason`. **BREAKING** (partial): replaces stdlib `ValueError` only on batch validation paths (`batch_edit`, `batch_rewrite`). Non-batch `ValueError` paths are unchanged.
- Re-export `ParagraphIndexError` and `BatchOperationError` from `docx_editor/__init__.py`.
- Document structured error fields and recovery patterns in `skills/docx/SKILL.md`.

## Capabilities

### New Capabilities
- `structured-errors`: Contract for LLM-facing exceptions — which errors exist, what fields they carry, what message format callers can rely on for recovery.

### Modified Capabilities
<!-- None. text-operations specifies search/replace mechanics, not error surface. -->

## Impact

- Code: `docx_editor/exceptions.py`, `docx_editor/track_changes.py` (raise sites in `_resolve_paragraph`, `_resolve_text`, batch validation), `docx_editor/comments.py` (`TextNotFoundError` raise site at line 110), `docx_editor/xml_editor.py` (`ParagraphRef.parse` error path), `docx_editor/__init__.py`.
- Tests: `tests/test_paragraph_hash.py`, `tests/test_batch_edit.py`.
- Docs: `skills/docx/SKILL.md` error-handling section.
- API consumers: `except IndexError` on bad paragraph refs and `except ValueError` on batch validation will no longer match. Both new exceptions inherit from `DocxEditError`, so catch-all handlers still work. Note in CHANGELOG/PR.
- Dependencies: None.
