# Change: Add Batch Edit with Reverse-Order Application

## Why

When an LLM issues multiple edits in a single turn, it must currently call `list_paragraphs()` after each edit to get fresh hashes. This creates N round-trips for N edits. Since intra-paragraph edits don't affect other paragraphs' indices or content, edits to different paragraphs are independent and can be validated upfront and applied in reverse paragraph order from a single snapshot.

## What Changes

- New `batch_edit()` method on `Document` that accepts a list of edit operations
- All paragraph hashes validated upfront — if any mismatch, the entire batch is rejected before any edits are applied (atomic all-or-nothing)
- Edits applied in reverse paragraph order so earlier paragraphs' hashes remain valid throughout
- Requires `paragraph` on every edit in the batch (batch mode is inherently paragraph-scoped)

## Impact

- Affected specs: `specs/text-operations`
- Affected code: `docx_editor/document.py`, `docx_editor/track_changes.py`
- **Non-breaking**: New method, existing API unchanged
- **Depends on**: `add-paragraph-hash-anchors` (uses ParagraphRef, HashMismatchError, _resolve_paragraph)
