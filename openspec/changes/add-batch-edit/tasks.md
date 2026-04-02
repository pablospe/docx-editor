# Tasks: Add Batch Edit with Reverse-Order Application

## 1. Core Implementation

- [x] 1.1 Add `EditOperation` dataclass with fields: `action` (replace/delete/insert_after/insert_before), `paragraph` (required), `find`/`text`/`anchor`/`replace_with` as appropriate, `occurrence` (default 0)
- [x] 1.2 Add `_validate_batch()` to `RevisionManager` — parse all ParagraphRefs, resolve all paragraphs and validate hashes upfront; raise `HashMismatchError` on first mismatch (no edits applied)
- [x] 1.3 Add `_apply_batch()` to `RevisionManager` — sort edits by paragraph index descending, apply each edit sequentially
- [x] 1.4 Add `batch_edit(operations: list[EditOperation]) -> list[int]` to `Document` — validate then apply, return list of change IDs

## 2. Tests

- [x] 2.1 Test: batch of 5 edits to different paragraphs — all succeed, all change IDs returned
- [x] 2.2 Test: batch with one stale hash — entire batch rejected, no edits applied
- [x] 2.3 Test: batch applied in reverse order — edits to P20 don't invalidate P5's hash
- [x] 2.4 Test: single `list_paragraphs()` call suffices for entire batch
- [x] 2.5 Test: batch with duplicate paragraph targets — both edits apply (same paragraph, different text)
- [x] 2.6 Test: empty batch returns empty list

## 3. Benchmark Update

- [ ] 3.1 Add batch mode to `benchmarks/hash_anchored_vs_plain.py` — compare N individual calls vs single batch_edit call
- [ ] 3.2 Measure: time for 10 edits (10x individual vs 1x batch)
