# Tasks: Add Hash-Anchored Paragraph References

## 1. Core Infrastructure

- [x] 1.1 Add `HashMismatchError` to `exceptions.py` with fields: `paragraph_index`, `expected_hash`, `actual_hash`, `paragraph_preview`
- [x] 1.2 Add `ParagraphRef` dataclass to `xml_editor.py` with `index: int`, `hash: str`, and `parse(ref: str) -> ParagraphRef` classmethod (validates `P{n}#{hash}` format)
- [x] 1.3 Add `compute_paragraph_hash(paragraph_element) -> str` function to `xml_editor.py` that computes `zlib.crc32` of `build_text_map(paragraph).text`, returns 4-char lowercase hex
- [x] 1.4 Export new types from `__init__.py`: `ParagraphRef`, `HashMismatchError`
- [x] 1.5 Write tests for `ParagraphRef.parse()` — valid refs, invalid format, edge cases
- [x] 1.6 Write tests for `compute_paragraph_hash()` — normal paragraph, empty paragraph, paragraph with tracked changes

## 2. Paragraph Listing

- [x] 2.1 Add `list_paragraphs(max_chars: int = 80) -> list[str]` to `Document` that returns `["P1#a7b2| Introduction to the...", "P2#f3c1| The committee has...", ...]`
- [x] 2.2 Write tests for `list_paragraphs()` — multi-paragraph doc, empty paragraphs, paragraphs with tracked changes, truncation behavior

## 3. Paragraph Resolution in RevisionManager

- [x] 3.1 Add `_resolve_paragraph(self, ref: ParagraphRef) -> Element` to `RevisionManager` — gets all `<w:p>` elements, validates index is in range, recomputes hash, raises `HashMismatchError` on mismatch, returns the `<w:p>` element
- [x] 3.2 Write tests for `_resolve_paragraph()` — valid ref, index out of range, hash mismatch (with helpful error message), hash after edit changes

## 4. Scoped Text Search

- [x] 4.1 Add `paragraph: str | None = None` parameter to `RevisionManager.replace_text()` — when set, parse ref, resolve paragraph, search only within that paragraph using `build_text_map` + `find_in_text_map` with paragraph-local occurrence
- [x] 4.2 Add `paragraph: str | None = None` parameter to `RevisionManager.suggest_deletion()`— same scoping logic
- [x] 4.3 Add `paragraph: str | None = None` parameter to `RevisionManager.insert_text_after()` and `insert_text_before()` — same scoping logic
- [x] 4.4 Write tests for scoped replace — find text in specific paragraph, occurrence is paragraph-local, text exists in other paragraphs but not targeted one
- [x] 4.5 Write tests for scoped delete — same patterns as 4.4
- [x] 4.6 Write tests for scoped insert_after/insert_before — same patterns

## 5. Document API

- [x] 5.1 Add `paragraph: str | None = None` parameter to `Document.replace()`, `Document.delete()`, `Document.insert_after()`, `Document.insert_before()` — pass through to RevisionManager
- [x] 5.2 Write integration tests for the full flow: `list_paragraphs()` → parse ref → `replace(paragraph="P2#f3c1")` → verify edit landed in correct paragraph

## 6. Staleness Detection Tests

- [x] 6.1 Test: edit paragraph 2, then try to use old hash for paragraph 2 → `HashMismatchError`
- [x] 6.2 Test: insert paragraph above, old P2 is now P3, using old `P2#hash` → `HashMismatchError` (different content at P2 now)
- [x] 6.3 Test: `HashMismatchError` message includes current hash so LLM can retry
- [x] 6.4 Test: multiple edits in sequence, each using fresh refs from `list_paragraphs()` after prior edit → all succeed

## 7. Benchmark: Hash-Anchored vs Plain Edits

- [ ] 7.1 Create benchmark script (`benchmarks/hash_anchored_vs_plain.py`) that compares both approaches on a multi-paragraph document:
  - **Speed**: Time per operation with and without `paragraph=` parameter (measure hash computation + resolution overhead)
  - **Accuracy**: Run a batch of N edits (e.g., 20 targeted replacements across a 50-paragraph doc) using both approaches and count how many land on the correct paragraph
- [ ] 7.2 Design accuracy test scenario: generate a document with repeated phrases across paragraphs, issue a sequence of edits that shift indices mid-batch, measure:
  - Plain `occurrence`-based: how many edits hit wrong paragraph after shifts
  - Hash-anchored: how many raise `HashMismatchError` (correctly rejected) vs silently wrong
- [ ] 7.3 Document benchmark results in `benchmarks/README.md` with a comparison table
