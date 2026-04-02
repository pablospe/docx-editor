## Context

The docx-editor library currently requires LLMs to specify exact search text when editing paragraphs. This is brittle: ambiguous matches, occurrence counting, and multi-step search-replace sequences are common failure modes. LLMs are better at producing desired output than crafting search queries. We need a method that accepts the desired paragraph text and automatically generates tracked changes.

This builds on the hash-anchored paragraph references (`add-paragraph-hash-anchors`) and batch edit infrastructure (`add-batch-edit`).

## Goals / Non-Goals

**Goals:**
- Let the LLM specify what a paragraph should say, not how to transform it
- Generate fine-grained `<w:del>` + `<w:ins>` tracked changes (not paragraph-level replace)
- Preserve existing formatting where text is unchanged
- Support batch rewriting of multiple paragraphs from one snapshot

**Non-Goals:**
- Formatting changes (bold, italic, etc.) — only text content changes
- Structural changes (splitting/merging paragraphs, adding/removing paragraphs)
- Sub-paragraph targeting — the method always rewrites the entire paragraph text
- Smart formatting inference for newly inserted text beyond inheriting adjacent run properties

## Decisions

### Diff Algorithm: `difflib.SequenceMatcher` with Word-Level Tokens

- **Decision**: Tokenize both old and new text into words and whitespace, then use `difflib.SequenceMatcher` to compute the minimal diff.
- **Why**: `difflib` is Python stdlib, well-tested, and `SequenceMatcher` produces human-readable diffs. Word-level diffing produces natural tracked changes (whole words are inserted/deleted, not individual characters).
- **Tokenization**: Split on word boundaries using `re.findall(r'\S+|\s+', text)` to preserve whitespace tokens. This ensures whitespace changes are tracked accurately.
- **Alternative considered**: Character-level diffing — produces noisy tracked changes where partial words are marked as changed. Word-level is more natural for document review.
- **Alternative considered**: `google-diff-match-patch` — external dependency, character-level by default. Not worth the dependency for this use case.
- **Alternative considered**: Line-level diffing — too coarse for paragraph editing where sentences may be reworded.

### Mapping Diff Hunks to XML Positions

- **Decision**: Use `build_text_map()` to get the paragraph's visible text with character-to-XML mappings. For each diff hunk that removes old text, locate the corresponding character range in the text map and apply `<w:del>` operations. For each hunk that adds new text, insert `<w:ins>` at the appropriate XML position.
- **How it works**:
  1. Call `build_text_map(paragraph)` to get `TextMap` with visible text and character positions
  2. Run `SequenceMatcher` on word-tokenized old text vs new text
  3. For `'equal'` opcodes: skip (text unchanged, XML untouched)
  4. For `'delete'` opcodes: find the character range in the text map, apply deletion (same as existing `delete_text` path)
  5. For `'insert'` opcodes: find the insertion point in the text map, insert new `<w:ins>` element
  6. For `'replace'` opcodes: combine delete + insert at the same position
- **Why**: This reuses the existing `build_text_map()` infrastructure. The text map already handles cross-boundary spans, tracked changes, and element splitting.
- **Application order**: Process hunks in reverse document order (right-to-left) so that earlier positions remain valid as modifications are applied.

### Formatting Preservation

- **Decision**: When inserting new text via `<w:ins>`, copy the `<w:rPr>` (run properties) from the nearest adjacent run at the insertion point.
- **Why**: If the user changes "Hello world" to "Hello beautiful world", the word "beautiful" should inherit the formatting of the surrounding text. This is the same behavior Word uses when typing new text.
- **Edge case**: If the paragraph has no runs (empty paragraph being filled), no `<w:rPr>` is applied (default formatting).

### Edge Cases

#### Empty Paragraphs
- Rewriting an empty paragraph inserts all new text as a single `<w:ins>` run.
- The text map for an empty paragraph is empty string `""`, so the diff is a single insert hunk.

#### Rewrite to Empty
- Rewriting a paragraph to empty text (`new_text=""`) deletes all visible text via `<w:del>` elements.
- The paragraph element itself is preserved (it's still a `<w:p>`, just with no visible text).

#### Paragraphs with Existing Tracked Changes
- `build_text_map()` already handles paragraphs with `<w:ins>` and `<w:del>` elements.
- Deletions of text that is inside an existing `<w:ins>` will undo the insertion (remove from `<w:ins>`) rather than creating nested `<w:del>` inside `<w:ins>`.
- This follows the same logic as the existing `replace_text` path for mixed-state editing.

#### No-Op Rewrite
- If `new_text` equals the current visible text, `SequenceMatcher` produces all `'equal'` opcodes.
- The method returns without modifying the XML. No tracked changes are generated.

### `batch_rewrite()` Design

- **Decision**: Follow the same pattern as `batch_edit()` from `add-batch-edit`:
  1. Accept a list of `(ref, new_text)` pairs
  2. Validate all paragraph hashes upfront — reject entire batch if any mismatch
  3. Apply rewrites in reverse paragraph order (highest index first)
- **Why**: Consistent with existing batch semantics. All refs are validated against the pre-edit state.
- **Multiple rewrites to same paragraph**: Not supported in a single batch (each paragraph can appear at most once). Unlike `batch_edit()` where multiple find-replace operations within a paragraph are meaningful, a rewrite specifies the complete final text — a second rewrite would overwrite the first.

## Risks / Trade-offs

- **Risk**: Word-level diffing may produce suboptimal hunks for heavily rewritten text (e.g., entire sentence reworded) -> **Mitigation**: `SequenceMatcher` uses a heuristic for "junk" detection that handles this reasonably. For extreme rewrites, the result is equivalent to delete-all + insert-all, which is correct if not minimal.
- **Risk**: Formatting inheritance from adjacent runs may not always match user expectations -> **Mitigation**: This is the same behavior as typing in Word. For formatting-aware edits, users should use `replace()` with explicit formatting parameters (future work).
- **Risk**: Large paragraphs may produce many small hunks -> **Mitigation**: `SequenceMatcher` naturally coalesces equal regions. The number of hunks is proportional to the number of changes, not the paragraph length.

## Open Questions

- Should `rewrite_paragraph()` return a summary of changes made (e.g., number of insertions/deletions)? This could help LLMs verify their edits were applied as expected.
