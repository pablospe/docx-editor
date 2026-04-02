# Change: Add Paragraph Rewrite via Automatic Diffing

## Why

Current edit methods (`replace`, `delete`) require the LLM to specify exact search text and handle occurrence disambiguation. This is error-prone: the LLM must craft precise `find` strings, deal with occurrence counts, and hope the search text is unambiguous. LLMs are much better at producing desired output text than constructing search-and-replace queries. A `rewrite_paragraph()` method lets the LLM simply state what the paragraph should say, and the system automatically diffs old vs new text to generate fine-grained tracked changes.

### Track Changes Quality is the Priority

The ultimate goal of this project is producing **high-quality tracked changes** that look natural when reviewed in Microsoft Word. This is not just about convenience — tracked changes are the core value proposition. Any new editing method must produce track changes that are at least as clean as the surgical methods.

`rewrite_paragraph()` uses word-level diffing (`difflib.SequenceMatcher`), which produces natural tracked changes for most edits — whole words are inserted/deleted, matching how a human reviewer would see changes in Word. For simple single-word swaps, the diff output is identical to what `replace()` produces (one `<w:del>` + one `<w:ins>`). For heavily rewritten text, the diff may produce more hunks than a hand-crafted sequence of surgical edits would, but each hunk is still a valid, reviewable tracked change.

### When to Use Rewrite vs Surgical Methods

Both methods produce proper tracked changes. The choice depends on the edit:

**Use `replace()` / `delete()` / `insert_after()` / `insert_before()` when:**
- Making a single, precise change (e.g., "30" → "60", delete one phrase)
- The change is simple and unambiguous — no occurrence confusion
- Token efficiency matters (only the changed word is sent, not the whole paragraph)
- You want guaranteed minimal track changes (exactly one `<w:del>` + `<w:ins>`)

**Use `rewrite_paragraph()` when:**
- Restructuring a sentence or rewriting multiple parts of a paragraph at once
- The search text would be ambiguous or hard to specify precisely
- You'd otherwise need to chain 3+ surgical edits on the same paragraph
- The paragraph is short enough that outputting the full text is not wasteful

A single-word change via `rewrite_paragraph()` produces identical track changes to `replace()` — it's a degenerate case. The difference is only in token cost (rewrite sends the full paragraph text).

## What Changes

- New `rewrite_paragraph(ref, new_text)` method on `Document` that accepts a hash-anchored paragraph reference and the complete desired paragraph text
- Internal word-level diffing using `difflib.SequenceMatcher` to compare old visible text vs new text
- Automatic mapping of diff hunks back to XML positions using existing `build_text_map()` / `find_in_text_map()`
- Fine-grained `<w:del>` + `<w:ins>` generation for each changed segment (unchanged text is preserved as-is)
- New `batch_rewrite()` method on `Document` for rewriting multiple paragraphs from a single snapshot
- Formatting preservation: new text inserted adjacent to a changed segment inherits the run properties of that segment

## Impact

- Affected specs: `specs/text-operations`
- Affected code: `docx_editor/document.py`, `docx_editor/track_changes.py`, `docx_editor/xml_editor.py`
- Affected docs: `skills/docx/SKILL.md`, `README.md`
- **Non-breaking**: New methods only. Existing `replace()`, `delete()`, `insert_after()`, `insert_before()` are unchanged.
- **Depends on**: `add-paragraph-hash-anchors` (uses ParagraphRef, HashMismatchError, `_resolve_paragraph`, `compute_paragraph_hash`)
- **Depends on**: `add-batch-edit` (batch_rewrite follows the same upfront-validation, reverse-order pattern)
