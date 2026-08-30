# Change: Add Hash-Anchored Paragraph References

## Why

When an LLM issues multiple edits in a single turn, earlier edits can shift paragraph indices, causing later edits to target the wrong paragraph or silently apply to stale content. The `occurrence` parameter is fragile for the same reason — it counts globally across all paragraphs and any insertion/deletion changes the count for subsequent operations.

Inspired by the "hash-anchored edit" approach (oh-my-openagent), we add content-derived hashes to paragraph references so that:
1. Edits can be **scoped to a specific paragraph** (disambiguation)
2. Edits against **stale content are rejected** before corrupting the document (staleness detection)
3. Multiple edits in one batch can all reference the **original document state** safely

## What Changes

- New `list_paragraphs()` method on `Document` that returns hash-tagged paragraph previews
- New `paragraph` parameter on `replace()`, `delete()`, `insert_after()`, `insert_before()` that accepts a hash-anchored reference like `"P3#a7b2"`
- New `ParagraphRef` dataclass to parse and validate `P{index}#{hash}` references
- New `HashMismatchError` exception when a paragraph's content has changed since the LLM last read it
- Hash computation utility using the paragraph's visible text (from `build_text_map`)
- The `occurrence` parameter becomes **paragraph-local** when `paragraph` is specified

## Impact

- Affected specs: `specs/text-operations`
- Affected code: `docx_editor/document.py`, `docx_editor/track_changes.py`, `docx_editor/xml_editor.py`, `docx_editor/exceptions.py`
- **Breaking**: `paragraph` parameter is now required on `replace()`, `delete()`, `insert_after()`, `insert_before()`.
