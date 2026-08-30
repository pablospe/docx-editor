# Design: Hash-Anchored Paragraph References

## Context

The docx-editor library is used by LLMs to edit Word documents via text-based search. When an LLM issues multiple edits in one turn, paragraph indices shift and occurrence counts become stale. We need a mechanism to detect this and reject stale edits rather than silently corrupting the document.

### Inspiration: oh-my-openagent Hash-Anchored Edit Tool

OmO tags every **line** of source code with `{line}#{2-char-hash}|{content}`, using xxHash32 mod 256 mapped to a 16-char alphabet. Edits reference these tags; the system recomputes hashes before applying and rejects mismatches.

For docx, the unit is **paragraphs** (not lines), and the hash needs to be computed from visible text (respecting tracked changes — insertions visible, deletions hidden).

## Goals / Non-Goals

**Goals:**
- Detect when an edit targets a paragraph whose content has changed since it was last read
- Scope text search to a single paragraph, eliminating global `occurrence` ambiguity
- Support batched edits that all reference the pre-edit document state
- Keep the API simple — one new optional parameter on existing methods

**Non-Goals:**
- Word-level or character-level anchoring (paragraph granularity is sufficient for docx)
- Automatic retry or re-mapping of stale references (the LLM should re-read and retry)
- Changing the internal XML editing mechanics
- Hashing for non-paragraph elements (tables, headers, footers)

## Decisions

### Hash Algorithm: CRC32 truncated to 4 hex chars

- **Decision**: Use `zlib.crc32(text.encode("utf-8")) & 0xFFFF` → 4-char lowercase hex string (e.g., `a7b2`)
- **Why**: Python stdlib (`zlib`), no dependencies. 65,536 buckets is more than enough — collisions are astronomically unlikely within a single document (typically <200 paragraphs). 4 hex chars are short enough for LLMs to handle without wasting tokens.
- **Alternative considered**: xxHash32 (like OmO) — requires `xxhash` dependency. Not worth it for this use case.
- **Alternative considered**: 2-char custom alphabet (like OmO's 256 buckets) — fewer buckets increases collision risk for identical short paragraphs. 4 hex chars are still very compact.
- **Alternative considered**: Full SHA256 — overkill, wastes tokens.

### Hash Input: Visible text from `build_text_map()`

- **Decision**: Hash is computed from `build_text_map(paragraph).text` — the same text the user sees.
- **Why**: This is exactly what `get_visible_text()` already uses. It includes pending insertions and excludes pending deletions. If a tracked change is accepted/rejected, the hash changes, which is the correct behavior (the paragraph content changed).
- **Edge case**: Empty paragraphs get hashed too (empty string → deterministic hash).

### Reference Format: `P{1-indexed}#{4-hex-hash}`

- **Decision**: `P1#a7b2` means "paragraph 1 (1-indexed) with content hash a7b2".
- **Why 1-indexed**: Matches what LLMs see in `list_paragraphs()` output and is more natural for non-programmers. Paragraphs are displayed as P1, P2, P3...
- **Regex**: `^P(\d+)#([0-9a-f]{4})$`

### Batch Edit Strategy: Validate all hashes before applying any edits

- **Decision**: When multiple edits include paragraph refs, validate ALL hashes against current state first. If any mismatch, reject the entire batch.
- **Why**: This matches the "all edits reference original state" contract. Partial application would leave the document in an inconsistent state.
- **Note**: This is a future concern — currently edits are applied one-at-a-time via individual method calls. The validation happens per-call. A future `batch_edit()` method could enforce atomic all-or-nothing semantics, but that's out of scope for this change.

### Where paragraph resolution lives

- **Decision**: Add a `_resolve_paragraph()` method to `RevisionManager` that takes a `ParagraphRef`, finds the `<w:p>` element by index, recomputes the hash, and either returns the element or raises `HashMismatchError`.
- **Why**: `RevisionManager` already owns text search and has access to the editor's DOM. Adding paragraph resolution here keeps text operations in one place.

### Integration with existing search

- **Decision**: When `paragraph` is specified, text search methods (`_get_nth_match`, `_find_across_boundaries`) are scoped to search only within that paragraph. The `occurrence` parameter becomes paragraph-local.
- **Why**: This is the simplest change — we already have `build_text_map(paragraph)` and `find_in_text_map()`. We just need to call them on a specific `<w:p>` instead of iterating all paragraphs.

## Risks / Trade-offs

- **Risk**: Hash collisions between different paragraphs with same hash → **Mitigation**: 65,536 buckets + paragraph index makes this a non-issue. Two paragraphs would need the same index AND the same hash to collide, which can't happen (index is unique).
- **Risk**: LLM forgets to call `list_paragraphs()` and uses wrong hashes → **Mitigation**: `HashMismatchError` gives a clear error message with the current hash, so the LLM can self-correct.
- **Risk**: Performance of recomputing hashes → **Mitigation**: `build_text_map()` is already called during search. Computing CRC32 on the text is negligible.

## Open Questions

- Should `list_paragraphs()` show the full paragraph text or a truncated preview? → **Proposed**: First ~80 chars with `...` if truncated. This keeps context window usage reasonable while giving enough text for the LLM to identify paragraphs.
- Should empty paragraphs (e.g., blank lines between sections) be included in the listing? → **Proposed**: Yes, they're real paragraphs and their indices matter for editing.
