# Design: Cross-Boundary Text Operations

## Context

The library searches for text within individual `<w:t>` elements. When text spans revision boundaries (`<w:ins>`, `<w:del>`), it cannot be found because no single element contains the complete string.

This is how Word represents a document with tracked changes:
```xml
<w:r><w:t>Exploratory Aim: </w:t></w:r>
<w:ins w:id="1" w:author="Alice">
  <w:r><w:t>To examine whether...</w:t></w:r>
</w:ins>
```

Users expect to search for "Aim: To" and find it.

## Goals / Non-Goals

**Goals:**
- Enable searching across element boundaries
- Enable replacing text spanning multiple `<w:t>` elements
- Handle mixed-state editing (text spanning revision boundaries) via atomic decomposition
- Provide read-only flattened text view for analysis

**Non-Goals:**
- Cross-paragraph operations (too complex for v1)
- Field codes, bookmarks, or non-text content

## Decisions

### Decision 1: Virtual Text Map Architecture

**What:** Build a per-paragraph flattened text view with position mapping back to source nodes.

**Why:** This is how Word itself handles search/replace. It's proven, incremental, and doesn't require rewriting the entire XML handling.

**Data structures:**
```python
@dataclass
class TextPosition:
    node: Element          # The <w:t> element
    offset_in_node: int    # Character offset within the node
    is_inside_ins: bool    # Inside <w:ins>?
    is_inside_del: bool    # Inside <w:del>? (excluded from visible text)

@dataclass
class TextMap:
    text: str                      # Concatenated visible text
    positions: list[TextPosition]  # One per character
```

### Decision 2: Atomic Decomposition for Mixed-State Editing

**What:** When a replace operation spans revision boundaries, decompose the match into segments by revision context and apply per-segment operations.

**Alternatives considered:**
| Option | Pros | Cons |
|--------|------|------|
| A. Reject with error | Explicit, safe | Blocks valid use cases |
| B. Implicit accept | "Just works" | Destroys revision history |
| C. Atomic decomposition (chosen) | Handles mixed state, no invalid XML | More complex implementation |

**Why C:** Users editing documents with existing tracked changes need to replace text that spans boundaries. Rejecting these operations is too limiting for real-world workflows (legal redlines, collaborative editing). Atomic decomposition handles this without producing invalid XML.

**Algorithm:**

Given a replace of "Aim: To" where "Aim: " is regular text and "To" is inside `<w:ins>To examine</w:ins>`:

1. **Decompose** — The text map classifies the match into segments:
   - Segment 1: "Aim: " → regular text (Node A)
   - Segment 2: "To" → inserted text (Node B, inside `<w:ins>`)

2. **Delete regular text** (Segment 1) — Standard logic: wrap "Aim: " in `<w:del>`

3. **Delete inserted text** (Segment 2) — Cannot wrap in `<w:del>` (would create invalid `<w:del><w:ins>...</w:ins></w:del>`). Instead:
   - Split the `<w:ins>` element: isolate "To" from " examine"
   - Remove the isolated `<w:ins>To</w:ins>` node entirely (undoing that part of the insertion)
   - The remaining `<w:ins> examine</w:ins>` stays intact

4. **Insert new text** — Place `<w:ins>Goal: </w:ins>` at the split point

**Segment types and their delete strategies:**

| Segment type | Delete strategy |
|-------------|----------------|
| Regular text | Wrap in `<w:del>` (standard) |
| Inside `<w:ins>` | Split insertion, remove target portion (undo partial insertion) |
| Inside `<w:del>` | Skip (already deleted, not in visible text) |

### Decision 3: Per-Paragraph Scope

**What:** Text maps are built per paragraph (`<w:p>`), not document-wide.

**Why:**
- OOXML structures content by paragraph
- Cross-paragraph edits are rare and complex
- Keeps memory bounded on large documents
- Matches Word's search behavior

## Risks / Trade-offs

| Risk | Mitigation |
|------|------------|
| Performance on large documents | Lazy per-paragraph map building; no pre-computation |
| Insertion splitting produces invalid XML | Test extensively with Word; validate output |
| Breaking existing behavior | New methods; existing API unchanged until validated |
| Edge cases in deeply nested revisions | Start with single-level nesting; document unsupported cases |

## Testing Strategy

**Tests-first (TDD) approach:** Write failing tests before implementing each phase. Tests serve as the specification and guide the implementation.

Each phase begins by writing unit tests that cover the expected behavior described in the spec scenarios. Implementation follows to make the tests pass.

## Migration Plan

1. **Phase 1:** Add `get_visible_text()` (read-only, no risk)
2. **Phase 2:** Add `find_text()` with boundary info (read-only, no risk)
3. **Phase 3:** Replace across multiple `<w:t>` elements in same revision context
4. **Phase 4:** Mixed-state editing via atomic decomposition (spans revision boundaries)

Rollback: Revert to per-element search if issues found.

## Open Questions

1. Should `get_visible_text()` return paragraph boundaries (e.g., `\n`)?
2. Performance target for text map building? (Need benchmarks)
3. Should we expose `find_all_text()` for multiple matches?
4. When deleting text inside `<w:ins>`, should we remove the node or wrap in `<w:del>` inside `<w:ins>`? (Word supports both; removing is simpler and preserves cleaner history)
