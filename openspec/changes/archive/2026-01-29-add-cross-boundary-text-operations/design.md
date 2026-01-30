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

### Goals
- Enable searching across element boundaries
- Enable replacing text spanning multiple `<w:t>` elements
- Handle mixed-state editing (text spanning revision boundaries) via atomic decomposition
- Provide read-only flattened text view for analysis

### Non-Goals
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

## Design Constraints

### Run Property Preservation

When splitting runs during cross-boundary operations, each new `<w:r>` fragment MUST carry the `<w:rPr>` from its source run. The current single-element `replace_text` already does this (`track_changes.py:122-142`). The cross-boundary version must apply the same pattern per-segment — each segment uses its own source run's `w:rPr`.

When a match spans runs with different formatting (e.g., bold + italic), each split fragment preserves its original formatting. The new inserted text inherits `w:rPr` from the first run in the match range.

### Word Compatibility

Every test that produces a modified `.docx` MUST include a round-trip validation: reopen the output file and verify it parses without errors. This catches malformed XML, missing namespaces, and invalid nesting that would cause Word to show a repair dialog.

## Testing Strategy

**Tests-first (TDD) approach:** Write failing tests before implementing each phase. Tests serve as the specification and guide the implementation.

Each phase begins by writing unit tests that cover the expected behavior described in the spec scenarios. Implementation follows to make the tests pass.

## Migration Plan

1. **Phase 1:** Add `get_visible_text()` (read-only, no risk)
2. **Phase 2:** Add `find_text()` with boundary info (read-only, no risk)
3. **Phase 3:** Replace across multiple `<w:t>` elements in same revision context
4. **Phase 4:** Mixed-state editing via atomic decomposition (spans revision boundaries)

Rollback: Revert to per-element search if issues found.

## Resolved Questions

1. **`get_visible_text()` paragraph boundaries:** Yes, paragraphs are joined with `\n`. Implemented in `document.py`.
2. **Performance target:** Deferred. Lazy per-paragraph map building keeps memory bounded. Benchmarks can be added when real-world usage reveals bottlenecks.
3. **`find_all_text()`:** Deferred. The `occurrence` parameter on `find_text()` covers the common case. A dedicated `find_all_text()` can be added later if needed.
4. **Deleting text inside `<w:ins>`:** Remove the node (undo the insertion). This produces cleaner history and avoids invalid `<w:del>` nesting inside `<w:ins>`. Implemented in `_remove_from_insertion()`.
