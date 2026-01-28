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
- Fail explicitly when operations cross revision boundaries
- Provide read-only flattened text view for analysis

**Non-Goals:**
- Cross-paragraph operations (too complex for v1)
- Implicit acceptance of existing revisions
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

### Decision 2: Explicit Failure on Revision Boundaries

**What:** Raise `RevisionBoundaryError` when replace operation spans existing `<w:ins>` or `<w:del>`.

**Alternatives considered:**
| Option | Pros | Cons |
|--------|------|------|
| A. Reject (chosen) | Explicit, safe, preserves history | User must handle edge case |
| B. Implicit accept | "Just works" | Destroys revision history |
| C. Nested revisions | Full fidelity | Very complex, edge cases |

**Why A:** Explicit is better than implicit. Users working with legal/enterprise documents need control over revision history. Option C could be added later.

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
| Complex nested revision structures | Explicit failure rather than corruption |
| Breaking existing behavior | New methods; existing API unchanged until validated |

## Migration Plan

1. **Phase 1:** Add `get_visible_text()` (read-only, no risk)
2. **Phase 2:** Add `find_text()` with boundary info (read-only, no risk)
3. **Phase 3:** Update `replace_text()` internals (risk: behavior change)
   - Add feature flag or new method name during transition
   - Deprecation warnings if behavior would differ

Rollback: Revert to per-element search if issues found.

## Open Questions

1. Should `get_visible_text()` return paragraph boundaries (e.g., `\n`)?
2. Performance target for text map building? (Need benchmarks)
3. Should we expose `find_all_text()` for multiple matches?
