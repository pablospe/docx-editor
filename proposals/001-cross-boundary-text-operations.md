# OpenSpec Proposal: Cross-Boundary Text Operations

**Status:** Draft
**Author:** Pablo
**Created:** 2026-01-28
**Issue:** [#1](https://github.com/pablospe/docx-editor/issues/1)

## Problem Statement

The library cannot find or edit text that spans across tracked change boundaries (insertions/deletions). This is a fundamental limitation reported by users working with documents containing existing revisions.

### Example

Given this document structure:
```xml
<w:r><w:t>Exploratory Aim: </w:t></w:r>
<w:ins w:id="1" w:author="Alice">
  <w:r><w:t>To examine whether...</w:t></w:r>
</w:ins>
```

| Search Query | Current Result |
|--------------|----------------|
| `"Exploratory Aim:"` | ✓ Found |
| `"To examine whether"` | ✓ Found |
| `"Exploratory Aim: To"` | ✗ NOT FOUND |

### Root Cause

The current implementation in `track_changes.py` searches within individual `<w:t>` elements:

```python
def _get_nth_match(self, text: str, occurrence: int):
    matches = self.editor.find_all_nodes(tag="w:t", contains=text)
```

This design assumes the entire search text exists within a **single `<w:t>` element**. It does not handle text spanning multiple runs or revision boundaries.

### Impact

- ✓ Works on documents with **no existing tracked changes**
- ✓ Works when edits target text **fully inside or fully outside** revision blocks
- ✗ Fails silently when edits **span boundaries**

## Questions to Resolve

### Q1: What should happen when replacing text that spans an existing insertion?

**Scenario:** Document has `"Hello <w:ins>world</w:ins>"`, user wants to replace `"Hello world"` with `"Hi there"`.

| Option | Behavior | Complexity |
|--------|----------|------------|
| A. Reject operation | Return error, require user to accept/reject existing changes first | Low |
| B. Accept-then-replace | Implicitly accept the insertion, then apply the replacement | Medium |
| C. Nested revisions | Delete original text + delete the insertion + insert new text | High |

**Recommendation:** Option A for v1 (explicit is better than implicit), with Option C as future enhancement.

### Q2: Should deleted text (`<w:delText>`) be included in visible text?

Word excludes deleted text from visible content. We should match this behavior.

**Recommendation:** Exclude `<w:delText>` from the flattened text view.

### Q3: How to handle replacement text placement?

When replacing `"Aim: To"` which spans a boundary:

```
Before: "Exploratory Aim: " + <ins>"To examine"</ins>
After:  "Exploratory " + <del>"Aim: "</del> + <ins>"Goal: "</ins> + <ins>"To examine"</ins>
```

**Recommendation:** Insert new text at the position of the first character being replaced.

## Proposed Solution

### Phase 1: Virtual Text Map (Core Infrastructure)

Create a flattened text representation that maps character positions back to source XML nodes.

```python
@dataclass
class TextPosition:
    """Maps a character position to its source XML node."""
    node: Element          # The <w:t> element
    offset_in_node: int    # Character offset within the node
    is_inside_ins: bool    # True if inside <w:ins>
    is_inside_del: bool    # True if inside <w:del> (should be excluded)

@dataclass
class TextMap:
    """Flattened text view of a paragraph."""
    text: str                      # Concatenated visible text
    positions: list[TextPosition]  # One per character in text

    def find(self, search: str, start: int = 0) -> int | None:
        """Find search string, return start index or None."""
        idx = self.text.find(search, start)
        return idx if idx >= 0 else None

    def get_nodes_for_range(self, start: int, end: int) -> list[TextPosition]:
        """Get all positions for a character range."""
        return self.positions[start:end]
```

**New method in `DocxXMLEditor`:**

```python
def build_text_map(self, paragraph: Element) -> TextMap:
    """Build a virtual text map for a paragraph.

    Concatenates all <w:t> text (excluding <w:delText>) and tracks
    the source node for each character position.
    """
```

### Phase 2: Cross-Boundary Search

Update `_get_nth_match()` to use the text map:

```python
def _get_nth_match_v2(self, text: str, occurrence: int) -> TextMapMatch | None:
    """Find text across element boundaries.

    Returns:
        TextMapMatch with start/end positions and affected nodes,
        or None if not found.
    """
    count = 0
    for para in self.editor.find_all_nodes(tag="w:p"):
        text_map = self.editor.build_text_map(para)
        start = 0
        while (idx := text_map.find(text, start)) is not None:
            count += 1
            if count == occurrence:
                return TextMapMatch(
                    paragraph=para,
                    text_map=text_map,
                    start=idx,
                    end=idx + len(text),
                )
            start = idx + 1
    return None
```

### Phase 3: Boundary-Aware Operations

#### Simple Case: Text Within Single Element

No change from current behavior.

#### Complex Case: Text Spans Multiple Elements

```python
def replace_text_v2(self, old: str, new: str, occurrence: int = 1) -> Revision:
    match = self._get_nth_match_v2(old, occurrence)
    if match is None:
        raise TextNotFoundError(old)

    nodes = match.get_affected_nodes()

    # Check if any nodes are inside existing tracked changes
    if any(n.is_inside_ins or n.is_inside_del for n in nodes):
        raise RevisionBoundaryError(
            f"Text '{old}' spans existing tracked changes. "
            "Accept or reject those changes first, or use replace_text_force()."
        )

    # Split and replace across nodes
    return self._replace_across_nodes(match, new)
```

**Node splitting logic:**

```python
def _replace_across_nodes(self, match: TextMapMatch, new_text: str) -> Revision:
    """Replace text that spans multiple <w:t> nodes.

    Strategy:
    1. Split first node: keep text before match, delete match portion
    2. Delete entire middle nodes
    3. Split last node: delete match portion, keep text after
    4. Insert new text at match start position
    """
```

## Implementation Plan

### Milestone 1: Read-Only Text Map (Low Risk)

1. Implement `TextPosition` and `TextMap` dataclasses
2. Implement `build_text_map()` in `DocxXMLEditor`
3. Add `get_visible_text()` method to `Document` class
4. Add comprehensive tests with revision-heavy documents

**Deliverable:** Users can get flattened text view for analysis.

### Milestone 2: Cross-Boundary Search (Medium Risk)

1. Implement `TextMapMatch` dataclass
2. Implement `_get_nth_match_v2()`
3. Add `find_text()` method that returns match info including boundary status
4. Tests for searching across all boundary types

**Deliverable:** Users can search across boundaries and get detailed match info.

### Milestone 3: Boundary-Aware Replace (Higher Risk)

1. Implement `_replace_across_nodes()` with node splitting
2. Add `RevisionBoundaryError` exception
3. Update `replace_text()` to use new implementation
4. Add `replace_text_force()` for users who want implicit accept behavior
5. Extensive tests with real-world documents

**Deliverable:** Full cross-boundary text replacement.

## Risks and Mitigations

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| Complex nested revision structures break assumptions | Medium | High | Comprehensive test suite with real documents; fail explicitly rather than corrupt |
| Performance degradation on large documents | Low | Medium | Build text maps lazily per-paragraph; cache if needed |
| Breaking existing API behavior | Medium | High | New methods (`replace_text_v2`) during transition; deprecation warnings |
| Edge cases in XML namespace handling | Low | Medium | Reuse existing `_parse_fragment()` infrastructure |

## Alternatives Considered

### A. Run Normalizer (Rejected)

Merge adjacent `<w:r>` elements before searching.

**Why rejected:** Only fixes formatting splits, not revision boundaries. Doesn't solve the core problem.

### B. Regex-Based XML Search (Rejected)

Search the raw XML string with regex.

**Why rejected:** Fragile, namespace issues, can't handle CDATA or entity encoding properly.

### C. Require Pre-Acceptance of Changes (Rejected)

Force users to accept all tracked changes before editing.

**Why rejected:** Destroys revision history, unacceptable for legal/enterprise workflows.

## Success Criteria

1. All existing tests continue to pass
2. New test: replace text spanning `<w:ins>` boundary with explicit error
3. New test: `get_visible_text()` returns correct flattened text
4. New test: search finds text spanning multiple `<w:t>` elements
5. Real-world document from issue #1 can be processed without silent failures

## Open Questions

1. Should we support cross-paragraph operations? (Tentative: No, too complex for v1)
2. Should `get_visible_text()` include field codes, bookmarks? (Tentative: No, text only)
3. Performance target for text map building? (Need benchmarks)

## Conclusion

**Are these problems solvable? Yes.**

The fundamental challenge is bridging OOXML's structural representation with users' mental model of continuous text. The proposed "Virtual Text Map" approach is:

- **Proven:** This is how Word itself handles search/replace
- **Incremental:** Can be implemented in phases without breaking existing functionality
- **Explicit:** Fails loudly on edge cases rather than corrupting documents

The main complexity is in Phase 3 (boundary-aware replace), but Phases 1-2 provide significant value on their own and can be shipped independently.

**Estimated scope:** Medium-sized feature, primarily touching `xml_editor.py` and `track_changes.py`.
