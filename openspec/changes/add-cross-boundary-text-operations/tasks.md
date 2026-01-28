# Tasks: Add Cross-Boundary Text Operations

**Approach:** Tests-first (TDD). Each phase starts by writing failing tests, then implementing to make them pass.

## 1. Core Infrastructure

- [ ] 1.1 Add unit tests for text map building (plain text, tracked changes, deleted text)
- [ ] 1.2 Add `TextPosition` dataclass to `xml_editor.py`
- [ ] 1.3 Add `TextMap` dataclass with `find()` and `get_nodes_for_range()` methods
- [ ] 1.4 Implement `build_text_map()` in `DocxXMLEditor`

## 2. Public API

- [ ] 2.1 Add integration tests for `get_visible_text()`
- [ ] 2.2 Add `get_visible_text()` method to `Document` class

## 3. Cross-Boundary Search

- [ ] 3.1 Add tests for searching across element and revision boundaries
- [ ] 3.2 Implement `TextMapMatch` dataclass
- [ ] 3.3 Implement `_get_nth_match_v2()` using text maps
- [ ] 3.4 Add `find_text()` method returning match info with boundary status

## 4. Boundary-Aware Replace (Same Revision Context)

- [ ] 4.1 Add tests for replace within single element (regression)
- [ ] 4.2 Add tests for replace across multiple `<w:t>` elements in same context
- [ ] 4.3 Implement `_replace_across_nodes()` with node splitting logic
- [ ] 4.4 Update `replace_text()` to use cross-boundary search

## 5. Mixed-State Editing (Atomic Decomposition)

- [ ] 5.1 Add tests for replace spanning regular text + `<w:ins>` boundary
- [ ] 5.2 Add tests for replace spanning `<w:ins>` + regular text boundary
- [ ] 5.3 Add tests for replace fully within `<w:ins>` (regression)
- [ ] 5.4 Add tests for `<w:ins>` node splitting (partial match within insertion)
- [ ] 5.5 Add tests for `w:rPr` preservation across split runs with different formatting
- [ ] 5.6 Add tests for output validity — round-trip reopen modified `.docx` to verify well-formed XML
- [ ] 5.7 Implement segment decomposition — classify match ranges by revision context (regular, inside `<w:ins>`, inside `<w:del>`)
- [ ] 5.8 Implement `<w:ins>` node splitting — isolate target text from remaining insertion
- [ ] 5.9 Implement per-segment delete strategies (wrap in `<w:del>` for regular text, remove node for inserted text)
- [ ] 5.10 Integrate atomic decomposition into `replace_text()` flow

## 6. Documentation

- [ ] 6.1 Update README with new capabilities
- [ ] 6.2 Add docstrings to new public methods
- [ ] 6.3 Add usage examples in docs/
