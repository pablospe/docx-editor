# Tasks: Add Cross-Boundary Text Operations

**Approach:** Tests-first (TDD). Each phase starts by writing failing tests, then implementing to make them pass.

## 1. Core Infrastructure

- [x] 1.1 Add unit tests for text map building (plain text, tracked changes, deleted text)
- [x] 1.2 Add `TextPosition` dataclass to `xml_editor.py`
- [x] 1.3 Add `TextMap` dataclass with `find()` and `get_nodes_for_range()` methods
- [x] 1.4 Implement `build_text_map()` in `DocxXMLEditor`

## 2. Public API

- [x] 2.1 Add integration tests for `get_visible_text()`
- [x] 2.2 Add `get_visible_text()` method to `Document` class

## 3. Cross-Boundary Search

- [x] 3.1 Add tests for searching across element and revision boundaries
- [x] 3.2 Implement `TextMapMatch` dataclass
- [x] 3.3 Implement `_find_across_boundaries()` using text maps
- [x] 3.4 Add `find_text()` method returning match info with boundary status

## 4. Boundary-Aware Replace (Same Revision Context)

- [x] 4.1 Add tests for replace within single element (regression)
- [x] 4.2 Add tests for replace across multiple `<w:t>` elements in same context
- [x] 4.3 Implement `_replace_across_nodes()` with node splitting logic
- [x] 4.4 Update `replace_text()` to use cross-boundary search

## 5. Mixed-State Editing (Atomic Decomposition)

- [x] 5.1 Add tests for replace spanning regular text + `<w:ins>` boundary
- [x] 5.2 Add tests for replace spanning `<w:ins>` + regular text boundary
- [x] 5.3 Add tests for replace fully within `<w:ins>` (regression)
- [x] 5.4 Add tests for `<w:ins>` node splitting (partial match within insertion)
- [x] 5.5 Add tests for `w:rPr` preservation across split runs with different formatting
- [x] 5.6 Add tests for output validity — round-trip reopen modified `.docx` to verify well-formed XML
- [x] 5.7 Implement segment decomposition — classify match ranges by revision context (regular, inside `<w:ins>`, inside `<w:del>`)
- [x] 5.8 Implement `<w:ins>` node splitting — isolate target text from remaining insertion
- [x] 5.9 Implement per-segment delete strategies (wrap in `<w:del>` for regular text, remove node for inserted text)
- [x] 5.10 Integrate atomic decomposition into `replace_text()` flow

## 6. Documentation

- [x] 6.1 Update README with new capabilities
- [x] 6.2 Add docstrings to new public methods
- [x] 6.3 Add usage examples in docs/
