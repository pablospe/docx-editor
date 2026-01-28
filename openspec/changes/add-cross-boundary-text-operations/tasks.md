# Tasks: Add Cross-Boundary Text Operations

## 1. Core Infrastructure

- [ ] 1.1 Add `TextPosition` dataclass to `xml_editor.py`
- [ ] 1.2 Add `TextMap` dataclass with `find()` and `get_nodes_for_range()` methods
- [ ] 1.3 Implement `build_text_map()` in `DocxXMLEditor`
- [ ] 1.4 Add unit tests for text map building with various document structures

## 2. Public API

- [ ] 2.1 Add `get_visible_text()` method to `Document` class
- [ ] 2.2 Add `RevisionBoundaryError` to `exceptions.py`
- [ ] 2.3 Add integration tests for `get_visible_text()`

## 3. Cross-Boundary Search

- [ ] 3.1 Implement `TextMapMatch` dataclass
- [ ] 3.2 Implement `_get_nth_match_v2()` using text maps
- [ ] 3.3 Add `find_text()` method returning match info with boundary status
- [ ] 3.4 Add tests for searching across revision boundaries

## 4. Boundary-Aware Replace

- [ ] 4.1 Implement `_replace_across_nodes()` with node splitting logic
- [ ] 4.2 Update `replace_text()` to use cross-boundary search
- [ ] 4.3 Add boundary detection and `RevisionBoundaryError` raising
- [ ] 4.4 Add tests for replace within single element (regression)
- [ ] 4.5 Add tests for replace across multiple `<w:t>` elements
- [ ] 4.6 Add tests for explicit error when spanning `<w:ins>`/`<w:del>`

## 5. Documentation

- [ ] 5.1 Update README with new capabilities
- [ ] 5.2 Add docstrings to new public methods
- [ ] 5.3 Add usage examples in docs/
