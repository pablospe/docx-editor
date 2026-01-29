# Change: Add Cross-Boundary Text Operations

## Why

Users cannot find or edit text that spans tracked change boundaries. When a document has existing revisions (insertions/deletions), any text search that crosses these boundaries silently fails. This is the #1 user-reported limitation ([Issue #1](https://github.com/pablospe/docx-editor/issues/1)).

**Example:** Given `"Exploratory Aim: " + <w:ins>"To examine"</w:ins>`, searching for `"Aim: To"` fails because no single `<w:t>` element contains it.

## What Changes

1. **Add Virtual Text Map** - Build flattened text view that maps character positions back to source XML nodes
2. **Add cross-boundary search** - Find text across `<w:t>`, `<w:ins>`, and `<w:del>` boundaries
3. **Add boundary-aware replace** - Replace text spanning multiple elements with proper node splitting
4. **Add mixed-state editing** - Handle replacements spanning revision boundaries using atomic decomposition (split operations per segment type)
5. **Add `get_visible_text()` API** - Let users access flattened text for analysis

## Impact

- **Affected specs:** text-operations (new capability spec)
- **Affected code:**
  - `docx_editor/xml_editor.py` - Add `build_text_map()` method
  - `docx_editor/track_changes.py` - Update `_get_nth_match()` and `replace_text()`
  - `docx_editor/document.py` - Add `get_visible_text()` public API
- **Breaking changes:** None (new functionality, existing behavior preserved)
