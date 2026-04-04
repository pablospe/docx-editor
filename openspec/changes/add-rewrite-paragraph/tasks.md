## 1. Core Diffing Infrastructure

- [x] 1.1 Implement word-level tokenizer that splits visible text into words and whitespace tokens
- [x] 1.2 Implement diff engine using `difflib.SequenceMatcher` on word tokens
- [x] 1.3 Implement hunk-to-XML mapper that converts diff hunks to text-map positions using `build_text_map()` and `find_in_text_map()`

## 2. Rewrite Method

- [x] 2.1 Add `rewrite_paragraph(ref, new_text)` to `RevisionManager`
- [x] 2.2 Integrate hash validation via `_resolve_paragraph()`
- [x] 2.3 For each diff hunk: generate `<w:del>` for removed text and `<w:ins>` for added text
- [x] 2.4 Preserve formatting by inheriting `<w:rPr>` from adjacent runs
- [x] 2.5 Expose `rewrite_paragraph()` on `Document` facade

## 3. Batch Rewrite

- [x] 3.1 Add `batch_rewrite(rewrites)` to `Document` that accepts a list of `(ref, new_text)` pairs
- [x] 3.2 Validate all paragraph hashes upfront (reject entire batch on any mismatch)
- [x] 3.3 Apply rewrites in reverse paragraph order

## 4. Edge Cases

- [x] 4.1 Handle empty paragraph rewrite (insert all new text)
- [x] 4.2 Handle rewrite to empty text (delete all paragraph text)
- [x] 4.3 Handle paragraphs with existing tracked changes
- [x] 4.4 Handle no-op rewrite (new_text equals old text)

## 5. Tests

- [x] 5.1 Unit tests for word-level tokenizer
- [x] 5.2 Unit tests for diff-to-hunk mapping
- [x] 5.3 Integration tests for `rewrite_paragraph()` with simple text changes
- [x] 5.4 Integration tests for `rewrite_paragraph()` with formatting preservation
- [x] 5.5 Integration tests for `rewrite_paragraph()` with existing tracked changes
- [x] 5.6 Integration tests for `batch_rewrite()` with multiple paragraphs
- [x] 5.7 Test hash mismatch rejection for rewrite operations
- [x] 5.8 Test no-op rewrite raises no error and produces no changes

## 6. Documentation

- [x] 6.1 Update `skills/docx/SKILL.md` with `rewrite_paragraph()` API and usage guidance (when to use rewrite vs surgical methods)
- [x] 6.2 Update `README.md` with `rewrite_paragraph()` in features list and quick start examples
