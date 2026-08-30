# text-operations Specification

## Purpose

Enable text search, replacement, and deletion across `<w:t>` element boundaries within a paragraph, including mixed revision contexts (regular text and tracked insertions).
## Requirements
### Requirement: Virtual Text Map

The system SHALL provide a flattened text view of document content that maps character positions back to their source XML elements.

The text map SHALL:
- Concatenate all visible text from `<w:t>` elements within a paragraph
- Exclude text inside `<w:delText>` elements (deleted content)
- Track whether each character position is inside a `<w:ins>` or `<w:del>` element
- Map each character position to its source `<w:t>` node and offset

#### Scenario: Build text map for paragraph with tracked changes

- **GIVEN** a paragraph containing: regular text "Hello ", inserted text "beautiful ", regular text "world"
- **WHEN** `build_text_map()` is called on the paragraph
- **THEN** the text map contains "Hello beautiful world"
- **AND** positions 0-5 map to the first `<w:t>` with `is_inside_ins=False`
- **AND** positions 6-15 map to the second `<w:t>` with `is_inside_ins=True`
- **AND** positions 16-20 map to the third `<w:t>` with `is_inside_ins=False`

#### Scenario: Deleted text excluded from visible text

- **GIVEN** a paragraph containing: "Hello " and deleted text "old " and "world"
- **WHEN** `build_text_map()` is called
- **THEN** the text map contains "Hello world" (deleted text excluded)

### Requirement: Cross-Boundary Text Search

The system SHALL find text that spans multiple XML elements within a paragraph.

#### Scenario: Search finds text spanning element boundary

- **GIVEN** a paragraph with "Exploratory Aim: " in one `<w:t>` and "To examine" in another
- **WHEN** searching for "Aim: To"
- **THEN** the search succeeds and returns match information
- **AND** the match indicates it spans multiple elements

#### Scenario: Search finds text spanning insertion boundary

- **GIVEN** a paragraph with "Hello " as regular text and "world" inside `<w:ins>`
- **WHEN** searching for "Hello world"
- **THEN** the search succeeds
- **AND** the match indicates it spans a revision boundary

### Requirement: Visible Text API

The system SHALL provide a public API to retrieve the flattened visible text of a document.

#### Scenario: Get visible text from document

- **GIVEN** a document with paragraphs containing mixed regular and tracked-change content
- **WHEN** `get_visible_text()` is called
- **THEN** the method returns a string containing all visible text
- **AND** deleted text is excluded
- **AND** inserted text is included

### Requirement: Boundary-Aware Text Replacement

The system SHALL replace text that spans multiple `<w:t>` elements within the same revision context, using proper node splitting.

#### Scenario: Replace text spanning multiple runs

- **GIVEN** a paragraph with "Hello " in one run and "world" in another (no revisions)
- **WHEN** `replace_text("Hello world", "Hi there")` is called
- **THEN** the replacement succeeds
- **AND** the original text is wrapped in `<w:del>`
- **AND** the new text is wrapped in `<w:ins>`

### Requirement: Mixed-State Editing

The system SHALL replace text that spans revision boundaries by decomposing the operation into per-segment atomic actions, one for each revision context (regular text, inside `<w:ins>`, inside `<w:del>`).

#### Scenario: Replace spanning regular text and insertion

- **GIVEN** a paragraph with "Exploratory Aim: " as regular text and "To examine" inside `<w:ins>`
- **WHEN** `replace_text("Aim: To", "Goal: To")` is called
- **THEN** the replacement succeeds
- **AND** "Aim: " is wrapped in `<w:del>` (standard deletion of regular text)
- **AND** "To" is removed from the `<w:ins>` element (undoing partial insertion)
- **AND** "Goal: To" is inserted as a new `<w:ins>` element
- **AND** the remaining " examine" stays inside the original `<w:ins>`

#### Scenario: Replace fully within insertion

- **GIVEN** a paragraph with "Hello beautiful world" entirely inside `<w:ins>`
- **WHEN** `replace_text("beautiful", "wonderful")` is called
- **THEN** the replacement succeeds
- **AND** "beautiful" is removed from the insertion (undoing partial insertion)
- **AND** "wonderful" is inserted as a new `<w:ins>` element

#### Scenario: Insertion node is split when partially matched

- **GIVEN** a paragraph with `<w:ins>To examine whether</w:ins>`
- **WHEN** a delete operation targets "To" (the first 2 characters)
- **THEN** the `<w:ins>` is split into two parts
- **AND** the `<w:ins>To</w:ins>` portion is removed
- **AND** `<w:ins> examine whether</w:ins>` remains intact

### Requirement: Paragraph Hash Computation

The system SHALL compute a content-derived hash for each paragraph using `zlib.crc32` of the paragraph's visible text (from `build_text_map`), truncated to 4 lowercase hex characters.

The hash SHALL:
- Use `zlib.crc32(text.encode("utf-8")) & 0xFFFF` formatted as 4-char lowercase hex
- Compute from the same visible text that `get_visible_text()` uses (insertions included, deletions excluded)
- Produce a deterministic hash for empty paragraphs

#### Scenario: Hash computed from visible text

- **GIVEN** a paragraph with visible text "Hello beautiful world"
- **WHEN** `compute_paragraph_hash()` is called on the paragraph
- **THEN** a 4-character lowercase hex string is returned
- **AND** calling it again on the same paragraph returns the same hash

#### Scenario: Hash changes when content changes

- **GIVEN** a paragraph with visible text "Hello world"
- **WHEN** a tracked change modifies the paragraph content
- **THEN** `compute_paragraph_hash()` returns a different hash than before the change

#### Scenario: Hash excludes deleted text

- **GIVEN** a paragraph containing "Hello " and deleted text "old " and "world"
- **WHEN** `compute_paragraph_hash()` is called
- **THEN** the hash is computed from "Hello world" (deleted text excluded)

### Requirement: Paragraph Reference Format

The system SHALL provide a `ParagraphRef` dataclass that parses and validates references in the format `P{1-indexed}#{4-hex-hash}`.

The `ParagraphRef` SHALL:
- Parse valid references matching `^P(\d+)#([0-9a-f]{4})$`
- Use 1-based paragraph indexing
- Raise `ValueError` for invalid reference formats

#### Scenario: Parse valid paragraph reference

- **GIVEN** a reference string "P3#a7b2"
- **WHEN** `ParagraphRef.parse("P3#a7b2")` is called
- **THEN** it returns a `ParagraphRef` with `index=3` and `hash="a7b2"`

#### Scenario: Reject invalid paragraph reference

- **GIVEN** an invalid reference string "paragraph3"
- **WHEN** `ParagraphRef.parse("paragraph3")` is called
- **THEN** a `ValueError` is raised

### Requirement: Paragraph Listing

The system SHALL provide a `list_paragraphs()` method on `Document` that returns hash-tagged paragraph previews.

The listing SHALL:
- Return a list of strings in the format `P{index}#{hash}| {preview_text}`, at most `limit` entries (default 200; `limit=None` for all) starting at 1-based `start` (default 1); when paragraphs remain beyond the window, the last entry is a notice line starting with `...` (e.g. `... 50 more paragraphs; use start=201 or limit=None`) rather than a paragraph
- Use 1-based paragraph indexing
- Truncate preview text to `max_chars` (default 80) with `...` suffix when truncated
- Include empty paragraphs (with empty preview text)

#### Scenario: List paragraphs with previews

- **GIVEN** a document with 3 paragraphs: "Introduction to the project", "", "The committee has decided to proceed"
- **WHEN** `list_paragraphs()` is called
- **THEN** the result is a list of 3 strings
- **AND** each string starts with `P{n}#{hash}|`
- **AND** the first string contains "Introduction to the project"
- **AND** the second string represents the empty paragraph
- **AND** the third string contains "The committee has decided to proceed"

#### Scenario: Truncate long paragraphs

- **GIVEN** a paragraph with visible text longer than 80 characters
- **WHEN** `list_paragraphs(max_chars=80)` is called
- **THEN** the preview is truncated to 80 characters followed by "..."

### Requirement: Paragraph-Scoped Text Operations

The system SHALL require a `paragraph` reference on `replace()`, `delete()`, `insert_after()`, and `insert_before()` methods that scopes the text search to a single paragraph.

For every call:
- The system SHALL parse the reference using `ParagraphRef`
- The system SHALL resolve the paragraph by index and validate its hash
- The `occurrence` parameter SHALL count matches within that paragraph only (paragraph-local)
- Text search SHALL only consider content within the specified paragraph

#### Scenario: Replace text scoped to a specific paragraph

- **GIVEN** a document where "the" appears in paragraphs 1, 2, and 3
- **WHEN** `replace("the", "THE", paragraph="P2#f3c1")` is called
- **THEN** only the first occurrence of "the" in paragraph 2 is replaced
- **AND** paragraphs 1 and 3 are unchanged

#### Scenario: Paragraph-local occurrence counting

- **GIVEN** a document where paragraph 2 contains "the" three times
- **WHEN** `replace("the", "THE", occurrence=2, paragraph="P2#f3c1")` is called
- **THEN** the second occurrence of "the" within paragraph 2 is replaced
- **AND** no other paragraphs are affected

#### Scenario: Text not found in scoped paragraph

- **GIVEN** a document where "specific" appears only in paragraph 1
- **WHEN** `replace("specific", "general", paragraph="P2#f3c1")` is called
- **THEN** the operation fails (text not found in the specified paragraph)

### Requirement: Staleness Detection

The system SHALL raise `HashMismatchError` when a paragraph reference's hash does not match the paragraph's current content hash.

The `HashMismatchError` SHALL include:
- The paragraph index
- The expected hash (from the reference)
- The actual hash (recomputed from current content)
- A preview of the paragraph's current content

#### Scenario: Reject edit with stale hash

- **GIVEN** a document where paragraph 2 has been modified since the LLM last called `list_paragraphs()`
- **WHEN** an edit is attempted using the old hash for paragraph 2
- **THEN** a `HashMismatchError` is raised
- **AND** the error message includes the current hash so the caller can retry

#### Scenario: Reject edit after paragraph shift

- **GIVEN** a document where a paragraph was inserted above paragraph 2, shifting old paragraph 2 to index 3
- **WHEN** an edit is attempted using the old `P2#{old_hash}`
- **THEN** a `HashMismatchError` is raised (the content at index 2 is now different)

#### Scenario: Successful sequential edits with fresh references

- **GIVEN** a document with multiple paragraphs
- **WHEN** the caller edits paragraph 2, then calls `list_paragraphs()` again, then edits paragraph 3 using the fresh reference
- **THEN** both edits succeed because each uses a current hash

### Requirement: Batch Edit Operations

The system SHALL provide a `batch_edit()` method on `Document` that accepts a list of edit operations and applies them atomically.

Each operation SHALL specify:
- `action`: one of `replace`, `delete`, `insert_after`, `insert_before`
- `paragraph`: a hash-anchored paragraph reference (required for batch mode)
- Action-specific fields: `find`/`replace_with` for replace, `text` for delete, `anchor`/`text` for insert

The system SHALL:
- Validate ALL paragraph hashes upfront before applying any edits
- Reject the entire batch with `HashMismatchError` if any hash is stale (no edits applied)
- Apply edits in reverse paragraph order (highest index first) so that earlier paragraphs' hashes remain valid
- Return a list of `EditResult` objects, one per operation, in operation order (`list[EditValidationResult]` when `dry_run=True`)

#### Scenario: Batch of edits to different paragraphs

- **GIVEN** a document with paragraphs P1 through P10
- **WHEN** `batch_edit()` is called with 3 edits targeting P3, P7, and P9
- **THEN** all 3 edits succeed
- **AND** a list of 3 `EditResult` objects is returned
- **AND** edits are applied in order P9, P7, P3 (reverse)

#### Scenario: Batch rejected on stale hash

- **GIVEN** a document where paragraph P5 has been modified since the last `list_paragraphs()` call
- **WHEN** `batch_edit()` is called with edits including a stale ref for P5
- **THEN** `HashMismatchError` is raised
- **AND** no edits from the batch are applied to the document

#### Scenario: Single snapshot suffices for entire batch

- **GIVEN** a document with 20 paragraphs
- **WHEN** `list_paragraphs()` is called once, and all refs are used in a single `batch_edit()` call
- **THEN** all edits succeed without needing to re-read paragraph hashes between edits

#### Scenario: Multiple edits to same paragraph

- **GIVEN** a batch with two edits targeting the same paragraph P5 (different text within it)
- **WHEN** `batch_edit()` is called
- **THEN** both edits are applied to P5
- **AND** the second edit uses the paragraph content as modified by the first edit

### Requirement: Paragraph Rewrite via Automatic Diffing

The system SHALL provide a `rewrite_paragraph(ref, new_text)` method on `Document` that accepts a hash-anchored paragraph reference and the complete desired text for that paragraph.

The method SHALL:
- Validate the paragraph hash using the same mechanism as other hash-anchored methods
- Retrieve the paragraph's current visible text via `build_text_map()`
- Compute a word-level diff between old and new text using `difflib.SequenceMatcher`
- Generate fine-grained `<w:del>` and `<w:ins>` tracked changes for each changed segment
- Preserve unchanged text and its formatting without modification
- Inherit run properties (`<w:rPr>`) from the nearest adjacent run for newly inserted text

#### Scenario: Rewrite paragraph with word-level changes

- **GIVEN** a document where paragraph 2 has visible text "The committee has decided to proceed with the plan"
- **WHEN** `rewrite_paragraph("P2#f3c1", "The board has decided to approve the plan")` is called
- **THEN** the operation succeeds
- **AND** "committee" is wrapped in `<w:del>` and "board" is wrapped in `<w:ins>`
- **AND** "proceed with" is wrapped in `<w:del>` and "approve" is wrapped in `<w:ins>`
- **AND** "The ", "has decided to ", and " the plan" remain unchanged in the XML

#### Scenario: Rewrite paragraph with additions only

- **GIVEN** a document where paragraph 1 has visible text "Hello world"
- **WHEN** `rewrite_paragraph("P1#ab12", "Hello beautiful world")` is called
- **THEN** the operation succeeds
- **AND** "beautiful " is wrapped in `<w:ins>`
- **AND** "Hello " and "world" remain unchanged in the XML

#### Scenario: Rewrite paragraph with deletions only

- **GIVEN** a document where paragraph 1 has visible text "Hello beautiful world"
- **WHEN** `rewrite_paragraph("P1#cd34", "Hello world")` is called
- **THEN** the operation succeeds
- **AND** "beautiful " is wrapped in `<w:del>`
- **AND** "Hello " and "world" remain unchanged in the XML

#### Scenario: Rewrite rejected with stale hash

- **GIVEN** a document where paragraph 3 has been modified since the last `list_paragraphs()` call
- **WHEN** `rewrite_paragraph("P3#old1", "New text")` is called with a stale hash
- **THEN** `HashMismatchError` is raised
- **AND** no changes are applied to the document

#### Scenario: Rewrite empty paragraph

- **GIVEN** a document where paragraph 4 is an empty paragraph
- **WHEN** `rewrite_paragraph("P4#e000", "New content for this paragraph")` is called
- **THEN** the operation succeeds
- **AND** the entire new text is wrapped in a single `<w:ins>` element

#### Scenario: Rewrite paragraph to empty text

- **GIVEN** a document where paragraph 2 has visible text "Remove this entirely"
- **WHEN** `rewrite_paragraph("P2#ff12", "")` is called
- **THEN** the operation succeeds
- **AND** all visible text is wrapped in `<w:del>` elements
- **AND** the `<w:p>` paragraph element is preserved

#### Scenario: No-op rewrite produces no changes

- **GIVEN** a document where paragraph 1 has visible text "Unchanged text"
- **WHEN** `rewrite_paragraph("P1#ab00", "Unchanged text")` is called
- **THEN** the operation succeeds
- **AND** no tracked changes are generated in the XML

#### Scenario: Rewrite paragraph with existing tracked changes

- **GIVEN** a paragraph containing "Hello " as regular text and "beautiful " inside `<w:ins>` and "world" as regular text
- **WHEN** `rewrite_paragraph(ref, "Hello wonderful world")` is called
- **THEN** "beautiful " is removed from the `<w:ins>` element (undoing partial insertion)
- **AND** "wonderful " is inserted as a new `<w:ins>` element
- **AND** "Hello " and "world" remain unchanged

#### Scenario: Formatting preserved on insertion

- **GIVEN** a paragraph where "Hello world" is formatted in bold
- **WHEN** `rewrite_paragraph(ref, "Hello beautiful world")` is called
- **THEN** the inserted "beautiful " inherits the bold formatting from the adjacent run

### Requirement: Batch Paragraph Rewrite

The system SHALL provide a `batch_rewrite(rewrites)` method on `Document` that accepts a list of paragraph rewrites and applies them atomically.

Each rewrite SHALL specify:
- `ref`: a hash-anchored paragraph reference
- `new_text`: the complete desired text for that paragraph

The method SHALL:
- Validate ALL paragraph hashes upfront before applying any rewrites
- Reject the entire batch with `HashMismatchError` if any hash is stale (no rewrites applied)
- Apply rewrites in reverse paragraph order (highest index first)
- Reject batches that contain duplicate paragraph references

#### Scenario: Batch rewrite of multiple paragraphs

- **GIVEN** a document with paragraphs P1 through P5
- **WHEN** `batch_rewrite()` is called with rewrites for P2, P4, and P5
- **THEN** all 3 rewrites succeed
- **AND** rewrites are applied in order P5, P4, P2 (reverse)

#### Scenario: Batch rejected on stale hash

- **GIVEN** a document where paragraph P3 has been modified since the last `list_paragraphs()` call
- **WHEN** `batch_rewrite()` is called with rewrites including a stale ref for P3
- **THEN** `HashMismatchError` is raised
- **AND** no rewrites from the batch are applied to the document

#### Scenario: Batch rejected on duplicate paragraph

- **GIVEN** a batch with two rewrites targeting the same paragraph P2
- **WHEN** `batch_rewrite()` is called
- **THEN** a `ValueError` is raised
- **AND** no rewrites are applied

#### Scenario: Single snapshot suffices for entire batch

- **GIVEN** a document with 10 paragraphs
- **WHEN** `list_paragraphs()` is called once, and refs are used in a single `batch_rewrite()` call
- **THEN** all rewrites succeed without needing to re-read paragraph hashes between rewrites
