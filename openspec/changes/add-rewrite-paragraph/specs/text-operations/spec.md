## ADDED Requirements

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
