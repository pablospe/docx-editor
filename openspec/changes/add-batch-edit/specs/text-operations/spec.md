## ADDED Requirements

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
- Return a list of change IDs corresponding to each operation

#### Scenario: Batch of edits to different paragraphs

- **GIVEN** a document with paragraphs P1 through P10
- **WHEN** `batch_edit()` is called with 3 edits targeting P3, P7, and P9
- **THEN** all 3 edits succeed
- **AND** a list of 3 change IDs is returned
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
