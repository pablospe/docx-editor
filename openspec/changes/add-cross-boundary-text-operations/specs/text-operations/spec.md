## ADDED Requirements

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

### Requirement: Revision Boundary Error

The system SHALL raise `RevisionBoundaryError` when a text operation would span existing tracked changes, rather than silently failing or implicitly accepting changes.

#### Scenario: Replace spanning insertion raises error

- **GIVEN** a paragraph with "Hello " as regular text and "world" inside `<w:ins>`
- **WHEN** `replace_text("Hello world", "Hi there")` is called
- **THEN** a `RevisionBoundaryError` is raised
- **AND** the error message indicates the text spans existing tracked changes
- **AND** the document is not modified

#### Scenario: Replace within insertion succeeds

- **GIVEN** a paragraph with "Hello beautiful world" entirely inside `<w:ins>`
- **WHEN** `replace_text("beautiful", "wonderful")` is called
- **THEN** the replacement succeeds (text is within single revision context)
