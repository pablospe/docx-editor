## ADDED Requirements

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
- Return a list of strings in the format `P{index}#{hash}| {preview_text}`
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

The system SHALL accept an optional `paragraph` parameter on `replace()`, `delete()`, `insert_after()`, and `insert_before()` methods that scopes the text search to a single paragraph.

When `paragraph` is specified:
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
