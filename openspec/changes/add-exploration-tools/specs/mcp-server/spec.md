## ADDED Requirements

### Requirement: Document Exploration Tools

The MCP server SHALL provide tools for exploring large documents without loading full text into context.

Tools SHALL include:
- `search_text(path, query, context_chars?)` - Search for text, returning matches with surrounding context and paragraph refs
- `get_paragraph_text(path, paragraphs)` - Read specific paragraphs in full by hash-anchored ref
- `get_document_info(path)` - Document overview with paragraph count, word count, and heading outline

#### Scenario: Search text with context

- **GIVEN** a document contains "force majeure" in paragraph P12#a3b4
- **WHEN** `search_text(path="doc.docx", query="force majeure", context_chars=100)` is called
- **THEN** the tool returns a list of matches
- **AND** each match includes the paragraph ref (e.g., "P12#a3b4")
- **AND** each match includes surrounding context text (up to 100 chars before and after)
- **AND** each match includes the paragraph index

#### Scenario: Search text not found

- **GIVEN** a document does not contain "nonexistent phrase"
- **WHEN** `search_text(path="doc.docx", query="nonexistent phrase")` is called
- **THEN** the tool returns an empty match list
- **AND** the response indicates zero matches found

#### Scenario: Get specific paragraphs by ref

- **GIVEN** a document has paragraphs P1#aa11, P2#bb22, P3#cc33
- **WHEN** `get_paragraph_text(path="doc.docx", paragraphs=["P1#aa11", "P3#cc33"])` is called
- **THEN** the tool returns the full text of paragraphs P1 and P3
- **AND** each result includes the paragraph ref and full text

#### Scenario: Get paragraph with invalid ref

- **GIVEN** a document does not contain paragraph "P99#dead"
- **WHEN** `get_paragraph_text(path="doc.docx", paragraphs=["P99#dead"])` is called
- **THEN** the tool returns an error for that specific paragraph
- **AND** valid paragraphs in the same request are still returned

#### Scenario: Get document info

- **GIVEN** a document has 150 paragraphs and headings at various levels
- **WHEN** `get_document_info(path="doc.docx")` is called
- **THEN** the tool returns paragraph count, word count, and a heading outline
- **AND** the heading outline includes heading text and level

#### Scenario: Get document info without headings

- **GIVEN** a document has no heading styles applied
- **WHEN** `get_document_info(path="doc.docx")` is called
- **THEN** the tool returns paragraph count and word count
- **AND** the heading outline is empty

## MODIFIED Requirements

### Requirement: Read Tools

The MCP server SHALL expose read-only document inspection tools.

Tools SHALL include:
- `find_text(path, text)` - Check if text exists in document
- `count_matches(path, text)` - Count occurrences of text
- `get_visible_text(path, max_chars?)` - Get visible text of document, auto-truncated for large docs
- `list_paragraphs(path, max_chars?, start?, limit?)` - List paragraphs with optional pagination

`get_visible_text` SHALL accept an optional `max_chars` parameter (default: 10000). When the document text exceeds `max_chars`, the response SHALL be truncated and include a hint directing the LLM to use `search_text`, `get_paragraph_text`, or `get_document_info` for targeted exploration.

`list_paragraphs` SHALL accept optional `start` and `limit` parameters for paginating through large documents. When omitted, all paragraphs are returned.

#### Scenario: Count matches via MCP

- **GIVEN** a document contains "30 days" three times
- **WHEN** `count_matches(path="doc.docx", text="30 days")` is called
- **THEN** the tool returns 3

#### Scenario: Get visible text for small document

- **GIVEN** a document has 500 characters of visible text
- **WHEN** `get_visible_text(path="doc.docx")` is called
- **THEN** the full visible text is returned

#### Scenario: Get visible text for large document

- **GIVEN** a document has 50000 characters of visible text
- **WHEN** `get_visible_text(path="doc.docx")` is called
- **THEN** the first 10000 characters are returned
- **AND** a hint is appended: the document was truncated and the LLM should use exploration tools

#### Scenario: Get visible text with custom max_chars

- **GIVEN** a document has 50000 characters of visible text
- **WHEN** `get_visible_text(path="doc.docx", max_chars=5000)` is called
- **THEN** the first 5000 characters are returned
- **AND** a truncation hint is appended

#### Scenario: List paragraphs with pagination

- **GIVEN** a document has 200 paragraphs
- **WHEN** `list_paragraphs(path="doc.docx", start=50, limit=20)` is called
- **THEN** paragraphs 50 through 69 are returned with their hash-anchored refs
- **AND** the response indicates total paragraph count and the returned range
