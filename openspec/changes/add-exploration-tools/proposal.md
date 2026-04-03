# Change: Add Document Exploration Tools for Large Documents

## Why

The current MCP tools assume the LLM can read the full document via `get_visible_text`. For large documents (100+ pages), this dumps too much text into the context window, degrading LLM performance and wasting tokens. Without exploration tools, the LLM has no way to orient itself in a large document, search for specific content, or read targeted sections.

The LLM will never voluntarily use chunked reading unless the full-text path is blocked or clearly discouraged. Smart truncation of `get_visible_text` combined with server instructions that guide the LLM toward exploration tools solves this.

## What Changes

- **MODIFIED:** `get_visible_text` - Auto-truncates large documents and appends a hint directing the LLM to use exploration tools
- **MODIFIED:** `list_paragraphs` - Adds pagination via `start`/`limit` parameters for browsing large documents
- **NEW:** `search_text(path, query, context_chars?)` - Full-text search returning matches with surrounding context and paragraph refs
- **NEW:** `get_paragraph_text(path, paragraphs)` - Read specific paragraphs in full by their hash-anchored refs
- **NEW:** `get_document_info(path)` - Document overview: paragraph count, word count, heading outline
- **MODIFIED:** `SERVER_INSTRUCTIONS` - Updated to guide LLMs toward exploration workflow for large docs

## Impact

- Affected specs: `mcp-server` (modified capability)
- Affected code:
  - `docx_editor_mcp/tools.py` (new tool functions + modified existing)
  - `docx_editor_mcp/server.py` (register new tools, update instructions)
  - `tests/test_mcp/test_tools.py` (new tests)
