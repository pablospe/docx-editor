## 1. Core Library Support

- [x] 1.1 Add `search_text(query, context_chars=100)` method to `Document` class
- [x] 1.2 Add `get_paragraph_text(paragraphs)` method to `Document` class
- [x] 1.3 Add `get_document_info()` method to `Document` class (paragraph count, word count, heading outline)
- [x] 1.4 Add `start`/`limit` parameters to `Document.list_paragraphs()`

## 2. MCP Tool Functions

- [x] 2.1 Add `search_text` tool function in `tools.py`
- [x] 2.2 Add `get_paragraph_text` tool function in `tools.py`
- [x] 2.3 Add `get_document_info` tool function in `tools.py`
- [x] 2.4 Modify `list_paragraphs` tool to accept `start`/`limit`
- [x] 2.5 Modify `get_visible_text` tool to accept `max_chars` and auto-truncate with hint

## 3. Server Registration and Instructions

- [x] 3.1 Register new tools in `server.py` (`_register_tools`)
- [x] 3.2 Update `SERVER_INSTRUCTIONS` to guide LLMs toward exploration workflow

## 4. Tests

- [x] 4.1 Test `search_text` (found, not found, multiple matches, context extraction)
- [x] 4.2 Test `get_paragraph_text` (single, multiple, invalid ref)
- [x] 4.3 Test `get_document_info` (with headings, without headings)
- [x] 4.4 Test `list_paragraphs` pagination (start/limit, out of range)
- [x] 4.5 Test `get_visible_text` truncation (small doc returns full, large doc truncates with hint)
