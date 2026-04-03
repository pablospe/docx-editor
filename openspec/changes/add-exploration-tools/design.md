## Context

LLMs using MCP tools will always prefer reading the full document if allowed, because they have no reason to use chunked exploration when `get_visible_text` returns everything. For large documents this wastes context and degrades quality. We need to make the exploration path the natural choice.

## Goals / Non-Goals

- Goals:
  - Make large document editing feasible via MCP without context window overflow
  - Guide the LLM to use exploration tools when the document is too large
  - Keep existing behavior unchanged for small documents
- Non-Goals:
  - Semantic search or NLP-based features
  - Page-based navigation (docx paragraphs don't map cleanly to pages)
  - Streaming or chunked responses

## Decisions

- **Smart truncation over hard limits**: `get_visible_text` returns full text for small docs (under `max_chars`), truncates with a hint for large docs. This preserves backward compatibility while naturally steering the LLM to exploration tools.
  - Alternatives: Hard error on large docs (too disruptive), always paginate (unnecessary for small docs)

- **Combine truncation with server instructions**: Both the truncated response and `SERVER_INSTRUCTIONS` tell the LLM about the exploration workflow. Belt and suspenders — the instructions help even before the LLM tries `get_visible_text`.
  - Alternatives: Only instructions (LLM may ignore), only truncation (no upfront guidance)

- **`max_chars=10000` default for `get_visible_text`**: Roughly 2500 tokens, enough for small docs but triggers truncation for anything substantial.

- **Paragraph-ref based `get_paragraph_text`**: Accepts a list of hash-anchored refs (from `list_paragraphs` or `search_text`) to read specific paragraphs in full. This is the primary "targeted read" tool.

- **`search_text` returns context + paragraph refs**: Each match includes surrounding text and its paragraph ref, so the LLM can immediately use the ref for editing without a separate `list_paragraphs` call.

## Risks / Trade-offs

- **LLM may still try to call `get_visible_text` repeatedly**: Mitigated by truncation hint explicitly saying "use search_text or get_paragraph_text instead"
- **`get_document_info` heading extraction depends on paragraph styles**: Not all docx files use heading styles consistently. The tool should return whatever headings exist, gracefully handling docs with no headings.

## Open Questions

- None currently
