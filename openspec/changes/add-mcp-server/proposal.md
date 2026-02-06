# Change: Add MCP Server for Persistent DOM Caching

## Why

When Claude Code edits a document through the skill, each tool invocation runs a separate Python script. The document is loaded, edited, saved, and the DOM is freed. For multi-edit sessions (common in interactive document review), this reload penalty compounds - loading a large document takes seconds, multiplied by each edit.

An MCP server keeps the document DOM in memory between tool calls, providing 10-30x faster performance for repeated operations on the same document.

## What Changes

- **NEW:** `docx_editor_mcp/` package - MCP server wrapping the core library
- **NEW:** Document cache with LRU eviction and external change detection
- **NEW:** MCP tools mirroring the library's public API
- **MODIFIED:** Skill documentation to reference MCP tools
- **NEW:** Optional dependency: `pip install docx-editor[mcp]`

## Impact

- Affected specs: `mcp-server` (new capability)
- Affected code:
  - `docx_editor_mcp/` (new package)
  - `skills/docx/SKILL.md` (documentation update)
  - `pyproject.toml` (optional dependency)

## Key Design Decisions

1. **Separate package** - MCP code lives in `docx_editor_mcp/`, keeping core library lean
2. **Both modes supported** - Skill-based (Python scripts) remains default; MCP is opt-in
3. **MCP Tool Search** - Claude Code's lazy tool loading prevents context bloat
4. **Explicit save only** - User controls when to persist changes
5. **External change detection** - Check file mtime before operations, warn if modified externally
