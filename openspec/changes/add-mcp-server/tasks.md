## 1. Core MCP Package

- [x] 1.1 Create `docx_editor_mcp/` directory structure
- [x] 1.2 Implement `cache.py` with `DocumentCache` and `CachedDocument` classes
- [x] 1.3 Implement path normalization (resolve ~, relative paths, symlinks)
- [x] 1.4 Implement mtime-based external change detection
- [x] 1.5 Implement LRU eviction with save-on-evict for dirty documents
- [x] 1.6 Implement session author memory with cross-platform default (`getpass.getuser()`)

## 2. MCP Server

- [x] 2.1 Implement `server.py` with MCP server initialization
- [x] 2.2 Add graceful shutdown with best-effort save

## 3. MCP Tools

- [x] 3.1 Implement document lifecycle tools: `open_document`, `save_document`, `close_document`, `reload_document`, `force_save`
- [x] 3.2 Implement track changes tools: `replace_text`, `delete_text`, `insert_after`, `insert_before`
- [x] 3.3 Implement comment tools: `add_comment`, `list_comments`, `reply_to_comment`, `resolve_comment`, `delete_comment`
- [x] 3.4 Implement revision tools: `list_revisions`, `accept_revision`, `reject_revision`, `accept_all`, `reject_all`
- [x] 3.5 Implement read tools: `find_text`, `count_matches`, `get_visible_text`

## 4. Packaging

- [x] 4.1 Add `[mcp]` optional dependency to `pyproject.toml` (depends on `mcp` package)
- [x] 4.2 Add `__main__.py` for `python -m docx_editor_mcp` invocation
- [x] 4.3 Add `[project.scripts]` entry point: `mcp-server-docx`
- [x] 4.4 Verify package installs correctly with `pip install .[mcp]`
- [ ] 4.5 Test `uvx mcp-server-docx` works from PyPI (requires PyPI publish)

## 5. Plugin Configuration

- [x] 5.1 Add `.mcp.json` to plugin root with uvx-based server config
- [ ] 5.2 Test plugin auto-configures MCP server on enable (requires manual verification)

## 6. Skill Update

- [x] 6.1 Rewrite `skills/docx/SKILL.md` to use MCP tools (MCP-first approach)
- [x] 6.2 Document available MCP tools and their parameters
- [x] 6.3 Add examples of common workflows using MCP tools

## 7. Testing

- [x] 7.1 Unit tests for `DocumentCache` (get, eviction, mtime detection) - tests/mcp/test_cache.py
- [x] 7.2 Integration tests for MCP tools - tests/mcp/test_tools.py
- [x] 7.3 Test external change detection behavior - tests/mcp/test_tools.py::TestExternalChangeDetection
- [x] 7.4 Test shutdown with dirty documents - tests/mcp/test_server.py::TestGracefulShutdown

## 8. Documentation

- [x] 8.1 Add MCP section to README
- [x] 8.2 Document MCP configuration for Claude Code
