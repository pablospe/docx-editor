## 1. Core MCP Package

- [ ] 1.1 Create `docx_editor_mcp/` directory structure
- [ ] 1.2 Implement `cache.py` with `DocumentCache` and `CachedDocument` classes
- [ ] 1.3 Implement path normalization (resolve ~, relative paths, symlinks)
- [ ] 1.4 Implement mtime-based external change detection
- [ ] 1.5 Implement LRU eviction with save-on-evict for dirty documents
- [ ] 1.6 Implement session author memory with cross-platform default (`getpass.getuser()`)

## 2. MCP Server

- [ ] 2.1 Implement `server.py` with MCP server initialization
- [ ] 2.2 Add graceful shutdown with best-effort save

## 3. MCP Tools

- [ ] 3.1 Implement document lifecycle tools: `open_document`, `save_document`, `close_document`, `reload_document`, `force_save`
- [ ] 3.2 Implement track changes tools: `replace_text`, `delete_text`, `insert_after`, `insert_before`
- [ ] 3.3 Implement comment tools: `add_comment`, `list_comments`, `reply_to_comment`, `resolve_comment`, `delete_comment`
- [ ] 3.4 Implement revision tools: `list_revisions`, `accept_revision`, `reject_revision`, `accept_all`, `reject_all`
- [ ] 3.5 Implement read tools: `find_text`, `count_matches`, `get_visible_text`

## 4. Packaging

- [ ] 4.1 Add `[mcp]` optional dependency to `pyproject.toml` (depends on `mcp` package)
- [ ] 4.2 Add `__main__.py` for `python -m docx_editor_mcp` invocation
- [ ] 4.3 Add `[project.scripts]` entry point: `mcp-server-docx`
- [ ] 4.4 Verify package installs correctly with `pip install .[mcp]`
- [ ] 4.5 Test `uvx mcp-server-docx` works from PyPI

## 5. Plugin Configuration

- [ ] 5.1 Add `.mcp.json` to plugin root with uvx-based server config
- [ ] 5.2 Test plugin auto-configures MCP server on enable

## 6. Skill Update

- [ ] 6.1 Rewrite `skills/docx/SKILL.md` to use MCP tools (MCP-first approach)
- [ ] 6.2 Document available MCP tools and their parameters
- [ ] 6.3 Add examples of common workflows using MCP tools

## 7. Testing

- [ ] 7.1 Unit tests for `DocumentCache` (get, eviction, mtime detection)
- [ ] 7.2 Integration tests for MCP tools
- [ ] 7.3 Test external change detection behavior
- [ ] 7.4 Test shutdown with dirty documents

## 8. Documentation

- [ ] 8.1 Add MCP section to README
- [ ] 8.2 Document MCP configuration for Claude Code
