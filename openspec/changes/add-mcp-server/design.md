## Context

Claude Code invokes the docx skill by running Python scripts via Bash. Each script execution loads the document, performs edits, saves, and exits. For interactive editing sessions where users make 5-10 edits on the same document, the reload overhead dominates.

MCP (Model Context Protocol) servers maintain state between tool calls, making them ideal for caching parsed document DOMs.

### Stakeholders
- Users doing interactive document editing with Claude Code
- The docx_editor library (must remain usable standalone)

## Goals / Non-Goals

**Goals:**
- Provide faster repeated operations on the same document
- Keep the core library independent of MCP
- Support both MCP and script-based workflows
- Detect external file modifications

**Non-Goals:**
- Multi-user/concurrent access (single-user tool)
- Memory-optimized large document handling
- Real-time file synchronization

## Decisions

### Decision: Separate `docx_editor_mcp/` package

The MCP server lives in a separate Python package within the same repo.

**Why:**
- Core library stays lean and usable without MCP
- Clear separation of concerns
- Optional installation via extras

**Alternatives considered:**
- Bake MCP into core library - Rejected: adds complexity for non-MCP users
- Separate repository - Rejected: harder to keep in sync, overkill for thin wrapper

### Decision: Document count-based cache (not memory-based)

Cache evicts based on document count (default: 10), not memory usage.

**Why:**
- Simple to implement and reason about
- Memory measurement in Python is non-trivial
- 10 documents is reasonable for interactive use

**Alternatives considered:**
- Memory-based eviction - Rejected: complex, requires tracking DOM memory usage
- No limit - Rejected: unbounded memory growth

### Decision: Explicit save only

Documents are only persisted when the user explicitly calls `save_document`.

**Why:**
- Predictable behavior
- User controls when changes are committed
- Matches the library's existing pattern

**Alternatives considered:**
- Auto-save on timer - Rejected: unpredictable, may save partial work
- Save on eviction - Still implemented as safety net, but not primary mechanism

### Decision: mtime-based external change detection

Before each operation, check if file's mtime differs from cached value.

**Why:**
- Simple, no external dependencies
- Catches most external modification scenarios
- Low overhead

**Alternatives considered:**
- File watcher (watchdog) - Rejected: adds dependency, complexity
- Hash-based detection - Rejected: expensive for large files
- Ignore external changes - Rejected: risk of data loss

### Decision: Keep both skill modes (script + MCP)

The skill continues to work via Python scripts. MCP is an optional upgrade.

**Why:**
- No breaking changes for existing users
- MCP is opt-in for users who need performance
- Lower barrier to entry

### Decision: Console scripts entry point

Provide `mcp-server-docx` command via `[project.scripts]` in pyproject.toml.

```toml
[project.scripts]
mcp-server-docx = "docx_editor_mcp:main"
```

**Why:**
- Standard Python packaging pattern for CLI tools
- Works with `uvx mcp-server-docx` for zero-install execution
- Follows MCP naming convention (`mcp-server-{name}`)
- Enables manual MCP add: `claude mcp add docx -- uvx mcp-server-docx`

**Alternatives considered:**
- Only `python -m docx_editor_mcp` - Rejected: less discoverable, harder to configure in mcp.json

### Decision: Plugin auto-configuration via .mcp.json

The docx-editor plugin includes `.mcp.json` to auto-configure the MCP server.

```json
{
  "docx": {
    "command": "uvx",
    "args": ["mcp-server-docx"]
  }
}
```

**Why:**
- Zero-config for plugin users - MCP server starts automatically
- Uses uvx for automatic package resolution
- No manual `claude mcp add` required

**Alternatives considered:**
- Manual configuration only - Rejected: adds friction for users
- Bundled binary - Rejected: harder to maintain, platform-specific

### Decision: Session author memory with hint

On first document open without explicit author, use system username and hint Claude to ask user.

```python
def get_author(self, explicit_author: str | None) -> tuple[str, bool]:
    """Returns (author, is_first_use)"""
    if explicit_author:
        self.session_author = explicit_author
        return explicit_author, False
    if self.session_author:
        return self.session_author, False
    default = getpass.getuser() or "Reviewer"
    self.session_author = default
    return default, True  # True = hint Claude to ask
```

**Why:**
- Cross-platform: `getpass.getuser()` works on Linux, macOS, Windows
- Non-blocking: silent default works if Claude doesn't ask
- Session-aware: author remembered for subsequent operations

### Decision: Path normalization

All paths normalized to absolute canonical form before cache lookup.

```python
def normalize_path(path: str) -> str:
    return os.path.realpath(os.path.expanduser(path))
```

**Why:**
- Prevents duplicate cache entries for same physical file
- Handles `~`, relative paths, symlinks consistently
- Simple, no edge cases

### Decision: MCP-first skill

The skill teaches Claude to use MCP tools directly, not Python scripts.

**Why:**
- Plugin auto-configures MCP - it's always available
- Simpler skill content
- Better user experience (faster, cached operations)

**Alternatives considered:**
- Document both modes - Rejected: adds complexity, confusing
- Detect and adapt - Rejected: unnecessary since plugin ensures MCP availability

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│ Claude Code                                                  │
│                                                              │
│   Skill loads on-demand → MCP tools via Tool Search         │
│                     ↓                                        │
│   Claude calls: delete_paragraph("file.docx", 28)           │
└──────────────────────┬──────────────────────────────────────┘
                       │ MCP Protocol (stdio)
                       ↓
┌─────────────────────────────────────────────────────────────┐
│ MCP Server (docx_editor_mcp)                                │
│                                                              │
│  ┌────────────────────────────────────────┐                 │
│  │ DocumentCache                          │                 │
│  │  - max_documents: 10                   │                 │
│  │  - LRU eviction                        │                 │
│  │  - mtime tracking per document         │                 │
│  │                                         │                 │
│  │  get(path) -> CachedDocument           │                 │
│  │  save(path) -> update mtime            │                 │
│  │  close(path) -> evict from cache       │                 │
│  └────────────────────────────────────────┘                 │
│                                                              │
│  MCP Tools:                                                  │
│  - Track changes: replace, delete, insert_after, etc.       │
│  - Comments: add_comment, list_comments, etc.               │
│  - Revisions: list_revisions, accept_revision, etc.         │
│  - Document: open_document, save_document, close_document   │
└─────────────────────────────────────────────────────────────┘
                       │
                       ↓
              ┌───────────────┐
              │ docx_editor   │  (core library, unchanged)
              └───────────────┘
                       │
                       ↓
                 ┌──────────┐
                 │ file.docx│
                 └──────────┘
```

## Package Structure

```
docx-edit/
├── docx_editor/              # Core library (existing, unchanged)
├── docx_editor_mcp/          # NEW: MCP server wrapper
│   ├── __init__.py
│   ├── server.py             # MCP server entry point
│   ├── cache.py              # DocumentCache, CachedDocument
│   └── tools.py              # MCP tool definitions
├── skills/docx/SKILL.md      # Updated: document MCP option
└── pyproject.toml            # Add [mcp] optional dependency
```

## Risks / Trade-offs

| Risk | Mitigation |
|------|------------|
| User forgets to save, loses work | Warn on close if dirty; save on eviction |
| External edits overwritten | mtime check before operations, return error |
| MCP server crashes with unsaved data | Best-effort save on shutdown |
| Context overhead from MCP tools | Tool Search lazy-loads definitions |

## Migration Plan

1. Implement `docx_editor_mcp/` package
2. Add `[mcp]` optional dependency to pyproject.toml
3. Update skill documentation with MCP setup instructions
4. Test with Claude Code MCP configuration
5. Document in README

**Rollback:** Remove `docx_editor_mcp/` from pyproject.toml extras. Core library unaffected.

## Resolved Questions

1. **CLI entry point?** → Yes, `mcp-server-docx` via `[project.scripts]`. Also supports `python -m docx_editor_mcp`.
2. **Plugin auto-configuration?** → Add `.mcp.json` to plugin root with uvx-based command.
