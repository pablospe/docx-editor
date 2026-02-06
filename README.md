# docx-editor

[![Release](https://img.shields.io/github/v/release/pablospe/docx-editor)](https://img.shields.io/github/v/release/pablospe/docx-editor)
[![Build status](https://img.shields.io/github/actions/workflow/status/pablospe/docx-editor/main.yml?branch=main)](https://github.com/pablospe/docx-editor/actions/workflows/main.yml?query=branch%3Amain)
[![codecov](https://codecov.io/gh/pablospe/docx-editor/branch/main/graph/badge.svg)](https://codecov.io/gh/pablospe/docx-editor)
[![Commit activity](https://img.shields.io/github/commit-activity/m/pablospe/docx-editor)](https://img.shields.io/github/commit-activity/m/pablospe/docx-editor)
[![License](https://img.shields.io/github/license/pablospe/docx-editor)](https://img.shields.io/github/license/pablospe/docx-editor)

Pure Python library for Word document track changes and comments, without requiring Microsoft Word.

> **Note:** The PyPI package is named `docx-editor` because `docx-edit` was too similar to an existing package.

- **Github repository**: <https://github.com/pablospe/docx-editor/>
- **Documentation**: <https://pablospe.github.io/docx-editor/>

## Features

- **Track Changes**: Replace, delete, and insert text with revision tracking
- **Cross-Boundary Editing**: Find and replace text spanning multiple XML elements and revision boundaries
- **Mixed-State Editing**: Atomic decomposition for text spanning `<w:ins>`/`<w:del>` boundaries
- **Comments**: Add, reply, resolve, and delete comments
- **Revision Management**: List, accept, and reject tracked changes
- **Cross-Platform**: Works on Linux, macOS, and Windows
- **No Dependencies**: Only requires `defusedxml` for secure XML parsing

## Installation

```bash
pip install docx-editor
```

## Claude Code Plugin

This repo includes a plugin for [Claude Code](https://claude.ai/claude-code) that enables AI-assisted Word document editing.

This plugin extends the [original Anthropic docx skill](https://github.com/anthropics/skills/tree/main/skills/docx) which requires Claude to manually manipulate OOXML. Instead, this plugin provides an interface (`docx-editor`) that handles all the complexity—Claude just calls simple Python methods like `doc.replace()` or `doc.add_comment()`, making document editing significantly faster and less error-prone.

### Install as plugin

```bash
# Add the marketplace
/plugin marketplace add pablospe/docx-editor

# Install the plugin
/plugin install docx-editor@docx-editor-marketplace

# Install dependencies
pip install docx-editor python-docx
```

### Manual install (alternative)

```bash
# Install dependencies
pip install docx-editor python-docx

# Copy skill to Claude Code skills directory
git clone https://github.com/pablospe/docx-editor /tmp/docx-editor
mkdir -p ~/.claude/skills
cp -r /tmp/docx-editor/skills/docx ~/.claude/skills/
rm -rf /tmp/docx-editor
```

Once installed, Claude Code can help you edit Word documents with track changes, comments, and revisions.

## Quick Start

```python
from docx_editor import Document

with Document.open("contract.docx") as doc:
    # Track changes
    doc.replace("30 days", "60 days")
    doc.insert_after("Section 5", "New clause")
    doc.delete("obsolete text")

    # Comments
    doc.add_comment("Section 5", "Please review")

    # Revision management
    revisions = doc.list_revisions()
    doc.accept_revision(revision_id=1)

    doc.save()
```

### Cross-Boundary Text Operations

Text in Word documents with tracked changes can span revision boundaries. `docx-editor` handles this transparently:

```python
from docx_editor import Document

with Document.open("reviewed.docx") as doc:
    # Get visible text (inserted text included, deleted excluded)
    text = doc.get_visible_text()

    # Find text across element boundaries
    match = doc.find_text("Aim: To")
    if match and match.spans_boundary:
        print("Text spans a revision boundary")

    # Replace works even across revision boundaries
    doc.replace("Aim: To", "Goal: To")

    doc.save()
```

## MCP Server (Optional)

For faster performance when making multiple edits, use the MCP (Model Context Protocol) server. It keeps documents loaded between operations, making repeated edits 10-30x faster.

### Installation

```bash
pip install docx-editor[mcp]
```

### Running the Server

```bash
# Via console script
mcp-server-docx

# Or via Python module
python -m docx_editor_mcp
```

### Claude Code Configuration

The docx-editor plugin automatically configures the MCP server when installed. For manual configuration:

```bash
claude mcp add docx -- uvx mcp-server-docx
```

Or add to your `.mcp.json`:

```json
{
  "docx": {
    "command": "uvx",
    "args": ["mcp-server-docx"]
  }
}
```

### Available MCP Tools

- **Document lifecycle**: `open_document`, `save_document`, `close_document`, `reload_document`, `force_save`
- **Track changes**: `replace_text`, `delete_text`, `insert_after`, `insert_before`
- **Comments**: `add_comment`, `list_comments`, `reply_to_comment`, `resolve_comment`, `delete_comment`
- **Revisions**: `list_revisions`, `accept_revision`, `reject_revision`, `accept_all`, `reject_all`
- **Read**: `find_text`, `count_matches`, `get_visible_text`
