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

- **Hash-Anchored Paragraph References**: `list_paragraphs()` returns stable, hash-based paragraph IDs for safe, unambiguous targeting
- **Batch Editing**: Atomic `batch_edit()` with upfront hash validation across all operations
- **Paragraph Rewrite**: `rewrite_paragraph()` with automatic word-level diffing — specify desired text, get fine-grained tracked changes
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
import os

author = os.environ.get("USER") or "Reviewer"
with Document.open("contract.docx", author=author) as doc:
    # Step 1: List paragraphs with hash-anchored references
    for p in doc.list_paragraphs():
        print(p)
    # Output: P1#a7b2| Introduction to the contract...
    #         P2#f3c1| The committee shall review...

    # Step 2: Edit — each method returns the new paragraph ref
    r = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
    doc.replace("net", "gross", paragraph=r)  # chain without list_paragraphs()
    doc.delete("obsolete text", paragraph="P5#d4e5")
    doc.insert_after("Section 5", " (as amended)", paragraph="P3#b2c4")

    # Rewrite entire paragraph (automatic word-level diff)
    doc.rewrite_paragraph("P2#f3c1",
        "The board shall approve the updated proposal.")

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
import os

author = os.environ.get("USER") or "Reviewer"
with Document.open("reviewed.docx", author="Editor") as doc:
    # Get visible text (inserted text included, deleted excluded)
    text = doc.get_visible_text()

    # List paragraphs to find hash-anchored references
    refs = doc.list_paragraphs()

    # Find text across element boundaries
    match = doc.find_text("Aim: To")
    if match and match.spans_boundary:
        print("Text spans a revision boundary")

    # Replace works even across revision boundaries
    doc.replace("Aim: To", "Goal: To", paragraph="P1#a7b2")

    doc.save()
```

### Batch Editing

Apply multiple edits atomically with upfront hash validation:

```python
from docx_editor import Document, EditOperation

with Document.open("contract.docx", author="Editor") as doc:
    refs = doc.list_paragraphs()
    doc.batch_edit([
        EditOperation(action="replace", find="old", replace_with="new", paragraph="P2#f3c1"),
        EditOperation(action="delete", text="remove this", paragraph="P5#d4e5"),
    ])
    doc.save()
```
