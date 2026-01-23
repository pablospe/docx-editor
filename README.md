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
- **Comments**: Add, reply, resolve, and delete comments
- **Revision Management**: List, accept, and reject tracked changes
- **Cross-Platform**: Works on Linux, macOS, and Windows
- **No Dependencies**: Only requires `defusedxml` for secure XML parsing

## Installation

```bash
pip install docx-editor
```

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

## Claude Code Skill

This repo includes a skill for [Claude Code](https://claude.ai/claude-code) that enables AI-assisted Word document editing.

This skill extends the [original Anthropic docx skill](https://github.com/anthropics/claude-code/tree/main/.claude/skills/docx) which requires Claude to manually manipulate OOXML. Instead, this skill provides an interface (`docx-editor`) that handles all the complexity—Claude just calls simple Python methods like `doc.replace()` or `doc.add_comment()`, making document editing significantly faster and less error-prone.

### Install the skill

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

---

Repository initiated with [fpgmaas/cookiecutter-uv](https://github.com/fpgmaas/cookiecutter-uv).
