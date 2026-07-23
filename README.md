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
- **Paragraph Location**: `get_paragraph_location(ref)` reports whether a paragraph lives in the body or inside a table cell — with `w:gridSpan`-aware logical column, row, table index, and nesting depth. `list_paragraph_locations()` returns `(ref, location)` for every paragraph in one batch pass, avoiding a per-paragraph table rescan
- **Batch Editing**: Atomic `batch_edit()` with upfront hash validation across all operations
- **Paragraph Rewrite**: `rewrite_paragraph()` with automatic word-level diffing — specify desired text, get fine-grained tracked changes
- **Track Changes**: Replace, delete, and insert text with revision tracking
- **Cross-Boundary Editing**: Find and replace text spanning multiple XML elements and revision boundaries
- **Mixed-State Editing**: Atomic decomposition for text spanning `<w:ins>`/`<w:del>` boundaries
- **Comments**: Add, reply, resolve, and delete comments
- **Revision Management**: List, accept, and reject tracked changes at three granularities — individual revisions, groups (one logical edit), and changesets (one whole `batch_edit`/`batch_rewrite` call); `EditResult` and `Revision` objects carry `group_id` and `changeset_id`
- **Session Mode**: Optional persistent kernel (`docx-session start/exec/eval/status/stop`) keeps documents open across many small commands — ideal for AI agents (`pip install "docx-editor[session]"`)
- **Cross-Platform**: Works on Linux, macOS, and Windows
- **No Dependencies**: Only requires `defusedxml` for secure XML parsing

## Installation

```bash
pip install docx-editor             # editing: track changes, comments, revisions
pip install "docx-editor[create]"   # + python-docx, for creating new documents
pip install "docx-editor[session]"  # + docx-session persistent CLI
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
pip install "docx-editor[create]"
```

### Manual install (alternative)

```bash
# Install dependencies
pip install "docx-editor[create]"

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
    # Repeated text? Pass occurrence= (0-based) to pick which match; omitting it
    # requires a unique match, else AmbiguousTextError.
    doc.replace("the", "The", paragraph="P4#c5d6", occurrence=2)

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

    # List paragraphs to find hash-anchored references. Large documents:
    # a bare call returns at most 200 paragraphs, ending with a
    # "... N more paragraphs; use start=201 or limit=None" notice.
    # Refs stay globally indexed across pages (page 2 starts at P201, not P1).
    page1 = doc.list_paragraphs()                 # up to P200, then the notice
    page2 = doc.list_paragraphs(start=201)        # next page, per the notice
    everything = doc.list_paragraphs(limit=None)  # uncapped, never a notice

    # Find text across element boundaries
    match = doc.find_text("Aim: To")
    if match and match.spans_revision:
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
        EditOperation.replace("old", "new", paragraph="P2#f3c1"),
        EditOperation.delete("remove this", paragraph="P5#d4e5"),
    ])
    doc.save()
```

Pass `dry_run=True` to validate every operation up front without touching the
document — `doc.batch_edit(ops, dry_run=True)` returns a list of
`EditValidationResult` instead of applying the edits.

### Saving into synced folders

`docx-editor` is safe to use inside cloud-synced folders (OneDrive, Dropbox,
Google Drive, iCloud) and while Word is running.

- **Atomic save.** `save()` writes the new document to a temporary file in the
  destination's own directory and promotes it with a single atomic rename, flushed
  to disk. The destination is never observed half-written, so a sync client can
  never upload a torn file. If the write (or `validate=True`) fails, the original
  is left exactly as it was — a failed validation can no longer destroy your
  document. The saved file keeps the original's permissions, and a symlinked
  destination is followed to the file it points at.

  Because the temp file is created next to the destination, saving needs write
  permission on the **containing directory**, not just on the document itself. If
  the directory is read-only, `save()` raises `PermissionError`.

  A **write-protected document is refused**, not silently replaced: `save()` raises
  `PermissionError` if the destination is read-only, even though the rename itself
  would have been permitted by the directory.

  An atomic rename replaces the file's inode, so state bound to the *old* inode
  does not survive it: the saved document keeps its **permissions**, but its
  ownership, POSIX ACLs, extended attributes, and any hardlinks to it do not carry
  over. This is inherent to atomic saving (every editor that writes this way behaves
  the same). If a document depends on an ACL or a hardlink, save to a new path.

- **Open-in-Word guard.** Before writing, `save()` checks for the `~$` owner
  (lock) file Word places next to any open document. If the destination looks
  open, it raises `DocumentOpenError` rather than racing Word's writes:

  ```python
  from docx_editor import Document, DocumentOpenError

  try:
      doc.save()
  except DocumentOpenError:
      # Someone has this document open in Word — close it and retry.
      ...
  ```

  If you are certain the `~$` file is a stale lock left by a crashed session, pass
  `force=True` to save anyway: `doc.save(force=True)`.

- **Limitation — remote co-authoring is undetectable.** The guard only sees a
  *local* `~$` file. A document being edited remotely (OneDrive/SharePoint
  co-authoring, or Word for the web) leaves **no local lock file**, so it cannot
  be detected from the filesystem. In that case, rely on the cloud provider's
  version history to recover if edits collide.
