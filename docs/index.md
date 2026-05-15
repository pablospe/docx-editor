# docx-editor

Pure Python library for Word document track changes and comments, without requiring Microsoft Word.

> **Note:** The PyPI package is named `docx-editor` because `docx-edit` was too similar to an existing package.

## Features

- **Hash-Anchored Paragraph References**: target edits with paragraph refs like `P2#f3c1`
- **Batch Editing**: apply multiple paragraph-scoped edits with upfront hash validation
- **Paragraph Rewrite**: rewrite a paragraph and generate tracked changes from the diff
- **Track Changes**: Replace, delete, and insert text with revision tracking
- **Comments**: Add, reply, resolve, and delete comments
- **Revision Management**: List, accept, and reject tracked changes
- **Cross-Boundary Editing**: Find and replace text spanning multiple XML elements
- **Cross-Platform**: Works on Linux, macOS, and Windows
- **No Dependencies**: Only requires `defusedxml` for secure XML parsing

## Installation

```bash
pip install docx-editor
```

## Quick Start

```python
from docx_editor import Document

with Document.open("contract.docx", author="Legal Team") as doc:
    # List paragraphs to get hash-anchored references.
    for paragraph in doc.list_paragraphs():
        print(paragraph)
    # Example output:
    # P1#a7b2| Introduction to the contract...
    # P2#f3c1| Payment is due within 30 days...

    # Edit methods require a paragraph ref and return the updated ref.
    ref = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
    ref = doc.insert_after("Payment", " terms", paragraph=ref)
    doc.delete("obsolete text", paragraph="P5#d4e5")

    # Comments and revision management.
    doc.add_comment("Section 5", "Please review")
    revisions = doc.list_revisions()
    if revisions:
        doc.accept_revision(revisions[0].id)

    doc.save()
```

## Context Manager

```python
from docx_editor import Document

with Document.open("contract.docx") as doc:
    ref = doc.list_paragraphs()[0].split("|", 1)[0]
    doc.replace("old term", "new term", paragraph=ref)
    doc.save()
# Automatically closes and cleans up workspace
```
