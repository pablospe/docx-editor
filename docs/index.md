# docx-edit

Pure Python library for Word document track changes and comments, without requiring Microsoft Word.

## Features

- **Track Changes**: Replace, delete, and insert text with revision tracking
- **Comments**: Add, reply, resolve, and delete comments
- **Revision Management**: List, accept, and reject tracked changes
- **Cross-Platform**: Works on Linux, macOS, and Windows
- **No Dependencies**: Only requires `defusedxml` for secure XML parsing

## Installation

```bash
pip install docx-edit
```

## Quick Start

```python
from docx_edit import Document

# Open and edit
doc = Document.open("contract.docx")
doc.replace("30 days", "60 days")           # Tracked replacement
doc.insert_after("Section 5", "New clause") # Tracked insertion
doc.delete("obsolete text")                 # Tracked deletion

# Comments
doc.add_comment("Section 5", "Please review")
doc.reply_to_comment(comment_id=0, "Approved")

# Revision management
revisions = doc.list_revisions()
doc.accept_revision(revision_id=1)
doc.reject_all(author="OtherUser")

# Save and close
doc.save()
doc.close()
```

## Context Manager

```python
from docx_edit import Document

with Document.open("contract.docx") as doc:
    doc.replace("old term", "new term")
    doc.save()
# Automatically closes and cleans up workspace
```
