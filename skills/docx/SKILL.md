---
name: docx
description: "Comprehensive document creation, editing, and analysis with support for tracked changes, comments, formatting preservation, and text extraction. When Claude needs to work with professional documents (.docx files) for: (1) Creating new documents, (2) Modifying or editing content, (3) Working with tracked changes, (4) Adding comments, or any other document tasks"
---

# DOCX creation, editing, and analysis

## Overview

This skill provides tools for creating, editing, and analyzing .docx files. For editing existing documents, use the **MCP tools** which provide fast, cached operations with tracked changes support.

## Workflow Decision Tree

```
What do you need to do?
|
+-- Read/Analyze Content
|   Use pandoc for text extraction (see "Reading and analyzing content")
|
+-- Navigate Document Structure (for large docs)
|   Use python-docx to explore before editing (see "Navigating document structure")
|
+-- Create New Document
|   Use python-docx (recommended, simpler)
|   Or docx-js for complex formatting (see docx-js.md)
|
+-- Edit Existing Document
    Use MCP tools (see "Editing with MCP Tools")
    - Tracked changes (redlining)
    - Comments (add, reply, resolve)
    - Accept/reject revisions
```

## Editing with MCP Tools

The docx-editor MCP server provides fast document editing with persistent caching. Documents stay loaded between operations, making repeated edits 10-30x faster.

### Document Lifecycle

```
open_document(path, author?)    # Load document into cache
save_document(path)             # Save changes to disk
close_document(path, force?)    # Remove from cache
reload_document(path)           # Discard changes, reload from disk
force_save(path)                # Save even if externally modified
```

**Author handling:**
- On first open without explicit author, uses system username
- The MCP server remembers author for the session
- If a default is used, consider asking the user for their preferred name

### Track Changes Tools

All edits are automatically tracked with author attribution.

```
replace_text(path, old_text, new_text, occurrence?)
delete_text(path, text, occurrence?)
insert_after(path, anchor, text, occurrence?)
insert_before(path, anchor, text, occurrence?)
```

- `occurrence` (0-indexed) selects which match to target when text appears multiple times
- Returns change ID, or -1 if edit was in-place (within existing tracked insertion)

### Comment Tools

```
add_comment(path, anchor_text, comment_text, occurrence?)
list_comments(path, author?)      # Filter by author optional
reply_to_comment(path, comment_id, reply_text)
resolve_comment(path, comment_id)
delete_comment(path, comment_id)
```

### Revision Tools

```
list_revisions(path, author?)     # Filter by author optional
accept_revision(path, revision_id)
reject_revision(path, revision_id)
accept_all(path, author?)         # Accept all, optionally by author
reject_all(path, author?)         # Reject all, optionally by author
```

### Read Tools

```
find_text(path, pattern)          # Find text, returns context
count_matches(path, pattern)      # Count occurrences
get_visible_text(path)            # Get all visible text (insertions included, deletions excluded)
```

### Example: Document Review Workflow

```
# 1. Open the document
open_document("/path/to/contract.docx", author="Pablo")

# 2. Check what we're working with
get_visible_text("/path/to/contract.docx")

# 3. Make changes (all tracked)
replace_text("/path/to/contract.docx", "30 days", "60 days")
delete_text("/path/to/contract.docx", "and any affiliates")
insert_after("/path/to/contract.docx", "termination.", " Notice must be in writing.")

# 4. Add a review comment
add_comment("/path/to/contract.docx", "indemnification", "Review with counsel")

# 5. Save when done
save_document("/path/to/contract.docx")
```

### External Change Detection

The MCP server tracks file modification times. If someone else edits the file while you have it open:
- `save_document` will fail with a warning
- Use `reload_document` to get the latest version (discards your changes)
- Use `force_save` to overwrite external changes (use carefully)

### Best Practices

**Target specific text:** Replace first occurrence by default. Use unique context:
```
# BAD - might match wrong location
replace_text(path, "the", "a")

# GOOD - unique context
replace_text(path, "the meeting was productive", "the conference was productive")
```

**Verify uniqueness first:**
```
count_matches(path, "the meeting was productive")
# If count > 1, add more context or use occurrence parameter
```

**Multiple matches:** Use `occurrence` parameter (0-indexed):
```
replace_text(path, "Section 1", "Article 1", occurrence=2)  # Third match
```

## Reading and analyzing content

### Text extraction

Convert the document to markdown using pandoc:

```bash
# Convert document to markdown with tracked changes
pandoc --track-changes=all path-to-file.docx -o output.md

# Options: --track-changes=accept/reject/all
```

### Raw XML access

For comments, complex formatting, or metadata, unpack the document:

```bash
unzip document.docx -d unpacked/
```

Key files:
* `word/document.xml` - Main document contents
* `word/comments.xml` - Comments referenced in document.xml
* `word/media/` - Embedded images and media

## Navigating document structure

Use **python-docx** to explore large documents before editing:

```python
from docx import Document

doc = Document('file.docx')

# List paragraphs with styles
for i, p in enumerate(doc.paragraphs):
    print(f"{i}: [{p.style.name}] {p.text[:50]}...")

# Find specific content
for i, p in enumerate(doc.paragraphs):
    if "target text" in p.text:
        print(f"Found at paragraph {i}: {p.text}")
```

## Creating a new Word document

### With python-docx (recommended)

```python
from docx import Document
from docx.shared import Pt, Inches

doc = Document()
doc.add_heading("Document Title", 0)
doc.add_paragraph("This is body text.")
doc.add_heading("Section 1", 1)
doc.add_paragraph("Section content here.")
doc.save("output.docx")
```

### With docx-js (for complex formatting)

For advanced formatting, read [`docx-js.md`](docx-js.md) for syntax and best practices.

## Parallel Processing

**Reading in parallel**: Safe. Multiple agents can read simultaneously.

**Editing in parallel**: NOT safe for the same document. MCP maintains document state - concurrent edits will conflict. Edit documents sequentially.

## Converting Documents to Images

```bash
# DOCX to PDF
soffice --headless --convert-to pdf document.docx

# PDF to images
pdftoppm -jpeg -r 150 document.pdf page
# Creates: page-1.jpg, page-2.jpg, etc.
```

## Dependencies

- **MCP server**: Installed with plugin (auto-configured)
- **python-docx**: `pip install python-docx` (for reading/creating)
- **pandoc**: `sudo apt-get install pandoc` (for text extraction)
- **LibreOffice**: `sudo apt-get install libreoffice` (for PDF conversion)
- **Poppler**: `sudo apt-get install poppler-utils` (for pdftoppm)
