---
name: docx
description: "Comprehensive document creation, editing, and analysis with support for tracked changes, comments, formatting preservation, and text extraction. When Claude needs to work with professional documents (.docx files) for: (1) Creating new documents, (2) Modifying or editing content, (3) Working with tracked changes, (4) Adding comments, or any other document tasks"
---

# DOCX creation, editing, and analysis

## Overview

A user may ask you to create, edit, or analyze the contents of a .docx file. A .docx file is essentially a ZIP archive containing XML files and other resources. You have different tools and workflows available for different tasks.

## Workflow Decision Tree

```
What do you need to do?
|
+-- Read/Analyze Content
|   Use pandoc for text extraction (see "Reading and analyzing content")
|
+-- Navigate Document Structure (for large docs or precise targeting)
|   Use python-docx to explore before editing (see "Navigating document structure")
|
+-- Create New Document
|   Use python-docx (recommended, simpler)
|   Or docx-js for complex formatting (see "Creating with docx-js")
|
+-- Edit Existing Document
    Use docx_editor Python library (see "Editing an existing Word document")
    - Tracked changes (redlining)
    - Comments (add, reply, resolve)
    - Accept/reject revisions
```

## Reading and analyzing content

### Text extraction

Convert the document to markdown using pandoc. Pandoc provides excellent support for preserving document structure and can show tracked changes:

```bash
# Convert document to markdown with tracked changes
pandoc --track-changes=all path-to-file.docx -o output.md

# Options: --track-changes=accept/reject/all
```

### Raw XML access

For comments, complex formatting, document structure, embedded media, and metadata, unpack the document:

```bash
unzip document.docx -d unpacked/
```

Key file structures:
* `word/document.xml` - Main document contents
* `word/comments.xml` - Comments referenced in document.xml
* `word/media/` - Embedded images and media files
* Tracked changes use `<w:ins>` (insertions) and `<w:del>` (deletions) tags

## Navigating document structure

Use **python-docx** to explore document structure before editing. This is useful for:
- Large documents that won't fit in context
- Finding the right text/context to target for edits
- Understanding document organization

```python
from docx import Document

doc = Document('file.docx')

# List all paragraphs with their styles
for i, p in enumerate(doc.paragraphs):
    print(f"{i}: [{p.style.name}] {p.text[:50]}...")

# Access tables
for t, table in enumerate(doc.tables):
    print(f"Table {t}:")
    for r, row in enumerate(table.rows):
        for c, cell in enumerate(row.cells):
            print(f"  [{r},{c}]: {cell.text[:30]}...")

# Find specific content
for i, p in enumerate(doc.paragraphs):
    if "target text" in p.text:
        print(f"Found at paragraph {i}: {p.text}")
```

## Creating a new Word document

### With python-docx (recommended)

Use **python-docx** for most document creation needs. It's simpler and keeps everything in Python.

```python
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

# Add title
title = doc.add_heading("Document Title", 0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Add paragraphs
doc.add_paragraph("This is body text.")

# Add heading and more content
doc.add_heading("Section 1", 1)
doc.add_paragraph("Section content here.")

# Add a table
table = doc.add_table(rows=2, cols=2)
table.cell(0, 0).text = "Header 1"
table.cell(0, 1).text = "Header 2"
table.cell(1, 0).text = "Data 1"
table.cell(1, 1).text = "Data 2"

# Add page break
doc.add_page_break()

# Save
doc.save("output.docx")
```

### With docx-js (for complex formatting)

For advanced formatting needs (precise spacing, complex table styling, detailed TOC), use **docx-js** (JavaScript/TypeScript).

**Workflow:**
1. **MANDATORY - READ ENTIRE FILE**: Read [`docx-js.md`](docx-js.md) (~350 lines) for syntax, critical formatting rules, and best practices.
2. Create a JavaScript/TypeScript file using Document, Paragraph, TextRun components
3. Export as .docx using Packer.toBuffer()

## Editing an existing Word document

Use the **docx_editor** Python library for all editing operations. It handles tracked changes, comments, and revisions with a simple API.

### Installation

```bash
pip install docx-editor python-docx
```

- **docx-editor**: Track changes, comments, and revisions ([PyPI](https://pypi.org/project/docx-editor/))
- **python-docx**: Reading document structure and creating new documents

### Author Name for Track Changes

**IMPORTANT**: Never use "Claude" or any AI name as the author. Use one of these approaches:

1. **Get system username** (recommended):
   ```python
   import os
   author = os.environ.get("USER") or os.environ.get("USERNAME") or "Reviewer"
   ```

2. **Ask the user** if you need a specific reviewer name

3. **Use "Reviewer"** as a generic fallback

### Basic Usage

```python
from docx_editor import Document
import os

# Get author from system username
author = os.environ.get("USER") or os.environ.get("USERNAME") or "Reviewer"

# Open document (supports context manager)
with Document.open("contract.docx", author=author) as doc:
    # Make changes (automatically tracked)
    doc.replace("old text", "new text")   # Tracked replacement
    doc.delete("text to delete")           # Tracked deletion
    doc.insert_after("anchor", "new text") # Tracked insertion after anchor
    doc.insert_before("anchor", "prefix")  # Tracked insertion before anchor

    doc.save()  # Overwrites original
    # or doc.save("reviewed.docx")  # Save to new file
# Workspace is cleaned up automatically on normal exit
# On exception, workspace is preserved for inspection
```

Without context manager:

```python
doc = Document.open("contract.docx", author=author)
# ... edits ...
doc.save()
doc.close()
```

### Track Changes API

```python
from docx_editor import Document
import os

author = os.environ.get("USER") or "Reviewer"
doc = Document.open("document.docx", author=author)

# Count occurrences before editing (verify uniqueness)
count = doc.count_matches("30 days")
print(f"Found {count} occurrences")

# Find text (returns TextMapMatch or None, works across element boundaries)
match = doc.find_text("30 days")

# Get all visible text (inserted text included, deleted text excluded)
visible = doc.get_visible_text()

# Replace text (creates tracked deletion + insertion)
doc.replace("30 days", "60 days")  # Replaces first occurrence
doc.replace("30 days", "90 days", occurrence=1)  # Replaces second occurrence

# Delete text (creates tracked deletion)
doc.delete("unnecessary clause")
doc.delete("duplicate text", occurrence=2)  # Delete third occurrence

# Insert text (creates tracked insertion)
doc.insert_after("Section 3.", " Additional terms apply.")
doc.insert_before("Section 3.", "See also: ")
doc.insert_after("Section 3.", " (revised)", occurrence=1)  # After second match

doc.save("edited.docx")
doc.close()
```

**Return values:** All track changes methods return an `int` change ID. Returns `-1` when the target text is already inside a tracked insertion (`<w:ins>`), because the edit is done in-place without creating new revision markup.

**Raises:** `TextNotFoundError` if the text is not found.

### Comments API

```python
from docx_editor import Document
import os

author = os.environ.get("USER") or "Reviewer"
doc = Document.open("document.docx", author=author)

# Add a comment anchored to text (returns comment ID)
doc.add_comment("ambiguous term", "Please clarify this term")

# List all comments (returns list[Comment] objects)
comments = doc.list_comments()
for c in comments:
    print(f"ID: {c.id}, Author: {c.author}, Text: {c.text}, Resolved: {c.resolved}")
    for reply in c.replies:
        print(f"  Reply: {reply.text}")

# Filter by author
my_comments = doc.list_comments(author="Reviewer")

# Reply to a comment (returns new comment ID)
doc.reply_to_comment(comment_id=1, reply="I agree, needs clarification")

# Resolve or delete comments (return True if found, False if not)
doc.resolve_comment(comment_id=1)
doc.delete_comment(comment_id=2)

doc.save()
doc.close()
```

### Revision Management API

```python
from docx_editor import Document
import os

author = os.environ.get("USER") or "Reviewer"
doc = Document.open("reviewed.docx", author=author)

# List all tracked revisions (returns list[Revision] objects)
revisions = doc.list_revisions()
for r in revisions:
    print(f"ID: {r.id}, Type: {r.type}, Author: {r.author}, Text: {r.text}")

# Filter by author
their_changes = doc.list_revisions(author="OtherUser")

# Accept or reject individual revisions (return True if found, False if not)
doc.accept_revision(revision_id=1)
doc.reject_revision(revision_id=2)

# Accept or reject all revisions (returns count of revisions processed)
doc.accept_all()
doc.reject_all()

# Accept/reject only specific author's revisions
doc.accept_all(author="Reviewer")
doc.reject_all(author="OtherUser")

doc.save()
doc.close()
```

## Redlining Workflow (Document Review)

For comprehensive document review with tracked changes:

### Step 1: Analyze the document

```bash
# Get readable text with any existing tracked changes
pandoc --track-changes=all contract.docx -o contract.md
```

Review the markdown to understand document structure and identify needed changes.

### Step 2: Plan your changes

Organize changes by section or type:
- Date changes
- Party name updates
- Term modifications
- Clause additions/removals

### Step 3: Implement changes

```python
from docx_editor import Document
import os

author = os.environ.get("USER") or "Reviewer"
doc = Document.open("contract.docx", author=author)

# Section 2 changes
doc.replace("30 days", "60 days")
doc.replace("January 1, 2024", "March 1, 2024")

# Section 5 changes
doc.delete("and any affiliates")
doc.insert_after("termination.", " Notice must be provided in writing.")

# Add review comments
doc.add_comment("indemnification clause", "Review with counsel")

doc.save("contract-reviewed.docx")
doc.close()
```

### Step 4: Verify changes

```bash
pandoc --track-changes=all contract-reviewed.docx -o verification.md
```

Check that all changes appear correctly in the output.

## Best Practices for AI Editing

### Targeting Specific Text

docx_editor replaces the **first occurrence** of text found. To target specific locations, use unique surrounding context:

```python
# BAD - ambiguous, might match wrong location
doc.replace("the", "a")

# GOOD - unique context ensures correct match
doc.replace("the meeting was productive", "the conference was productive")
```

### Verifying Edits (Critical for Large Documents)

Since you may not read the entire document, your "unique" text might not actually be unique. **Always verify edits**:

**Step 1: Count occurrences BEFORE editing**
```python
from docx_editor import Document
import os

author = os.environ.get("USER") or "Reviewer"
doc = Document.open('file.docx', author=author)
count = doc.count_matches("the meeting was productive")
print(f"Found {count} occurrences")
```

**Step 2: If multiple matches, either add more context OR use occurrence parameter**
```python
# Option A: Use more surrounding context for uniqueness
doc.replace("In Q3, the meeting was productive and resulted",
            "In Q3, the conference was productive and resulted")

# Option B: Target specific occurrence (0-indexed)
doc.replace("the meeting was productive",
            "the conference was productive",
            occurrence=2)  # Replace the 3rd match
```

**Step 3: Verify AFTER editing**
```python
# Check the revision was created in the expected location
revisions = doc.list_revisions()
for r in revisions:
    print(f"{r.type}: '{r.text[:50]}...'")
```

If the edit appears in an unexpected location, reject it and retry with more context or different occurrence.

### Workflow for Large Documents

1. **Explore structure first** with python-docx:
   ```python
   from docx import Document
   doc = Document('large-file.docx')
   for i, p in enumerate(doc.paragraphs):
       if "keyword" in p.text:
           print(f"{i}: {p.text[:80]}...")
   ```

2. **Count occurrences** of your target text to ensure uniqueness

3. **Add surrounding context** if multiple matches exist

4. **Edit with docx_editor** using that unique context

5. **Verify the edit** was made in the correct location

### Complementary Tools

| Task                    | Tool                                            |
| ----------------------- | ----------------------------------------------- |
| Read/navigate structure | python-docx                                     |
| Create new documents    | python-docx (or docx-js for complex formatting) |
| Edit with track changes | docx_editor                                       |
| Comments & revisions    | docx_editor                                       |
| Text extraction         | pandoc                                          |

### Parallel Processing with Subagents

**Reading in parallel**: Safe! Multiple subagents can read the same document simultaneously.

**Pattern for large documents** (map-reduce style):
1. Get document structure with python-docx (paragraph count, headings)
2. Spawn parallel subagents to summarize chunks
3. Main agent reads summaries
4. "Focus" on interesting sections with detailed reads

```
Subagents (parallel):
  - Agent 1: summarize paragraphs 0-100
  - Agent 2: summarize paragraphs 101-200
  - Agent 3: summarize paragraphs 201-300
           ↓
Main agent: reads summaries → identifies interesting section
           ↓
Focus: detailed read of paragraphs 150-180
```

Benefits:
- **Speed**: Parallel reads
- **Small context**: Each agent sees only their chunk
- **Cost-effective**: Use smaller models for simple tasks

**Model recommendations:**

| Task                               | Recommended | Why                           |
| ---------------------------------- | ----------- | ----------------------------- |
| Quick overview / triage            | Haiku       | Fast, cheap, gets main points |
| Standard summarization             | **Sonnet**  | Best quality/cost balance     |
| Detailed document analysis         | Opus        | Catches nuances others miss   |
| Legal/contract review              | Opus        | Every detail matters          |
| Bulk document processing           | Haiku       | Cost-effective at scale       |
| Simple API calls (resolve comment) | Haiku       | Just execution                |

**Key insight**: Sonnet is typically the good default for summarization tasks - good quality without Opus cost. Use Haiku for bulk/speed, Opus when every detail matters.

If unsure, ask the user: "Should I use Opus (best), Sonnet (recommended) or Haiku (faster/cheaper) for this task?"

**Editing in parallel**: NOT safe for the same document. docx_editor uses a shared workspace - concurrent edits will overwrite each other. Edit documents sequentially, or use different files.

### Limitations

- **Text in shapes/text boxes**: May not be accessible via standard paragraph iteration
- **Charts**: Text inside charts is embedded in separate XML, not easily editable
- **Concurrent editing**: Not supported on same document (use sequential access)
- **Most edits**: Are in paragraphs and tables, which are well supported

## Converting Documents to Images

To visually analyze Word documents, convert them to images:

```bash
# Step 1: Convert DOCX to PDF
soffice --headless --convert-to pdf document.docx

# Step 2: Convert PDF pages to JPEG images
pdftoppm -jpeg -r 150 document.pdf page
# Creates: page-1.jpg, page-2.jpg, etc.
```

Options for pdftoppm:
- `-r 150`: Resolution in DPI (adjust for quality/size)
- `-jpeg` or `-png`: Output format
- `-f N`: First page to convert
- `-l N`: Last page to convert

## Code Style Guidelines

When generating code for DOCX operations:
- Write concise code
- Avoid verbose variable names and redundant operations
- Avoid unnecessary print statements

## Dependencies

Required dependencies (install if not available):

- **docx_editor**: `pip install docx-editor` (for track changes, comments, revisions)
- **python-docx**: `pip install python-docx` (for reading structure and creating documents)
- **pandoc**: `sudo apt-get install pandoc` (for text extraction to markdown)
- **docx** (npm): `npm install -g docx` (optional, for complex document formatting)
- **LibreOffice**: `sudo apt-get install libreoffice` (for PDF conversion)
- **Poppler**: `sudo apt-get install poppler-utils` (for pdftoppm)
