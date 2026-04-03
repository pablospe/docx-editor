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
    # Step 1: List paragraphs with hash-anchored references
    for p in doc.list_paragraphs():
        print(p)
    # Output: P1#a7b2| Introduction to the contract...
    #         P2#f3c1| The committee shall review...

    # Step 2: Edit using paragraph references (safe, unambiguous)
    doc.replace("old text", "new text", paragraph="P2#f3c1")
    doc.delete("text to delete", paragraph="P5#d4e5")
    doc.insert_after("anchor", "new text", paragraph="P3#b2c4")
    doc.insert_before("anchor", "prefix", paragraph="P3#b2c4")

    doc.save()  # Overwrites original
    # or doc.save("reviewed.docx")  # Save to new file
# Workspace is cleaned up automatically on normal exit
# On exception, workspace is preserved for inspection
```

Without context manager:

```python
doc = Document.open("contract.docx", author=author)
refs = doc.list_paragraphs()
# ... edits using paragraph references ...
doc.save()
doc.close()
```

### Track Changes API

```python
from docx_editor import Document
import os

author = os.environ.get("USER") or "Reviewer"
doc = Document.open("document.docx", author=author)

# List paragraphs to get hash-anchored references
for p in doc.list_paragraphs():
    print(p)
# Output: P1#a7b2| Introduction...
#         P2#f3c1| The payment term is 30 days...
#         P3#b2c4| Section 3. Terms and conditions...

# Find text (returns TextMapMatch or None, works across element boundaries)
match = doc.find_text("30 days")

# Get all visible text (inserted text included, deleted text excluded)
visible = doc.get_visible_text()

# Replace text (creates tracked deletion + insertion)
doc.replace("30 days", "60 days", paragraph="P2#f3c1")

# Delete text (creates tracked deletion)
doc.delete("unnecessary clause", paragraph="P5#d4e5")

# Insert text (creates tracked insertion)
doc.insert_after("Section 3.", " Additional terms apply.", paragraph="P3#b2c4")
doc.insert_before("Section 3.", "See also: ", paragraph="P3#b2c4")

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

# List paragraphs to get hash-anchored references
for p in doc.list_paragraphs():
    print(p)

# Section 2 changes (using paragraph references from list_paragraphs)
doc.replace("30 days", "60 days", paragraph="P4#a1b2")
doc.replace("January 1, 2024", "March 1, 2024", paragraph="P5#c3d4")

# Section 5 changes
doc.delete("and any affiliates", paragraph="P12#e5f6")
doc.insert_after("termination.", " Notice must be provided in writing.", paragraph="P14#g7h8")

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

### Hash-Anchored Paragraph References

The `list_paragraphs()` method returns stable, hash-based paragraph references that eliminate ambiguity when targeting text. Each reference includes a paragraph number and a content hash:

```python
from docx_editor import Document
import os

author = os.environ.get("USER") or "Reviewer"
with Document.open("file.docx", author=author) as doc:
    # Step 1: List paragraphs — each has a unique hash anchor
    for p in doc.list_paragraphs():
        print(p)
    # Output: P1#a7b2| Introduction to the contract...
    #         P2#f3c1| The committee shall review all...
    #         P3#b2c4| The meeting was productive...

    # Step 2: Edit returns the new ref — use it for follow-up edits
    result = doc.replace("the meeting was productive",
                         "the conference was productive",
                         paragraph="P3#b2c4")
    # result.ref = "P3#d5e6" — fresh hash, ready for the next edit
    doc.save()
```

The `paragraph` argument is **required** for all edit methods. If the paragraph content has changed since you called `list_paragraphs()`, a `HashMismatchError` is raised — preventing edits to the wrong location.

**Every edit method returns the new paragraph ref.** Use `.ref` to chain edits without calling `list_paragraphs()` again:

```python
# Chain 3 edits on the same paragraph — no list_paragraphs() between them:
r1 = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
r2 = doc.replace("Manager", "Director", paragraph=r1.ref)
r3 = doc.delete("draft ", paragraph=r2.ref)
# r3.ref is the final hash for paragraph 2
```

### Batch Editing

For multiple independent edits, use `batch_edit()`:

```python
from docx_editor import Document, EditOperation

with Document.open("file.docx", author=author) as doc:
    refs = doc.list_paragraphs()
    results = doc.batch_edit([
        EditOperation(action="replace", find="old term", replace_with="new term", paragraph="P2#f3c1"),
        EditOperation(action="delete", text="remove this", paragraph="P5#d4e5"),
        EditOperation(action="insert_after", anchor="Section 5", text=" (amended)", paragraph="P3#b2c4"),
    ])
    # results[0].ref = "P2#c3d4" — fresh ref for paragraph 2
    doc.save()
```

If any hash is stale, the entire batch is rejected before any edits are applied.

### Paragraph Rewrite (Fallback for Structural Edits)

**Default: always use surgical methods** (`replace`, `delete`, `insert_after`, `insert_before`, `batch_edit`).

**Use `rewrite_paragraph()` only when the edit cannot be decomposed into independent find→replace pairs.** This happens when:
- **Sentence restructuring** — the grammar or clause order changes, not just word swaps
- **Reordering** — words, items, or clauses move to different positions
- **Intertwined changes** — edits overlap or depend on each other so they can't be applied independently

**Use surgical methods when** each change is an independent substitution, even if there are many of them. Five independent word swaps → `batch_edit`, not `rewrite_paragraph`.

**Examples — surgical is correct:**

```python
# Single word swap — use replace():
doc.replace("30", "60", paragraph="P2#f3c1")

# Multiple independent swaps — use batch_edit():
# "CFO" → "Finance Director", "audit committee" → "board", "December 31st" → "January 15th"
doc.batch_edit([
    EditOperation(action="replace", find="CFO", replace_with="Finance Director", paragraph="P5#a7b2"),
    EditOperation(action="replace", find="audit committee", replace_with="board", paragraph="P5#a7b2"),
    EditOperation(action="replace", find="December 31st", replace_with="January 15th", paragraph="P5#a7b2"),
])
```

**Examples — rewrite is correct:**

```python
# Rephrasing (sentence structure changes completely):
# "The committee recommends that the timeline be extended by three months"
# → "The board has approved a three-month extension"
new_ref = doc.rewrite_paragraph("P5#a7b2",
    "The board has approved a three-month extension for further stakeholder review.")
# new_ref = "P5#d6e7" — fresh ref for follow-up edits

# Reordering items in a list:
# "final report, executive summary, and presentation slides"
# → "presentation slides, final report, and executive summary"
new_ref = doc.rewrite_paragraph("P3#c4d5",
    "Deliverables include the presentation slides, final report, and executive summary.")
```

**Batch rewrite** for multiple paragraphs at once:

```python
with Document.open("contract.docx", author=author) as doc:
    refs = doc.list_paragraphs()
    doc.batch_rewrite([
        (refs[1].split("|")[0], "Rephrased paragraph 2 text here."),
        (refs[4].split("|")[0], "Restructured paragraph 5 text here."),
    ])
    doc.save()
```

### Workflow for Large Documents

1. **List paragraphs** with hash-anchored references:
   ```python
   from docx_editor import Document
   doc = Document.open("large-file.docx", author="Reviewer")
   for p in doc.list_paragraphs():
       print(p)
   ```

2. **Identify target paragraphs** by scanning the output for relevant content

3. **Edit with paragraph references** — the hash ensures you target the correct location:
   ```python
   doc.replace("old text", "new text", paragraph="P42#c3d4")
   ```

4. **Verify** with `list_revisions()` if needed

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
