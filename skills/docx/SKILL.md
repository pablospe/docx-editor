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

## Python Library API (Direct Usage)

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
    # Each method returns the new paragraph ref for chaining
    new_ref = doc.replace("old text", "new text", paragraph="P2#f3c1")
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

# All edit methods return the new paragraph ref as a plain string.
# Use it for follow-up edits on the same paragraph:
new_ref = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
doc.replace("net", "gross", paragraph=new_ref)  # chain without list_paragraphs()

# Delete text (creates tracked deletion)
doc.delete("unnecessary clause", paragraph="P5#d4e5")

# Insert text (creates tracked insertion)
doc.insert_after("Section 3.", " Additional terms apply.", paragraph="P3#b2c4")
doc.insert_before("Section 3.", "See also: ", paragraph="P3#b2c4")

# To accept/reject a specific edit, use list_revisions() to get the change ID:
revisions = doc.list_revisions()
doc.accept_revision(revisions[-1].id)

doc.save("edited.docx")
doc.close()
```

**Return values:** All edit methods return the new paragraph reference as a plain `str` (e.g., `"P2#c3d4"`). Use this for follow-up edits on the same paragraph without calling `list_paragraphs()` again. To get change IDs for accept/reject, use `doc.list_revisions()`.

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
# Each edit returns the new ref — chain edits on the same paragraph
r = doc.replace("30 days", "60 days", paragraph="P4#a1b2")
doc.replace("net", "gross", paragraph=r)  # second edit on same paragraph
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
    # returns "P3#d5e6" — fresh hash, ready for the next edit
    doc.save()
```

The `paragraph` argument is **required** for all edit methods. If the paragraph content has changed since you called `list_paragraphs()`, a `HashMismatchError` is raised — preventing edits to the wrong location.

**Every edit method returns the new paragraph ref as a plain string.** Chain edits without calling `list_paragraphs()` again:

```python
# Chain 3 edits on the same paragraph — no list_paragraphs() between them:
r1 = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
r2 = doc.replace("Manager", "Director", paragraph=r1)
r3 = doc.delete("draft ", paragraph=r2)
# r3 is "P2#xxxx" — the final hash for paragraph 2
```

### Batch Editing

For multiple independent edits, use `batch_edit()`:

```python
from docx_editor import Document, EditOperation

with Document.open("file.docx", author=author) as doc:
    refs = doc.list_paragraphs()
    new_refs = doc.batch_edit([
        EditOperation(action="replace", find="old term", replace_with="new term", paragraph="P2#f3c1"),
        EditOperation(action="delete", text="remove this", paragraph="P5#d4e5"),
        EditOperation(action="insert_after", anchor="Section 5", text=" (amended)", paragraph="P3#b2c4"),
    ])
    # new_refs[0] = "P2#c3d4" — fresh ref for paragraph 2
    doc.save()
```

If any hash is stale, the entire batch is rejected before any edits are applied.

### Error Handling & Recovery

All LLM-facing errors inherit from `DocxEditError` and carry structured fields so you can retry in-loop without re-reading the document. Catch the specific class or the base — both work.

| Error                  | Fields                                                                                  | Recovery                                                                                         |
| ---------------------- | --------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| `HashMismatchError`    | `paragraph_index`, `expected_hash`, `actual_hash`, `paragraph_preview`                  | Retry with `P{paragraph_index}#{actual_hash}`.                                                   |
| `TextNotFoundError`    | `search_text`, `paragraph_ref`, `paragraph_preview`, `occurrence`, `total_occurrences`  | Use `paragraph_preview` to pick a substring that actually appears; if `total_occurrences` is set, retry with an `occurrence` < `total_occurrences`. |
| `ParagraphIndexError`  | `index`, `total_paragraphs`                                                             | Clamp to `1..total_paragraphs` or call `list_paragraphs()` to pick a valid ref.                 |
| `BatchOperationError`  | `operation_index`, `reason`                                                             | Fix the op at `operations[operation_index]` (or drop it) and retry the batch.                    |

```python
from docx_editor import (
    BatchOperationError,
    HashMismatchError,
    ParagraphIndexError,
    TextNotFoundError,
)

try:
    doc.replace("stale text", "new text", paragraph="P3#olda")
except HashMismatchError as e:
    doc.replace("stale text", "new text", paragraph=f"P{e.paragraph_index}#{e.actual_hash}")
except TextNotFoundError as e:
    # e.paragraph_preview shows the current paragraph content for recovery
    ...
except ParagraphIndexError as e:
    # Clamp to a valid 1-indexed paragraph number (guard the empty-doc case)
    if e.total_paragraphs == 0:
        raise  # no paragraphs to retry against
    safe_idx = max(1, min(e.index, e.total_paragraphs))
    ref = doc.list_paragraphs()[safe_idx - 1].split("|")[0]
    doc.replace("stale text", "new text", paragraph=ref)
except BatchOperationError as e:
    # Drop or fix the failing op and retry the batch
    ops.pop(e.operation_index)
    doc.batch_edit(ops)
```

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
import os
author = os.environ.get("USER") or "Reviewer"
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
