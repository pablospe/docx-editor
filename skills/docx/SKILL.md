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
    - Review received redlines (see "Reviewing Someone Else's Redlines")
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
pip install docx-editor             # editing: track changes, comments, revisions
pip install "docx-editor[create]"   # + python-docx, for creating new documents
pip install "docx-editor[session]"  # + docx-session persistent CLI
```

- **Editing** (track changes, comments, revisions) needs only the base `pip install docx-editor` ([PyPI](https://pypi.org/project/docx-editor/)) — python-docx is NOT bundled
- **Creating** new documents and reading structure uses python-docx, which the `[create]` extra pulls in
- **Session mode** (`docx-session`) needs the `[session]` extra

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
# Workspace is cleaned up automatically on normal exit — WITHOUT saving.
# There is no dirty check: unsaved edits are silently discarded, so always
# call save() before the block ends.
# On exception, the workspace is preserved for inspection in the user cache
# dir (~/.cache/docx-editor/<hash>/ on Linux; error messages print the exact path)
```

Without context manager:

```python
doc = Document.open("contract.docx", author=author)
refs = doc.list_paragraphs()
# ... edits using paragraph references ...
doc.save()   # save first —
doc.close()  # close() deletes the workspace without saving; unsaved edits are lost
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

# Find text (returns SearchResult or None, works across element boundaries).
# Chain into an edit with paragraph=match.paragraph_ref plus
# occurrence=match.paragraph_occurrence (edits count occurrences per paragraph).
match = doc.find_text("30 days")
# Optionally scope to one paragraph — occurrence then counts within it:
match = doc.find_text("30 days", paragraph="P2#f3c1")

# Enumerate EVERY match in one call (returns list[SearchResult], [] if none).
# Each result plugs straight into a follow-up edit — no occurrence math needed.
# Edit them all in one atomic batch; reversed() puts same-paragraph ops in the
# required DESCENDING occurrence order, safe however the matches are spread:
ops = [EditOperation.replace(r.text, "60 days", paragraph=r.paragraph_ref,
                             occurrence=r.paragraph_occurrence)
       for r in reversed(doc.find_all("30 days"))]
doc.batch_edit(ops)
# Optionally scope to one paragraph: doc.find_all("term", paragraph="P2#f3c1")
# One-at-a-time doc.replace() per result also works when every paragraph has
# at most one match. Several matches in ONE paragraph: an edit invalidates the
# paragraph's remaining refs and shifts the occurrence numbers of the matches
# after it — re-run find_all after each edit, or batch in DESCENDING occurrence
# order as above (an edit never shifts the matches before it; ascending
# mis-targets, and descending is not valid for self-overlapping search strings
# like "aa" in "aaaa").

# Get all visible text (inserted text included, deleted text excluded)
visible = doc.get_visible_text()

# Inverse view: deleted text included, inserted text excluded. Read-only —
# refs, hashes, and edits keep working on the visible view. For revisions
# inside paragraphs this equals what reject_all() would leave.
original = doc.get_original_text()

# Structural location: table cell, list item (numId/ilvl; style-inherited
# numbering resolved), heading context, and section index. Base conventions
# are MIXED — read the comments:
loc = doc.get_paragraph_location("P3#b2c4")
if loc.table:  # table.index, row, col are all 1-based
    print(f"table {loc.table.index} r{loc.table.row} c{loc.table.col}")
if loc.list:  # ilvl is 0-based (0 == top level)
    print(f"list numId={loc.list.num_id} level={loc.list.ilvl}")
if loc.outline_level is not None:  # 0-based; 0 == Heading 1; None == body text
    print(f"heading level {loc.outline_level + 1}: style={loc.style}")
print(" > ".join(loc.heading_path))  # e.g. "Chapter one > Termination"
print(loc.section)  # 1-based section index; sectPr-carrying paragraph closes its section

# One-pass batch variant: [(ref, ParagraphLocation), ...] for every paragraph
locations = doc.list_paragraph_locations()

# All edit methods return an EditResult — a str subclass whose value is the
# new paragraph ref. Use it anywhere a ref string is expected:
new_ref = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
doc.replace("net", "gross", paragraph=new_ref)  # chain without list_paragraphs()

# occurrence is 0-based everywhere (0 = first). Edit methods count occurrences
# within the target paragraph. find_text and add_comment count within the
# paragraph when paragraph= is given, and document-wide when paragraph=None.
# find_all bridges the two conventions: each result's paragraph_occurrence is
# already in edit-method coordinates.
#
# OMITTING occurrence means "the target is unique in the search scope".
# If it matches more than once, the edit raises AmbiguousTextError instead of
# silently editing the first match. Pick a match explicitly:
doc.replace("thirty", "sixty", paragraph="P4#a1d2", occurrence=1)  # 2nd "thirty"
doc.replace("thirty", "sixty", paragraph="P4#a1d2", occurrence=0)  # 1st, chosen on purpose

# Delete text (creates tracked deletion)
doc.delete("unnecessary clause", paragraph="P5#d4e5")

# Insert text (creates tracked insertion)
doc.insert_after("Section 3.", " Additional terms apply.", paragraph="P3#b2c4")
doc.insert_before("Section 3.", "See also: ", paragraph="P3#b2c4")

# To accept/reject a specific edit as a unit, use its revision group:
result = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
doc.reject_group(result.group_id)   # undo the whole edit (del + ins together)

doc.save("edited.docx")
doc.close()
```

**Return values:** All edit methods return an `EditResult` — a `str` subclass whose value is the new paragraph reference (e.g., `"P2#c3d4"`); use it for follow-up edits on the same paragraph without calling `list_paragraphs()` again. It also carries `group_id` (the revision group holding every revision the edit created — pass to `accept_group()`/`reject_group()`) and `revision_ids` (the members' change ids). `group_id` is `None` when the edit created no new revisions (e.g. text spliced into one of your own pending insertions). A single edit routinely creates **more than two** revisions: a replace whose span crosses run boundaries (e.g. part of the text is bold) creates one deletion per run plus the insertion, and a rewrite creates one revision per diff hunk — resolve them via `group_id`, never by guessing id pairs.

**Multi-author documents:** Editing inside *another* author's pending insertion preserves their proposal, matching Word: deletions nest a `<w:del>` under your authorship inside their `<w:ins>`, and replacements/insertions put your text in your own sibling `<w:ins>` (splitting theirs when needed) instead of silently rewriting it. `accept_all(author=...)` / `reject_all(author=...)` then resolve each author's changes independently. Only your own pending insertions are edited in place.

**Raises:** `TextNotFoundError` if the text is not found (or the requested `occurrence` is out of range — the error then reports `total_occurrences`). `AmbiguousTextError` if `occurrence` is omitted and the target matches more than once in the search scope. `ValueError` for search/anchor text that is not a non-empty string, replacement/insertion text that is not a string, a non-string `paragraph` ref, or an `occurrence` that is not a non-negative integer — all rejected up front, before any change is made.

### Per-section change report

Group revisions by heading by joining `list_paragraph_locations()` with
`list_revisions()` on `paragraph_ref`. The fixture (python-docx, i.e. the
`[create]` extra) also demonstrates `outline_level`/`section`/`heading_path`:

```python
from docx import Document as DocxDocument            # fixture: [create] extra
from docx.enum.section import WD_SECTION

d = DocxDocument()
d.add_heading("Chapter one", 1)
d.add_heading("Termination", 2)
d.add_paragraph("The term is thirty days.")
d.add_section(WD_SECTION.NEW_PAGE)                   # section 1 -> 2
d.add_heading("Chapter two", 1)
d.add_paragraph("Payment is due in thirty days.")
d.save("report.docx")

from docx_editor import Document
doc = Document.open("report.docx", author="Reviewer")
for r in reversed(doc.find_all("thirty days")):
    doc.replace(r.text, "sixty days", paragraph=r.paragraph_ref,
                occurrence=r.paragraph_occurrence)

locs = dict(doc.list_paragraph_locations())          # snapshot AFTER the edits
report = {}
for r in doc.list_revisions():
    loc = locs.get(r.paragraph_ref)                  # None ref: e.g. table-row revisions
    key = (" > ".join(loc.heading_path) or "(front matter)") if loc else "(unlocated)"
    report.setdefault(key, []).append(r)
for heading, revs in report.items():
    print(f"{heading}: {len(revs)} revision(s)")     # "Chapter one > Termination: 2 ..."
    for r in revs:
        print(f"  [{r.type}] {r.author}: {r.text}")
```

The join works because both APIs report the same hash-anchored refs for the
current state — so take the locations snapshot *after* editing. Keep the
report read-only (`r.type`/`r.author`/`r.text`): deletion `occurrence` values
must not feed the anchor APIs (see the revision API notes below).

### Comments API

```python
from docx_editor import Document
import os

author = os.environ.get("USER") or "Reviewer"
doc = Document.open("document.docx", author=author)

# Add a comment anchored to text (returns the new comment's ID).
# Full signature: add_comment(anchor_text, comment, *, paragraph=None, occurrence=None)
# paragraph=None searches the whole document and counts occurrence document-wide;
# pass paragraph= to scope both the search and the occurrence count to it.
# Like the edit methods, an omitted occurrence requires a unique anchor
# (AmbiguousTextError otherwise).
# The anchor is located with the same visible-text search the edit
# methods use, so anchors spanning run boundaries are found.
cid = doc.add_comment("ambiguous term", "Please clarify this term")
cid2 = doc.add_comment("term", "Anchored to the 2nd 'term' in P3",
                       paragraph="P3#b2c4", occurrence=1)

# Comment IDs are allocated sequentially starting at 0 (in a document with no
# existing comments) — always use the ID returned by add_comment, don't guess.

# List all comments (returns list[Comment] objects)
comments = doc.list_comments()
for c in comments:
    print(f"ID: {c.id}, Author: {c.author}, Text: {c.text}, Resolved: {c.resolved}")
    for reply in c.replies:
        print(f"  Reply: {reply.text}")

# Filter by author
my_comments = doc.list_comments(author="Reviewer")

# Reply to a comment (returns the reply's new comment ID)
doc.reply_to_comment(cid, "I agree, needs clarification")

# Resolve or delete comments (return True if found, False if not)
doc.resolve_comment(cid)
doc.delete_comment(cid2)

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
    # Location: which paragraph the revision lives in, and where in it
    print(f"  at {r.paragraph_ref} occurrence={r.occurrence}")
    # Nesting: e.g. a foreign deletion inside another author's insertion
    print(f"  nested_under={r.nested_under} contains_ids={r.contains_ids}")
    # Group: revisions created by the same edit in this session share it
    print(f"  group_id={r.group_id}")

# Filter by author
their_changes = doc.list_revisions(author="OtherUser")

# Filter to one paragraph (ref from list_paragraphs()). The filter validates
# the ref exactly like the edit methods — a stale ref raises HashMismatchError
# (malformed: ValueError; out of range: ParagraphIndexError) — so re-list
# paragraphs after edits before filtering:
para_changes = doc.list_revisions(paragraph="P3#a7b2")

# For INSERTIONS, r.paragraph_ref/r.occurrence plug straight into the anchor
# APIs (occurrence is 0-based, same convention as replace/delete/add_comment).
# Filter on occurrence is not None — even an insertion can be unlocatable
# (see the None cases below):
r = next(r for r in doc.list_revisions() if r.type == "insertion" and r.occurrence is not None)
doc.add_comment(r.text, "please confirm", paragraph=r.paragraph_ref, occurrence=r.occurrence)

# Accept or reject individual revisions (return True if found, False if not).
# Use IDs from list_revisions() — don't guess numbering.
doc.accept_revision(revisions[0].id)
doc.reject_revision(their_changes[0].id)

# Accept or reject one edit's revisions as a unit (returns count processed).
# Every edit method returns an EditResult carrying group_id; prefer
# groups over per-id calls for rewrite_paragraph — one rewrite creates many
# revisions (one per diff hunk), and accepting only some of them by id
# garbles the paragraph. Raises RevisionError for an unknown group id.
result = doc.rewrite_paragraph("P3#a7b2", "Entirely new paragraph text.")
doc.reject_group(result.group_id)   # undo the whole rewrite, or:
# doc.accept_group(result.group_id) # apply it in full

# Groups are in-memory and per-open-Document: after close()/reopen,
# revisions report group_id=None and old group ids raise RevisionError.
# save() keeps groups alive (the Document stays open). Foreign/pre-session
# revisions never have a group — resolve those by id or author.

# Accept or reject all revisions (returns count of revisions processed)
doc.accept_all()
doc.reject_all()

# Accept/reject only specific author's revisions
doc.accept_all(author="Reviewer")
doc.reject_all(author="OtherUser")

doc.save()
doc.close()
```

**`occurrence` is None when targeting-by-text does not apply**: empty revision
text (paragraph-mark markers), a host insertion whose text was partly consumed
by a nested deletion, or a nested deletion itself (its text never existed in
the original document). Never treat `None` as `0` — in every `None` case,
`nested_under`/`contains_ids` describe the revision instead.

**Deletion occurrences count in the original (pre-revision) text**, where the
deleted span still exists — they locate the deletion for reporting, but must
NOT be passed to replace/delete/add_comment, which search the visible text
(an earlier inserted copy of the same text would shift the count and silently
anchor the wrong span). Only insertion occurrences feed the anchor APIs.

**Nested revisions**: when a reviewer deletes text inside another author's
pending insertion (Word and this library both produce this), the deletion
nests inside the insertion. The host insertion's `text` reports the full text
it originally inserted; `contains_ids` lists the nested revision ids, and the
nested deletion's `nested_under` points back at its host.

- *Accepting* the insertion unwraps it in place — a deletion another author
  nested inside it survives as an independent pending deletion.
- *Rejecting* the insertion removes everything inside it — nested deletions
  disappear with it.
- `accept_all()` / `reject_all()` resolve nesting fully (they re-scan until no
  revisions remain), and `author=` filters process each author's changes
  independently.

Predict the outcome, then verify with `get_markup_text()`.

### Inline markup view (verify redlines without pandoc)

```python
# Every paragraph on one line; revisions wrapped inline:
#   Keep [ins#4:Reviewer]added[/ins][del#3:Reviewer]removed[/del]
# Nesting renders naturally:
#   [ins#1:A]kept [del#9:B]gone[/del][/ins]
print(doc.get_markup_text())
```

A human/agent verification view, not a parseable format (author names are not
escaped; tabs/breaks are not rendered; text inside a drawing's text box
appears both inline and again as its own line, same as `get_text()`).

### Reviewing Someone Else's Redlines

When a document arrives with pending tracked changes from other authors:

**Anchors match visible text.** `find_text`, the edit methods, and
`add_comment` all search the view with pending insertions included and pending
deletions excluded — target the text as it reads with all changes showing
(exactly what `get_visible_text()` returns).

**Predict, then verify.** `get_visible_text()` is the document as it will read
after `accept_all()`; `get_original_text()` is the document as it will read
after `reject_all()` (for revisions inside paragraphs). Snapshot before acting,
compare after:

```python
import os
author = os.environ.get("USER") or "Reviewer"

doc = Document.open("redlined.docx", author=author)
expected = doc.get_visible_text()   # the accept-all outcome
doc.accept_all(author="OtherUser")  # resolve their changes only
assert doc.get_visible_text() == expected
doc.save("resolved.docx")
doc.close()
```

To act on a subset, target revisions by ID: `list_revisions()` (optionally
`author=`- or `paragraph=`-filtered), pick by `.text`/`.type`/`.paragraph_ref`,
then `accept_revision(id)` / `reject_revision(id)`. For edits made in the
current session, prefer `accept_group()`/`reject_group()` with the
`EditResult.group_id` — it resolves everything one edit created, in one call.

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

Refs are **1-based** global indexes (`P1`, not `P0`). For large documents, page through paragraphs to save tokens — `paragraph_count()` gives the total for bounds, and `start`/`limit` return a slice whose refs keep their global index. You pick the page size (there's no built-in limit); with `limit=N` the next page starts at `start + N`. Pass `max_chars=0` to get bare refs (`P1#a7b2`) with no preview text or `| ` separator:

```python
total = doc.paragraph_count()                        # cheap bounds check
page_size = 50                                        # caller's choice
page2 = doc.list_paragraphs(start=1 + page_size, limit=page_size)  # refs P51..P100
refs_only = doc.list_paragraphs(max_chars=0)         # "P1#a7b2", no preview
```

The `paragraph` argument is **required** for all edit methods. If the paragraph content has changed since you called `list_paragraphs()`, a `HashMismatchError` is raised — preventing edits to the wrong location.

**Every edit method returns an `EditResult` — a `str` subclass whose value is the new paragraph ref (plus `group_id`/`revision_ids`).** Chain edits without calling `list_paragraphs()` again:

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
    ops = [
        EditOperation.replace("old term", "new term", paragraph="P2#f3c1"),
        EditOperation.delete("remove this", paragraph="P5#d4e5"),
        EditOperation.insert_after("Section 5", " (amended)", paragraph="P3#b2c4"),
    ]
    new_refs = doc.batch_edit(ops)
    # new_refs[0] = "P2#c3d4" — fresh ref for paragraph 2
    doc.save()
```

Build operations with the typed constructors (`EditOperation.replace/.delete/.insert_after/.insert_before` — same signatures as the `Document` methods). They validate arguments immediately and raise `ValueError` with a field-specific message, instead of failing later at apply time.

**Single-exception contract:** whatever goes wrong with an operation — stale
hash, malformed ref, missing text, ambiguous target — `batch_edit` (and
`batch_rewrite`) raise **only** `BatchOperationError`, with `operation_index`
naming the failing op and `original` (also `__cause__`) holding the underlying
typed exception. The batch is atomic: nothing is applied on failure.

**Pre-flight with `dry_run=True`** — validates every operation without applying
anything and returns `list[EditValidationResult]` (fields: `index`, `paragraph`,
`valid`, `error`), one per operation in input order. The document is left
unchanged (`ops` is the `EditOperation` list built above):

```python
results = doc.batch_edit(ops, dry_run=True)
if all(r.valid for r in results):
    new_refs = doc.batch_edit(ops)
else:
    for r in results:
        if not r.valid:
            print(f"op {r.index} on {r.paragraph}: {r.error}")
```

Each operation is validated independently against the current document —
`dry_run` does **not** simulate the sequential effects of multiple operations
on the same paragraph.

**Multiple operations on the same paragraph** apply sequentially in input
order: each operation's find/anchor text and `occurrence` resolve against the
paragraph's visible text *as left by the previous operations in the batch* (a
tracked delete removes text from that view; an insert adds to it). Across
different paragraphs, operations are applied in reverse document order — a
behavior that keeps one `list_paragraphs()` snapshot valid for the whole batch.

### Error Handling & Recovery

All LLM-facing errors inherit from `DocxEditError` and carry structured fields so you can retry in-loop without re-reading the document. Catch the specific class or the base — both work.

| Error                  | Fields                                                                                  | Recovery                                                                                         |
| ---------------------- | --------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| `HashMismatchError`    | `paragraph_index`, `expected_hash`, `actual_hash`, `paragraph_preview`                  | Retry with `P{paragraph_index}#{actual_hash}`.                                                   |
| `TextNotFoundError`    | `search_text`, `paragraph_ref`, `paragraph_preview`, `occurrence`, `total_occurrences`  | Use `paragraph_preview` to pick a substring that actually appears; if `total_occurrences` is set, retry with an `occurrence` < `total_occurrences`. |
| `AmbiguousTextError`   | `search_text`, `paragraph_ref`, `paragraph_preview`, `total_occurrences`                | The target matched more than once with no `occurrence` given. Enumerate with `find_all()` and pick, or pass an explicit `occurrence` (0-based).      |
| `ParagraphIndexError`  | `index`, `total_paragraphs`                                                             | Clamp to `1..total_paragraphs` or call `list_paragraphs()` to pick a valid ref.                 |
| `BatchOperationError`  | `operation_index`, `reason`, `original`                                                 | Fix the op at `operations[operation_index]` (or drop it) and retry the batch; `original` is the underlying typed error (e.g. use its `actual_hash` to re-target a stale ref), or None for batch-level rules with no underlying exception (a missing paragraph ref, an element that is not an `EditOperation`, or a duplicate paragraph in `batch_rewrite` — `batch_edit` allows repeats, applied sequentially). Batch methods never raise the inner types directly. |
| `RevisionError`        | `revision_id`, `group_id` (set when the error is about that id, else None)              | Unknown `group_id` passed to `accept_group()`/`reject_group()`. Groups are in-memory and per-open-Document: use the `EditResult.group_id` you were handed in this session; after `close()`/reopen resolve by revision id or author instead. |
| `CommentError`         | `comment_id` (set when a comment id was targeted, else None)                            | Replying to a nonexistent comment id — call `list_comments()` and retry with a real id. With `comment_id=None` the arguments themselves were invalid; fix them per the message. |
| `DocumentClosedError`  | `path`                                                                                  | The `Document` was used after `close()`. Reopen with `Document.open(e.path)`; edits not saved before the `close()` were discarded with the workspace. |
| `DocumentNotFoundError`| `path`                                                                                  | The file doesn't exist at `path` — fix the path (typo, wrong cwd) before retrying. |
| `WorkspaceSyncError`   | `workspace_path`, `source_path`                                                         | Workspace and source disagree (unsaved edits from a previous session, or the source changed on disk). **Do not retry blindly** — `force_recreate=True` (open) / `force=True` (save) DISCARDS one side; to rescue the workspace's edits first, save them elsewhere: `Workspace(source, create=False).save("rescued.docx")`. |
| `DocumentOpenError`    | `path`, `owner_file`                                                                    | **Do not retry blindly.** The destination is open in Word. Stop and tell the user to close it. Only pass `force=True` if the user confirms the `~$` lock is stale (crashed session). |

```python
from docx_editor import (
    AmbiguousTextError,
    BatchOperationError,
    HashMismatchError,
    ParagraphIndexError,
    RevisionError,
    TextNotFoundError,
)

try:
    doc.replace("stale text", "new text", paragraph="P3#olda")
except HashMismatchError as e:
    doc.replace("stale text", "new text", paragraph=f"P{e.paragraph_index}#{e.actual_hash}")
except TextNotFoundError as e:
    # e.paragraph_preview shows the current paragraph content for recovery
    ...
except AmbiguousTextError as e:
    # Target matched e.total_occurrences times — enumerate and pick
    r = doc.find_all("stale text", paragraph=e.paragraph_ref)[0]
    doc.replace("stale text", "new text", paragraph=r.paragraph_ref,
                occurrence=r.paragraph_occurrence)
except ParagraphIndexError as e:
    # Clamp to a valid 1-indexed paragraph number (guard the empty-doc case)
    if e.total_paragraphs == 0:
        raise  # no paragraphs to retry against
    safe_idx = max(1, min(e.index, e.total_paragraphs))
    ref = doc.list_paragraphs()[safe_idx - 1].split("|")[0]
    doc.replace("stale text", "new text", paragraph=ref)

# Batch recovery — BatchOperationError is the ONLY exception batch_edit raises
# per failing op, so this loop handles every failure mode:
while ops:
    try:
        doc.batch_edit(ops)
        break
    except BatchOperationError as e:
        if isinstance(e.original, HashMismatchError):
            # Re-target the stale op instead of dropping it
            op = ops[e.operation_index]
            op.paragraph = f"P{e.original.paragraph_index}#{e.original.actual_hash}"
        else:
            ops.pop(e.operation_index)  # drop the failing op and retry
```

**Saving safely (`DocumentOpenError`).** `save()` is atomic — it writes to a temp
file in the destination's directory and renames, so a failed save (including
`validate=True`) never corrupts or deletes the existing document, and the saved
file keeps its original permissions. Before writing it also checks for Word's `~$`
lock file next to the destination and raises `DocumentOpenError` if the document
looks open. This one is **not** an in-loop retry: it means a human has the file
open, so stop and tell the user to close it in Word. Use `e.path` and
`e.owner_file` to tell them exactly which file. `force=True` bypasses the guard and
is only for a confirmed-stale lock left by a crashed session — never reach for it
just to make the error go away.

Two limits worth knowing: the guard sees only *local* locks, so remote co-authoring
in OneDrive/SharePoint or Word-for-the-web leaves no local file and cannot be
detected; and because the save writes a temp file next to the destination, it needs
**write permission on the containing directory**, not just on the document.

A directory-permission problem surfaces as a plain `PermissionError`, not
`DocumentOpenError` — do not tell the user to close Word for that one. So does a
**write-protected document**: `save()` refuses it rather than replacing it, so offer
to save under a new path instead.

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
    EditOperation.replace("CFO", "Finance Director", paragraph="P5#a7b2"),
    EditOperation.replace("audit committee", "board", paragraph="P5#a7b2"),
    EditOperation.replace("December 31st", "January 15th", paragraph="P5#a7b2"),
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

### Session Mode (persistent Python for multi-step editing)

For 3+ operations on the same document — iterative review conversations, large
documents, exploratory editing — use a persistent session instead of one-off
scripts. The document (and all your variables) stays open between commands:

```bash
# Requires: pip install "docx-editor[session]"
docx-session start

# Use ABSOLUTE paths: the kernel keeps the cwd it was started in, which is not
# necessarily the cwd of a later exec. `start` prints the cwd it captured.
docx-session exec "from docx_editor import Document; doc = Document.open('/abs/path/contract.docx', author='Reviewer')"
docx-session exec "paras = doc.list_paragraphs(); print('\n'.join(str(p) for p in paras[:20]))"
docx-session exec "ref = doc.replace('30 days', '45 days', paragraph='P2#f3c1'); ref"
docx-session exec "doc.add_comment('45 days', 'Extended per negotiation.', paragraph=ref)"
docx-session exec "doc.save(); doc.close()"

# Read data out as JSON: eval takes an *expression* and prints one JSON envelope
# on stdout — {"status", "value", "serialized", "stdout", "stderr", "traceback"}.
# Prefer it over exec + print(json.dumps(...)) for structured reads.
docx-session eval "[str(p) for p in doc.list_paragraphs()[:5]]"

# Multi-line or quote-heavy code: '-' reads the code from stdin (exec and eval;
# also use it for expressions starting with '-', which argparse reads as a flag)
docx-session exec - <<'PY'
for p in doc.list_paragraphs():
    if "deadline" in str(p):
        print(p)
PY

# Check whether a session is running: prints "running"/"not running" plus
# pid/state/connection-file detail lines; exit code 0 if running, 3 if not
docx-session status

docx-session stop

# Two documents at once: --session-file (goes AFTER the subcommand; default
# ~/.cache/docx-editor/kernel.json) selects an independent kernel per file.
# Every start/exec/status/stop must pass the same file for its session:
docx-session start --session-file /tmp/kernel-a.json
docx-session start --session-file /tmp/kernel-b.json
docx-session exec --session-file /tmp/kernel-a.json "..."
docx-session stop --session-file /tmp/kernel-a.json
docx-session stop --session-file /tmp/kernel-b.json
```

Rules:

- **Always `docx-session stop` when the editing task is done** — don't leave kernels running.
- Exit code 1 means the code raised: the traceback is on stderr, the session survives — fix the call and continue (introspect with `docx-session exec "import inspect; print(inspect.signature(doc.replace))"` when unsure).
- Exit code 2 means timeout — the kernel is alive and your code is still running. Exit code 3 means no session is running.
- Exit code 4 means the kernel died mid-exec or is unreachable: its state (open documents, variables) is lost. Recover with `docx-session stop` then `start`, and re-open your documents.
- `eval` output: `"serialized": false` means the value wasn't JSON-serializable and `value` holds its `repr` string. On exit 4 the envelope has `"status": "dead"`; on exit 3 (no session) there is no envelope at all.
- `status` reports `state: busy` while an exec is in flight — a busy session is healthy, don't stop/restart it.
- Variables persist between `exec` calls: keep refs returned by edits in Python variables instead of re-running `list_paragraphs()`.
- Use absolute paths inside `exec` — the kernel's cwd is whatever `start` captured.
- A `exec` sent while the kernel is still busy **queues** behind the running one; `--timeout` covers the whole wait. A timeout does not cancel the running code.
- The session is non-interactive: `input()` (and anything reading stdin) raises `StdinNotImplementedError` rather than hanging.
- `doc.save()` raises `WorkspaceSyncError` if the file changed on disk while the session held it open (e.g. the user edited it in Word). Ask the user before retrying with `doc.save(force=True)` — force overwrites their changes.
- A session that saved to a different path (or whose save failed) and never called `doc.close()` leaves the workspace flagged as holding unsaved changes; the next `Document.open()` of the same source raises `WorkspaceSyncError` instead of silently carrying those edits over. `Document.open(path, force_recreate=True)` recovers but **discards** those edits — to rescue them first, save the orphaned workspace to a new file: `from docx_editor.workspace import Workspace; Workspace("contract.docx", create=False).save("rescued.docx")` (deep import — `Workspace` is not exported at package root), then reopen the source with `force_recreate=True`.
- `Document.open()` raises `WorkspaceLockedError` if a live session (another process, or an unclosed `Document` in this one) already holds the document's workspace. Close the other session, or use `Document.open(path, force_recreate=True)` to take the workspace over, discarding its unsaved edits. Stale locks from dead processes are reclaimed automatically.
- Concurrent sessions via `--session-file` must edit *different* documents — the same document still shares one workspace (see "Editing in parallel" below).
- For a single edit, a one-off script is still fine — session mode pays off with repeated operations.

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

**Editing in parallel**: NOT safe for the same document. The workspace is keyed by the document's absolute path in the user cache dir, so two processes editing the same file share one workspace and will overwrite each other. Edit the same document sequentially. Different files never collide (each gets its own workspace), so editing distinct documents in parallel is fine.

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

- **docx_editor**: `pip install docx-editor` (track changes, comments, revisions — editing needs nothing else)
- **python-docx**: `pip install "docx-editor[create]"` (adds python-docx, for reading structure and creating documents)
- **docx-session**: `pip install "docx-editor[session]"` (persistent session CLI)
- **pandoc**: `sudo apt-get install pandoc` (for text extraction to markdown)
- **docx** (npm): `npm install -g docx` (optional, for complex document formatting)
- **LibreOffice**: `sudo apt-get install libreoffice` (for PDF conversion)
- **Poppler**: `sudo apt-get install poppler-utils` (for pdftoppm)
