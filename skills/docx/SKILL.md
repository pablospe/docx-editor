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
|   CAUTION: clean documents only — python-docx is blind to tracked changes
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

> **Redlined documents: python-docx is blind to tracked changes.**
> `paragraph.text` silently drops everything inside `<w:ins>`/`<w:del>` — a
> redlined document reads as mangled prose with no error. If the document has
> (or will receive) tracked changes, read it with docx_editor
> (`list_paragraphs()` / `get_visible_text()`) or `pandoc --track-changes=all`.
>
> **Never carry paragraph indexes between libraries.** python-docx
> `doc.paragraphs` lists body-level paragraphs only (table cells excluded);
> docx_editor `P{i}` refs number every paragraph including table cells — the
> numberings diverge at the first table (a 3000-paragraph doc can have 3027
> `P` refs), so `P{i}` ≠ `paragraphs[i-1]`. Each `ParagraphInfo` says which
> side of that divergence it is on: `info.in_table` is `True` exactly for the
> paragraphs python-docx skips.

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

# Add lists (built-in styles)
doc.add_paragraph("First bullet", style="List Bullet")
doc.add_paragraph("Second bullet", style="List Bullet")
doc.add_paragraph("First numbered item", style="List Number")

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

On a modern Debian/Ubuntu (or any interpreter marked [PEP-668](https://peps.python.org/pep-0668/)
externally-managed), a bare `pip install` into the system Python is refused —
install into a virtualenv:

```bash
python3 -m venv .venv && . .venv/bin/activate   # if this errors "ensurepip is not
                                                # available", first: apt install python3-venv
pip install docx-editor                          # add [create] / [session] as needed

# or with uv:   uv venv && uv pip install docx-editor
# or, for just the docx-session CLI:  pipx install "docx-editor[session]"
```

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
# On exception, the workspace dir (doc.workspace_path) is kept — but it holds
# only state already flushed by save(); unsaved tracked-change edits live in
# memory and are LOST, and the next open() succeeds with no trace of them.
# (add_comment is the exception: the first comment writes comment parts into
# the workspace and flags it, so an unsaved comment can instead leave the
# next open() raising WorkspaceSyncError — recover with force_recreate=True.)
# To keep work when a step fails, catch the exception INSIDE the with block
# and doc.save("rescued.docx") — once the block exits, the document is
# closed. (Swallow the exception and the block exits normally: the workspace
# is cleaned up and the next open() is clean. Re-raise after the rescue save,
# though, and the kept workspace stays flagged as diverged — the next open()
# of the source then raises WorkspaceSyncError; force_recreate=True discards
# it, and rescued.docx already holds your edits.)
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
match = doc.find_text("30 days")
# Optionally scope to one paragraph — occurrence then counts within it:
match = doc.find_text("30 days", paragraph="P2#f3c1")

# Chain straight into an edit: pass the SearchResult itself as the target and it
# supplies the text, the paragraph AND the occurrence. Do NOT also pass
# paragraph=/occurrence= with it (that raises — it would contradict the match):
doc.replace(match, "60 days")
# Every text-targeting method takes one: replace, delete, insert_after,
# insert_before, add_comment, and the EditOperation constructors. The explicit
# spelling stays valid — it is the SAME single edit the line above just made,
# not a follow-up (re-running it would raise HashMismatchError on the now-stale
# match):
#   doc.replace(match.text, "60 days", paragraph=match.paragraph_ref,
#               occurrence=match.paragraph_occurrence)

# Just need the count? doc.count_matches("30 days") returns the document-wide
# occurrence total (no paragraph scope) without building SearchResults.

# Enumerate EVERY match in one call (returns list[SearchResult], [] if none).
# Edit them all in one atomic batch; reversed() puts same-paragraph ops in the
# required DESCENDING occurrence order, safe however the matches are spread:
ops = [EditOperation.replace(r, "60 days") for r in reversed(doc.find_all("30 days"))]
doc.batch_edit(ops)
# Optionally scope to one paragraph: doc.find_all("term", paragraph="P2#f3c1")
# One-at-a-time doc.replace() per result also works when every paragraph has
# at most one match. Several matches in ONE paragraph: an edit invalidates the
# paragraph's remaining refs and shifts the occurrence numbers of the matches
# after it — re-run find_all after each edit, or batch in DESCENDING occurrence
# order as above (an edit never shifts the matches before it; ascending
# mis-targets, and descending is not valid for self-overlapping search strings
# like "aa" in "aaaa").
# SearchResults print compactly — SearchResult(P3#a7b2 occ=0 '30 days') — so
# printing a whole find_all() list is cheap. Each result also carries
# paragraph_index (the int inside paragraph_ref); never string-parse refs.

# Show the paragraphs AROUND a match ("what's the surrounding section?").
# window=2 (default) returns up to 5 ParagraphInfo records, clamped at the
# document edges; each prints as "P{i}#{hash}| full text". Note the re-find:
# the edits above made the earlier `match` stale, and a stale ref raises
# HashMismatchError wherever it is used — search again after editing.
match = doc.find_text("60 days")
for info in doc.context(match.paragraph_ref, window=2):
    print(info)

# Fetch ONE paragraph by number — single-item counterpart to
# list_paragraphs_structured(). 1-based (P1 is index=1), returns ParagraphInfo
# with full untruncated text; ParagraphIndexError when out of range.
# Both get_paragraph() and context() are O(document) despite their fixed-size
# output (they walk every w:p). Fine per lookup; for many paragraphs call
# list_paragraphs_structured(limit=None) ONCE and index that list instead:
info = doc.get_paragraph(match.paragraph_index)

# Every ParagraphInfo also carries the paragraph's cheap structural facts:
print(info.style)          # raw w:pStyle id, e.g. "Heading1" (None = unstyled)
print(info.outline_level)  # 0-based; 0 == Heading 1; None == body text
print(info.in_table)       # True = the paragraph lives in a table cell
# So a table of contents needs ONE call, no list_paragraph_locations():
toc = [(i.outline_level, i.text) for i in doc.list_paragraphs_structured(limit=None)
       if i.outline_level is not None]
# in_table also makes the P-index divergence self-describing: table-cell
# paragraphs DO get P{i} refs (python-docx's doc.paragraphs skips them).
# Table coordinates, list numbering, heading paths and section indexes still
# need get_paragraph_location()/list_paragraph_locations() below.

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
if loc.table:  # table.index, row, col are all 1-based (body tables only)
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

# A whitespace-only fix is written as the minimal revision — a pure insertion
# here, a pure deletion the other way round, never an invisible del+ins pair:
doc.replace("clause  2", "clause 2", paragraph="P5#d4e5")

# To accept/reject a specific edit as a unit, use its revision group:
result = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
doc.reject_group(result.group_id)   # undo the whole edit (del + ins together)

# Attach the WHY to the redline: note= anchors a comment on exactly the
# revisions this edit creates, and result.comment_id is that comment's id.
# The comment is deleted when the edit is resolved — accept OR reject.
result = doc.replace("30 days", "60 days", paragraph="P2#f3c1",
                     note="Aligns with the master agreement (§4.2).")

doc.save("edited.docx")
doc.close()
```

**Return values:** All edit methods return an `EditResult` — a `str` subclass whose value is the new paragraph reference (e.g., `"P2#c3d4"`); use it for follow-up edits on the same paragraph without calling `list_paragraphs()` again. It also carries `group_id` (the revision group holding every revision the edit created — pass to `accept_group()`/`reject_group()`), `revision_ids` (the members' change ids), and `comment_id` (the rationale comment created by `note=`, else `None`). `group_id` is `None` when the edit created no new revisions (e.g. text spliced into one of your own pending insertions). A single edit routinely creates **more than two** revisions: a replace whose (trimmed) span crosses run boundaries (e.g. part of the text is bold) creates one deletion per run plus the insertion, and a rewrite creates one revision per diff hunk — resolve them via `group_id`, never by guessing id pairs.

**Replace granularity:** `replace()` trims words shared by `find` and `replace_with` at either end, so only the changed words are written as revisions — a replace that only adds or only removes words becomes a pure insertion or deletion. The insertion carries the formatting that covers the most characters of the replaced span — runs sharing identical formatting tally together (ties → earliest-seen formatting). **Accepting** a replace that straddled mixed formatting therefore leaves the replacement uniformly in that one majority format, while each deletion run keeps its own original formatting — so a **reject** restores the pre-edit mix. Replacing text with itself is a **no-op**: no revisions are written and the returned `EditResult` equals the input ref with `group_id=None` and `revision_ids=()` — check that triple to detect it.

**Multi-author documents:** Editing inside *another* author's pending insertion preserves their proposal, matching Word: deletions nest a `<w:del>` under your authorship inside their `<w:ins>`, and replacements/insertions put your text in your own sibling `<w:ins>` (splitting theirs when needed) instead of silently rewriting it. `accept_all(author=...)` / `reject_all(author=...)` then resolve each author's changes independently. Only your own pending insertions are edited in place.

**Amending your own pending text:** an edit landing wholly inside one of *your own* pending insertions **amends that insertion** rather than tracking a change against it — the new text is spliced in at the match position, whether it covers part of the insertion or all of it. Your unsaved text was never in the document, so there is nothing to counter-propose: no `<w:del>`/`<w:ins>` pair is written, no new revision is created, and the `EditResult` comes back with `group_id=None` and `revision_ids=()` (with an updated ref). The amended text lives on inside the insertion holding the **end** of the match, which keeps its id and its group — so **to undo an amendment, reject the group of the insertion it amended**, not the amending call's (there isn't one). `list_revisions()` shows that insertion carrying its current, amended text. A match spanning two of your own adjacent insertions consolidates into that one; any insertion it consumed whole is dropped.

**Rationale notes (`note=`):** five edit methods — `replace`, `delete`, `insert_after`, `insert_before`, `rewrite_paragraph` — plus each `EditOperation` of a `batch_edit` take an optional `note=` (`split_paragraph` and `batch_rewrite` do not). The note becomes a Word comment whose range brackets exactly the revisions that edit created, so the reviewer sees the change and the reason for it together; the comment's id comes back on `EditResult.comment_id`. Five rules:

1. **One comment per operation, deduplicated per call.** Twenty ops of one `batch_edit` sharing a note string produce **one** comment (anchored on the first of them), and all twenty results carry that same `comment_id`; twenty different notes produce twenty comments. Two separate calls with the same note are two proposals, so two comments.
2. **The note dies with the revision it explains** — on **accept** as well as reject, through every verb (`accept_revision`/`reject_revision`, `accept_group`/`reject_group`, `accept_changeset`/`reject_changeset`, `accept_all`/`reject_all`). So a pipeline ending in `accept_all()` ships a clean document, not one full of agent rationale, and any replies threaded under a note go with it. The test is whether the revisions still exist, not which call removed them: a later edit that amends the annotated insertion out of existence ends the note too, as does rejecting *another author's* insertion that carried a revision of yours away inside it. `accept_all(author="Someone Else")` leaves notes whose revisions survive it alone. A note's range never points at text nobody changed: when the edit its markers bracket is resolved while another edit it explains is still pending, the whole thread — replies included — moves onto that surviving redline; when nothing it explains can carry a marker any more, the note goes with the anchor it lost. **For rationale that must survive resolution, use `add_comment()`** — plain comments are document-scoped, note comments are revision-scoped.
3. **`comment_id` is an ordinary live comment id** — `reply_to_comment()`, `resolve_comment()` and `delete_comment()` all take it. Deleting it yourself is final and takes the whole thread: the note stops tracking its revisions.
4. **Nothing to anchor warns, never fails.** A no-op `replace`, an amendment to your own pending insertion, an unchanged `rewrite_paragraph`, and a bare `"\n"` split (its only revision is the paragraph mark, which cannot host a comment anchor) create no revision to bracket: the **edit still applies**, `comment_id` is `None`, and an `UnanchoredNoteWarning` names the cause. An operation sharing its note text with a sibling of the same call that *did* anchor it reports that shared `comment_id` instead, and warns nothing. Silence it with `warnings.filterwarnings("ignore", category=UnanchoredNoteWarning)`.
5. **The link is per-session.** The comment survives `save()` as an ordinary OOXML comment, but the note-to-revision link is per-open-`Document` like `group_id` — after a reopen, rejecting the (inferred) group leaves the comment behind; `delete_comment()` removes it.

A note must be a non-empty string with no control characters (`\n` included — a comment body is a single `<w:t>`), and it is validated **before** the edit runs, so a bad note never leaves an applied edit with a dropped rationale. In a batch it fails the whole batch atomically (`BatchOperationError`), and `dry_run=True` reports the row. Note that a `note=` edit writes the comment parts into the workspace, so it carries the same workspace side effect `add_comment()` does (see the workspace note under Basic Usage above).

```python
# A redline that explains itself:
result = doc.replace("30 days", "60 days", paragraph="P2#f3c1",
                     note="Aligns payment terms with the master agreement (§4.2).")

# One rationale over several redlines: one comment, three results carrying its id.
results = doc.batch_edit([
    EditOperation.replace("Manager", "Director", paragraph="P3#a7b2", note="Title change per HR."),
    EditOperation.replace("manager", "director", paragraph="P5#c4d8", note="Title change per HR."),
    EditOperation.delete("(interim)", paragraph="P7#e1f9", note="Title change per HR."),
])
assert len({r.comment_id for r in results}) == 1

doc.accept_all()          # clean deliverable: revisions applied, notes gone
```

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
    doc.replace(r, "sixty days")                     # the match pins ref+occurrence

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

# Resolve or delete comments (return True if found, False if not).
# delete_comment takes the whole thread: replies go with their parent.
doc.resolve_comment(cid)
doc.delete_comment(cid2)

doc.save()
doc.close()
```

### Revision Management API

```python
from docx_editor import Document, EditOperation
import os

author = os.environ.get("USER") or "Reviewer"
doc = Document.open("reviewed.docx", author=author)

# List the document's tracked insertions and deletions (list[Revision]).
# Other revision types are not listed here — see list_unhandled_revisions().
revisions = doc.list_revisions()
for r in revisions:
    print(f"ID: {r.id}, Type: {r.type}, Author: {r.author}, Text: {r.text}")
    # Location: which paragraph the revision lives in, and where in it
    print(f"  at {r.paragraph_ref} occurrence={r.occurrence}")
    # Nesting: e.g. a foreign deletion inside another author's insertion
    print(f"  nested_under={r.nested_under} contains_ids={r.contains_ids}")
    # Group: revisions from the same logical edit share it. group_source
    # says how it was established: "recorded" (edit made in this session)
    # or "inferred" (reconstructed at parse time for revisions already in
    # the file). None only for ungroupable revisions (missing author/date,
    # outside any paragraph, duplicated id, or a mid-session split half
    # of a foreign insertion).
    print(f"  group_id={r.group_id} group_source={r.group_source}")
    # Changeset: the whole edit CALL the group belongs to (see below).
    # changeset_source mirrors group_source ("recorded" vs "inferred").
    print(f"  changeset_id={r.changeset_id} changeset_source={r.changeset_source}")

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

# Accept or reject a whole edit CALL as one changeset (returns count
# processed). The tiers nest: revision < group < changeset. One
# batch_edit/batch_rewrite is ONE changeset that may span several groups, so
# accept_changeset resolves the entire call at once while accept_group resolves
# a single logical edit within it. Every EditResult carries changeset_id
# alongside group_id; Revision objects carry changeset_id/changeset_source with
# the same recorded-vs-inferred semantics as the group fields. changeset ids
# are per-open-Document (renumbered on each open, like group ids). Raises
# RevisionError for an unknown changeset id.
results = doc.batch_edit([
    EditOperation.replace("old", "new", paragraph="P4#1a2b"),
    EditOperation.delete("stale clause", paragraph="P6#3c4d"),
])
doc.reject_changeset(results[0].changeset_id)   # undo the whole call, or:
# doc.accept_changeset(results[0].changeset_id) # apply it in full

# Group ids are in-memory and per-open-Document, renumbered on each open.
# Pre-existing revisions (previous sessions, foreign reviewers, Word
# round-trips) get INFERRED groups reconstructed at parse time: contiguous
# same-paragraph revisions sharing identical w:author + w:date are one
# group (r.group_source == "inferred"; session edits are "recorded"). So
# accept_group/reject_group work after reopen too — but always take the
# group_id from THIS session's list_revisions()/EditResult; a stale id
# from a previous session may resolve to a different group. Own edits
# stamp collision-bumped whole-second dates (two changesets by one author
# never share a second within one open session); all ops of one
# batch_edit/batch_rewrite call share one date (one changeset). Revisions
# from other sources — foreign authors, or own edits from a previous
# session — still merge on identical author + date. save() keeps groups
# alive (the Document stays open).
# Revision ids, unlike group ids, ARE stable: they are the w:id attributes
# stored in the document XML and survive save()/close()/reopen — resolving
# by revision id in a later session is always safe.

# Accept or reject all insertions and deletions. The return value counts the
# revisions processed and behaves as that int in comparisons, arithmetic and
# f-strings; it also carries what could NOT be processed.
result = doc.accept_all()
doc.reject_all()

# ALWAYS check this before telling a human "all changes accepted", and check
# the result of THIS call — .unhandled describes the call it came from.
# accept_all/reject_all resolve w:ins/w:del only. A Word redline whose
# revisions are format changes (w:pPrChange, w:rPrChange) or drag-and-drop
# moves (w:moveFrom/w:moveTo) returns 0 — not because there was nothing to
# do, but because nothing there could be resolved.
if result.unhandled:
    print(result.unhandled_types)   # {'w:moveFrom': 8, 'w:moveTo': 8, ...}
    for row in doc.list_unhandled_revisions():
        print(f"still pending: {row.tag} by {row.author} @{row.paragraph_ref}")
    # Report these to the human as STILL PENDING — the document is not fully
    # adjudicated. An UnhandledRevisionWarning is also emitted.

# Accept/reject only specific author's revisions. On a filtered call
# .unhandled counts only that author's marks.
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
- `accept_all()` / `reject_all()` resolve nesting fully **for insertions and
  deletions** (they re-scan until no `w:ins`/`w:del` remain), and `author=`
  filters process each author's changes independently. Other revision types
  are not resolved at all — see the `result.unhandled` recipe above.

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
escaped; tabs/breaks are not rendered — unlike `get_visible_text()`, where a
tab mark is a `\t`; text inside a drawing's text box does not appear at all,
same as `get_visible_text()` — see Limitations).

### Reviewing Someone Else's Redlines

When a document arrives with pending tracked changes from other authors:

**Anchors match visible text.** `find_text`, the edit methods, and
`add_comment` all search the view with pending insertions included and pending
deletions excluded — target the text as it reads with all changes showing
(exactly what `get_visible_text()` returns, tab marks as `\t` included).

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

To act on a subset, prefer whole groups: `accept_group()`/`reject_group()`
resolves everything one edit created in one call. For this session's edits
take the `EditResult.group_id`; for pre-existing revisions (a reviewer
session over someone else's redlines, or your own after reopen) take the
`group_id` from `list_revisions()` — those groups are inferred at parse time
(`group_source == "inferred"`), so one logical edit still resolves as a
unit. Fall back to per-id `accept_revision(id)` / `reject_revision(id)`
(pick by `.text`/`.type`/`.paragraph_ref`) only for revisions inside a group
you don't want to resolve whole.

**Caveat — group separability doesn't survive a reopen.** Two edits to the
**same paragraph** made in one `batch_edit`/`batch_rewrite` share that call's
single changeset date, so after a save + reopen they reconstruct as **one**
inferred group — `accept_group`/`reject_group` can no longer resolve them apart
(the changeset tier is unaffected: the whole call stays one exact changeset). If
you may need to accept one and reject the other in a *later* session, make them
**separate edit calls** rather than same-paragraph ops in one batch — distinct
calls get distinct collision-bumped dates and so survive reopen as distinct
groups (and changesets). Different-paragraph ops never merge; only same-paragraph
ones do (see the inferred-group rule under Revision Management API).

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

Refs are **1-based** global indexes (`P1`, not `P0`). A bare `list_paragraphs()` call returns at most **200 paragraphs** (the default `limit`). Whenever paragraphs remain beyond the returned window — default or explicit `limit` — the last list entry is a truncation notice instead of a paragraph, e.g. `"... 50 more paragraphs; use start=201 or limit=None"`, telling you the next `start`. Notice lines always begin with `...` and never match the `P{i}#{hash}` ref shape, so filter them with `entry.startswith("...")` when consuming entries as refs. `paragraph_count()` gives the total for bounds; `start`/`limit` return a slice whose refs keep their global index; `limit=None` removes the cap. Pass `max_chars=0` to get bare refs (`P1#a7b2`) with no preview text or `| ` separator:

```python
page1 = doc.list_paragraphs()                        # up to P200, then "... N more" notice
page2 = doc.list_paragraphs(start=201)               # next page, per the notice
refs = [e for e in page1 if not e.startswith("...")]  # drop the notice line
everything = doc.list_paragraphs(limit=None)         # uncapped, never a notice
refs_only = doc.list_paragraphs(max_chars=0, limit=None)  # every "P1#a7b2", no preview, no notice
```

(`list_paragraphs_structured()` has the same 200-record default cap but appends **no notice** — every entry stays a typed `ParagraphInfo` carrying `index`, `ref`, full `text`, plus `style`/`outline_level`/`in_table`. Detect truncation by checking whether the last record's `index` is still below `paragraph_count()` (robust for any `start`), or pass `limit=None`.)

**Reaching the end of the pagination:** a `start` past the last paragraph is a valid request with an empty answer — both listing methods return `[]`, exactly like `lst[500:]` on a 300-item list, and **do not raise**. So `while page := doc.list_paragraphs(start=s, limit=200)` terminates cleanly. (`start=0` or negative *does* raise `ValueError`: those are invalid inputs, not empty ranges.) To know where the end is rather than discovering it, use `paragraph_count()` up front or follow the truncation notice's `start=` hint.

The `paragraph` argument is **required** whenever the target is a plain string — `replace`, `delete`, `insert_after`, `insert_before` (`add_comment` may search document-wide). Pass a `SearchResult` as the target instead and it supplies the ref itself. Either way, if the paragraph content has changed since the ref was minted, a `HashMismatchError` is raised — preventing edits to the wrong location.

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

Build operations with the typed constructors (`EditOperation.replace/.delete/.insert_after/.insert_before` — same signatures as the `Document` methods, including the SearchResult form: `EditOperation.replace(match, "60 days")`). They validate arguments immediately and raise `ValueError` with a field-specific message, instead of failing later at apply time.

**Single-exception contract:** whatever goes wrong with an operation — stale
hash, malformed ref, missing text, ambiguous target — `batch_edit` (and
`batch_rewrite`) raise **only** `BatchOperationError`, with `operation_index`
naming the failing op and `original` (also `__cause__`) holding the underlying
typed exception. The batch is atomic: nothing is applied on failure.

**Pre-flight with `dry_run=True`** — validates every operation without applying
anything and returns `list[EditValidationResult]` (fields: `index`, `paragraph`,
`valid`, `error`, `current_ref`), one per operation in input order. The document
is left unchanged (`ops` is the `EditOperation` list built above):

```python
results = doc.batch_edit(ops, dry_run=True)
if all(r.valid for r in results):
    new_refs = doc.batch_edit(ops)
else:
    for r in results:
        if not r.valid:
            print(f"op {r.index} on {r.paragraph}: {r.error}")
```

`current_ref` is the recovery field for the one failure you can fix
mechanically: a **stale hash**. It holds the ref for that paragraph's *current*
content, so rebuild the operation with it — never regex the hash out of `error`.
It is `None` on every other row (valid ones included), which is also how you
tell a stale hash apart from a missing target:

```python
for r in doc.batch_edit(ops, dry_run=True):
    if r.current_ref:                                  # stale hash → retry
        ops[r.index] = EditOperation.replace("30 days", "60 days", paragraph=r.current_ref)
    elif not r.valid:                                  # text/ref problem → re-search
        print(f"op {r.index} needs attention: {r.error}")
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

### Paragraph Structure (splitting on `\n`)

A `\n` in edit text means a **tracked paragraph split** at that point. There is
no way to embed a literal newline in a paragraph's text — Word paragraphs are
inherently single-line, and a raw `\n` in a run would be an invisible,
unreviewable artifact. So the library gives `\n` its universal meaning:

- **Mechanism.** The first paragraph's paragraph *mark* is flagged as an
  inserted revision; the tail (everything after the split point) moves into a
  new following paragraph as unchanged content. **Accepting** keeps the split;
  **rejecting** removes the mark and **rejoins** the two paragraphs.
- **Where it works.** Any content text: `replace(find, "a\nb")`,
  `insert_after`/`insert_before`, `rewrite_paragraph`, and the same inside a
  `batch_edit` `EditOperation`. Multiple `\n`s make multiple splits.
- **One unit.** A `\n`-containing operation is ONE revision group covering the
  deletion, the inserted runs in every resulting paragraph, and each inserted
  mark — `reject_group` reverts the whole split atomically; one call is still
  one changeset.

```python
# Split a paragraph while replacing text; both halves are tracked.
r = doc.replace("...year.", "...year.\nA new clause follows.", paragraph="P4#a1b2")
r.refs            # ("P4#…", "P5#…") — every resulting paragraph, in order
r.refs[0] == r    # True — the string value is always the first paragraph
doc.reject_group(r.group_id)   # rejoins into the original single paragraph

# Explicit sugar for a pure split (no text change), cut before an anchor:
res = doc.split_paragraph("P4#a1b2", before="However,")
first_ref, second_ref = res.refs
```

**`EditResult.refs`** carries every resulting paragraph ref (length 1 for a
normal edit, ≥2 for a split). A split shifts the index of every later
paragraph, so **re-resolve** stale refs (`list_paragraphs()`/`find_text()`)
before reusing them — this is the same re-resolve discipline that already
applies after any edit.

**Rejected characters.** Every other C0 control character is rejected at all
text inputs with a teaching `ValueError` (`CommentError` for comment text):
carriage return (`\r`), NUL, DEL, etc. — they would enter the document as
invisible literals. Two are special. `\n` means a split. `\t` is how a
`<w:tab/>` mark reads in the visible and original text views (not in
`get_markup_text()`), so **search and anchor text may
contain it** (`find_text("Name\tValue")`, `insert_after("Name\t", "x")`,
`add_comment("a\tb", ...)`), while **content text may not** — nothing writes a
tracked tab. A `replace`/`delete` target containing `\t` raises `ValueError`
(a tab can be matched but not removed yet: edit the text beside it), and
`rewrite_paragraph` needs `new_text` to hold the same number of `\t` as the
paragraph has tab marks — the text between tabs is rewritten segment by
segment, so `rewrite_paragraph(ref, info.text.replace(...))` keeps working on
tab-bearing paragraphs. Tabs count in paragraph hashes, so refs of
tab-bearing paragraphs differ from those computed before tabs were mapped
(refs are session-scoped anyway).

### Error Handling & Recovery

All LLM-facing errors inherit from `DocxEditError` and carry structured fields so you can retry in-loop without re-reading the document. Catch the specific class or the base — both work.

| Error                  | Fields                                                                                  | Recovery                                                                                         |
| ---------------------- | --------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| `HashMismatchError`    | `paragraph_index`, `expected_hash`, `actual_hash`, `paragraph_preview`                  | Retry with `P{paragraph_index}#{actual_hash}`.                                                   |
| `TextNotFoundError`    | `search_text`, `paragraph_ref`, `paragraph_preview`, `occurrence`, `total_occurrences`  | Use `paragraph_preview` to pick a substring that actually appears (a tab mark is spelled `\t` there — search with a real tab); if `total_occurrences` is set, retry with an `occurrence` < `total_occurrences`. |
| `AmbiguousTextError`   | `search_text`, `paragraph_ref`, `paragraph_preview`, `total_occurrences`                | The target matched more than once with no `occurrence` given. Enumerate with `find_all()` and pick, or pass an explicit `occurrence` (0-based).      |
| `ParagraphIndexError`  | `index`, `total_paragraphs`                                                             | Clamp to `1..total_paragraphs` and retry with `get_paragraph(i)`, or call `list_paragraphs()` to pick a valid ref. |
| `BatchOperationError`  | `operation_index`, `reason`, `original`                                                 | Fix the op at `operations[operation_index]` (or drop it) and retry the batch; `original` is the underlying typed error (e.g. use its `actual_hash` to re-target a stale ref), or None for batch-level rules with no underlying exception (a missing paragraph ref, an element that is not an `EditOperation`, or a duplicate paragraph in `batch_rewrite` — `batch_edit` allows repeats, applied sequentially). Batch methods never raise the inner types directly. |
| `RevisionError`        | `revision_id`, `group_id`, `changeset_id` (whichever the error is about is set, the rest None) | Unknown `group_id` passed to `accept_group()`/`reject_group()`, or unknown `changeset_id` passed to `accept_changeset()`/`reject_changeset()`. Group and changeset ids are per-open-Document and renumbered on each open (recorded for this session's edits, inferred for pre-existing revisions): use an id from this session's `EditResult` or `list_revisions()`, never one saved from a previous session. |
| `CommentError`         | `comment_id` (set when a comment id was targeted, else None)                            | Replying to a nonexistent comment id — call `list_comments()` and retry with a real id. With `comment_id=None` the arguments themselves were invalid; fix them per the message. A `\r`/other control character in `anchor_text` or `comment_text`, or a `\t` in `comment_text`, also raises this — strip it (comments hold no newlines either). |
| `ValueError`           | *(builtin — message names the field)*                                                  | Bad argument caught before any mutation: a malformed ref, a bad `occurrence`, a **control character** (`\r`, NUL, …) in any text, a `\t` in content text, or a `\t` in a `replace`/`delete` target. Only `\n` is allowed in content (it means a tracked paragraph split) and only `\t` in search/anchor text (it matches a tab mark). Remove the control character — never smuggle layout as `\t`/`\r`; to edit around a tab, target the text beside it. |
| `DocumentClosedError`  | `path`                                                                                  | The `Document` was used after `close()`. Reopen with `Document.open(e.path)`; edits not saved before the `close()` were discarded with the workspace. |
| `DocumentNotFoundError`| `path`                                                                                  | The file doesn't exist at `path` — fix the path (typo, wrong cwd) before retrying. |
| `InvalidDocumentError` | `path`                                                                                  | The file at `path` is not a valid .docx — wrong suffix, a directory, empty/truncated/not a ZIP, missing word/document.xml, or malformed XML. Not an in-loop retry: the message names which check failed; fix or re-export the input file. |
| `WorkspaceSyncError`   | `workspace_path`, `source_path`                                                         | Workspace and source disagree (unsaved edits from a previous session, or the source changed on disk). **Do not retry blindly** — `force_recreate=True` (open) / `force=True` (save) / `Document.discard_workspace(path)` DISCARDS one side; to rescue the workspace's edits first, save them elsewhere: `Workspace(source, create=False).save("rescued.docx")`. After a crashed script, `Document.discard_workspace(path)` once beats `force_recreate=True` on every open. |
| `WorkspaceLockedError` | `pid`, `lock_path`                                                                      | A live session already holds this document's workspace — another process, or an unclosed `Document` in THIS one. Close it (or stop that process) and retry; `Document.open(path, force_recreate=True)` or `Document.discard_workspace(path)` takes the workspace over but DISCARDS the holder's unsaved edits — confirm the holder is gone first. Locks left by dead processes are reclaimed automatically and never raise. |
| `DocumentOpenError`    | `path`, `owner_file`                                                                    | **Do not retry blindly.** The destination is open in Word. Stop and tell the user to close it. Only pass `force=True` if the user confirms the `~$` lock is stale (crashed session). |
| `DocumentProtectedError` | `path`, `mode`                                                                        | The document enforces Word's *Restrict Editing* with a mode that locks the body text — `mode` is `readOnly`, `forms` or `comments`. **Not an in-loop retry:** the author asked for the content not to be edited, so tell the user, and only pass `Document.open(path, allow_protected=True)` if they confirm it (the protection stays in the saved file; with `mode="comments"`, `add_comment()` is what that mode permits). A document enforcing `trackedChanges`, or one whose protection is switched off, opens normally and never raises — and so does one using Word's *Password to modify* / *Always Open Read-Only* (`w:writeProtection`), which is a different element this guard does not read. |

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
    doc.replace(r, "new text")  # the picked match pins ref + occurrence
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

**The track-changes switch.** A save that leaves a revision you authored also turns
Word's Track Changes switch on in the saved file (`<w:trackRevisions/>` in
`word/settings.xml`). Your redline is visible either way; the switch is what keeps
the *recipient's* own typing tracked, so the next round of edits can still be told
apart from yours. What counts is the document's state, not whether this session
edited: your own pending redline reopened from an earlier session still turns the
switch on, because it is still waiting for a reply. Nothing else is touched: a
document holding no revision of yours — one you did not redline, or one whose
revisions you accepted — is saved with its settings as they were. Pass `doc.save(path, track_changes=False)` to opt out; it never removes a
switch the document already had. A document that turns tracking *off* explicitly
(`<w:trackRevisions w:val="false"/>`) is left exactly as its author configured it and
the save emits a `UserWarning` saying the recipient's edits will not be tracked —
if the user wants tracking on anyway, `save(path, track_changes=True)` overrides it.

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
# on stdout — {"status", "value", "serialized", "stdout", "stderr", "traceback",
# "error"}. Prefer it over exec + print(json.dumps(...)) for structured reads.
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
- Exit code 2 means timeout, in one of two flavours the stderr line names (and `eval`'s JSON envelope reports as `"started"`): **your code ran too long** (`started: true` — "kernel still running"; raise `--timeout` or make the code faster), or **it never left the queue** (`started: false` — "still queued … never started"; the kernel is busy with an *earlier* command, and nothing of yours executed, so re-sending later is safe). Exit code 3 means no session is running.
- **Exit code 0 is not always "it ran".** If a previous command *raised* while your request was queued behind it, the kernel discards yours without executing it. `exec` then exits **0** with no output and `Warning: the kernel discarded this request without running it` on stderr — so read stderr on exit 0 (or check `started` on an `ExecResult`) before believing an edit applied. `eval` exits **1** with `The kernel discarded this request without running it` on stderr and **no JSON envelope**. Either way nothing of yours happened: re-send it.
- Exit code 4 means the kernel died mid-exec or is unreachable: its state (open documents, variables) is lost. Recover with `docx-session stop` then `start`, and re-open your documents.
- `eval` output: library dataclasses (SearchResult, ParagraphInfo, ParagraphLocation, Revision, Comment) arrive as real JSON objects (`"serialized": true`, datetimes as ISO strings, tuples as lists) — access fields directly, never string-parse a repr. `"serialized": false` means the value wasn't JSON-serializable and `value` holds its `repr` string. On exit 1 the envelope carries `"error": {"type", "message", <structured recovery fields — e.g. actual_hash, total_occurrences>}` plus a compact, machine-path-stripped `traceback` (library frames appear as relative `docx_editor/...` paths and the eval line as `<docx-session eval>`; only absolute machine paths are stripped) (`"error"` is null for the rare raise that bypasses the kernel-side capture, e.g. SystemExit — fall back to `traceback`). On exit 4 the envelope has `"status": "dead"`; on exit 3 (no session) there is no envelope at all.
- `eval` of an edit call returns the `EditResult` as a plain JSON *string* (it is a `str` subclass, so its value is just the new ref) — `group_id`/`revision_ids` are lost in transit. Keep the result in a kernel-side variable and eval a dict projection when you need them: `docx-session exec "r = doc.replace(...)"` then `docx-session eval "{'ref': str(r), 'group_id': r.group_id}"`.
- `status` reports `state: busy` while an exec is in flight — a busy session is healthy, don't stop/restart it. The report is a point-in-time snapshot: a sub-second exec can already read `idle` by the time you check, so never conclude from `idle` that code didn't run.
- Variables persist between `exec` calls: keep refs returned by edits in Python variables instead of re-running `list_paragraphs()`.
- **Never one `exec` per edit.** Each `docx-session` call is a fresh CLI round-trip (~250 ms of subprocess + IPC overhead) that dwarfs the edit itself — 50 edits sent as 50 `exec` calls take ~12 s, the same 50 batched into one `exec` is a single round-trip (~0.3 s: the ~250 ms overhead plus the in-kernel loop), ~40x less overhead. Loop inside a single `exec`, or use `batch_edit`.
- **Project fields kernel-side for large reads.** `eval` serializes the whole value to JSON (~150 chars per `SearchResult`), so a big `find_all` returns tens of KB of context. Return only the fields you need — e.g. `docx-session eval "[(r.paragraph_ref, r.paragraph_occurrence) for r in doc.find_all('30 days')]"` — instead of the full objects (same idea as the `EditResult` projection above).
- Use absolute paths inside `exec` — the kernel's cwd is whatever `start` captured.
- A `exec` sent while the kernel is still busy **queues** behind the running one; `--timeout` covers the whole wait. A timeout does not cancel the running code. Tell the two timeouts apart with `started` / the stderr line (see exit code 2 above) before deciding whether re-sending is safe.
- The session is non-interactive: `input()` (and anything reading stdin) raises `StdinNotImplementedError` rather than hanging.
- `doc.save()` raises `WorkspaceSyncError` if the file changed on disk while the session held it open (e.g. the user edited it in Word). Ask the user before retrying with `doc.save(force=True)` — force overwrites their changes.
- A session that saved to a different path (or whose save failed) and never called `doc.close()` leaves the workspace flagged as holding unsaved changes; the next `Document.open()` of the same source raises `WorkspaceSyncError` instead of silently carrying those edits over. Two recoveries, both **discarding** those edits: `Document.open(path, force_recreate=True)` (per-open) or `Document.discard_workspace(path)` (once, then open normally — the fix for a crashed script, since it also clears a leftover lock and returns `False` when there was nothing to delete, so it is safe to call unconditionally). To rescue the edits first, save the orphaned workspace to a new file: `from docx_editor.workspace import Workspace; Workspace("contract.docx", create=False).save("rescued.docx")` (deep import — `Workspace` is not exported at package root), then discard.
- `Document.open()` raises `WorkspaceLockedError` if a live session (another process, or an unclosed `Document` in this one) already holds the document's workspace. Close the other session, or use `Document.open(path, force_recreate=True)` / `Document.discard_workspace(path)` to take the workspace over, discarding its unsaved edits — make sure the holder is really gone first. Stale locks from dead processes are reclaimed automatically.
- Concurrent sessions via `--session-file` must edit *different* documents — a second session opening the same document raises `WorkspaceLockedError` (see "Editing in parallel" below).
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

**Editing in parallel**: NOT possible for the same document — the workspace is keyed by the document's absolute path and advisory-locked: while a live session holds it, a second `Document.open()` raises `WorkspaceLockedError` naming the holder's `pid` and `lock_path` (no silent clobbering; the holder can be another process or an unclosed `Document` in this same process). Edit sequentially, or take the workspace over with `Document.open(path, force_recreate=True)` — that DISCARDS the holder's unsaved edits. Locks left by dead processes are reclaimed automatically. Different files never collide (each gets its own workspace), so editing distinct documents in parallel is fine.

### Limitations

**What this library will not do.** It redlines and reviews documents that already
exist, and deliberately stops there — so when a request falls outside, reach for
the right tool instead of asking for a feature that is not coming. **Redaction** is
a permanent refusal: a tool built to preserve history cannot honestly promise
removal. **Creating documents and editing structure** — tables, images, styles,
sections, TOCs, fields, content controls, mail merge — belongs to
[python-docx](https://python-docx.readthedocs.io/), and format conversion to
[pandoc](https://pandoc.org/). **Regex find/replace** is out (matching runs over
text that is deliberately anchored to runs and revisions here), and so is
**diffing two arbitrary documents**: the version that fits this domain is
`list_revisions()` plus `get_visible_text()`/`get_original_text()`, and
`rewrite_paragraph()`'s own word-level diff. There is no **schema validator**
either — the contract that matters is "opens in Word with zero repair prompts",
which is tested by round-tripping real documents. Text in shapes and text boxes is
excluded rather than half-editable, and headers, footers, footnotes and endnotes
wait for a real demand.

- **Text in shapes/text boxes**: Excluded, deliberately — text boxes are not an editing surface. Text-box content (`w:txbxContent`) appears in no paragraph listing, no text view, no search and no paragraph hash, and no ref addresses it. Word normally stores a box twice (an `mc:Choice` copy and an `mc:Fallback` copy), and a correct edit would have to write both copies, so an addressable box paragraph would let one write update a single copy and desynchronize the pair. The exclusion is uniform either way — a box stored once is excluded too. To read a box's text, go through HTML: `soffice --headless --convert-to html file.docx` then `pandoc file.html -t plain` (pandoc may render a `[ShapeN]` label beside a box's text, from the placeholder LibreOffice exports for a named shape — ignore those). Its `txt:Text` filter and pandoc reading the `.docx` directly both drop text boxes silently, so neither tells you anything is missing. A revision *inside* a box is still listed with `paragraph_ref`/`occurrence` left `None`, and `accept_all()`/`reject_all()` always resolve it. Anything narrower depends on the storage: a twice-stored box lists the revision once per copy, so one `accept_revision()`/`accept_group()` call resolves only one of them; copies with distinct ids and identical author/date share an inferred changeset that `accept_changeset()` resolves, while copies sharing a `w:id` are ungroupable (`group_id` and `changeset_id` both `None`), so no group- or changeset-keyed call reaches them and `accept_all()`/`reject_all()` is the single call that takes both. Because an all-text-box file (a poster, a flyer, a certificate) therefore reads as blank — `get_visible_text()` returns only the separators between its host paragraphs — check `doc.has_textbox_content` before reporting a document as having no text: `if not doc.get_visible_text().strip() and doc.has_textbox_content:`.
- **Tabs**: A tab mark (`<w:tab/>`) is one `\t` in the visible and original text views, in search and in hashes (`get_markup_text()` does not render it), and edits may land on either side of it — but no edit writes, deletes or replaces one yet: `replace`/`delete` targets containing `\t` are rejected, `rewrite_paragraph` must keep the paragraph's tab marks (same count; the text between them is what changes), and `w:br`/`w:ptab` are not mapped at all (ISSUES.md #6).
- **Charts**: Text inside charts is embedded in separate XML, not easily editable
- **Concurrent editing**: Not supported — a second open of the same document raises `WorkspaceLockedError`; use sequential access
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
