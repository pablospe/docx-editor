# Quick Start

This guide covers the essential usage patterns for docx-editor.

## Opening Documents

```python
from docx_editor import Document

doc = Document.open("contract.docx", author="Legal Team")
# ... make changes ...
doc.save()
doc.close()
```

The recommended approach is to use a context manager so the temporary workspace is cleaned up automatically:

```python
from docx_editor import Document

with Document.open("contract.docx", author="Legal Team") as doc:
    # ... make changes ...
    doc.save()
```

If the source `.docx` was modified outside this library since the workspace was created, `Document.open()` raises `WorkspaceSyncError` instead of silently discarding the workspace. The same error is raised when a leftover workspace holds unsaved changes from a previous session — for example, one that saved to a different path (or whose save failed) and never called `close()` — since adopting it would carry those edits into the new session. The error message includes the workspace path. Pass `force_recreate=True` to acknowledge the divergence and re-unpack from the current source:

```python
doc = Document.open("contract.docx", force_recreate=True)
```

## Workspace location

Opening a document unpacks it into a workspace directory. By default this lives under the platform user cache — `~/.cache/docx-editor/` on Linux (or `$XDG_CACHE_HOME`), `~/Library/Caches/docx-editor/` on macOS, `%LOCALAPPDATA%\docx-editor\` on Windows — in a subfolder named after a hash of the document's absolute path.

Override the base directory with the `DOCX_EDITOR_WORKSPACE_DIR` environment variable, or per-call with `workspace_dir=`:

```python
# Keep the workspace next to the document (relative path → resolved against
# the document's directory); useful for inspecting the unpacked XML.
doc = Document.open("contract.docx", workspace_dir=".docx")
```

The default location moved from the old `.docx/<stem>/` folder next to the document. Leftover `.docx/` folders from older versions are no longer used and can be deleted.

## Paragraph References

Tracked edit methods are scoped to a hash-anchored paragraph reference. Start by listing paragraphs:

```python
with Document.open("contract.docx") as doc:
    for paragraph in doc.list_paragraphs():
        print(paragraph)
```

Output looks like:

```text
P1#a7b2| Introduction to the contract...
P2#f3c1| Payment is due within 30 days...
P3#b2c4| Section 5 describes review obligations...
```

A bare `list_paragraphs()` call returns at most 200 paragraphs. When more remain, the last entry is a truncation notice (e.g. `"... 50 more paragraphs; use start=201 or limit=None"`) telling you the next `start`; notice lines always begin with `...` and never look like refs. Refs keep their global 1-based index across pages (page 2 starts at `P201`, not `P1`):

```python
with Document.open("contract.docx") as doc:
    page1 = doc.list_paragraphs()                 # up to P200, then "... N more" notice
    page2 = doc.list_paragraphs(start=201)        # next page, per the notice
    everything = doc.list_paragraphs(limit=None)  # uncapped, never a notice
```

Use the `P{index}#{hash}` part as the `paragraph=` argument. Edit methods return a new paragraph ref after the hash changes, so keep the returned value when chaining edits in the same paragraph.

When the search text appears more than once in the target paragraph, pass `occurrence=` (0-based, so `occurrence=1` is the second match) to `replace()`, `delete()`, `insert_after()`, `insert_before()`, and `add_comment()`. Omitting it requires a unique match — otherwise the call raises `AmbiguousTextError`.

## Track Changes

### Replace Text

```python
with Document.open("contract.docx") as doc:
    ref = "P2#f3c1"
    ref = doc.replace("30 days", "60 days", paragraph=ref)
    doc.save()
```

This creates a tracked deletion of `30` and a tracked insertion of `60`. `replace()` trims the common prefix and suffix and tracks only the differing words, so the unchanged surrounding word (` days`) is left intact rather than deleted and re-inserted.

### Delete Text

```python
with Document.open("contract.docx") as doc:
    ref = "P5#d4e5"
    doc.delete("obsolete clause", paragraph=ref)
    doc.save()
```

### Insert Text

```python
with Document.open("contract.docx") as doc:
    ref = "P3#b2c4"
    ref = doc.insert_after("Section 5", " (as amended)", paragraph=ref)
    doc.insert_before("review", "legal ", paragraph=ref)
    doc.save()
```

### Rewrite a Paragraph

```python
with Document.open("contract.docx") as doc:
    ref = "P2#f3c1"
    doc.rewrite_paragraph(ref, "Payment is due within 60 days after invoice receipt.")
    doc.save()
```

`rewrite_paragraph()` compares the current paragraph text to the desired text and creates tracked changes from the diff.

### Batch Edits

```python
from docx_editor import Document, EditOperation

with Document.open("contract.docx") as doc:
    new_refs = doc.batch_edit([
        EditOperation.replace("30 days", "60 days", paragraph="P2#f3c1"),
        EditOperation.delete("obsolete clause", paragraph="P5#d4e5"),
        EditOperation.insert_after("Section 5", " (as amended)", paragraph="P3#b2c4"),
    ])
    print(new_refs)
    doc.save()
```

To validate a batch before applying it, pass `dry_run=True`. This checks every operation up front and returns a list of `EditValidationResult` objects (each with `valid` and `error`) without modifying the document:

```python
with Document.open("contract.docx") as doc:
    results = doc.batch_edit([
        EditOperation.replace("30 days", "60 days", paragraph="P2#f3c1"),
        EditOperation.delete("obsolete clause", paragraph="P5#d4e5"),
    ], dry_run=True)
    for r in results:
        print(r.index, r.valid, r.error)
```

## Comments

### Add a Comment

```python
comment_id = doc.add_comment("Section 5", "Please review this section")
print(f"Created comment with ID: {comment_id}")
```

### Reply to a Comment

```python
reply_id = doc.reply_to_comment(comment_id=0, reply="I agree with this change")
```

### List Comments

```python
comments = doc.list_comments()
for comment in comments:
    print(f"Comment {comment.id}: {comment.text}")
    print(f"  Author: {comment.author}")
    print(f"  Resolved: {comment.resolved}")

    for reply in comment.replies:
        print(f"  Reply: {reply.text}")
```

Filter by author:

```python
my_comments = doc.list_comments(author="Legal Team")
```

### Resolve or Delete a Comment

```python
doc.resolve_comment(comment_id=0)
doc.delete_comment(comment_id=0)
```

## Revision Management

### List Revisions

```python
revisions = doc.list_revisions()
for rev in revisions:
    print(f"{rev.type}: '{rev.text}' by {rev.author}")
```

Filter by author:

```python
my_revisions = doc.list_revisions(author="Legal Team")
```

### Accept or Reject a Revision

```python
# For insertions, accept keeps the inserted content.
# For deletions, accept permanently removes the deleted content.
doc.accept_revision(revision_id=1)

# For insertions, reject removes the inserted content.
# For deletions, reject restores the deleted content.
doc.reject_revision(revision_id=1)
```

### Accept or Reject a Group or Changeset

Revisions form three nested tiers: an individual **revision**, the **group** of revisions from one logical edit (e.g. a single `replace()` or `rewrite_paragraph()`), and the **changeset** of all groups from one whole call (`batch_edit()` / `batch_rewrite()`). Accept or reject a whole tier in one call:

```python
from docx_editor import EditOperation

# Every edit method returns an EditResult carrying group_id and changeset_id.
result = doc.rewrite_paragraph("P2#f3c1", "Payment is due within 90 days.")
doc.accept_group(result.group_id)           # resolve one logical edit as a unit

# One batch_edit is one changeset that may span several groups:
refs = doc.batch_edit([
    EditOperation.replace("30 days", "60 days", paragraph="P2#f3c1"),
    EditOperation.delete("obsolete clause", paragraph="P5#d4e5"),
])
doc.reject_changeset(refs[0].changeset_id)  # undo the whole batch in one call
```

Every `Revision` also carries `group_id`/`group_source` and `changeset_id`/`changeset_source` — the source is `"recorded"` for edits made in this session and `"inferred"` for revisions reconstructed when an existing document is reopened. Group and changeset ids are renumbered each time the document is opened, so always resolve them with an id from the current session's `EditResult` or `list_revisions()`.

### Accept or Reject All

```python
count = doc.accept_all()
print(f"Accepted {count} revisions")

count = doc.reject_all(author="OtherUser")
print(f"Rejected {count} revisions")
```

## Saving and Closing

```python
doc.save()                  # Save to original path
doc.save("contract_v2.docx") # Save to a new path
doc.close()                 # Delete workspace
doc.close(cleanup=False)    # Keep workspace (in the cache dir) for inspection
```

## Complete Example

```python
from docx_editor import Document

with Document.open("contract.docx", author="Legal Review") as doc:
    for paragraph in doc.list_paragraphs():
        print(paragraph)

    payment_ref = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
    doc.insert_after("payment terms", " (net 60)", paragraph=payment_ref)
    doc.delete("penalty clause", paragraph="P7#a10b")

    doc.add_comment("Section 5", "Needs legal review")

    for rev in doc.list_revisions():
        print(f"{rev.type}: {rev.text}")

    doc.accept_all(author="Senior Counsel")
    doc.save()
```
