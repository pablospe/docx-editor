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

If the source `.docx` was modified outside this library since the workspace was created, `Document.open()` raises `WorkspaceSyncError` instead of silently discarding the workspace. Pass `force_recreate=True` to acknowledge the divergence and re-unpack from the current source:

```python
doc = Document.open("contract.docx", force_recreate=True)
```

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

For large documents, page through paragraphs with `start`/`limit`. You choose the page size; refs keep their global 1-based index (with a page size of 50, page 2 starts at `P51`, not `P1`):

```python
with Document.open("contract.docx") as doc:
    total = doc.paragraph_count()
    page_size = 50
    for start in range(1, total + 1, page_size):
        for paragraph in doc.list_paragraphs(start=start, limit=page_size):
            print(paragraph)
```

Use the `P{index}#{hash}` part as the `paragraph=` argument. Edit methods return a new paragraph ref after the hash changes, so keep the returned value when chaining edits in the same paragraph.

## Track Changes

### Replace Text

```python
with Document.open("contract.docx") as doc:
    ref = "P2#f3c1"
    ref = doc.replace("30 days", "60 days", paragraph=ref)
    doc.save()
```

This creates a tracked deletion of `30 days` and a tracked insertion of `60 days`.

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
        EditOperation(action="replace", find="30 days", replace_with="60 days", paragraph="P2#f3c1"),
        EditOperation(action="delete", text="obsolete clause", paragraph="P5#d4e5"),
        EditOperation(action="insert_after", anchor="Section 5", text=" (as amended)", paragraph="P3#b2c4"),
    ])
    print(new_refs)
    doc.save()
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
doc.close(cleanup=False)    # Keep workspace for inspection
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
