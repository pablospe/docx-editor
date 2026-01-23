# Quick Start

This guide covers the essential usage patterns for docx-editor.

## Opening Documents

### Basic Open

```python
from docx_editor import Document

doc = Document.open("contract.docx")
# ... make changes ...
doc.save()
doc.close()
```

### With Custom Author

Track changes are attributed to the author name:

```python
doc = Document.open("contract.docx", author="Legal Team")
```

### Force Recreate Workspace

If you need a fresh workspace (discarding any pending changes):

```python
doc = Document.open("contract.docx", force_recreate=True)
```

### Using Context Manager

The recommended approach for automatic cleanup:

```python
from docx_editor import Document

with Document.open("contract.docx") as doc:
    doc.replace("old term", "new term")
    doc.save()
# Workspace is automatically cleaned up
```

## Track Changes

### Replace Text

Replace text with a tracked deletion and insertion:

```python
doc.replace("30 days", "60 days")
```

This creates:

- A tracked deletion of "30 days"
- A tracked insertion of "60 days"

### Delete Text

Mark text as deleted (strikethrough in Word):

```python
doc.delete("obsolete clause")
```

### Insert Text

Insert new text after or before an anchor:

```python
# Insert after a phrase
doc.insert_after("Section 5", " (as amended)")

# Insert before a phrase
doc.insert_before("Section 6", "New clause: ")
```

## Comments

### Add a Comment

Attach a comment to specific text in the document:

```python
comment_id = doc.add_comment("Section 5", "Please review this section")
print(f"Created comment with ID: {comment_id}")
```

### Reply to a Comment

Add a reply to an existing comment thread:

```python
reply_id = doc.reply_to_comment(comment_id=0, reply="I agree with this change")
```

### List Comments

Get all comments in the document:

```python
comments = doc.list_comments()
for comment in comments:
    print(f"Comment {comment.id}: {comment.text}")
    print(f"  Author: {comment.author}")
    print(f"  Resolved: {comment.resolved}")

    # Check for replies
    for reply in comment.replies:
        print(f"  Reply: {reply.text}")
```

Filter by author:

```python
my_comments = doc.list_comments(author="Legal Team")
```

### Resolve a Comment

Mark a comment as resolved:

```python
doc.resolve_comment(comment_id=0)
```

### Delete a Comment

Remove a comment from the document:

```python
doc.delete_comment(comment_id=0)
```

## Revision Management

### List Revisions

Get all tracked changes:

```python
revisions = doc.list_revisions()
for rev in revisions:
    print(f"{rev.type}: '{rev.text}' by {rev.author}")
```

Filter by author:

```python
my_revisions = doc.list_revisions(author="Legal Team")
```

### Accept a Revision

Accept a specific revision by ID:

```python
# For insertions: keeps the inserted content
# For deletions: permanently removes the deleted content
doc.accept_revision(revision_id=1)
```

### Reject a Revision

Reject a specific revision by ID:

```python
# For insertions: removes the inserted content
# For deletions: restores the deleted content
doc.reject_revision(revision_id=1)
```

### Accept/Reject All

Accept or reject all revisions at once:

```python
# Accept all revisions
count = doc.accept_all()
print(f"Accepted {count} revisions")

# Reject all revisions by a specific author
count = doc.reject_all(author="OtherUser")
print(f"Rejected {count} revisions")
```

## Saving and Closing

### Save to Original Path

```python
doc.save()
```

### Save to New Path

```python
doc.save("contract_v2.docx")
```

### Close Without Cleanup

Keep the workspace for inspection:

```python
doc.close(cleanup=False)
```

### Close With Cleanup (Default)

Delete the workspace folder:

```python
doc.close()  # cleanup=True is the default
```

## Complete Example

```python
from docx_editor import Document

# Open document with custom author
with Document.open("contract.docx", author="Legal Review") as doc:
    # Make tracked changes
    doc.replace("30 days", "60 days")
    doc.insert_after("payment terms", " (net 60)")
    doc.delete("penalty clause")

    # Add review comments
    doc.add_comment("Section 5", "Needs legal review")

    # Review existing changes
    for rev in doc.list_revisions():
        print(f"{rev.type}: {rev.text}")

    # Accept changes from a trusted author
    doc.accept_all(author="Senior Counsel")

    # Save changes
    doc.save()
```
