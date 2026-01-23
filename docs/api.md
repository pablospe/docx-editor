# API Reference

## Document

The main entry point for docx-editor. Provides methods for opening documents, making tracked changes, managing comments, and handling revisions.

```python
from docx_editor import Document
```

### Opening Documents

#### `Document.open(path, author=None, force_recreate=False)`

Open a Word document for editing.

**Parameters:**

- `path` (str | Path): Path to the .docx file
- `author` (str, optional): Author name for tracked changes. Defaults to system username.
- `force_recreate` (bool): If True, delete existing workspace and create fresh. Defaults to False.

**Returns:** Document instance ready for editing

**Example:**

```python
doc = Document.open("contract.docx")
doc = Document.open("contract.docx", author="Legal Team")
```

### Properties

#### `author`

Get the author name for tracked changes.

```python
print(doc.author)  # "Legal Team"
```

#### `source_path`

Get the path to the source document.

```python
print(doc.source_path)  # Path("/path/to/contract.docx")
```

### Track Changes Methods

#### `replace(find, replace_with)`

Replace text with tracked changes.

**Parameters:**

- `find` (str): Text to find and replace
- `replace_with` (str): Replacement text

**Returns:** The change ID of the insertion (int)

**Example:**

```python
doc.replace("30 days", "60 days")
```

#### `delete(text)`

Mark text as deleted with tracked changes.

**Parameters:**

- `text` (str): Text to mark as deleted

**Returns:** The change ID of the deletion (int)

**Example:**

```python
doc.delete("obsolete clause")
```

#### `insert_after(anchor, text)`

Insert text after anchor with tracked changes.

**Parameters:**

- `anchor` (str): Text to find as insertion point
- `text` (str): Text to insert after the anchor

**Returns:** The change ID of the insertion (int)

**Example:**

```python
doc.insert_after("Section 5", " (as amended)")
```

#### `insert_before(anchor, text)`

Insert text before anchor with tracked changes.

**Parameters:**

- `anchor` (str): Text to find as insertion point
- `text` (str): Text to insert before the anchor

**Returns:** The change ID of the insertion (int)

**Example:**

```python
doc.insert_before("Section 6", "New clause: ")
```

### Comment Methods

#### `add_comment(anchor_text, comment)`

Add a comment anchored to specific text.

**Parameters:**

- `anchor_text` (str): Text to attach the comment to
- `comment` (str): The comment content

**Returns:** The comment ID (int)

**Example:**

```python
doc.add_comment("Section 5", "Please review this section")
```

#### `reply_to_comment(comment_id, reply)`

Add a reply to an existing comment.

**Parameters:**

- `comment_id` (int): ID of the comment to reply to
- `reply` (str): The reply content

**Returns:** The new comment ID for the reply (int)

**Example:**

```python
doc.reply_to_comment(0, "I agree with this change")
```

#### `list_comments(author=None)`

List all comments in the document.

**Parameters:**

- `author` (str, optional): If provided, filter by author name

**Returns:** List of Comment objects (with replies nested)

**Example:**

```python
comments = doc.list_comments()
for c in comments:
    print(f"{c.author}: {c.text}")
```

#### `resolve_comment(comment_id)`

Mark a comment as resolved.

**Parameters:**

- `comment_id` (int): ID of the comment to resolve

**Returns:** True if resolved, False if not found (bool)

**Example:**

```python
doc.resolve_comment(0)
```

#### `delete_comment(comment_id)`

Delete a comment from the document.

**Parameters:**

- `comment_id` (int): ID of the comment to delete

**Returns:** True if deleted, False if not found (bool)

**Example:**

```python
doc.delete_comment(0)
```

### Revision Management Methods

#### `list_revisions(author=None)`

List all tracked changes in the document.

**Parameters:**

- `author` (str, optional): If provided, filter by author name

**Returns:** List of Revision objects

**Example:**

```python
revisions = doc.list_revisions()
for r in revisions:
    print(f"{r.type}: {r.text} by {r.author}")
```

#### `accept_revision(revision_id)`

Accept a revision by ID.

- For insertions: keeps the inserted content
- For deletions: permanently removes the deleted content

**Parameters:**

- `revision_id` (int): ID of the revision to accept

**Returns:** True if accepted, False if not found (bool)

**Example:**

```python
doc.accept_revision(1)
```

#### `reject_revision(revision_id)`

Reject a revision by ID.

- For insertions: removes the inserted content
- For deletions: restores the deleted content

**Parameters:**

- `revision_id` (int): ID of the revision to reject

**Returns:** True if rejected, False if not found (bool)

**Example:**

```python
doc.reject_revision(1)
```

#### `accept_all(author=None)`

Accept all revisions.

**Parameters:**

- `author` (str, optional): If provided, only accept revisions by this author

**Returns:** Number of revisions accepted (int)

**Example:**

```python
count = doc.accept_all()
print(f"Accepted {count} revisions")
```

#### `reject_all(author=None)`

Reject all revisions.

**Parameters:**

- `author` (str, optional): If provided, only reject revisions by this author

**Returns:** Number of revisions rejected (int)

**Example:**

```python
count = doc.reject_all(author="OtherUser")
```

### Save and Close Methods

#### `save(path=None, validate=False)`

Save the document.

**Parameters:**

- `path` (str | Path, optional): Output path. Defaults to original source path.
- `validate` (bool): If True, validate with LibreOffice before saving. Defaults to False.

**Returns:** Path to the saved document (Path)

**Example:**

```python
doc.save()  # Save to original path
doc.save("contract_v2.docx")  # Save to new path
```

#### `close(cleanup=True)`

Close the document and clean up workspace.

**Parameters:**

- `cleanup` (bool): If True, delete the workspace folder. Defaults to True.

**Example:**

```python
doc.close()  # Clean up workspace
doc.close(cleanup=False)  # Keep workspace for inspection
```

---

## Comment

Represents a document comment.

```python
from docx_editor import Comment
```

### Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `id` | int | The comment ID |
| `text` | str | The comment content |
| `author` | str | The comment author |
| `date` | datetime or None | When the comment was created |
| `resolved` | bool | Whether the comment is resolved |
| `replies` | list[Comment] | Nested replies to this comment |

### Example

```python
comments = doc.list_comments()
for comment in comments:
    print(f"ID: {comment.id}")
    print(f"Text: {comment.text}")
    print(f"Author: {comment.author}")
    print(f"Date: {comment.date}")
    print(f"Resolved: {comment.resolved}")
    print(f"Replies: {len(comment.replies)}")
```

---

## Revision

Represents a tracked change (insertion or deletion).

```python
from docx_editor import Revision
```

### Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `id` | int | The revision ID |
| `type` | str | Either "insertion" or "deletion" |
| `author` | str | The revision author |
| `date` | datetime or None | When the revision was made |
| `text` | str | The inserted or deleted text |

### Example

```python
revisions = doc.list_revisions()
for rev in revisions:
    symbol = "+" if rev.type == "insertion" else "-"
    print(f"{symbol} {rev.text} (by {rev.author})")
```

---

## Exceptions

### `TextNotFoundError`

Raised when the specified text is not found in the document.

```python
from docx_editor.exceptions import TextNotFoundError

try:
    doc.replace("nonexistent text", "new text")
except TextNotFoundError as e:
    print(f"Text not found: {e}")
```

### `CommentError`

Raised when a comment operation fails.

```python
from docx_editor.exceptions import CommentError

try:
    doc.reply_to_comment(999, "reply")
except CommentError as e:
    print(f"Comment error: {e}")
```

### `RevisionError`

Raised when a revision operation fails.

```python
from docx_editor.exceptions import RevisionError
```

### `WorkspaceExistsError`

Raised when attempting to create a workspace that already exists.

```python
from docx_editor.exceptions import WorkspaceExistsError
```

### `WorkspaceSyncError`

Raised when the workspace is out of sync with the source document.

```python
from docx_editor.exceptions import WorkspaceSyncError
```
