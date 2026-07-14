# API Reference

## Document

The main entry point for docx-editor. Provides methods for opening documents, making tracked changes, managing comments, and handling revisions.

```python
from docx_editor import Document
```

### Opening Documents

#### `Document.open(path, author=None, force_recreate=False, workspace_dir=None)`

Open a Word document for editing.

**Parameters:**

- `path` (str | Path): Path to the .docx file
- `author` (str, optional): Author name for tracked changes. Defaults to system username.
- `force_recreate` (bool): If True, delete any existing workspace (stale or in-sync) before opening — whatever XML it holds is discarded — and re-unpack from the current source. Use this to recover from `WorkspaceSyncError`. Defaults to False.
- `workspace_dir` (str | Path, optional): Base directory for the workspace. Overrides the `DOCX_EDITOR_WORKSPACE_DIR` environment variable and the platform cache default (see [Workspace location](#workspace-location)). A relative path resolves against the document's directory, so `workspace_dir=".docx"` keeps the workspace next to the file. Defaults to None.

**Returns:** Document instance ready for editing

**Raises:**

- `WorkspaceSyncError`: If the source `.docx` was modified since the workspace was created, or if a leftover workspace holds unsaved changes from a previous session (one that saved to a different path, or whose save failed, and never closed). Pass `force_recreate=True` to discard the workspace and re-unpack from the current source. The workspace is never deleted silently. The error message includes the workspace path.
- `WorkspaceError`: If the workspace directory cannot be created (e.g. the base is not writable), the home directory backing the default cache cannot be determined, or an existing workspace was unpacked from a different document. The message names the override to set.

**Example:**

```python
doc = Document.open("contract.docx")
doc = Document.open("contract.docx", author="Legal Team")
```

#### Workspace location

When you open a document, its unpacked OOXML contents are stored in a workspace directory. By default this lives under the platform user cache, in a subfolder `docx-editor/<hash>` where `<hash>` is derived from the document's absolute path:

| Platform | Default base directory |
| --- | --- |
| Linux | `$XDG_CACHE_HOME/docx-editor` (falls back to `~/.cache/docx-editor`) |
| macOS | `~/Library/Caches/docx-editor` |
| Windows | `%LOCALAPPDATA%\docx-editor` (falls back to `~\AppData\Local\docx-editor`) |

To override the location:

- Set the `DOCX_EDITOR_WORKSPACE_DIR` environment variable to a base directory, or
- Pass `workspace_dir=` to `Document.open()` (takes precedence over the environment variable).

Both overrides are tilde-expanded, and an empty value counts as unset. A **relative** override resolves against the document's directory, so `workspace_dir=".docx"` reproduces the old next-to-file layout (handy for debugging). An **absolute** override is used as-is. (A relative `XDG_CACHE_HOME` / `%LOCALAPPDATA%` is ignored, per the XDG spec.)

Cleanup semantics are unchanged: the workspace persists until `close()` is called, and `close(cleanup=False)` preserves it for inspection. `close()` removes only the document's own workspace folder — the base directory is shared and is never deleted, so it is safe to point `workspace_dir` at a directory you also use for other things.

The workspace directory is created with owner-only permissions (`0o700`), since it holds the document's plaintext in a shared cache location. Use `doc.workspace_path` to locate it.

> **Note:** The default location moved from the old `.docx/<stem>/` folder next to the document to the platform cache. Workspaces created by older versions are no longer found and are simply ignored. Delete leftover `.docx/` folders, or pass `workspace_dir=".docx"` to keep using the old layout.

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

#### `workspace_path`

Get the path to this document's workspace folder. Since the workspace lives in the user cache by default, this is how you locate the unpacked XML — for example after `close(cleanup=False)`, or when a workspace was preserved because an exception was raised.

```python
print(doc.workspace_path)  # Path("/home/you/.cache/docx-editor/0bebafb463a87cfa")
```

### Track Changes Methods

#### `paragraph_count()`

Return the total number of paragraphs in the document. A cheap bounds check for pagination — avoids building the full `list_paragraphs()` result just to learn the count.

**Returns:** Total number of paragraphs (the highest valid 1-based ref index).

**Example:**

```python
count = doc.paragraph_count()
```

#### `list_paragraphs(max_chars=80, *, start=1, limit=None)`

List paragraphs with hash-anchored references. Refs are **1-based global** indexes (`P1`, `P2`, …) and stay correct across pages — a slice starting at paragraph 51 emits `P51#…`, not `P1#…`.

**Parameters:**

- `max_chars` (int): Maximum preview length (must be `>= 0`). Use `0` for refs only (e.g. `P1#a7b2`), with no preview or `| ` separator.
- `start` (int): 1-based index of the first paragraph to return (default 1). A `start` beyond the last paragraph yields an empty list.
- `limit` (int | None): Maximum number of paragraphs to return, or `None` for all paragraphs from `start` onward (default `None`).

**Returns:** List of strings in the form `P{index}#{hash}| preview text`, or bare `P{index}#{hash}` (no `| ` separator) when `max_chars=0`.

**Example:**

```python
refs = doc.list_paragraphs()
for ref in refs:
    print(ref)

# Page through a large document — you pick the page size; refs stay
# globally indexed (with size 50, page 2 starts at P51)
count = doc.paragraph_count()
page_size = 50
for start in range(1, count + 1, page_size):
    for ref in doc.list_paragraphs(start=start, limit=page_size):
        print(ref)  # process this page of refs
```

#### `get_paragraph_location(ref)`

Report whether a paragraph lives in the document body or inside a table cell.

**Parameters:**

- `ref` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`

**Returns:** `ParagraphLocation`. `location.in_table` is `False` for body paragraphs; `True` when the paragraph is inside a `<w:tc>` cell, in which case `location.table` carries the 1-based table index, row, `w:gridSpan`-aware logical column, and nesting depth.

**Example:**

```python
loc = doc.get_paragraph_location("P3#a7b2")
if loc.in_table:
    cell = loc.table
    print(f"table {cell.index} r{cell.row} c{cell.col} (depth {cell.depth})")
```

#### `list_paragraph_locations()`

Batch counterpart to `get_paragraph_location()`: pair every paragraph with its structural location in one pass, precomputing table indices once instead of rescanning the table hierarchy per ref.

**Returns:** List of `(ref, ParagraphLocation)` tuples in document order, where `ref` is the same `P{index}#{hash}` token emitted by `list_paragraphs()`.

**Example:**

```python
for ref, loc in doc.list_paragraph_locations():
    if loc.in_table:
        cell = loc.table
        print(f"{ref}: table {cell.index} r{cell.row} c{cell.col} (depth {cell.depth})")
```

#### `get_visible_text()`

Get flattened visible document text. Inserted text is included and deleted text is excluded.

**Returns:** Visible text with paragraphs separated by newlines (str)

**Example:**

```python
text = doc.get_visible_text()
```

#### `get_original_text()`

Get flattened original (pre-revision) document text. Deleted text is included and inserted text is excluded — the inverse of `get_visible_text()`. For intra-paragraph revisions this equals what `get_visible_text()` would return after `reject_all()`, without modifying the document (paragraph-level revisions such as inserted paragraph marks only affect line boundaries). Read-only: paragraph references and editing operations keep working on the visible view.

**Returns:** Original text with paragraphs separated by newlines (str)

**Example:**

```python
text = doc.get_original_text()
```

#### `find_text(text, occurrence=0)`

Find text in the document, including text spanning XML element boundaries.

**Parameters:**

- `text` (str): Text to search for
- `occurrence` (int): Which occurrence to return, counted document-wide. Defaults to 0.

**Returns:** [`SearchResult`](#searchresult), or None if not found

**Example:**

```python
match = doc.find_text("Aim: To")
if match:
    if match.spans_revision:
        print("Text spans a tracked-revision boundary")
    # The ref + in-paragraph occurrence pin the exact match for a follow-up edit
    doc.replace(
        "Aim: To", "Goal: To",
        paragraph=match.paragraph_ref, occurrence=match.paragraph_occurrence,
    )
```

#### `count_matches(text)`

Count visible text matches across the document.

**Parameters:**

- `text` (str): Text to search for

**Returns:** Number of occurrences found (int)

**Example:**

```python
if doc.count_matches("Section 5") > 1:
    print("Use paragraph refs and occurrence to target the intended match")
```

#### `replace(find, replace_with, *, paragraph, occurrence=0)`

Replace text with tracked changes.

**Parameters:**

- `find` (str): Text to find and replace
- `replace_with` (str): Replacement text
- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`
- `occurrence` (int): Which occurrence within the paragraph. Defaults to 0.

**Returns:** Updated paragraph reference (str)

**Example:**

```python
ref = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
doc.replace("net", "gross", paragraph=ref)
```

#### `delete(text, *, paragraph, occurrence=0)`

Mark text as deleted with tracked changes.

**Parameters:**

- `text` (str): Text to mark as deleted
- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`
- `occurrence` (int): Which occurrence within the paragraph. Defaults to 0.

**Returns:** Updated paragraph reference (str)

**Example:**

```python
ref = doc.delete("obsolete clause", paragraph="P5#d4e5")
```

#### `insert_after(anchor, text, *, paragraph, occurrence=0)`

Insert text after anchor with tracked changes.

**Parameters:**

- `anchor` (str): Text to find as insertion point
- `text` (str): Text to insert after the anchor
- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`
- `occurrence` (int): Which occurrence within the paragraph. Defaults to 0.

**Returns:** Updated paragraph reference (str)

**Example:**

```python
ref = doc.insert_after("Section 5", " (as amended)", paragraph="P3#b2c4")
```

#### `insert_before(anchor, text, *, paragraph, occurrence=0)`

Insert text before anchor with tracked changes.

**Parameters:**

- `anchor` (str): Text to find as insertion point
- `text` (str): Text to insert before the anchor
- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`
- `occurrence` (int): Which occurrence within the paragraph. Defaults to 0.

**Returns:** Updated paragraph reference (str)

**Example:**

```python
ref = doc.insert_before("Section 6", "New clause: ", paragraph="P4#a7b2")
```

#### `rewrite_paragraph(ref, new_text)`

Rewrite a paragraph using tracked changes generated from a word-level diff.

**Parameters:**

- `ref` (str): Paragraph reference from `list_paragraphs()`
- `new_text` (str): Desired paragraph text

**Returns:** Updated paragraph reference (str)

**Example:**

```python
ref = doc.rewrite_paragraph("P2#f3c1", "Payment is due within 60 days after invoice receipt.")
```

#### `batch_edit(operations)`

Apply multiple edits after validating paragraph hashes up front.

**Parameters:**

- `operations` (list[EditOperation]): Edit operations to apply

**Returns:** Updated paragraph references in input order (list[str])

**Example:**

```python
from docx_editor import EditOperation

new_refs = doc.batch_edit([
    EditOperation.replace("old", "new", paragraph="P2#f3c1"),
    EditOperation.delete("remove this", paragraph="P5#d4e5"),
])
```

Prefer the typed constructors ([`EditOperation`](#editoperation)) — they validate
arguments when the operation is built, so mistakes fail fast instead of at apply time.

#### `batch_rewrite(rewrites)`

Rewrite multiple paragraphs after validating paragraph hashes up front.

**Parameters:**

- `rewrites` (list[tuple[str, str]]): Pairs of paragraph ref and desired text

**Returns:** Updated paragraph references in input order (list[str])

**Example:**

```python
new_refs = doc.batch_rewrite([
    ("P1#a7b2", "Updated first paragraph."),
    ("P3#c3d4", "Updated third paragraph."),
])
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

After saving to a different path (or a save that fails), the workspace is flagged as holding unsaved changes; a later `Document.open()` of the source raises `WorkspaceSyncError` until the workspace is saved back to the source or discarded with `force_recreate=True`. See [`WorkspaceSyncError`](#workspacesyncerror) below.

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

## EditOperation

A single edit operation for `batch_edit()`. Build operations with the typed
constructors below — they validate arguments at construction time with the same
rules `batch_edit()` applies, so mistakes surface immediately. The raw
`EditOperation(action=..., ...)` form remains supported.

```python
from docx_editor import EditOperation
```

### Constructors

#### `EditOperation.replace(find, replace_with, *, paragraph, occurrence=0)`

- `find` (str): Text to find and replace. Must be non-empty.
- `replace_with` (str): Replacement text. Empty string is allowed (replacing
  with nothing is a valid tracked deletion).

#### `EditOperation.delete(text, *, paragraph, occurrence=0)`

- `text` (str): Text to mark as deleted. Must be non-empty.

#### `EditOperation.insert_after(anchor, text, *, paragraph, occurrence=0)`

#### `EditOperation.insert_before(anchor, text, *, paragraph, occurrence=0)`

- `anchor` (str): Text to find as insertion point. Must be non-empty.
- `text` (str): Text to insert.

All constructors also take:

- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`
- `occurrence` (int): Which occurrence within the paragraph. Defaults to 0. Must be >= 0.

**Raises:** `ValueError` at construction time if the paragraph ref is malformed,
`occurrence` is negative, a search-target argument (`find`, delete `text`,
`anchor`) is empty, or a payload argument (`replace_with`, insert `text`) is
`None` — payloads may be empty strings, search targets may not. Each
signature mirrors the corresponding `Document` method 1:1, so
`doc.replace(...)` translates mechanically to `EditOperation.replace(...)`.

### Example

```python
new_refs = doc.batch_edit([
    EditOperation.replace("30 days", "60 days", paragraph="P2#f3c1"),
    EditOperation.delete("obsolete clause", paragraph="P5#d4e5"),
    EditOperation.insert_after("Section 5", " (as amended)", paragraph="P7#b1c2"),
])
```

---

## SearchResult

The result of `Document.find_text()`. Carries no XML/DOM internals.

```python
from docx_editor import SearchResult
```

### Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `start` | int | Start offset of the match in the containing paragraph's visible text |
| `end` | int | Exclusive end offset, same coordinate space |
| `text` | str | The matched text |
| `paragraph_ref` | str | Hash-anchored ref like `P3#a7b2`, usable as the `paragraph=` argument of follow-up edits |
| `paragraph_occurrence` | int | Occurrence index of this match within its paragraph, usable as the `occurrence=` argument of follow-up edits |
| `spans_revision` | bool | True if the match crosses a tracked-revision boundary |

`start`/`end` are offsets within the matched paragraph's visible text, **not**
document-wide offsets. Coordinate systems differ between search and edit:
`find_text`'s `occurrence` counts matches document-wide, while edit methods
count within one paragraph — `paragraph_occurrence` bridges the two, so always
pass it alongside `paragraph_ref` when chaining into an edit. `paragraph_ref`
is computed at search time and — like refs from `list_paragraphs()` — goes
stale once that paragraph is edited.

### Example

```python
match = doc.find_text("30 days")
if match:
    doc.replace(
        "30 days", "60 days",
        paragraph=match.paragraph_ref, occurrence=match.paragraph_occurrence,
    )
```

---

## Deprecated internals

The text-map machinery (`TextMap`, `TextMapMatch`, `TextPosition`,
`build_text_map`, `find_in_text_map`) is no longer part of the public API:
these names have been removed from `docx_editor.__all__`, and accessing them
via the top-level package emits a `DeprecationWarning`. They will be removed
from the package namespace in the next release.

Use `Document.find_text()` / [`SearchResult`](#searchresult) instead. If you
genuinely need the internals (raw DOM positions), import them from
`docx_editor.xml_editor`.

---

## Exceptions

### `TextNotFoundError`

Raised when the specified text is not found in the document.

```python
from docx_editor.exceptions import TextNotFoundError

try:
    doc.replace("nonexistent text", "new text", paragraph="P2#f3c1")
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

Raised when the workspace is out of sync with the source document: the source changed on disk since the workspace was created, or the workspace holds unsaved changes from a previous session that the source never received.

```python
from docx_editor.exceptions import WorkspaceSyncError
```
