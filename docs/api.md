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

- `WorkspaceSyncError`: If the source `.docx` was modified since the workspace was created, or if a leftover workspace holds unsaved changes from a previous session — any session that made edits without a final successful `save()` back to the source (it saved to a different path, its save failed, or it closed with `close(cleanup=False)`). Pass `force_recreate=True` to discard the workspace and re-unpack from the current source. The workspace is never deleted silently. The error message includes the workspace path.
- `WorkspaceLockedError`: If a live session — another process, or an unclosed `Document` in this one — already holds the document's workspace. Close the other session, or pass `force_recreate=True` to take the workspace over and discard its unsaved edits. Locks left by dead processes are reclaimed silently.
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

Get the path to this document's workspace folder. Since the workspace lives in the user cache by default, this is how you locate the unpacked XML — for example after `close(cleanup=False)`, or when a workspace was preserved because an exception was raised. Either way the workspace holds the last state flushed by `save()`: tracked-change edits made but not saved live only in memory and are **not** in it. (A first `add_comment()` is the exception — it writes comment-part scaffolding into the workspace and flags it as diverged before any save; the unsaved comment text itself is still memory-only.)

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

#### `list_paragraphs(max_chars=80, *, start=1, limit=200)`

List paragraphs with hash-anchored references. Refs are **1-based global** indexes (`P1`, `P2`, …) and stay correct across pages — a slice starting at paragraph 51 emits `P51#…`, not `P1#…`.

**Changed in 0.6.1:** a bare call now returns at most 200 paragraphs (previously all of them). Whenever paragraphs remain beyond the returned window — default or explicit `limit` — the last list entry is a **truncation notice** instead of a paragraph, e.g. `"... 50 more paragraphs; use start=201 or limit=None"`. Notice lines always start with `...` and never match the `P{index}#{hash}` ref shape; filter them with `entry.startswith("...")` when consuming entries as refs. Pass `limit=None` for the full, notice-free listing.

**Parameters:**

- `max_chars` (int): Maximum preview length (must be `>= 0`). Use `0` for refs only (e.g. `P1#a7b2`), with no preview or `| ` separator.
- `start` (int): 1-based index of the first paragraph to return (default 1). A `start` beyond the last paragraph yields an empty list.
- `limit` (int | None): Maximum number of paragraphs to return (default 200), or `None` for all paragraphs from `start` onward.

**Returns:** List of strings in the form `P{index}#{hash}| preview text`, or bare `P{index}#{hash}` (no `| ` separator) when `max_chars=0` — plus one trailing `... N more paragraphs; use start=… or limit=None` notice when the window did not reach the end of the document.

**Example:**

```python
page1 = doc.list_paragraphs()          # up to P200, then "... N more" notice
page2 = doc.list_paragraphs(start=201)  # next page, per the notice
refs = [e for e in page1 if not e.startswith("...")]  # drop the notice line
everything = doc.list_paragraphs(limit=None)  # uncapped, never a notice
```

`list_paragraphs_structured()` (same `start`/`limit` semantics, returns typed `ParagraphInfo` records with full untruncated text) shares the 200-record default cap but appends **no notice** — every entry stays a `ParagraphInfo`. Detect truncation by checking whether the last record's `index` is still below `paragraph_count()` (robust for any `start`), or pass `limit=None`.

#### `get_paragraph(index)`

Return one paragraph as a structured `ParagraphInfo` record — the single-item counterpart to `list_paragraphs_structured()`. The returned record is identical to the one that method would emit for the same paragraph, without building a list.

**Parameters:**

- `index` (int): 1-based paragraph index (`P1` is `index=1`). Must be in `1 .. paragraph_count()`.

**Returns:** `ParagraphInfo` (index, hash-anchored ref, full untruncated text) for the paragraph at `index`.

**Raises:** [`ParagraphIndexError`](#exceptions) if `index` is out of range (`< 1` or greater than `paragraph_count()`).

**Example:**

```python
info = doc.get_paragraph(1)
print(info.ref, info.text)  # "P1#a7b2" "Full paragraph text..."
```

#### `context(ref, window=2)`

Return the paragraphs surrounding `ref`, in document order — the "show me the section around this match" helper. Fetches the referenced paragraph plus up to `window` paragraphs on each side, clamped at the document edges (no padding, no wrap-around).

**Parameters:**

- `ref` (str): Paragraph reference (e.g. `P3#a7b2`) from `list_paragraphs()`, `find_text()`/`find_all()`, or an edit result.
- `window` (int): Paragraphs to include on *each side* of the referenced one (default 2, so up to 5 records). Must be `>= 0`; `0` returns just the referenced paragraph.

**Returns:** List of `ParagraphInfo` records (index, ref, full text) — identical to what `list_paragraphs_structured()` would emit for the same span.

**Raises:** `ValueError` if `ref` is malformed or `window < 0`; [`ParagraphIndexError`](#exceptions) / [`HashMismatchError`](#exceptions) for an out-of-range or stale `ref`.

**Example:**

```python
match = doc.find_text("Termination")
for info in doc.context(match.paragraph_ref, window=2):
    print(info)  # "P{i}#{hash}| full paragraph text"
```

#### `get_paragraph_location(ref)`

Report whether a paragraph lives in the document body or inside a table cell, whether it is a list item, its heading context (style, outline level, and the chain of headings above it), and its section index.

**Parameters:**

- `ref` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`

**Returns:** `ParagraphLocation`. `location.in_table` is `False` for body paragraphs; `True` when the paragraph is inside a `<w:tc>` cell, in which case `location.table` carries the 1-based table index, row, `w:gridSpan`-aware logical column, and nesting depth. `location.list` is a `ListItem(num_id, ilvl)` for list paragraphs, `None` otherwise: a direct `w:pPr/w:numPr` wins when present — including Word's `numId=0` "numbering disabled" marker, which reports `None` with no style fallback — otherwise the numbering defined by the paragraph's style applies, with `w:basedOn` inheritance chains resolved. Rendered display numbers (e.g. "7.2(a)") are not computed.

`location.style` is the raw `w:pStyle` style id (e.g. `"Heading1"`), `None` when the paragraph carries no explicit style — no name resolution against `word/styles.xml`. `location.outline_level` is the 0-based outline level (`0` == Heading 1, so a document heading level is `outline_level + 1`): a direct `w:outlineLvl` on the paragraph wins, and the spec's `w:val="9"` marker means body text (`None`); otherwise the level defined by the paragraph's style applies, with `w:basedOn` inheritance chains resolved. `location.heading_path` is the chain of nearest preceding headings that contains the paragraph, outermost first (e.g. `("Chapter one", "Termination")`), built from each heading's current visible text; a heading's own path lists only its ancestors, never itself. Headings inside table cells participate in document order. `location.section` is the paragraph's 1-based section index: a paragraph carrying a direct `w:pPr/w:sectPr` closes a section and belongs to the section it closes, the next paragraph starts the following one, and the body-level `w:sectPr` defines the final section — single-section documents report `1` everywhere.

**Example:**

```python
loc = doc.get_paragraph_location("P3#a7b2")
if loc.in_table:
    cell = loc.table
    print(f"table {cell.index} r{cell.row} c{cell.col} (depth {cell.depth})")
if loc.list:
    print(f"list numId={loc.list.num_id} level={loc.list.ilvl}")
if loc.outline_level is not None:
    print(f"heading level {loc.outline_level + 1}: style={loc.style}")
print(" > ".join(loc.heading_path))  # e.g. "Chapter one > Termination"
print(f"section {loc.section}")
```

#### `list_paragraph_locations()`

Batch counterpart to `get_paragraph_location()`: pair every paragraph with its structural location in one pass, precomputing table indexes, style outline levels, style numbering, heading paths, and section indexes once instead of rescanning the document per ref.

**Returns:** List of `(ref, ParagraphLocation)` tuples in document order, where `ref` is the same `P{index}#{hash}` token emitted by `list_paragraphs()`. Each location carries the same table, list, style, outline-level, heading-path, and section info as `get_paragraph_location()`.

**Example:**

```python
for ref, loc in doc.list_paragraph_locations():
    if loc.in_table:
        cell = loc.table
        print(f"{ref}: table {cell.index} r{cell.row} c{cell.col} (depth {cell.depth})")
    if loc.list:
        print(f"{ref}: list numId={loc.list.num_id} level={loc.list.ilvl}")
    if loc.heading_path:
        print(f"{ref}: under {' > '.join(loc.heading_path)}")
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

#### `find_text(text, occurrence=0, paragraph=None)`

Find text in the document, including text spanning XML element boundaries.

**Parameters:**

- `text` (str): Text to search for (must be non-empty)
- `occurrence` (int): Which occurrence to return, 0-based (0 = first). Counted document-wide when `paragraph` is None, and within the paragraph when scoped. Defaults to 0.
- `paragraph` (str, optional): Paragraph reference (e.g. `P2#f3c1`) to scope the search — the same scoping `find_all` offers. `None` searches the whole document. Defaults to None.

**Returns:** [`SearchResult`](#searchresult), or None if not found

**Raises:** `ValueError` if `text` is empty or `paragraph` is malformed; [`ParagraphIndexError`](#exceptions) / [`HashMismatchError`](#exceptions) for an out-of-range or stale `paragraph`.

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

#### `find_all(text, paragraph=None)`

Find every match of `text`, in document order. One call replaces the N+1
`find_text` probes needed to enumerate N hits, and each result carries exactly
what a follow-up edit needs.

**Parameters:**

- `text` (str): Text to search for (must be non-empty)
- `paragraph` (str, optional): Paragraph reference (e.g. `P2#f3c1`) to scope the search. `None` searches the whole document. Defaults to None.

**Returns:** list of [`SearchResult`](#searchresult), empty when nothing matches (no-match is not an error for an enumeration API)

**Raises:** `ValueError` if `text` is empty or `paragraph` is malformed; [`ParagraphIndexError`](#exceptions) / [`HashMismatchError`](#exceptions) for an out-of-range or stale `paragraph`.

**Example:**

```python
# Edit every match in one atomic batch. reversed() puts same-paragraph ops in
# the required descending occurrence order, so this is safe however the
# matches are distributed:
ops = [
    EditOperation.replace(r.text, "60 days",
                          paragraph=r.paragraph_ref, occurrence=r.paragraph_occurrence)
    for r in reversed(doc.find_all("30 days"))
]
doc.batch_edit(ops)
```

Editing one match at a time (`doc.replace(...)` per result) also works when
every paragraph holds at most one match. With several matches in one
paragraph, an edit invalidates the paragraph's remaining refs and shifts the
occurrence numbers of the matches after it — either re-run `find_all` after
each edit, or batch the same-paragraph ops in **descending** occurrence order
as above; an edit never shifts the matches before it. (Ascending order
mis-targets; descending is not valid for search strings that overlap
themselves, e.g. `"aa"` in `"aaaa"`.)

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

#### `replace(find, replace_with, *, paragraph, occurrence=None)`

Replace text with tracked changes. When the target sits inside another author's pending insertion, that insertion is preserved: the matched text gets a nested `<w:del>` under your authorship and the replacement lands in your own sibling `<w:ins>` (Word's behavior), instead of silently rewriting the other author's proposal.

Words shared by `find` and `replace_with` at either end are trimmed first, so only the changed words become revisions — a replace that only adds or only removes words is written as a pure insertion or deletion. The replacement insertion carries the formatting (`rPr`) that covers the most characters of the replaced span (runs sharing identical formatting tally together), ties breaking to the earliest-seen formatting. When `replace_with` equals the found text, the call is a **no-op**: no revisions are created and the returned `EditResult` equals the input `paragraph` ref with `group_id=None` and `revision_ids=()` — that triple is how callers detect the no-op.

**Parameters:**

- `find` (str): Text to find and replace
- `replace_with` (str): Replacement text
- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`
- `occurrence` (int | None): Which occurrence within the paragraph, 0-based (0 = first). Omitted → the target must be unique within the paragraph; if it matches more than once, [`AmbiguousTextError`](#ambiguoustexterror) is raised instead of silently editing the first match.

**Returns:** Updated paragraph reference ([`EditResult`](#editresult) — a `str` subclass also carrying the edit's `group_id`/`revision_ids`)

**Example:**

```python
ref = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
doc.replace("net", "gross", paragraph=ref)
```

#### `delete(text, *, paragraph, occurrence=None)`

Mark text as deleted with tracked changes. Deleting text inside another author's pending insertion nests a `<w:del>` under your authorship inside their `<w:ins>`, preserving their proposal; only your own pending insertions are edited in place.

**Parameters:**

- `text` (str): Text to mark as deleted
- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`
- `occurrence` (int | None): Which occurrence within the paragraph, 0-based (0 = first). Omitted → the target must be unique within the paragraph; if it matches more than once, [`AmbiguousTextError`](#ambiguoustexterror) is raised instead of silently editing the first match.

**Returns:** Updated paragraph reference ([`EditResult`](#editresult) — a `str` subclass also carrying the edit's `group_id`/`revision_ids`)

**Example:**

```python
ref = doc.delete("obsolete clause", paragraph="P5#d4e5")
```

#### `insert_after(anchor, text, *, paragraph, occurrence=None)`

Insert text after anchor with tracked changes. An anchor inside another author's pending insertion produces your own sibling `<w:ins>` (splitting theirs when the anchor falls mid-content) rather than splicing your words into their proposal.

**Parameters:**

- `anchor` (str): Text to find as insertion point
- `text` (str): Text to insert after the anchor
- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`
- `occurrence` (int | None): Which occurrence within the paragraph, 0-based (0 = first). Omitted → the target must be unique within the paragraph; if it matches more than once, [`AmbiguousTextError`](#ambiguoustexterror) is raised instead of silently editing the first match.

**Returns:** Updated paragraph reference ([`EditResult`](#editresult) — a `str` subclass also carrying the edit's `group_id`/`revision_ids`)

**Example:**

```python
ref = doc.insert_after("Section 5", " (as amended)", paragraph="P3#b2c4")
```

#### `insert_before(anchor, text, *, paragraph, occurrence=None)`

Insert text before anchor with tracked changes. Foreign pending insertions are treated the same as in `insert_after()`.

**Parameters:**

- `anchor` (str): Text to find as insertion point
- `text` (str): Text to insert before the anchor
- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`
- `occurrence` (int | None): Which occurrence within the paragraph, 0-based (0 = first). Omitted → the target must be unique within the paragraph; if it matches more than once, [`AmbiguousTextError`](#ambiguoustexterror) is raised instead of silently editing the first match.

**Returns:** Updated paragraph reference ([`EditResult`](#editresult) — a `str` subclass also carrying the edit's `group_id`/`revision_ids`)

**Example:**

```python
ref = doc.insert_before("Section 6", "New clause: ", paragraph="P4#a7b2")
```

#### `rewrite_paragraph(ref, new_text)`

Rewrite a paragraph using tracked changes generated from a word-level diff.

A rewrite typically produces many revisions (one per diff hunk), none of which
is a self-contained edit — accepting only some of them by id garbles the
paragraph. All of one rewrite's revisions therefore share a revision group;
resolve them as a unit with [`accept_group()`](#accept_groupgroup_id) /
[`reject_group()`](#reject_groupgroup_id).

**Parameters:**

- `ref` (str): Paragraph reference from `list_paragraphs()`
- `new_text` (str): Desired paragraph text

**Returns:** Updated paragraph reference ([`EditResult`](#editresult) — a `str` subclass also carrying the edit's `group_id`/`revision_ids`; `group_id` is `None` when `new_text` equals the current text, or when every change landed inside your own pending insertions and was merged in place)

**Example:**

```python
result = doc.rewrite_paragraph("P2#f3c1", "Payment is due within 60 days after invoice receipt.")
doc.reject_group(result.group_id)  # changed your mind — undo the whole rewrite
```

#### `batch_edit(operations, *, dry_run=False)`

Apply multiple edits after validating paragraph hashes up front. If any hash is
stale, the entire batch is rejected before any edits are applied.

**Parameters:**

- `operations` (list[EditOperation]): Edit operations to apply
- `dry_run` (bool): If True, validate every operation without applying any edits and return one [`EditValidationResult`](#editvalidationresult) per operation, in input order; the document is left unchanged. Each operation is validated independently against the current document — sequential effects between multiple operations on the same paragraph are **not** simulated. Defaults to False.

**Returns:** Updated paragraph references in input order (list of [`EditResult`](#editresult)) — each operation that creates revisions gets its own revision group, so one op can be accepted and another rejected (`group_id` is `None` for an op that created no new revisions, e.g. text spliced into one of your own pending insertions); with `dry_run=True`, a list of [`EditValidationResult`](#editvalidationresult) instead

**Raises:** [`BatchOperationError`](#batchoperationerror) — the only exception a non-dry-run batch raises for a failing operation, whatever the underlying cause (stale hash, malformed ref, missing text, ambiguous target). `operation_index` names the failing op; `original` (also `__cause__`) holds the underlying typed exception. The batch is atomic: nothing is applied on failure.

**Example:**

```python
from docx_editor import EditOperation

ops = [
    EditOperation.replace("old", "new", paragraph="P2#f3c1"),
    EditOperation.delete("remove this", paragraph="P5#d4e5"),
]

# Pre-flight the batch, then apply
results = doc.batch_edit(ops, dry_run=True)
if all(r.valid for r in results):
    new_refs = doc.batch_edit(ops)
```

Prefer the typed constructors ([`EditOperation`](#editoperation)) — they validate
arguments when the operation is built, so mistakes fail fast instead of at apply time.

Multiple operations on the same paragraph apply sequentially in input order:
each operation's find/anchor text and `occurrence` resolve against the
paragraph's visible text as left by the previous operations in the batch (a
tracked delete removes text from that view; an insert adds to it). Across
different paragraphs, operations are applied in reverse document order — a
behavior that keeps one `list_paragraphs()` snapshot valid for the whole batch.

#### `batch_rewrite(rewrites)`

Rewrite multiple paragraphs after validating paragraph hashes up front.

**Parameters:**

- `rewrites` (list[tuple[str, str]]): Pairs of paragraph ref and desired text

**Returns:** Updated paragraph references in input order (list of [`EditResult`](#editresult)); each rewrite gets its own revision group (`group_id` is `None` for a rewrite that made no change or whose changes fully merged into your own pending insertions)

**Raises:** [`BatchOperationError`](#batchoperationerror) — same single-exception contract as `batch_edit()`.

**Example:**

```python
new_refs = doc.batch_rewrite([
    ("P1#a7b2", "Updated first paragraph."),
    ("P3#c3d4", "Updated third paragraph."),
])
```

### Comment Methods

#### `add_comment(anchor_text, comment, *, paragraph=None, occurrence=None)`

Add a comment anchored to specific text. Anchors are located with the same
visible-text search used by `count_matches()` and the tracked-change edit
methods, so anchors that span `w:t` run boundaries (formatting changes,
smart-quote splits, `w:ins` wrappers) are found.

**Parameters:**

- `anchor_text` (str): Text to attach the comment to
- `comment` (str): The comment content
- `paragraph` (str, optional): Paragraph reference (e.g. `P3#a7b2`) to scope the search. `None` searches the whole document. Defaults to None.
- `occurrence` (int | None): Which occurrence to anchor to, 0-based (0 = first), counted within `paragraph` when given and document-wide otherwise. Omitted → the anchor must be unique in the search scope, else [`AmbiguousTextError`](#ambiguoustexterror).

**Returns:** The comment ID (int). IDs are allocated sequentially starting at 0 in a document with no existing comments — always use the returned ID rather than guessing.

**Example:**

```python
cid = doc.add_comment("Section 5", "Please review this section")
doc.add_comment("term", "Note on the 2nd 'term'", paragraph="P3#a7b2", occurrence=1)
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

#### `list_revisions(author=None, paragraph=None)`

List all tracked changes in the document.

**Parameters:**

- `author` (str, optional): If provided, filter by author name
- `paragraph` (str, optional): If provided, only return revisions in this paragraph (hash-anchored ref from `list_paragraphs()`, e.g. `"P3#a7b2"`)

**Returns:** List of Revision objects

**Raises:** The `paragraph` ref is validated exactly like in the edit methods — `ValueError` (malformed ref), [`ParagraphIndexError`](#exceptions) (index out of range), [`HashMismatchError`](#exceptions) (stale hash).

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
- Nested revisions: accepting an insertion unwraps it in place, so a deletion another author nested inside it survives as an independent pending deletion

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
- Nested revisions: rejecting an insertion removes everything inside it — deletions another author nested inside it disappear with it

**Parameters:**

- `revision_id` (int): ID of the revision to reject

**Returns:** True if rejected, False if not found (bool)

**Example:**

```python
doc.reject_revision(1)
```

#### `accept_group(group_id)`

Accept every revision created by one logical edit operation.

Each edit method (`replace()`, `delete()`, `insert_after()`, `insert_before()`,
`rewrite_paragraph()`, and every operation of a batch) registers the revisions
it creates as one **revision group**; the returned [`EditResult`](#editresult)
carries the `group_id`, and `list_revisions()` stamps it on each member
revision. Accepting the group applies the whole edit — the safe alternative to
resolving a multi-revision edit (especially a rewrite) revision by revision,
which garbles the text if only some are applied.

Group ids are **in-memory and per-open-Document**, renumbered on each open.
Edits made through the open Document **record** their group
(`Revision.group_source == "recorded"`). For revisions already in the file —
previous sessions, foreign reviewers, Word round-trips — nothing is persisted
in the `.docx` (Word has no grouping concept and strips unknown markup), so
groups are **inferred** at parse time instead
(`Revision.group_source == "inferred"`): contiguous revisions in the same
paragraph sharing identical `w:author` + `w:date` reconstruct as one group.
That heuristic almost always matches the original logical edits — one caveat:
`w:date` has second precision, so two edits to the *same paragraph* within the
same second merge into one inferred group. When Word already resolved part of
a former edit, the remainder reconstructs as a smaller (rump) group, and
`accept_group()`/`reject_group()` handle it fine. Revisions missing an author
or date, sitting outside any paragraph (e.g. table-row markers), or sharing a
duplicated id stay ungrouped (`group_id=None`), as does the trailing half
created when an edit splits a *foreign* author's pending insertion mid-session
(foreign grouping is best-effort); revisions with non-numeric ids are omitted
from `list_revisions()` entirely (no id-keyed operation could target them).

Never carry a `group_id` across sessions: reopening renumbers groups from 1,
so a stale id from a previous session may silently resolve to a *different*
group rather than raise. Always take group ids from the current session's
`EditResult` or `list_revisions()`. `save()` does not invalidate groups (the
Document stays open and revision ids are preserved).

**Parameters:**

- `group_id` (int): Group id from an `EditResult` (or a `Revision.group_id`)

**Returns:** Number of revisions accepted (int). Members already resolved individually are skipped (and not counted).

**Raises:** [`RevisionError`](#revisionerror) if the group id is unknown to this open Document.

**Example:**

```python
result = doc.rewrite_paragraph(ref, "New text.")
doc.accept_group(result.group_id)  # apply the whole rewrite
```

#### `reject_group(group_id)`

Reject every revision created by one logical edit operation — the counterpart
of [`accept_group()`](#accept_groupgroup_id). Rejecting the group undoes the
whole edit, restoring the exact pre-edit text (deletions restored, insertions
removed). Same group semantics and lifetime as `accept_group()` — including
recorded vs inferred groups and per-open renumbering.

**Parameters:**

- `group_id` (int): Group id from an `EditResult` (or a `Revision.group_id`)

**Returns:** Number of revisions rejected (int). Members already resolved individually are skipped (and not counted).

**Raises:** [`RevisionError`](#revisionerror) if the group id is unknown.

**Example:**

```python
result = doc.rewrite_paragraph(ref, "New text.")
doc.reject_group(result.group_id)  # undo the whole rewrite
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

#### `save(path=None, validate=False, force=False)`

Save the document.

**Parameters:**

- `path` (str | Path, optional): Output path. Defaults to original source path.
- `validate` (bool): If True, validate with LibreOffice before saving. Defaults to False.
- `force` (bool): If True, skip save-time safety checks. By default `save()` refuses to overwrite the source if it changed on disk since it was opened (raising [`WorkspaceSyncError`](#workspacesyncerror)), or to write a destination that appears open in Word — a `~$` owner file exists next to it (raising [`DocumentOpenError`](#documentopenerror)). Pass `force=True` only for a confirmed-stale lock left by a crashed session. Defaults to False.

**Returns:** Path to the saved document (Path)

After saving to a different path (or a save that fails), the workspace is flagged as holding unsaved changes; a later `Document.open()` of the source raises `WorkspaceSyncError` until the workspace is saved back to the source or discarded with `force_recreate=True`. See [`WorkspaceSyncError`](#workspacesyncerror) below.

**Example:**

```python
doc.save()  # Save to original path
doc.save("contract_v2.docx")  # Save to new path
```

#### `close(cleanup=True)`

Close the document and clean up workspace. Releases the advisory workspace lock in both cleanup modes — closing is what frees the document for another session to open (see [`WorkspaceLockedError`](#workspacelockederror)).

> **Warning:** closing without saving discards unsaved edits. There is no dirty check: `close()` (the default `cleanup=True`, including normal context-manager exit) deletes the workspace and everything not yet written by `save()` is silently lost. `close(cleanup=False)` keeps the workspace on disk, but a later `Document.open()` of the same source raises `WorkspaceSyncError` if it holds unsaved changes, rather than silently carrying them over.

Any operation on the document after `close()` raises [`DocumentClosedError`](#documentclosederror).

**Parameters:**

- `cleanup` (bool): If True, delete the workspace folder. Defaults to True.

**Example:**

```python
doc.save()   # persist edits first —
doc.close()  # close() alone discards anything unsaved

# Or, instead of the above: keep the workspace on disk for inspection
doc.close(cleanup=False)
```

---

## Comment

Represents a document comment. IDs are allocated sequentially starting at 0 in
a document with no existing comments — always use the ID returned by
`add_comment()` / `reply_to_comment()` rather than assuming a numbering scheme.

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
| `group_id` | int or None | Revision group this revision belongs to (see [`accept_group()`](#accept_groupgroup_id)): recorded for this session's edits, inferred by reconstruction for revisions already in the file; None only for ungroupable revisions (missing author/date, outside any paragraph, duplicated id, or a mid-session split half of a foreign insertion) |
| `group_source` | str or None | Provenance of `group_id`: `"recorded"` (created through this open Document) or `"inferred"` (reconstructed at parse time from same-paragraph contiguity + identical author and date); None iff `group_id` is None |

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

#### `EditOperation.replace(find, replace_with, *, paragraph, occurrence=None)`

- `find` (str): Text to find and replace. Must be non-empty.
- `replace_with` (str): Replacement text. Empty string is allowed (replacing
  with nothing is a valid tracked deletion).

#### `EditOperation.delete(text, *, paragraph, occurrence=None)`

- `text` (str): Text to mark as deleted. Must be non-empty.

#### `EditOperation.insert_after(anchor, text, *, paragraph, occurrence=None)`

#### `EditOperation.insert_before(anchor, text, *, paragraph, occurrence=None)`

- `anchor` (str): Text to find as insertion point. Must be non-empty.
- `text` (str): Text to insert.

All constructors also take:

- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`
- `occurrence` (int | None): Which occurrence within the paragraph, 0-based (0 = first). Omitted → the target must be unique within the paragraph at apply time; if it matches more than once, the batch fails with a [`BatchOperationError`](#batchoperationerror) wrapping an [`AmbiguousTextError`](#ambiguoustexterror). Must be >= 0 when given.

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

## EditResult

The return value of every tracked-edit method (`replace()`, `delete()`,
`insert_after()`, `insert_before()`, `rewrite_paragraph()`, and the elements of
`batch_edit()` / `batch_rewrite()` results). A `str` **subclass** — the string
value is the new hash-anchored paragraph reference (e.g. `"P2#c3d4"`), so an
`EditResult` works unchanged anywhere a ref string is expected — with the
edit's revision-group info attached.

```python
from docx_editor import EditResult
```

### Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `group_id` | int or None | Revision group holding every revision this edit created, for [`accept_group()`](#accept_groupgroup_id) / [`reject_group()`](#reject_groupgroup_id). None when the edit created no new revisions (e.g. text spliced into one of your own pending insertions, a no-change rewrite, or a rewrite whose changes all merged into your own pending insertions). Valid only while this Document stays open — after reopen the same revisions belong to a freshly inferred group with a new id. |
| `revision_ids` | tuple[int, ...] | The `w:id`s of the group's member revisions, in creation order; `()` when `group_id` is None |

### Example

```python
result = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
print(str(result))         # "P2#c3d4" — the new paragraph ref
print(result.group_id)     # 1
print(result.revision_ids) # (0, 1) — the del and the ins

doc.replace("net", "gross", paragraph=result)  # usable as a plain ref
doc.reject_group(result.group_id)              # undo the first edit entirely
```

---

## EditValidationResult

The outcome of validating one `EditOperation` in a `batch_edit(ops, dry_run=True)`
call. One result is returned per operation, in input order.

```python
from docx_editor import EditValidationResult
```

### Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `index` | int | 0-based position of the operation in the input list |
| `paragraph` | str or None | The operation's paragraph ref (`None` if it was missing) |
| `valid` | bool | True if the operation would apply cleanly |
| `error` | str or None | Human-readable reason when not valid |

### Example

```python
from docx_editor import EditOperation

ops = [
    EditOperation.replace("old", "new", paragraph="P2#f3c1"),
    EditOperation.delete("remove this", paragraph="P5#d4e5"),
]
results = doc.batch_edit(ops, dry_run=True)
for r in results:
    if not r.valid:
        print(f"op {r.index} on {r.paragraph}: {r.error}")
if all(r.valid for r in results):
    new_refs = doc.batch_edit(ops)
```

---

## SearchResult

The result of `Document.find_text()` (a single match, or None) and
`Document.find_all()` (a list of them). Carries no XML/DOM internals.

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
| `paragraph_index` | int | 1-based index of the containing paragraph — the same integer embedded in `paragraph_ref`, so you never string-parse the ref |

`start`/`end` are offsets within the matched paragraph's visible text, **not**
document-wide offsets. Coordinate systems differ between search and edit:
`find_text`'s `occurrence` counts matches document-wide (unless scoped with
`paragraph=`), while edit methods count within one paragraph —
`paragraph_occurrence` bridges the two, so always pass it alongside
`paragraph_ref` when chaining into an edit. `paragraph_ref`
is computed at search time and — like refs from `list_paragraphs()` — goes
stale once that paragraph is edited.

`repr()`/`str()` are compact one-liners —
`SearchResult(P3#a7b2 occ=0 '30 days')`, with a trailing `spans_rev` marker
when `spans_revision` is true — so printing a whole `find_all()` list stays
cheap. Matched text longer than 60 characters is elided with `...` in the
display only; every field, including the full `text`, remains accessible as
an attribute.

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

Raised when the specified text is not found in the document, or when an
explicit `occurrence` is out of range — the error then carries `occurrence`
and `total_occurrences` and its message reports the actual count instead of
claiming the text is absent. Other structured fields: `search_text`,
`paragraph_ref`, `paragraph_preview`.

```python
from docx_editor.exceptions import TextNotFoundError

try:
    doc.replace("nonexistent text", "new text", paragraph="P2#f3c1")
except TextNotFoundError as e:
    print(f"Text not found: {e}")
```

### `AmbiguousTextError`

Raised when an edit target matches more than once in the search scope and no
`occurrence` was given. Not a `TextNotFoundError` subclass — the text *was*
found. Structured fields: `search_text`, `paragraph_ref` (None when
document-wide), `paragraph_preview` (None when document-wide),
`total_occurrences`.

```python
from docx_editor.exceptions import AmbiguousTextError

try:
    doc.replace("term", "clause", paragraph="P2#f3c1")
except AmbiguousTextError as e:
    r = doc.find_all("term", paragraph=e.paragraph_ref)[0]
    doc.replace("term", "clause", paragraph=r.paragraph_ref,
                occurrence=r.paragraph_occurrence)
```

### `BatchOperationError`

The only exception `batch_edit()` / `batch_rewrite()` raise for a failing
operation. Structured fields: `operation_index` (0-based position of the
failing op), `reason` (human-readable message), `original` (the underlying
typed exception, also set as `__cause__`; `None` for batch-level rule
violations that have no underlying exception, e.g. a missing paragraph ref or
a duplicate paragraph in `batch_rewrite`).

```python
from docx_editor.exceptions import BatchOperationError, HashMismatchError

while ops:
    try:
        doc.batch_edit(ops)
        break
    except BatchOperationError as e:
        if isinstance(e.original, HashMismatchError):
            op = ops[e.operation_index]
            op.paragraph = f"P{e.original.paragraph_index}#{e.original.actual_hash}"
        else:
            ops.pop(e.operation_index)
```

### `CommentError`

Raised when a comment operation fails. Structured field: `comment_id` (the comment id the operation targeted, e.g. the parent id of a failed reply; `None` when no comment id applies).

```python
from docx_editor.exceptions import CommentError

try:
    doc.reply_to_comment(999, "reply")
except CommentError as e:
    print(f"Comment {e.comment_id} not found")
```

### `RevisionError`

Raised when a revision operation fails — most commonly an unknown group id passed to `accept_group()` / `reject_group()`. Group ids are per-open-`Document` and renumbered on each open (recorded for this session's edits, inferred by reconstruction for revisions already in the file), so always use a group id from the current session's `EditResult` or `list_revisions()` — a stale id from a previous session may raise this, or worse, silently resolve to a different group. Structured fields: `revision_id` and `group_id` — set when the error is about that specific id (`group_id` for unknown-group errors), `None` otherwise.

```python
from docx_editor.exceptions import RevisionError
```

### `DocumentNotFoundError`

Raised by `Document.open()` when the source file does not exist. Structured field: `path` (the path that did not exist).

```python
from docx_editor.exceptions import DocumentNotFoundError
```

### `DocumentClosedError`

Raised when any operation is attempted on a `Document` after `close()`. Closing discards the workspace (unless `cleanup=False`), so the object cannot keep serving reads or edits — reopen the source to continue. Structured field: `path` (the source path of the closed document).

```python
from docx_editor import Document
from docx_editor.exceptions import DocumentClosedError

try:
    doc.get_visible_text()
except DocumentClosedError as e:
    doc = Document.open(e.path)
```

### `WorkspaceExistsError`

Raised when attempting to create a workspace that already exists.

```python
from docx_editor.exceptions import WorkspaceExistsError
```

### `WorkspaceSyncError`

Raised when the workspace is out of sync with the source document: the source changed on disk since the workspace was created, or the workspace holds unsaved changes from a previous session that the source never received. Structured fields: `workspace_path` and `source_path`.

`Document.open(path, force_recreate=True)` recovers but discards the workspace's unsaved edits. To rescue them first, save the orphaned workspace to a new file (`Workspace` is not exported at the package root — use the deep import):

```python
from docx_editor.exceptions import WorkspaceSyncError
from docx_editor.workspace import Workspace

try:
    doc = Document.open("contract.docx")
except WorkspaceSyncError:
    Workspace("contract.docx", create=False).save("rescued.docx")  # rescue unsaved edits
    doc = Document.open("contract.docx", force_recreate=True)
```

### `DocumentOpenError`

Raised by `save()` when the destination appears open in another program. Word writes a `~$` owner (lock) file next to any document it has open; if that stub exists at save time, saving would race Word's own writes, so `save()` refuses unless `force=True`. Also raised when the OS denies the final replace (on Windows, Word holding the file open is exactly this case) — `force=True` cannot suppress that one. The exception carries `path` (the destination) and `owner_file` (the `~$` file that triggered the guard, or `None` when the OS denied the replace) attributes.

```python
from docx_editor.exceptions import DocumentOpenError

try:
    doc.save()
except DocumentOpenError as e:
    print(f"Close {e.path} in Word first (lock: {e.owner_file})")
```

### `WorkspaceLockedError`

Raised when opening a document whose workspace is locked by a live session — another process (or another `Document` object in the same process) already has it open. Two sessions sharing one workspace would silently overwrite each other's saves. Close the other session, or pass `force_recreate=True` to take the workspace over and discard its unsaved edits. Locks left behind by dead processes are reclaimed automatically and never raise. The exception carries `pid` and `lock_path` attributes.

```python
from docx_editor.exceptions import WorkspaceLockedError

try:
    doc = Document.open("contract.docx")
except WorkspaceLockedError as e:
    print(f"Held by pid {e.pid}")
```
