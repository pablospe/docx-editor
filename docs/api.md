# API Reference

## Document

The main entry point for docx-editor. Provides methods for opening documents, making tracked changes, managing comments, and handling revisions.

```python
from docx_editor import Document
```

### Opening Documents

#### `Document.open(path, author=None, force_recreate=False, workspace_dir=None, *, allow_protected=False)`

Open a Word document for editing.

**Parameters:**

- `path` (str | Path): Path to the .docx file
- `author` (str, optional): Author name for tracked changes. Defaults to system username.
- `force_recreate` (bool): If True, delete any existing workspace (stale or in-sync) before opening — whatever XML it holds is discarded — and re-unpack from the current source. Use this to recover from `WorkspaceSyncError`. Defaults to False.
- `workspace_dir` (str | Path, optional): Base directory for the workspace. Overrides the `DOCX_EDITOR_WORKSPACE_DIR` environment variable and the platform cache default (see [Workspace location](#workspace-location)). A relative path resolves against the document's directory, so `workspace_dir=".docx"` keeps the workspace next to the file. Defaults to None.
- `allow_protected` (bool): If True, open a document whose editing protection is enforced (Word's *Restrict Editing*) instead of raising [`DocumentProtectedError`](#documentprotectederror). The protection itself is left in place, so the saved document reaches Word still restricted. Defaults to False.

**Returns:** Document instance ready for editing

**Raises:**

- `WorkspaceSyncError`: If the source `.docx` was modified since the workspace was created, or if a leftover workspace holds unsaved changes from a previous session — any session that made edits without a final successful `save()` back to the source (it saved to a different path, its save failed, or it closed with `close(cleanup=False)`). Pass `force_recreate=True` to discard the workspace and re-unpack from the current source. The workspace is never deleted silently. The error message includes the workspace path.
- `WorkspaceLockedError`: If a live session — another process, or an unclosed `Document` in this one — already holds the document's workspace. Close the other session, or pass `force_recreate=True` to take the workspace over and discard its unsaved edits. Locks left by dead processes are reclaimed silently.
- `WorkspaceError`: If the workspace directory cannot be created (e.g. the base is not writable), the home directory backing the default cache cannot be determined, or an existing workspace was unpacked from a different document. The message names the override to set.
- `DocumentProtectedError`: If the document enforces an editing protection that locks its body text — `readOnly`, `forms` or `comments`. Carries `path` and `mode`. Unprotect it in Word, or pass `allow_protected=True`. A document enforcing `trackedChanges` opens normally: that mode asks for exactly what this library does.

**Example:**

```python
doc = Document.open("contract.docx")
doc = Document.open("contract.docx", author="Legal Team")
```

#### `Document.discard_workspace(path, *, workspace_dir=None)`

Delete a document's workspace so the next `open()` starts clean — the recovery
call for a **crashed or killed script**. That session's workspace stays behind
flagged as holding unsaved changes, so every later `open()` of the document
raises [`WorkspaceSyncError`](#workspacesyncerror) (or
[`WorkspaceLockedError`](#workspacelockederror) if its advisory lock is still on
disk) until the workspace is dealt with. One call resets that state, instead of
passing `force_recreate=True` on every open or deleting the cache directory by
hand.

**This discards whatever unsaved edits the workspace holds** — the same
destruction as `open(force_recreate=True)`, without opening. To rescue them
first, save the orphaned workspace elsewhere (see
[`WorkspaceSyncError`](#workspacesyncerror)). The advisory lock sidecar goes with
the workspace, **including a lock held by a live session**, so make sure that
process is really gone: two sessions sharing one workspace silently overwrite
each other's saves.

**Parameters:**

- `path` (str | Path): The .docx whose workspace should be deleted. Need not exist — a workspace outlives its source, so this still cleans up after a moved or deleted document.
- `workspace_dir` (str | Path, optional): Base directory override, matching `open()`'s argument. Must be the same value the workspace was created with, or a different workspace is targeted (and nothing is found).

**Returns:** `True` if a workspace was deleted, `False` if there was none — so it is safe to call unconditionally.

**Example:**

```python
# Idempotent reset before a fresh run:
Document.discard_workspace("contract.docx")
doc = Document.open("contract.docx")
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

Get the path to this document's workspace folder. Since the workspace lives in the user cache by default, this is how you locate the unpacked XML — for example after `close(cleanup=False)`, or when a workspace was preserved because an exception was raised. Either way the workspace holds the last state flushed by `save()`: tracked-change edits made but not saved live only in memory and are **not** in it. (A first `add_comment()` — or a first `note=` edit, which creates a comment too — is the exception: it writes comment-part scaffolding into the workspace and flags it as diverged before any save; the unsaved comment text itself is still memory-only.)

```python
print(doc.workspace_path)  # Path("/home/you/.cache/docx-editor/0bebafb463a87cfa")
```

#### `has_textbox_content`

Whether any paragraph is hidden inside a drawing's text box (`w:txbxContent`).

Text boxes are not an editing surface: their paragraphs are absent from every listing, their text from every view, search and hash. That makes an all-text-box document — a poster, a flyer, a certificate — read as blank, which is indistinguishable from a genuinely empty document without this flag. Exactly the complement of the paragraph exclusion: `True` means at least one `<w:p>` was excluded, so any text those paragraphs carry is not reachable from here. To read it, go through HTML — `soffice --headless --convert-to html file.docx` then `pandoc file.html -t plain` — rather than reporting the document as empty; pandoc may render a `[ShapeN]` label beside a box's text, from the placeholder LibreOffice exports for a named shape — ignore those. Its `txt:Text` filter and pandoc reading the `.docx` directly both drop text boxes silently.

**Returns:** `True` if any paragraph was excluded as text-box content (bool)

**Example:**

```python
if not doc.get_visible_text().strip() and doc.has_textbox_content:
    print("Text lives in text boxes — not editable through refs")
```

### Track Changes Methods

#### `paragraph_count()`

Return the total number of paragraphs in the document. A cheap bounds check for pagination — avoids building the full `list_paragraphs()` result just to learn the count.

Paragraphs inside a drawing's text box (`w:txbxContent`) are not counted — text boxes are excluded from the ref index space entirely, so no ref ever addresses one. A document whose content lives only in text boxes therefore counts just the host paragraphs its boxes are anchored in; check [`has_textbox_content`](#has_textbox_content) before reporting it as empty.

**Returns:** Total number of paragraphs (the highest valid 1-based ref index).

**Example:**

```python
count = doc.paragraph_count()
```

#### `list_paragraphs(max_chars=80, *, start=1, limit=200)`

List paragraphs with hash-anchored references. Refs are **1-based global** indexes (`P1`, `P2`, …) and stay correct across pages — a slice starting at paragraph 51 emits `P51#…`, not `P1#…`. Text-box content is excluded, as it is from [`paragraph_count()`](#paragraph_count).

**Changed in 0.6.1:** a bare call now returns at most 200 paragraphs (previously all of them). Whenever paragraphs remain beyond the returned window — default or explicit `limit` — the last list entry is a **truncation notice** instead of a paragraph, e.g. `"... 50 more paragraphs; use start=201 or limit=None"`. Notice lines always start with `...` and never match the `P{index}#{hash}` ref shape; filter them with `entry.startswith("...")` when consuming entries as refs. Pass `limit=None` for the full, notice-free listing.

**Parameters:**

- `max_chars` (int): Maximum preview length (must be `>= 0`). Use `0` for refs only (e.g. `P1#a7b2`), with no preview or `| ` separator.
- `start` (int): 1-based index of the first paragraph to return (default 1). A `start` beyond the last paragraph is a valid request with an empty answer — it yields `[]` and does **not** raise, exactly like `lst[500:]` on a shorter list, so a loop-until-empty walk terminates cleanly. (`start=0` or negative *does* raise `ValueError`: those are invalid inputs, not empty ranges.) To know where the end is up front, use `paragraph_count()` or follow the truncation notice's `start=` hint.
- `limit` (int | None): Maximum number of paragraphs to return (default 200), or `None` for all paragraphs from `start` onward.

**Returns:** List of strings in the form `P{index}#{hash}| preview text`, or bare `P{index}#{hash}` (no `| ` separator) when `max_chars=0` — plus one trailing `... N more paragraphs; use start=… or limit=None` notice when the window did not reach the end of the document.

**Example:**

```python
page1 = doc.list_paragraphs()          # up to P200, then "... N more" notice
page2 = doc.list_paragraphs(start=201)  # next page, per the notice
refs = [e for e in page1 if not e.startswith("...")]  # drop the notice line
everything = doc.list_paragraphs(limit=None)  # uncapped, never a notice
```

`list_paragraphs_structured()` (same `start`/`limit` semantics, returns typed [`ParagraphInfo`](#paragraphinfo) records with full untruncated text plus `style`/`outline_level`/`in_table`) shares the 200-record default cap but appends **no notice** — every entry stays a `ParagraphInfo`. Detect truncation by checking whether the last record's `index` is still below `paragraph_count()` (robust for any `start`), or pass `limit=None`.

#### `get_paragraph(index)`

Return one paragraph as a structured [`ParagraphInfo`](#paragraphinfo) record — the single-item counterpart to `list_paragraphs_structured()`. The returned record is identical to the one that method would emit for the same paragraph, without building a list.

**Parameters:**

- `index` (int): 1-based paragraph index (`P1` is `index=1`). Must be in `1 .. paragraph_count()`.

**Returns:** `ParagraphInfo` (index, hash-anchored ref, full untruncated text, `style`, `outline_level`, `in_table`) for the paragraph at `index`.

**Raises:** [`ParagraphIndexError`](#paragraphindexerror) if `index` is out of range (`< 1` or greater than `paragraph_count()`).

**Performance:** fixed-size output, but **O(document)** work — every call walks all `<w:p>` elements to reach one. (`word/styles.xml` is parsed once per open `Document`, not per call.) Fine for a one-off lookup; for many paragraphs call `list_paragraphs_structured(limit=None)` once and index that list instead of looping over this. The same applies to `context()`.

**Example:**

```python
info = doc.get_paragraph(1)
print(info.ref, info.text)  # "P1#a7b2" "Full paragraph text..."
print(info.style, info.outline_level, info.in_table)  # "Heading1" 0 False
```

#### `context(ref, window=2)`

Return the paragraphs surrounding `ref`, in document order — the "show me the section around this match" helper. Fetches the referenced paragraph plus up to `window` paragraphs on each side, clamped at the document edges (no padding, no wrap-around).

**Parameters:**

- `ref` (str): Paragraph reference (e.g. `P3#a7b2`) from `list_paragraphs()`, `find_text()`/`find_all()`, or an edit result.
- `window` (int): Paragraphs to include on *each side* of the referenced one (default 2, so up to 5 records). Must be `>= 0`; `0` returns just the referenced paragraph.

**Returns:** List of [`ParagraphInfo`](#paragraphinfo) records — identical to what `list_paragraphs_structured()` would emit for the same span, structural fields included.

**Raises:** `ValueError` if `ref` is malformed or `window < 0`; [`ParagraphIndexError`](#paragraphindexerror) / [`HashMismatchError`](#hashmismatcherror) for an out-of-range or stale `ref`.

**Performance:** O(document) despite the fixed-size output — see `get_paragraph()` above.

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

**Returns:** `ParagraphLocation`. `location.in_table` is `False` for body paragraphs; `True` when the paragraph is inside a `<w:tc>` cell, in which case `location.table` carries the 1-based table index (body tables only — a table inside a text box is not counted), row, `w:gridSpan`-aware logical column, and nesting depth. `location.list` is a `ListItem(num_id, ilvl)` for list paragraphs, `None` otherwise: a direct `w:pPr/w:numPr` wins when present — including Word's `numId=0` "numbering disabled" marker, which reports `None` with no style fallback — otherwise the numbering defined by the paragraph's style applies, with `w:basedOn` inheritance chains resolved. Rendered display numbers (e.g. "7.2(a)") are not computed.

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

Batch counterpart to `get_paragraph_location()`: pair every paragraph with its structural location in one pass, precomputing table indexes, style outline levels, style numbering, heading paths, and section indexes once instead of rescanning the document per ref. Text-box content is excluded, as it is from [`paragraph_count()`](#paragraph_count).

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

Get flattened visible document text. Inserted text is included and deleted text is excluded. A tab mark (`<w:tab/>`) renders as one `\t` character — the same coordinate space `find_text()` searches and `SearchResult.start`/`end` index. A pending move counts like an insertion at its destination and a deletion at its source: its `w:moveTo` text is included and its `w:moveFrom` text excluded, so paragraph hashes and the refs built on them reflect the moved text at its destination only. Text inside a drawing's text box is excluded too — it belongs to the box, not to any addressable paragraph. A document whose content lives entirely in text boxes therefore returns nothing but the separators between its host paragraphs; check [`has_textbox_content`](#has_textbox_content) before reporting it as empty.

**Returns:** Visible text with paragraphs separated by newlines (str)

**Example:**

```python
text = doc.get_visible_text()
```

#### `get_original_text()`

Get flattened original (pre-revision) document text. Deleted text is included and inserted text is excluded — the inverse of `get_visible_text()` — and a pending move's text appears at its source (`w:moveFrom`) only. For intra-paragraph revisions this equals what `get_visible_text()` would return after `reject_all()`, without modifying the document (paragraph-level revisions such as inserted paragraph marks only affect line boundaries). Text inside a drawing's text box is excluded, exactly as in `get_visible_text()`. Read-only: paragraph references and editing operations keep working on the visible view.

**Returns:** Original text with paragraphs separated by newlines (str)

**Example:**

```python
text = doc.get_original_text()
```

#### `find_text(text, occurrence=0, paragraph=None)`

Find text in the document, including text spanning XML element boundaries. Text-box content is excluded, as it is from [`paragraph_count()`](#paragraph_count). A tab mark is `\t` in the searchable text: `"Name\tValue"` matches `Name<tab>Value`, while `"NameValue"` does not (the `TextNotFoundError` message and its `paragraph_preview` spell the tab as the two characters `\t`, so it cannot be mistaken for a space).

**Parameters:**

- `text` (str): Text to search for (must be non-empty; may contain `\t` to match a tab mark)
- `occurrence` (int): Which occurrence to return, 0-based (0 = first). Counted document-wide when `paragraph` is None, and within the paragraph when scoped. Defaults to 0.
- `paragraph` (str, optional): Paragraph reference (e.g. `P2#f3c1`) to scope the search — the same scoping `find_all` offers. `None` searches the whole document. Defaults to None.

**Returns:** [`SearchResult`](#searchresult), or None if not found

**Raises:** `ValueError` if `text` is empty or `paragraph` is malformed; [`ParagraphIndexError`](#paragraphindexerror) / [`HashMismatchError`](#hashmismatcherror) for an out-of-range or stale `paragraph`.

**Example:**

```python
match = doc.find_text("Aim: To")
if match:
    if match.spans_revision:
        print("Text spans a tracked-revision boundary")
    # Pass the match itself: it pins the text, the paragraph and the occurrence
    doc.replace(match, "Goal: To")
```

#### `find_all(text, paragraph=None)`

Find every match of `text`, in document order. One call replaces the N+1
`find_text` probes needed to enumerate N hits, and each result carries exactly
what a follow-up edit needs. Text-box content is excluded, as it is from [`paragraph_count()`](#paragraph_count).

**Parameters:**

- `text` (str): Text to search for (must be non-empty)
- `paragraph` (str, optional): Paragraph reference (e.g. `P2#f3c1`) to scope the search. `None` searches the whole document. Defaults to None.

**Returns:** list of [`SearchResult`](#searchresult), empty when nothing matches (no-match is not an error for an enumeration API)

**Raises:** `ValueError` if `text` is empty or `paragraph` is malformed; [`ParagraphIndexError`](#paragraphindexerror) / [`HashMismatchError`](#hashmismatcherror) for an out-of-range or stale `paragraph`.

**Example:**

```python
# Edit every match in one atomic batch. reversed() puts same-paragraph ops in
# the required descending occurrence order, so this is safe however the
# matches are distributed:
ops = [EditOperation.replace(r, "60 days") for r in reversed(doc.find_all("30 days"))]
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

Count visible text matches across the document. Text-box content is excluded, as it is from [`paragraph_count()`](#paragraph_count). Tab marks are `\t` in the searched text, as in `find_text()`.

**Parameters:**

- `text` (str): Text to search for

**Returns:** Number of occurrences found (int)

**Example:**

```python
if doc.count_matches("Section 5") > 1:
    print("Use paragraph refs and occurrence to target the intended match")
```

#### `replace(find, replace_with, *, paragraph=None, occurrence=None, note=None)`

Replace text with tracked changes. When the target sits inside another author's pending insertion, that insertion is preserved: the matched text gets a nested `<w:del>` under your authorship and the replacement lands in your own sibling `<w:ins>` (Word's behavior), instead of silently rewriting the other author's proposal.

Words shared by `find` and `replace_with` at either end are trimmed first, so only the changed words become revisions — a replace that only adds or only removes words is written as a pure insertion or deletion. When the trimmed spans are **whitespace only** on both sides, trimming continues character by character, so a spacing fix (`"clause  2"` → `"clause 2"`) becomes a pure deletion rather than an invisible deletion/insertion pair. Word spans are never trimmed character-wise: `"30 days"` → `"60 days"` stays a whole-word replacement. The replacement insertion carries the formatting (`rPr`) that covers the most characters of the replaced span (runs sharing identical formatting tally together), ties breaking to the earliest-seen formatting. **Accepting** a replace that straddled mixed formatting therefore leaves the replacement uniformly in that one majority format, while each deletion run keeps its own original formatting — so a **reject** restores the pre-edit mix. When `replace_with` equals the found text, the call is a **no-op**: no revisions are created and the returned `EditResult` equals the input `paragraph` ref with `group_id=None` and `revision_ids=()` — that triple is how callers detect the no-op.

A replace landing wholly inside **your own** pending insertion **amends** that insertion instead of tracking a change against it: the text is spliced in at the match position, whether the match covers part of the insertion or all of it. Your own unsaved text was never in the document, so there is nothing to counter-propose — no revision is created, and the `EditResult` comes back with `group_id=None` and `revision_ids=()` (with an updated ref). To undo an amendment, reject the group of the insertion it amended — the one holding the end of the match, which keeps its id and its group. A match spanning two of your own adjacent insertions consolidates into that one, dropping any insertion it consumed whole.

**Parameters:**

- `find` (str | [`SearchResult`](#searchresult)): Text to find and replace, or a match from `find_text()`/`find_all()` — which also supplies `paragraph` and `occurrence` (pass neither with it). Must not contain `\t`: a tab mark can be matched but not replaced yet — target the text beside it (`ValueError`, ROADMAP.md #6)
- `replace_with` (str): Replacement text
- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`. Required unless `find` is a `SearchResult`.
- `occurrence` (int | None): Which occurrence within the paragraph, 0-based (0 = first). Omitted → the target must be unique within the paragraph; if it matches more than once, [`AmbiguousTextError`](#ambiguoustexterror) is raised instead of silently editing the first match.
- `note` (str | None): Rationale for this edit, anchored as a comment on the revisions it creates — see [Rationale notes](#rationale-notes)

**Returns:** Updated paragraph reference ([`EditResult`](#editresult) — a `str` subclass also carrying the edit's `group_id`/`changeset_id`/`revision_ids`/`comment_id`)

**Warns:** [`UnanchoredNoteWarning`](#unanchorednotewarning) if `note=` was given but the edit created no revision to anchor it on — the edit still applies.

**Example:**

```python
ref = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
doc.replace("net", "gross", paragraph=ref)
if match := doc.find_text("Manager"):
    doc.replace(match, "Director")          # straight from a search
```

#### `delete(text, *, paragraph=None, occurrence=None, note=None)`

Mark text as deleted with tracked changes. Deleting text inside another author's pending insertion nests a `<w:del>` under your authorship inside their `<w:ins>`, preserving their proposal; only your own pending insertions are edited in place.

**Parameters:**

- `text` (str | [`SearchResult`](#searchresult)): Text to mark as deleted, or a match from `find_text()`/`find_all()` — which also supplies `paragraph` and `occurrence` (pass neither with it). Must not contain `\t`: a tab mark can be matched but not deleted yet — target the text beside it (`ValueError`, ROADMAP.md #6)
- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`. Required unless `text` is a `SearchResult`.
- `occurrence` (int | None): Which occurrence within the paragraph, 0-based (0 = first). Omitted → the target must be unique within the paragraph; if it matches more than once, [`AmbiguousTextError`](#ambiguoustexterror) is raised instead of silently editing the first match.
- `note` (str | None): Rationale for this edit, anchored as a comment on the revisions it creates — see [Rationale notes](#rationale-notes)

**Returns:** Updated paragraph reference ([`EditResult`](#editresult) — a `str` subclass also carrying the edit's `group_id`/`changeset_id`/`revision_ids`/`comment_id`)

**Warns:** [`UnanchoredNoteWarning`](#unanchorednotewarning) if `note=` was given but the edit created no revision to anchor it on — the edit still applies.

**Example:**

```python
ref = doc.delete("obsolete clause", paragraph="P5#d4e5")
if match := doc.find_text("obsolete clause"):
    doc.delete(match)                       # same edit, from a search
```

#### `insert_after(anchor, text, *, paragraph=None, occurrence=None, note=None)`

Insert text after anchor with tracked changes. An anchor inside another author's pending insertion produces your own sibling `<w:ins>` (splitting theirs when the anchor falls mid-content) rather than splicing your words into their proposal.

**Parameters:**

- `anchor` (str | [`SearchResult`](#searchresult)): Text to find as insertion point, or a match from `find_text()`/`find_all()` — which also supplies `paragraph` and `occurrence` (pass neither with it)
- `text` (str): Text to insert after the anchor
- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`. Required unless `anchor` is a `SearchResult`.
- `occurrence` (int | None): Which occurrence within the paragraph, 0-based (0 = first). Omitted → the target must be unique within the paragraph; if it matches more than once, [`AmbiguousTextError`](#ambiguoustexterror) is raised instead of silently editing the first match.
- `note` (str | None): Rationale for this edit, anchored as a comment on the revisions it creates — see [Rationale notes](#rationale-notes)

**Returns:** Updated paragraph reference ([`EditResult`](#editresult) — a `str` subclass also carrying the edit's `group_id`/`changeset_id`/`revision_ids`/`comment_id`)

**Warns:** [`UnanchoredNoteWarning`](#unanchorednotewarning) if `note=` was given but the edit created no revision to anchor it on — the edit still applies.

**Example:**

```python
ref = doc.insert_after("Section 5", " (as amended)", paragraph="P3#b2c4")
if match := doc.find_text("Section 5"):
    doc.insert_after(match, " (as amended)")
```

#### `insert_before(anchor, text, *, paragraph=None, occurrence=None, note=None)`

Insert text before anchor with tracked changes. Foreign pending insertions are treated the same as in `insert_after()`.

**Parameters:**

- `anchor` (str | [`SearchResult`](#searchresult)): Text to find as insertion point, or a match from `find_text()`/`find_all()` — which also supplies `paragraph` and `occurrence` (pass neither with it)
- `text` (str): Text to insert before the anchor
- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`. Required unless `anchor` is a `SearchResult`.
- `occurrence` (int | None): Which occurrence within the paragraph, 0-based (0 = first). Omitted → the target must be unique within the paragraph; if it matches more than once, [`AmbiguousTextError`](#ambiguoustexterror) is raised instead of silently editing the first match.
- `note` (str | None): Rationale for this edit, anchored as a comment on the revisions it creates — see [Rationale notes](#rationale-notes)

**Returns:** Updated paragraph reference ([`EditResult`](#editresult) — a `str` subclass also carrying the edit's `group_id`/`changeset_id`/`revision_ids`/`comment_id`)

**Warns:** [`UnanchoredNoteWarning`](#unanchorednotewarning) if `note=` was given but the edit created no revision to anchor it on — the edit still applies.

**Example:**

```python
ref = doc.insert_before("Section 6", "New clause: ", paragraph="P4#a7b2")
if match := doc.find_text("Section 6"):
    doc.insert_before(match, "New clause: ")
```

> **Passing a `SearchResult`** to `replace`/`delete`/`insert_after`/`insert_before`/`add_comment` (or an `EditOperation` constructor) is exactly equivalent to spelling out its `text`, `paragraph_ref` and `paragraph_occurrence`. Supplying `paragraph=` or `occurrence=` *as well* raises `ValueError` rather than picking a winner, and a match whose paragraph was edited in between raises [`HashMismatchError`](#hashmismatcherror) like any other stale ref.

> **Newlines split paragraphs.** A `\n` in any content text (`replace`'s
> `replace_with`, `insert_after`/`insert_before`'s `text`, `rewrite_paragraph`'s
> `new_text`, or a `batch_edit` op) is a **tracked paragraph split**: the first
> paragraph's mark is flagged inserted and the tail moves into a new paragraph.
> Accepting keeps the split; rejecting rejoins. All other C0 control characters
> (`\r`, NUL, …) are rejected with a teaching `ValueError`. See
> [`split_paragraph`](#split_paragraphref-before-occurrencenone) and [`EditResult`](#editresult).

> **Tabs are searchable, not writable.** A `<w:tab/>` mark is one `\t`
> character in the visible and original text views — hence in search, offsets
> and hashes; `get_markup_text()` does not render it — so search and anchor
> text may contain `\t`:
> `insert_after("Name\t", "x")` lands right after the tab, and a `\n` split
> may fall on either side of one. Content text (`replace_with`, insert `text`,
> `note`, comment bodies) still rejects `\t` — nothing writes a tracked tab —
> and a `replace`/`delete` target that contains `\t` is rejected with a
> `ValueError`: a tab can be matched but not removed yet (ROADMAP.md #6).
> Because tabs take part in paragraph hashes, refs of tab-bearing paragraphs
> differ from those computed before tabs were mapped (refs are session-scoped
> anyway).

#### Rationale notes

`note=` on `replace`, `delete`, `insert_after`, `insert_before`,
`rewrite_paragraph` and each `batch_edit` operation attaches the *why* of a
redline to the redline itself: the note becomes a Word comment whose range
brackets exactly the revisions that edit created, so a reviewer opening the
document sees the change and its rationale together. Five rules govern it.

1. **One comment per operation, deduplicated per call.** Twenty operations of
   one `batch_edit` sharing the same note text produce **one** comment,
   anchored on the first of them, and all twenty `EditResult`s report that same
   `comment_id`. Twenty different notes produce twenty comments. Dedupe scope
   is a single call — two separate `replace(note="X")` calls are two proposals
   at two moments, so two comments.

2. **The note does not outlive the revision it explains** (within the session
   that made it — see rule 5). When the last revision a note covers is
   resolved, the comment is deleted — on **accept**
   as well as on **reject**, and through every verb: `accept_revision`,
   `reject_revision`, `accept_group`, `reject_group`, `accept_changeset`,
   `reject_changeset`, `accept_all`, `reject_all`. A pipeline that calls
   `accept_all()` to produce a clean deliverable therefore does not ship agent
   rationale as live comments. Any replies threaded under the note go with it.

   The test is whether the revisions still exist, not which call removed them,
   so two cases that involve no accept or reject of yours also end the note: a
   later edit that amends the annotated insertion out of existence, and a
   rejection of *another author's* insertion that carries away a revision of
   yours nested inside it — including one made by `reject_all(author=...)`
   naming them. An `accept_all(author=...)` naming somebody else leaves notes
   whose revisions survive it untouched.

   A note's range never points at text nobody changed: when the edit its
   markers bracket is resolved while another edit it explains is still
   pending, the whole thread — replies included — moves onto that surviving
   redline. When no edit it explains can carry a comment marker any more, the
   note goes with the anchor it lost.

   For rationale that must survive resolution, use
   [`add_comment()`](#add_commentanchor_text-comment-paragraphnone-occurrencenone)
   instead: plain comments are document-scoped, note comments are
   revision-scoped.

3. **`EditResult.comment_id`** holds the id of the comment created, or `None`
   when no note was given. It is an ordinary, live comment id — usable with
   `reply_to_comment()`, `resolve_comment()` and `delete_comment()`. Deleting
   it yourself is final and takes the whole thread: the note stops tracking
   the revisions it explained, so resolving them later neither resurrects nor
   re-anchors it.

4. **Nothing to anchor warns, never fails.** Some edits create no revision a
   comment could bracket: a no-op `replace` (`find` equals `replace_with`), an
   edit that amends one of your own pending insertions, a rewrite that found no
   differences, and a bare `"\n"` whose only revision is the paragraph mark
   (which lives in `w:pPr/w:rPr`, where a comment marker cannot go). In each
   case the **edit still applies**, `comment_id` is `None`, and an
   `UnanchoredNoteWarning` names which cause it was — a dropped rationale is
   never silent. An operation that shares its note text with a sibling
   operation of the same call that *did* anchor it is not a dropped rationale:
   it reports that shared `comment_id` and warns nothing. To silence the
   category:

   ```python
   warnings.filterwarnings("ignore", category=UnanchoredNoteWarning)
   ```

5. **The link is per-session.** The comment itself is an ordinary OOXML comment
   and survives `save()` unchanged. The note-to-revision *link* does not — it
   is per-open-`Document`, exactly like `group_id` and `changeset_id`. After a
   reopen a note comment is just a comment: rejecting the (freshly inferred)
   group leaves it in place, and `delete_comment()` removes it.

   A `note=` edit writes the comment parts into the workspace as
   [`add_comment()`](#add_commentanchor_text-comment-paragraphnone-occurrencenone)
   does, where a bare tracked-change edit stays in memory until `save()` — so
   it carries that method's workspace side effect too.

A note is validated **before** the edit runs, so a bad one never leaves an
applied edit with a dropped rationale behind it: it must be a non-empty string
with no control characters (`\n` included — a comment body is a single
`<w:t>`), or `ValueError`. In a batch, an invalid note fails the whole batch
atomically with [`BatchOperationError`](#batchoperationerror), and `dry_run=True`
reports the offending row.

**Example:**

```python
from docx_editor import EditOperation

result = doc.replace(
    "30 days", "60 days",
    paragraph="P2#f3c1",
    note="Aligns payment terms with the master agreement (§4.2).",
)
print(result.comment_id)     # 0 — the rationale, anchored on the redline

# One rationale covering several redlines: one comment, three results.
results = doc.batch_edit([
    EditOperation.replace("Manager", "Director", paragraph="P3#a7b2", note="Title change per HR."),
    EditOperation.replace("manager", "director", paragraph="P5#c4d8", note="Title change per HR."),
    EditOperation.delete("(interim)", paragraph="P7#e1f9", note="Title change per HR."),
])
assert len({r.comment_id for r in results}) == 1

doc.reject_group(result.group_id)   # the redline and its rationale go together
doc.accept_all()                    # clean deliverable: no revisions, no notes
```

#### `split_paragraph(ref, *, before, occurrence=None)`

Split a paragraph into two with a tracked paragraph break, cutting immediately
before `before`. Explicit sugar for the `\n`-means-split behavior (equivalent to
`insert_before(before, "\n", ...)`): the paragraph mark is flagged as an inserted
revision and the tail (from `before` on) moves into a new following paragraph.
Accepting keeps the split; rejecting the group rejoins the two paragraphs.

**Parameters:**

- `ref` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`
- `before` (str, keyword-only): Text to split before; the break lands at its start. Must be a non-empty string in the paragraph.
- `occurrence` (int | None): Which occurrence of `before` within the paragraph, 0-based. Omitted → `before` must be unique in the paragraph, else [`AmbiguousTextError`](#ambiguoustexterror).

**Returns:** The first paragraph's reference ([`EditResult`](#editresult)); its `refs` tuple carries the refs of **both** resulting paragraphs.

**Example:**

```python
result = doc.split_paragraph("P2#f3c1", before="However,")
first_ref, second_ref = result.refs
doc.reject_group(result.group_id)  # rejoin
```

#### `rewrite_paragraph(ref, new_text, *, note=None)`

Rewrite a paragraph using tracked changes generated from a word-level diff.

A rewrite typically produces many revisions (one per diff hunk), none of which
is a self-contained edit — accepting only some of them by id garbles the
paragraph. All of one rewrite's revisions therefore share a revision group;
resolve them as a unit with [`accept_group()`](#accept_groupgroup_id) /
[`reject_group()`](#reject_groupgroup_id).

**Parameters:**

- `ref` (str): Paragraph reference from `list_paragraphs()`
- `new_text` (str): Desired paragraph text. Must hold the same number of `\t` as the paragraph has tab marks: the text between consecutive tabs is rewritten segment by segment, so `rewrite_paragraph(ref, info.text.replace(...))` works on tab-bearing paragraphs and words may move across a tab, while a rewrite that adds or removes a tab raises `ValueError` (ROADMAP.md #6)
- `note` (str | None): Rationale for the rewrite, anchored as one comment spanning its first through last revision — see [Rationale notes](#rationale-notes)

**Returns:** Updated paragraph reference ([`EditResult`](#editresult) — a `str` subclass also carrying the edit's `group_id`/`changeset_id`/`revision_ids`; `group_id` is `None` when `new_text` equals the current text, or when every change landed inside your own pending insertions and amended them in place — undo those by rejecting the group of the amended insertion; `comment_id` carries a `note=`'s comment)

**Warns:** [`UnanchoredNoteWarning`](#unanchorednotewarning) if `note=` was given but the edit created no revision to anchor it on — the edit still applies.

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
- `note` on each [`EditOperation`](#editoperation) (str | None): Rationale for that operation — see [Rationale notes](#rationale-notes)
- `dry_run` (bool): If True, validate every operation without applying any edits and return one [`EditValidationResult`](#editvalidationresult) per operation, in input order; the document is left unchanged. Each operation is validated independently against the current document — sequential effects between multiple operations on the same paragraph are **not** simulated. Defaults to False.

**Returns:** Updated paragraph references in input order (list of [`EditResult`](#editresult)) — each operation that creates revisions gets its own revision group, so one op can be accepted and another rejected (`group_id` is `None` for an op that created no new revisions, e.g. text spliced into one of your own pending insertions); with `dry_run=True`, a list of [`EditValidationResult`](#editvalidationresult) instead. Operations carrying `note=` also carry the resulting `comment_id`; operations of one call sharing the same note text share one comment.

**Warns:** [`UnanchoredNoteWarning`](#unanchorednotewarning) if `note=` was given but the edit created no revision to anchor it on — the edit still applies.

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
smart-quote splits, `w:ins` wrappers) are found. Text-box content is excluded, as it is from [`paragraph_count()`](#paragraph_count).

**Parameters:**

- `anchor_text` (str | [`SearchResult`](#searchresult)): Text to attach the comment to, or a match from `find_text()`/`find_all()` — which also supplies `paragraph` and `occurrence` (pass neither with it). May contain `\t` (a tab mark) anywhere, including at either end — the comment range brackets the tab
- `comment` (str): The comment content
- `paragraph` (str, optional): Paragraph reference (e.g. `P3#a7b2`) to scope the search. `None` searches the whole document. Defaults to None.
- `occurrence` (int | None): Which occurrence to anchor to, 0-based (0 = first), counted within `paragraph` when given and document-wide otherwise. Omitted → the anchor must be unique in the search scope, else [`AmbiguousTextError`](#ambiguoustexterror).

**Returns:** The comment ID (int). IDs are allocated sequentially starting at 0 in a document with no existing comments — always use the returned ID rather than guessing.

**Example:**

```python
cid = doc.add_comment("Section 5", "Please review this section")
doc.add_comment("term", "Note on the 2nd 'term'", paragraph="P3#a7b2", occurrence=1)
doc.add_comment(doc.find_all("term")[1], "Same 2nd 'term', from the search")
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

Delete a comment, and every reply threaded under it, from the document. A reply
is linked to its parent only by `w15:paraIdParent`, so the thread goes as a
unit — left behind, a reply would point at a paragraph id no part of the
document still holds.

**Parameters:**

- `comment_id` (int): ID of the comment to delete

**Returns:** True if deleted, False if not found (bool)

**Example:**

```python
doc.delete_comment(0)
```

### Revision Management Methods

#### `list_revisions(author=None, paragraph=None)`

List the document's tracked revisions.

Five `type` values: `"insertion"`, `"deletion"`, the two halves of a content
move (`"move_from"`/`"move_to"` — two rows, as in Word's revision pane) and
`"property_change"` (a `w:pPrChange`: the paragraph's previous properties;
`text` is `""`). Every row can be passed to `accept_revision()` /
`reject_revision()`. Every other revision type in the OOXML schema — run,
section and table property changes, table-structure revisions,
`w:numberingChange` and the custom-XML range marks — is listed by
[`list_unhandled_revisions()`](#list_unhandled_revisionsauthornone) instead,
because none of it can be. A move's range marks (`w:moveFromRangeStart` etc.)
are scaffolding: never listed, swept once the move they bracket is resolved.

**Parameters:**

- `author` (str, optional): If provided, filter by author name
- `paragraph` (str, optional): If provided, only return revisions in this paragraph (hash-anchored ref from `list_paragraphs()`, e.g. `"P3#a7b2"`)

**Returns:** List of Revision objects

**Raises:** The `paragraph` ref is validated exactly like in the edit methods — `ValueError` (malformed ref), [`ParagraphIndexError`](#paragraphindexerror) (index out of range), [`HashMismatchError`](#hashmismatcherror) (stale hash).

**Example:**

```python
revisions = doc.list_revisions()
for r in revisions:
    print(f"{r.type}: {r.text} by {r.author}")
# insertion: speedy by Reviewer
# move_from: relocated clause by Ann
# move_to: relocated clause by Ann
# property_change:  by Ann
```

#### `accept_revision(revision_id)`

Accept a revision by ID.

- For insertions: keeps the inserted content
- For deletions: permanently removes the deleted content
- For `move_to`: keeps the moved text at its destination
- For `move_from`: removes the moved text from its source
- For `property_change`: keeps the paragraph's current properties and drops the record of the previous ones
- Nested revisions: accepting an insertion unwraps it in place, so a deletion another author nested inside it survives as an independent pending deletion

**A move is two rows.** `accept_all()`/`reject_all()` and the inferred
changeset both halves share (`accept_changeset(rev.changeset_id)`) resolve
them together, which is what keeps the text in exactly one place. Resolving
one half by id is allowed — Word allows it too — but it is your call:
accepting the `move_to` and rejecting the `move_from` duplicates the text, the
inverse loses it. A lone half in a damaged file behaves as what it structurally
is: a `move_from` alone is a deletion, a `move_to` alone an insertion.

**Parameters:**

- `revision_id` (int): ID of the revision to accept

**Returns:** True if accepted, False if not found (bool)

**Note:** any `note=` rationale left with no live revision to explain is deleted with it, replies included — see [Rationale notes](#rationale-notes).

**Example:**

```python
doc.accept_revision(1)
```

#### `reject_revision(revision_id)`

Reject a revision by ID.

- For insertions: removes the inserted content
- For deletions: restores the deleted content
- For `move_to`: removes the moved text from its destination — including any edit of your own made inside it (a foreign `w:moveTo` is not split around your edits the way a foreign insertion is; a known gap). Edit after resolving the move, or reject it by id first
- For `move_from`: restores the moved text at its source
- For `property_change`: restores the paragraph's recorded previous properties. A record with none (LibreOffice writes a self-closing `w:pPrChange` for "previously no properties") clears them; the recorded style id is restored verbatim even when the document defines no such style (Word falls back to Normal for it)
- Nested revisions: rejecting an insertion removes everything inside it — deletions another author nested inside it disappear with it

Resolve both halves of a move together — see `accept_revision()`.

**Parameters:**

- `revision_id` (int): ID of the revision to reject

**Returns:** True if rejected, False if not found (bool)

**Note:** any `note=` rationale left with no live revision to explain is deleted with it, replies included — see [Rationale notes](#rationale-notes).

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
That heuristic matches this library's own edits exactly: each changeset (one
edit call, or one whole `batch_edit`/`batch_rewrite` call) stamps a
collision-bumped whole-second date, so within one open session two changesets
by the same author never share a second; all ops of one batch call share one
date (one changeset) and same-paragraph batch ops merge by design. The
collision counter is per-session (not seeded from dates already in the file),
so own writes from a *previous* session merge like foreign revisions: any
revisions with identical `w:author` + `w:date` merge (`w:date` has second
precision). When Word already resolved part of a former edit, the remainder
reconstructs as a smaller (rump) group, and
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

**Note:** any `note=` rationale left with no live revision to explain is deleted with it, replies included — see [Rationale notes](#rationale-notes).

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

**Note:** any `note=` rationale left with no live revision to explain is deleted with it, replies included — see [Rationale notes](#rationale-notes).

**Example:**

```python
result = doc.rewrite_paragraph(ref, "New text.")
doc.reject_group(result.group_id)  # undo the whole rewrite
```

#### `accept_changeset(changeset_id)`

Accept every revision created by **one whole call** — the *intent* tier, one
level above a group:

```
one call (a single edit, or an entire batch_edit / batch_rewrite)
  = one changeset  ⊇  one-or-more groups  ⊇  revisions
```

A single edit is a one-group changeset; a whole `batch_edit`/`batch_rewrite` is
one changeset over all the groups it created. Each returned
[`EditResult`](#editresult) carries the `changeset_id`, and `list_revisions()`
stamps `changeset_id`/`changeset_source` on each member revision. Accepting the
changeset applies the entire call. **There is no tier above this** — the model
stops at three: revision < group < changeset.

A changeset is the `(author, date)` **equivalence class over groups** — a
*global* class, not a contiguous run: a `batch_edit` whose ops land in different
paragraphs is one changeset even though its groups are non-contiguous. Edits
made here **record** their changeset (`changeset_source == "recorded"`);
revisions already in the file get **inferred** changesets, partitioning the
reconstructed groups by identical `w:author` + `w:date`
(`changeset_source == "inferred"`). Same lifetime and caveats as groups:
changeset ids are in-memory and per-open-`Document`, renumbered on each open, so
always take one from the current session's `EditResult` or `list_revisions()`.
Rump-tolerant — after Word has resolved part of the changeset, the remainder
resolves fine.

**Parameters:**

- `changeset_id` (int): Changeset id from an `EditResult` (or a `Revision.changeset_id`)

**Returns:** Number of revisions accepted across the changeset's groups (int). Members already resolved individually are skipped (and not counted).

**Raises:** [`RevisionError`](#revisionerror) if the changeset id is unknown to this open Document.

**Note:** any `note=` rationale left with no live revision to explain is deleted with it, replies included — see [Rationale notes](#rationale-notes).

**Example:**

```python
results = doc.batch_edit([...])           # one changeset over several groups
doc.accept_changeset(results[0].changeset_id)  # accept the whole batch at once
```

#### `reject_changeset(changeset_id)`

Reject every revision created by one whole call — the counterpart of
[`accept_changeset()`](#accept_changesetchangeset_id). Rejecting the changeset
undoes the entire call (every group), restoring the exact pre-call text. Same
changeset semantics and lifetime as `accept_changeset()` — including recorded
vs inferred changesets and per-open renumbering.

**Parameters:**

- `changeset_id` (int): Changeset id from an `EditResult` (or a `Revision.changeset_id`)

**Returns:** Number of revisions rejected across the changeset's groups (int). Members already resolved individually are skipped (and not counted).

**Raises:** [`RevisionError`](#revisionerror) if the changeset id is unknown to this open Document.

**Note:** any `note=` rationale left with no live revision to explain is deleted with it, replies included — see [Rationale notes](#rationale-notes).

**Example:**

```python
results = doc.batch_edit([...])
doc.reject_changeset(results[0].changeset_id)  # undo the whole batch
```

#### `accept_all(author=None)`

Accept every listed revision.

Resolves insertions, deletions, content moves (both halves as a unit — the
text ends up at its destination exactly once, range marks swept) and
paragraph-property changes (the record is dropped). Every other revision type
is left pending and reported on the result rather than silently ignored — see
[`ResolveResult`](#resolveresult) for the counting rule and
[`list_unhandled_revisions()`](#list_unhandled_revisionsauthornone) for what is
in that set. A moved or deleted paragraph mark resolves approximately: the
marker is dropped without merging or splitting paragraphs.

Only `word/document.xml` is inspected; headers, footers and footnotes are the
container-parts epic (ROADMAP.md #30).

**Parameters:**

- `author` (str, optional): If provided, only accept revisions by this author

**Returns:** [`ResolveResult`](#resolveresult) — an `int` whose value is the
number of revisions accepted, carrying `.unhandled` and `.unhandled_types`.

**Warns:** [`UnhandledRevisionWarning`](#unhandledrevisionwarning) if
`.unhandled` is nonzero.

**Note:** any `note=` rationale left with no live revision to explain is deleted with it, replies included — see [Rationale notes](#rationale-notes).

**Example:**

```python
result = doc.accept_all()
print(f"Accepted {result} revisions")
if result.unhandled:
    print(f"Still pending: {result.unhandled_types}")
    # -> Still pending: {'w:rPrChange': 3, 'w:cellIns': 1}
```

#### `reject_all(author=None)`

Reject every listed revision.

Resolves insertions, deletions, content moves (both halves as a unit — the
text is back at its source exactly once) and paragraph-property changes (the
recorded previous properties are restored); every other revision type is left
pending and reported exactly as in `accept_all()`.

**Parameters:**

- `author` (str, optional): If provided, only reject revisions by this author

**Returns:** [`ResolveResult`](#resolveresult) — an `int` whose value is the
number of revisions rejected, carrying `.unhandled` and `.unhandled_types`.

**Warns:** [`UnhandledRevisionWarning`](#unhandledrevisionwarning) if
`.unhandled` is nonzero.

**Note:** any `note=` rationale left with no live revision to explain is deleted with it, replies included — see [Rationale notes](#rationale-notes).

**Example:**

```python
result = doc.reject_all(author="OtherUser")
if result.unhandled:
    print(f"Still pending: {result.unhandled_types}")
```

#### `list_unhandled_revisions(author=None)`

List the revision types this library does not accept or reject.

The complement of `list_revisions()`: everything in the OOXML revision schema
except insertions, deletions, moves and paragraph-property changes — run,
section and table property changes (`w:rPrChange`, `w:sectPrChange`, the table
`*PrChange` family), table-structure revisions (`w:cellIns`, `w:cellDel`,
`w:cellMerge`), `w:numberingChange` and the custom-XML range marks. These
survive open/edit/save unchanged and are left pending by
`accept_all()`/`reject_all()`. A move's range marks are never listed here:
they are swept with the move they bracket. A handled-type mark that carries no numeric `w:id` (a nonconforming producer) is listed here rather than omitted: nothing id-keyed can resolve it, and it must not vanish from both listings.

Call this before telling a human "all changes accepted": on a run-format-only
redline `accept_all()` returns 0 because there was nothing it *could* accept,
not because there was nothing to do.

Rows are [`UnhandledRevision`](#unhandledrevision), deliberately not `Revision`
objects — they carry nothing `accept_revision()` could act on.

Only `word/document.xml` is inspected (ROADMAP.md #30).

**Parameters:**

- `author` (str, optional): If provided, filter by author name. Marks with no
  `w:author` attribute read as `"Unknown"`, so they match only
  `author="Unknown"`, are excluded from every other filtered call, and are
  included in an unfiltered one.

**Returns:** List of `UnhandledRevision` in document order

**Example:**

```python
result = doc.accept_all()
if result.unhandled:
    for row in doc.list_unhandled_revisions():
        print(f"still pending: {row.tag} by {row.author} @{row.paragraph_ref}")
```

### Save and Close Methods

#### `save(path=None, validate=False, force=False, *, track_changes=None)`

Save the document.

**Parameters:**

- `path` (str | Path, optional): Output path. Defaults to original source path.
- `validate` (bool): If True, validate with LibreOffice before saving. Defaults to False.
- `force` (bool): If True, skip save-time safety checks. By default `save()` refuses to overwrite the source if it changed on disk since it was opened (raising [`WorkspaceSyncError`](#workspacesyncerror)), or to write a destination that appears open in Word — a `~$` owner file exists next to it (raising [`DocumentOpenError`](#documentopenerror)). Pass `force=True` only for a confirmed-stale lock left by a crashed session. Defaults to False.
- `track_changes` (bool, optional): Whether to turn Word's track-changes switch (`<w:trackRevisions/>` in `word/settings.xml`) on in the saved file. `None` (the default) writes it exactly when the document carries a pending revision authored under this session's author name; `True` writes it either way (warning if the document has no `word/settings.xml` to write it into); `False` leaves `settings.xml` alone. See below.

**Returns:** Path to the saved document (Path)

**The track-changes switch.** A saved redline shows up in Word whether or not the switch is on, but the switch is what keeps the *recipient's* own typing tracked — without it the next round of edits arrives untracked and the two rounds can no longer be told apart. So a save of a document holding a revision authored by this session also writes `<w:trackRevisions/>`. The test is the document's state rather than whether this session edited, so a pending redline of ours reopened from an earlier session counts too — it is still awaiting a reply. A document holding no revision of ours (one not redlined, or one whose revisions were accepted) is saved untouched; `track_changes=False` opts out entirely and never removes a flag the document already had. A document that turns tracking *off* explicitly (`<w:trackRevisions w:val="false"/>`) is respected rather than overridden: the element is left as it is and the save emits a `UserWarning` saying the recipient's edits will not be tracked. `track_changes=True` overrides that setting.

After saving to a different path (or a save that fails), the workspace is flagged as holding unsaved changes; a later `Document.open()` of the source raises `WorkspaceSyncError` until the workspace is saved back to the source or discarded — with `force_recreate=True` on the next open, or once with [`Document.discard_workspace(path)`](#documentdiscard_workspacepath-workspace_dirnone). See [`WorkspaceSyncError`](#workspacesyncerror) below.

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

Represents a tracked change: an insertion, a deletion, one half of a content
move, or a paragraph-property change.

```python
from docx_editor import Revision
```

### Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `id` | int | The revision ID |
| `type` | str | `"insertion"`, `"deletion"`, `"move_from"`, `"move_to"` or `"property_change"`. The two halves of a move are two rows (Word's revision pane shows "Moved from"/"Moved to" the same way), sharing an inferred `changeset_id` when they carry the same author and date; resolve them together — see [`accept_revision()`](#accept_revisionrevision_id). `"property_change"` is a `w:pPrChange`: the paragraph's previous properties |
| `author` | str | The revision author |
| `date` | datetime or None | When the revision was made |
| `text` | str | The inserted, deleted or moved text; `""` for a property change or a paragraph-mark marker |
| `paragraph_ref` | str or None | Hash-anchored reference (`"P{i}#{hash}"`) of the containing paragraph; None when the revision sits in no addressable paragraph — outside any `<w:p>` (e.g. a `<w:trPr>` row marker), or inside a drawing's text box — still listed, and `accept_all()`/`reject_all()` always resolve it. Anything narrower depends on how the box is stored: Word normally writes a box twice (an `mc:Choice` copy and an `mc:Fallback` copy), so the revision is listed once per copy and one `accept_revision()`/`accept_group()` call resolves only the copy it lands on. Copies with distinct ids and identical author/date join one inferred changeset, which `accept_changeset()` resolves — along with every other group carrying that author and the identical raw `w:date` string. Copies sharing a `w:id` are ungroupable (`group_id` and `changeset_id` both None), so no group- or changeset-keyed call can reach them and `accept_all()`/`reject_all()` is the single call that takes both |
| `occurrence` | int or None | 0-based occurrence index of `text` within the containing paragraph, counted in the view where the revision's text lives (the visible view for insertions and `move_to` halves, the original pre-revision view for deletions and `move_from` halves). For insertions and `move_to` it plugs directly into the `occurrence=` parameter of the anchor APIs; None whenever targeting-by-text does not apply (empty text, a host insertion partly consumed by a nested deletion, a nested deletion, or a None `paragraph_ref`) |
| `nested_under` | int or None | id of the nearest enclosing revision (e.g. a foreign deletion inside another author's pending insertion), else None |
| `contains_ids` | tuple[int, ...] | ids of the revisions nested inside this one, in document order (empty tuple when none). Both nesting fields report *structural* containment and so, unlike `text`, still cross into a text box — accepting a host insertion does not resolve the box's own revisions |
| `group_id` | int or None | Revision group this revision belongs to (see [`accept_group()`](#accept_groupgroup_id)): recorded for this session's edits, inferred by reconstruction for revisions already in the file; None only for ungroupable revisions (missing author/date, outside any paragraph, duplicated id, or a mid-session split half of a foreign insertion) |
| `group_source` | str or None | Provenance of `group_id`: `"recorded"` (created through this open Document) or `"inferred"` (reconstructed at parse time from same-paragraph contiguity + identical author and date); None iff `group_id` is None |
| `changeset_id` | int or None | Changeset (one whole call) this revision's group belongs to (see [`accept_changeset()`](#accept_changesetchangeset_id) / [`reject_changeset()`](#reject_changesetchangeset_id)) — the `(author, date)` class over groups; None iff `group_id` is None |
| `changeset_source` | str or None | Provenance of `changeset_id`: `"recorded"` or `"inferred"`; None iff `changeset_id` is None |

### Example

```python
revisions = doc.list_revisions()
symbols = {"insertion": "+", "deletion": "-", "move_from": "<", "move_to": ">", "property_change": "¶"}
for rev in revisions:
    print(f"{symbols[rev.type]} {rev.text} (by {rev.author})")
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

#### `EditOperation.replace(find, replace_with, *, paragraph=None, occurrence=None, note=None)`

- `find` (str | [`SearchResult`](#searchresult)): Text to find and replace. Must be non-empty.
- `replace_with` (str): Replacement text. Empty string is allowed (replacing
  with nothing is a valid tracked deletion).

#### `EditOperation.delete(text, *, paragraph=None, occurrence=None, note=None)`

- `text` (str | [`SearchResult`](#searchresult)): Text to mark as deleted. Must be non-empty.

#### `EditOperation.insert_after(anchor, text, *, paragraph=None, occurrence=None, note=None)`

#### `EditOperation.insert_before(anchor, text, *, paragraph=None, occurrence=None, note=None)`

- `anchor` (str | [`SearchResult`](#searchresult)): Text to find as insertion point. Must be non-empty.
- `text` (str): Text to insert.

All constructors also take:

- `paragraph` (str): Paragraph reference from `list_paragraphs()`, such as `P2#f3c1`. Required unless the search target is a `SearchResult`, which supplies it.
- `occurrence` (int | None): Which occurrence within the paragraph, 0-based (0 = first). Omitted → the target must be unique within the paragraph at apply time; if it matches more than once, the batch fails with a [`BatchOperationError`](#batchoperationerror) wrapping an [`AmbiguousTextError`](#ambiguoustexterror). Must be >= 0 when given.
- `note` (str | None): Rationale for this operation, anchored as a comment on the revisions it creates. Operations of one `batch_edit` call sharing the same note text share one comment — see [Rationale notes](#rationale-notes).

**Raises:** `ValueError` at construction time if the paragraph ref is missing or
malformed, `occurrence` is negative, a search-target argument (`find`, delete
`text`, `anchor`) is empty, a payload argument (`replace_with`, insert `text`)
is `None` — payloads may be empty strings, search targets may not — or the
search target is a `SearchResult` and `paragraph`/`occurrence` was given too,
or `note` is neither `None` nor a non-empty control-character-free string.
Each signature mirrors the corresponding `Document` method 1:1, so
`doc.replace(...)` translates mechanically to `EditOperation.replace(...)`.

The dataclass fields stay plain strings whichever form you use, so an operation
built from a match is indistinguishable from a hand-written one.

### Example

```python
new_refs = doc.batch_edit([
    EditOperation.replace("30 days", "60 days", paragraph="P2#f3c1"),
    EditOperation.delete("obsolete clause", paragraph="P5#d4e5"),
    EditOperation.insert_after("Section 5", " (as amended)", paragraph="P7#b1c2"),
])

# Or straight from a search — no refs or occurrences to carry:
doc.batch_edit([EditOperation.replace(m, "60 days") for m in reversed(doc.find_all("30 days"))])
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
| `changeset_id` | int or None | Changeset (one whole call) this edit's group belongs to, for [`accept_changeset()`](#accept_changesetchangeset_id) / [`reject_changeset()`](#reject_changesetchangeset_id). Every `EditResult` from one `batch_edit`/`batch_rewrite` shares it; None iff `group_id` is None. Per-open-`Document`, like `group_id`. |
| `revision_ids` | tuple[int, ...] | The `w:id`s of the group's member revisions, in creation order; `()` when `group_id` is None |
| `refs` | tuple[str, ...] | Every resulting paragraph ref, in document order. `(str(self),)` for a normal edit; for a `\n` paragraph split it carries the first paragraph (== the string value) plus one ref per new paragraph the split created. A split shifts later indexes, so re-resolve stale refs before reuse. |
| `comment_id` | int or None | Id of the comment holding this edit's `note=` rationale, anchored on the revisions above — a live comment id, usable with `reply_to_comment()`/`delete_comment()`. `None` when no `note=` was given, and `None` with an `UnanchoredNoteWarning` when a note was given but there was nothing to anchor it on — unless a sibling operation of the same call recorded that same note text, in which case this carries the shared id and nothing warns. See [Rationale notes](#rationale-notes). |

### Example

```python
result = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
print(str(result))          # "P2#c3d4" — the new paragraph ref
print(result.group_id)      # 1
print(result.changeset_id)  # 1 — a single edit is a one-group changeset
print(result.revision_ids)  # (0, 1) — the del and the ins
print(result.refs)          # ("P2#c3d4",) — one ref (a split would give more)
print(result.comment_id)    # None — no note= was passed

# A newline splits the paragraph; refs covers both halves:
split = doc.replace("end.", "end.\nNew paragraph.", paragraph="P2#c3d4")
print(split.refs)           # ("P2#…", "P3#…")

doc.replace("net", "gross", paragraph=result)  # usable as a plain ref
doc.reject_group(result.group_id)              # undo the first edit entirely
```

---

## ResolveResult

The result of `accept_all()` / `reject_all()`: the number of revisions resolved,
plus what could not be.

Subclasses `int`, and the int value *is* the resolved count — so
`count = doc.accept_all()` keeps working in comparisons, arithmetic, f-strings
and `json.dumps`. `isinstance(result, int)` is True; the concrete type is
`ResolveResult`, so an exact-type check (`type(result) is int`) is not.

```python
from docx_editor import ResolveResult
```

### Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| *(the int value)* | int | Number of revisions resolved (insertions, deletions, move halves, paragraph-property changes) |
| `unhandled` | int | How many revision elements this library never resolves are still in the document. On an `author=`-filtered call this counts only that author's marks, matching the scope of the claim being made. `0` on a redline made of insertions, deletions, moves and paragraph-property changes. |
| `unhandled_types` | dict[str, int] | Tag → count for those elements, e.g. `{"w:rPrChange": 3, "w:cellIns": 1}`. Empty when `unhandled` is 0. |

Both are counted **after** resolution, which is the honest measure of the claim
being made ("everything is resolved"): a foreign mark inside a rejected
insertion's subtree is removed with it, so it correctly does not appear. It is
not a census of what the document held on entry.

They count what is still *pending*, so a mark recorded inside a change record —
a `w:cellIns` in a `w:tcPrChange`'s historical `w:tcPr` — is not counted a
second time alongside the change itself.

### Example

```python
result = doc.accept_all()
print(result)             # "2" — str/format stay the plain int
print(result + 1)         # 3
print(result.unhandled)   # 4
print(result.unhandled_types)
# {'w:rPrChange': 3, 'w:sectPrChange': 1}
```

---

## UnhandledRevision

One revision element this library does not accept or reject, as returned by
[`list_unhandled_revisions()`](#list_unhandled_revisionsauthornone).

```python
from docx_editor import UnhandledRevision
```

### Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `tag` | str | The element's tag — one of `UNHANDLED_REVISION_TAGS`, or a `HANDLED_REVISION_TAGS` tag whose element carries no numeric `w:id` (nothing id-keyed can resolve it, so it is reported here rather than omitted) — e.g. `"w:rPrChange"` |
| `id` | int or None | The element's `w:id`, or None when it carries none or a non-numeric one. Unlike `Revision`, an id-less mark is still listed — nothing here is targeted by id. |
| `author` | str | `w:author`, or `"Unknown"` when the attribute is absent — matching `Revision`. `w:tblGridChange` and the range `*End` marks carry only `w:id` in the schema, so they always read as `"Unknown"`. |
| `date` | datetime or None | Parsed `w:date`, or None when absent or unparseable |
| `paragraph_ref` | str or None | Hash-anchored ref of the containing `<w:p>`, or None when the mark sits in no addressable paragraph — outside any paragraph (e.g. a `w:tblPrChange` in a table's properties, or a `w:sectPrChange` in a section break), or inside a drawing's text box (where a mark is listed once per stored copy, exactly as `Revision` is) |

### Example

```python
for row in doc.list_unhandled_revisions():
    print(row)
# UnhandledRevision(rPrChange 2 @P1#6c81 by Bob)
# UnhandledRevision(sectPrChange 3 by Ann)
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
| `current_ref` | str or None | Ref for the paragraph's **current** content — set only when the row failed on a stale hash |

`current_ref` is the recovery field for the one failure a caller can fix
mechanically. When an operation carries `"P7#a7b2"` but the paragraph now
hashes to `c4d8`, the row reports `current_ref="P7#c4d8"` — rebuild the
operation with it instead of parsing the hash out of `error`. It is `None` for
every other outcome: valid rows, malformed refs, out-of-range indexes, missing
or ambiguous target text, and elements that are not an `EditOperation`. So
`if row.current_ref:` is how you spot the stale-hash rows.

It clears the **hash** check only. Validation stops at the first failure, so a
row can report a `current_ref` *and* have a target that has since moved or
become ambiguous — the rebuilt operation is re-validated like any other, and a
target that no longer fits fails loudly and atomically (a
[`BatchOperationError`](#batchoperationerror) wrapping `TextNotFoundError` or
`AmbiguousTextError`, with nothing applied). Dry-run the repaired batch, or
re-search the paragraph, when the content may have changed materially rather
than just shifted.

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

Repair the stale-hash rows and retry, with no message parsing:

```python
for row in doc.batch_edit(ops, dry_run=True):
    if row.current_ref:
        ops[row.index] = EditOperation.replace("old", "new", paragraph=row.current_ref)
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
document-wide offsets; a tab mark counts as one character (`\t`), exactly as
in `get_visible_text()`. Coordinate systems differ between search and edit:
`find_text`'s `occurrence` counts matches document-wide (unless scoped with
`paragraph=`), while edit methods count within one paragraph —
`paragraph_occurrence` bridges the two, so always pass it alongside
`paragraph_ref` when chaining into an edit — or skip the bookkeeping entirely
and pass the `SearchResult` itself as the edit's target (see below).
`paragraph_ref` is computed at search time and — like refs from
`list_paragraphs()` — goes stale once that paragraph is edited.

**A SearchResult can stand in for its own three fields.** Every
text-targeting method — `replace`, `delete`, `insert_after`, `insert_before`,
`add_comment`, and the `EditOperation` constructors — accepts a `SearchResult`
where it expects the target text, and takes `paragraph` and `occurrence` from
the match. Passing `paragraph=`/`occurrence=` *as well* raises `ValueError`
(it would contradict the match), and a stale match raises
[`HashMismatchError`](#hashmismatcherror) like any other stale ref.

`find_text()` returns `None` when there is no match, and that `None` is **not**
accepted as a target (a missing match must not become a silent no-op) — it
raises `ValueError` (`CommentError` from `add_comment`, which validates its
anchor before the ref). Check the result first, as the examples below do, rather
than piping `find_text()` straight into an edit in code that must not crash.

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
    doc.replace(match, "60 days")          # the match pins paragraph + occurrence

    # The SAME single edit, spelled out — not a follow-up. Running both raises
    # HashMismatchError, because the first one made `match` stale:
    #   doc.replace(match.text, "60 days",
    #               paragraph=match.paragraph_ref,
    #               occurrence=match.paragraph_occurrence)

# Sweep every match in one atomic batch (reversed() keeps same-paragraph ops in
# the required descending-occurrence order):
from docx_editor import EditOperation

doc.batch_edit([EditOperation.replace(m, "60 days") for m in reversed(doc.find_all("30 days"))])
```

---

## ParagraphInfo

The structured paragraph record returned by
`Document.list_paragraphs_structured()`, `Document.get_paragraph()` and
`Document.context()`. All three emit identical records for the same paragraph.

```python
from docx_editor import ParagraphInfo
```

### Attributes

| Attribute | Type | Description |
|-----------|------|-------------|
| `index` | int | 1-based paragraph index (`P1` is `index=1`) |
| `ref` | str | Hash-anchored ref `P{index}#{hash}`, usable as any `paragraph=` argument |
| `text` | str | Full visible text, never truncated |
| `in_table` | bool | True when the paragraph sits inside a `<w:tc>` table cell |
| `style` | str or None | Raw `w:pPr/w:pStyle/@w:val` style id (e.g. `"Heading1"`), None when unstyled |
| `outline_level` | int or None | 0-based outline level (0 == Heading 1), None for body text |

`in_table`, `style` and `outline_level` carry exactly the meanings
[`ParagraphLocation`](#get_paragraph_locationref) defines for them — a direct
`w:outlineLvl` wins over the style's level, `w:val="9"` means body text, and a
`w:pPrChange` revision record is never read as current formatting. They are the
*cheap* structural facts: list numbering, table coordinates, heading paths and
section indexes still need `get_paragraph_location()` /
`list_paragraph_locations()`.

`in_table` also makes the paragraph-index divergence self-describing: table-cell
paragraphs do get `P{i}` refs, and they are exactly the ones python-docx's
`doc.paragraphs` skips.

`str(info)` renders `"P{i}#{hash}| {text}"` — the same delimiter format as
`list_paragraphs()`, always with the full text.

### Example

```python
# Table of contents from one call — no list_paragraph_locations() needed:
toc = [
    (info.outline_level, info.text)
    for info in doc.list_paragraphs_structured(limit=None)
    if info.outline_level is not None
]

# Skip table-cell paragraphs while editing body prose:
body = [info for info in doc.list_paragraphs_structured(limit=None) if not info.in_table]
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
`paragraph_ref`, `paragraph_preview` (a tab mark is spelled `\t` in every
preview, so it cannot pass for a space).

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

### `HashMismatchError`

Raised when a `paragraph` ref's hash no longer matches the paragraph's current
content — the paragraph changed (usually by an earlier edit) since the ref was
listed, so the ref is stale. Structured fields: `paragraph_index` (1-based),
`expected_hash` (the hash in the stale ref), `actual_hash` (the paragraph's
current hash), and `paragraph_preview` (the current text, a tab mark spelled `\t`). Recover by retrying
with the fresh ref `P{paragraph_index}#{actual_hash}`, or re-list paragraphs.

```python
from docx_editor.exceptions import HashMismatchError

try:
    doc.replace("old", "new", paragraph="P2#f3c1")
except HashMismatchError as e:
    doc.replace("old", "new", paragraph=f"P{e.paragraph_index}#{e.actual_hash}")
```

### `ParagraphIndexError`

Raised when a paragraph index is out of range — `< 1` or greater than
`paragraph_count()` (for example `get_paragraph(0)`). Structured fields:
`index` (the offending index) and `total_paragraphs` (how many the document
has). Clamp to `1..total_paragraphs` and retry, or call `list_paragraphs()`
to pick a valid ref.

```python
from docx_editor.exceptions import ParagraphIndexError

try:
    para = doc.get_paragraph(index)
except ParagraphIndexError as e:
    para = doc.get_paragraph(max(1, min(e.index, e.total_paragraphs)))
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

Raised when a revision operation fails — most commonly an unknown group id passed to `accept_group()` / `reject_group()`, or an unknown changeset id passed to `accept_changeset()` / `reject_changeset()`. Group and changeset ids are per-open-`Document` and renumbered on each open (recorded for this session's edits, inferred by reconstruction for revisions already in the file), so always use an id from the current session's `EditResult` or `list_revisions()` — a stale id from a previous session may raise this, or worse, silently resolve to a different group/changeset. Structured fields: `revision_id`, `group_id`, and `changeset_id` — set when the error is about that specific id (`group_id` for unknown-group errors, `changeset_id` for unknown-changeset errors), `None` otherwise.

```python
from docx_editor.exceptions import RevisionError
```

### `DocumentNotFoundError`

Raised by `Document.open()` when the source file does not exist. Structured field: `path` (the path that did not exist).

```python
from docx_editor.exceptions import DocumentNotFoundError
```

### `InvalidDocumentError`

Raised by `Document.open()` when the source path is not a valid `.docx` — wrong
suffix, a directory, an empty/truncated file, not a ZIP, missing
`word/document.xml`, or malformed XML. Structured field: `path` (the source
path). Not an in-loop retry: the message names which check failed; fix or
re-export the input file.

```python
from docx_editor.exceptions import InvalidDocumentError
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

`Document.open(path, force_recreate=True)` and
[`Document.discard_workspace(path)`](#documentdiscard_workspacepath-workspace_dirnone)
both recover but discard the workspace's unsaved edits. To rescue them first,
save the orphaned workspace to a new file (`Workspace` is not exported at the
package root — use the deep import):

```python
from docx_editor.exceptions import WorkspaceSyncError
from docx_editor.workspace import Workspace

try:
    doc = Document.open("contract.docx")
except WorkspaceSyncError:
    Workspace("contract.docx", create=False).save("rescued.docx")  # rescue unsaved edits
    doc = Document.open("contract.docx", force_recreate=True)
```

After a crashed script, one `Document.discard_workspace("contract.docx")` is
usually the better fix than carrying `force_recreate=True` on every subsequent
open — it also clears a lock left behind by the dead process.

### `DocumentOpenError`

Raised by `save()` when the destination appears open in another program. Word writes a `~$` owner (lock) file next to any document it has open; if that stub exists at save time, saving would race Word's own writes, so `save()` refuses unless `force=True`. Also raised when the OS denies the final replace (on Windows, Word holding the file open is exactly this case) — `force=True` cannot suppress that one. The exception carries `path` (the destination) and `owner_file` (the `~$` file that triggered the guard, or `None` when the OS denied the replace) attributes.

```python
from docx_editor.exceptions import DocumentOpenError

try:
    doc.save()
except DocumentOpenError as e:
    print(f"Close {e.path} in Word first (lock: {e.owner_file})")
```

### `DocumentProtectedError`

Raised by `Document.open()` when the document's editing protection is enforced and its mode locks the body text. Word's *Restrict Editing* writes `<w:documentProtection>` into `word/settings.xml`; `readOnly`, `forms` and `comments` all mean the author asked for the content not to be edited, so opening refuses rather than producing a file Word reopens with the same lock over unexpected edits. A protection configured but switched off (`w:enforcement="0"`) never raises, and neither does the `trackedChanges` mode — it enforces exactly what this library already does. An enforcement value outside the schema's on/off spellings fails closed and raises. Only `w:documentProtection` is read: `w:writeProtection` (*Password to modify*, *Always Open Read-Only*) restricts saving rather than editing and is out of scope. The exception carries `path` (the document) and `mode` (the raw `w:edit` value) attributes.

Recovery: unprotect the document in Word (Review > Restrict Editing > Stop Protection), or open it anyway with `allow_protected=True`, which leaves the protection in the saved file.

```python
from docx_editor import Document
from docx_editor.exceptions import DocumentProtectedError

try:
    doc = Document.open("contract.docx")
except DocumentProtectedError as e:
    print(f"{e.path} is protected ({e.mode})")
    doc = Document.open(e.path, allow_protected=True)
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

## Warnings

### `UnhandledRevisionWarning`

Emitted by `accept_all()` / `reject_all()` when the document still holds
revision elements outside the types those verbs resolve (insertions,
deletions, moves, paragraph-property changes) — run/section/table property
changes, table-structure revisions, `w:numberingChange`, custom-XML range
marks. Without it,
`accept_all()` returning 0 on a run-format-only redline reads as "there was
nothing to accept" rather than "nothing here could be accepted".

Inspect what remains with
[`list_unhandled_revisions()`](#list_unhandled_revisionsauthornone), or read
the counts off the returned [`ResolveResult`](#resolveresult).

```python
import warnings
from docx_editor import UnhandledRevisionWarning

warnings.filterwarnings("ignore", category=UnhandledRevisionWarning)
```

### `UnanchoredNoteWarning`

Emitted by an edit method when `note=` was given but the edit created no
revision a comment could bracket — a no-op `replace`, an edit that amended one
of your own pending insertions, a rewrite that found no differences, or a bare
`"\n"` whose only revision is the paragraph mark. The edit itself still
applies; only the note is dropped, and `EditResult.comment_id` is `None`. The
message names which of those causes applied.

Without it a dropped rationale would be indistinguishable from a recorded one.
See [Rationale notes](#rationale-notes); use
[`add_comment()`](#add_commentanchor_text-comment-paragraphnone-occurrencenone)
to attach the text as an ordinary, document-scoped comment instead.

```python
import warnings
from docx_editor import UnanchoredNoteWarning

warnings.filterwarnings("ignore", category=UnanchoredNoteWarning)
```
