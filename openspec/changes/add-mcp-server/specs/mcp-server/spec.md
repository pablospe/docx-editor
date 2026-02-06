## ADDED Requirements

### Requirement: Document Cache

The MCP server SHALL maintain an in-memory cache of open documents to avoid reloading between operations.

The cache SHALL:
- Store parsed `Document` instances keyed by normalized absolute file path
- Normalize all paths to canonical form (resolve `~`, relative paths, symlinks)
- Limit the number of cached documents (default: 10)
- Use LRU (Least Recently Used) eviction when the limit is reached
- Save dirty documents before evicting them
- Track the last access time for each cached document

#### Scenario: Cache hit on repeated access

- **GIVEN** a document "contract.docx" was previously opened via MCP
- **WHEN** another MCP tool accesses "contract.docx"
- **THEN** the cached `Document` instance is returned
- **AND** no file I/O occurs

#### Scenario: Cache miss triggers load

- **GIVEN** a document "new.docx" is not in the cache
- **WHEN** an MCP tool accesses "new.docx"
- **THEN** the document is loaded from disk
- **AND** the document is added to the cache

#### Scenario: LRU eviction when cache full

- **GIVEN** the cache contains 10 documents (max limit)
- **AND** document "old.docx" has the oldest last-access time
- **WHEN** a new document "new.docx" is accessed
- **THEN** "old.docx" is evicted from the cache
- **AND** if "old.docx" has unsaved changes, it is saved before eviction
- **AND** "new.docx" is loaded into the cache

#### Scenario: Path normalization prevents duplicate cache entries

- **GIVEN** a document exists at "/home/user/docs/contract.docx"
- **WHEN** `open_document(path="~/docs/contract.docx")` is called
- **AND** later `replace_text(path="/home/user/docs/contract.docx", ...)` is called
- **THEN** both operations use the same cached document instance
- **AND** only one cache entry exists

### Requirement: External Change Detection

The MCP server SHALL detect when a cached document has been modified externally and prevent accidental overwrites.

#### Scenario: Detect external modification

- **GIVEN** a document "contract.docx" is cached with mtime T1
- **AND** the file is modified externally (mtime becomes T2)
- **WHEN** an MCP edit tool is called on "contract.docx"
- **THEN** the operation fails with an error message
- **AND** the error indicates the file was modified externally
- **AND** suggests using `reload_document` or `force_save`

#### Scenario: mtime updated after save

- **GIVEN** a document "contract.docx" is cached
- **WHEN** `save_document` is called
- **THEN** the cached mtime is updated to match the file's new mtime

#### Scenario: Reload document after external change

- **GIVEN** a document "contract.docx" is cached with unsaved changes
- **AND** the file was modified externally
- **WHEN** `reload_document(path="contract.docx")` is called
- **THEN** the cached document is discarded
- **AND** the document is reloaded from disk
- **AND** the tool returns a warning that unsaved changes were discarded

#### Scenario: Force save overwrites external changes

- **GIVEN** a document "contract.docx" is cached
- **AND** the file was modified externally
- **WHEN** `force_save(path="contract.docx")` is called
- **THEN** the cached document is saved to disk (overwriting external changes)
- **AND** the cached mtime is updated

### Requirement: Session Author Memory

The MCP server SHALL remember the author name for the session and hint Claude to ask the user on first use.

The author resolution order SHALL be:
1. Explicit `author` parameter on tool call
2. Previously set session author
3. System username via `getpass.getuser()` (cross-platform)
4. "Reviewer" as ultimate fallback

#### Scenario: First document open hints Claude to ask

- **GIVEN** no author has been set for this session
- **WHEN** `open_document(path="doc.docx")` is called without author
- **THEN** the document is opened with system username as author
- **AND** the response includes a hint: "Author set to 'pablo' (system default). Use author parameter to change."
- **AND** the session author is remembered

#### Scenario: Session author remembered

- **GIVEN** `open_document(path="doc1.docx", author="Legal Team")` was called
- **WHEN** `open_document(path="doc2.docx")` is called without author
- **THEN** "Legal Team" is used as the author (from session memory)

### Requirement: Document Lifecycle Tools

The MCP server SHALL provide tools to manage document lifecycle.

Tools SHALL include:
- `open_document(path, author?)` - Open document, set author for track changes
- `save_document(path)` - Save document to disk
- `close_document(path)` - Remove document from cache
- `reload_document(path)` - Discard cached changes, reload from disk
- `force_save(path)` - Save even when external changes detected

#### Scenario: Open document with author

- **WHEN** `open_document(path="/path/to/doc.docx", author="Reviewer")` is called
- **THEN** the document is loaded (or retrieved from cache)
- **AND** the author is set for subsequent track changes operations
- **AND** the tool returns success with document info

#### Scenario: Open document with default author

- **WHEN** `open_document(path="/path/to/doc.docx")` is called without author
- **THEN** the author defaults to the session author or system username

#### Scenario: Save document

- **GIVEN** a document is open and has unsaved changes
- **WHEN** `save_document(path="/path/to/doc.docx")` is called
- **THEN** the document is saved to disk
- **AND** the dirty flag is cleared
- **AND** the cached mtime is updated

#### Scenario: Close document

- **GIVEN** a document is in the cache
- **WHEN** `close_document(path="/path/to/doc.docx")` is called
- **THEN** the document is removed from the cache
- **AND** if dirty, a warning is returned (but document is still closed)

### Requirement: Track Changes Tools

The MCP server SHALL expose the library's track changes functionality as MCP tools.

Tools SHALL include:
- `replace_text(path, old_text, new_text, occurrence?)` - Replace text with tracking
- `delete_text(path, text, occurrence?)` - Delete text with tracking
- `insert_after(path, anchor, text, occurrence?)` - Insert after anchor with tracking
- `insert_before(path, anchor, text, occurrence?)` - Insert before anchor with tracking

#### Scenario: Replace text via MCP

- **GIVEN** a document is open via MCP
- **WHEN** `replace_text(path="doc.docx", old_text="30 days", new_text="60 days")` is called
- **THEN** the replacement is tracked (deletion + insertion markup)
- **AND** the document is marked dirty
- **AND** the tool returns success with change ID

#### Scenario: Text not found

- **WHEN** `delete_text(path="doc.docx", text="nonexistent")` is called
- **AND** the text does not exist in the document
- **THEN** the tool returns an error indicating text not found

### Requirement: Comment Tools

The MCP server SHALL expose the library's comment functionality as MCP tools.

Tools SHALL include:
- `add_comment(path, anchor_text, comment_text)` - Add comment anchored to text
- `list_comments(path, author?)` - List all comments, optionally filtered by author
- `reply_to_comment(path, comment_id, reply_text)` - Reply to existing comment
- `resolve_comment(path, comment_id)` - Mark comment as resolved
- `delete_comment(path, comment_id)` - Delete a comment

#### Scenario: Add comment via MCP

- **GIVEN** a document is open via MCP
- **WHEN** `add_comment(path="doc.docx", anchor_text="ambiguous term", comment_text="Please clarify")` is called
- **THEN** a comment is added anchored to "ambiguous term"
- **AND** the document is marked dirty
- **AND** the tool returns the new comment ID

### Requirement: Revision Tools

The MCP server SHALL expose the library's revision management functionality as MCP tools.

Tools SHALL include:
- `list_revisions(path, author?)` - List all tracked revisions
- `accept_revision(path, revision_id)` - Accept a specific revision
- `reject_revision(path, revision_id)` - Reject a specific revision
- `accept_all(path, author?)` - Accept all revisions, optionally filtered by author
- `reject_all(path, author?)` - Reject all revisions, optionally filtered by author

#### Scenario: List revisions via MCP

- **GIVEN** a document has tracked changes from multiple authors
- **WHEN** `list_revisions(path="doc.docx")` is called
- **THEN** all revisions are returned with id, type, author, and text

#### Scenario: Accept revision via MCP

- **GIVEN** a document has a tracked insertion with id=5
- **WHEN** `accept_revision(path="doc.docx", revision_id=5)` is called
- **THEN** the revision is accepted (insertion becomes regular text)
- **AND** the document is marked dirty

### Requirement: Read Tools

The MCP server SHALL expose read-only document inspection tools.

Tools SHALL include:
- `find_text(path, text)` - Check if text exists in document
- `count_matches(path, text)` - Count occurrences of text
- `get_visible_text(path)` - Get full visible text of document

#### Scenario: Count matches via MCP

- **GIVEN** a document contains "30 days" three times
- **WHEN** `count_matches(path="doc.docx", text="30 days")` is called
- **THEN** the tool returns 3

### Requirement: Graceful Shutdown

The MCP server SHALL attempt to save dirty documents on shutdown.

#### Scenario: Shutdown with dirty documents

- **GIVEN** the cache contains dirty documents
- **WHEN** the MCP server receives a shutdown signal
- **THEN** all dirty documents are saved
- **AND** the server exits cleanly

#### Scenario: Shutdown save failure

- **GIVEN** a dirty document cannot be saved (e.g., disk full)
- **WHEN** the MCP server shuts down
- **THEN** the error is logged
- **AND** the server continues shutting down (best-effort)
