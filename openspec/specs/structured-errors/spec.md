# structured-errors Specification

## Purpose
TBD - created by archiving change add-structured-error-types. Update Purpose after archive.
## Requirements
### Requirement: LLM-Facing Errors Inherit From Shared Base

Every exception class raised by docx-editor during an edit operation SHALL inherit from `DocxEditError` so that LLM consumers can use a single catch-all handler.

#### Scenario: DocxEditError catches all edit-time errors

- **WHEN** an edit operation raises `TextNotFoundError`, `HashMismatchError`, `ParagraphIndexError`, or `BatchOperationError`
- **THEN** a handler using `except DocxEditError` matches the exception

### Requirement: TextNotFoundError Carries Structured Recovery Fields

The system SHALL raise `TextNotFoundError` with the following instance attributes so that callers can recover without re-reading the document:

- `search_text: str` — the text that was not found (always present)
- `paragraph_ref: str | None` — the paragraph reference the search was scoped to, or `None` if the search was unscoped
- `paragraph_preview: str | None` — the current visible text of the scoped paragraph, or `None` if unscoped. MUST be truncated to 80 characters (with `"..."` suffix) to match `HashMismatchError`'s truncation behavior.
- `occurrence: int | None` — the 1-indexed occurrence number the caller requested (e.g. via `occurrence=5`), or `None` if not an occurrence-based lookup
- `total_occurrences: int | None` — the actual number of occurrences found, or `None` if not an occurrence-based lookup

The error message SHALL include `search_text`. When `paragraph_ref` is present, the message SHALL also include it. When `paragraph_preview` is present, the message SHALL include the current paragraph content so that a caller reading only the message can still see what the paragraph actually contains.

All raise sites MUST use the new structured signature — including `docx_editor/comments.py` (line 110), not just `track_changes.py`.

#### Scenario: Scoped search failure carries paragraph ref and preview

- **WHEN** `doc.replace("missing_text", "x", paragraph="P1#<hash>")` is called and `"missing_text"` is not in paragraph 1
- **THEN** a `TextNotFoundError` is raised
- **AND** `err.search_text == "missing_text"`
- **AND** `err.paragraph_ref` contains `"P1#<hash>"`
- **AND** `err.paragraph_preview` equals the current visible text of paragraph 1
- **AND** `str(err)` contains the search text, the paragraph ref, and the current paragraph content

#### Scenario: Occurrence-based search failure carries occurrence counts

- **WHEN** `doc.replace("some_text", "x", occurrence=5)` is called but only 3 occurrences of `"some_text"` exist
- **THEN** a `TextNotFoundError` is raised
- **AND** `err.search_text == "some_text"`
- **AND** `err.occurrence == 5`
- **AND** `err.total_occurrences == 3`
- **AND** `str(err)` mentions the requested occurrence and the total found

#### Scenario: Unscoped search failure still carries search text

- **WHEN** `doc.replace("missing_text_anywhere", "x")` is called without a `paragraph` argument and the text is absent from the document
- **THEN** a `TextNotFoundError` is raised
- **AND** `err.search_text == "missing_text_anywhere"`
- **AND** `err.paragraph_ref is None`
- **AND** `err.paragraph_preview is None`
- **AND** `str(err)` contains the search text

### Requirement: ParagraphIndexError For Out-Of-Range Paragraph References

The system SHALL raise `ParagraphIndexError` (inheriting from `DocxEditError`) — not stdlib `IndexError` — when a caller supplies a paragraph reference whose 1-indexed paragraph number exceeds the document's paragraph count. The exception SHALL carry:

- `index: int` — the index the caller supplied (1-indexed, as written in the ref)
- `total_paragraphs: int` — the number of paragraphs the document currently has

The message SHALL state the invalid index and the valid range so that a caller reading only the message can correct the reference.

#### Scenario: Out-of-range paragraph index raises ParagraphIndexError

- **WHEN** `doc.replace("x", "y", paragraph="P999#0000")` is called on a document with fewer than 999 paragraphs
- **THEN** a `ParagraphIndexError` is raised (not stdlib `IndexError`)
- **AND** `err.index == 999`
- **AND** `err.total_paragraphs` equals the document's actual paragraph count
- **AND** `str(err)` names the invalid index and the valid range

#### Scenario: ParagraphIndexError is a DocxEditError

- **WHEN** the above out-of-range call raises
- **THEN** `isinstance(err, DocxEditError)` is `True`

### Requirement: BatchOperationError For Batch Validation Failures

The system SHALL raise `BatchOperationError` (inheriting from `DocxEditError`) — not stdlib `ValueError` — when `batch_edit` or `batch_rewrite` validation rejects an operation. The exception SHALL carry:

- `operation_index: int` — the 0-indexed position of the failing operation in the input list
- `reason: str` — a human-readable description of why the operation was rejected

The message SHALL identify the operation by index and include the reason.

Additionally, `batch_edit` SHALL wrap each `_apply_single_edit` call in a `try/except ValueError` that re-raises as `BatchOperationError(i, str(e))`. This covers the `ValueError` sites inside `_apply_single_edit` (e.g. `"replace requires 'find' and 'replace_with'"`) that fire exclusively from batch context but do not carry `operation_index` on their own.

Non-batch `ValueError` raise sites (e.g. malformed ref strings, invalid constructor arguments outside batch paths) SHALL continue to raise `ValueError`. This change is scoped to batch validation paths only.

#### Scenario: Invalid batch operation identifies its position

- **WHEN** `doc.batch_edit(ops)` is called with a validation-failing operation at position 1 (e.g. a replace op missing its `find` argument)
- **THEN** a `BatchOperationError` is raised (not stdlib `ValueError`)
- **AND** `err.operation_index == 1`
- **AND** `err.reason` describes the validation failure
- **AND** `str(err)` identifies operation 1

#### Scenario: Non-batch ValueError is unchanged

- **WHEN** a non-batch code path raises `ValueError` (e.g. a malformed paragraph ref syntax outside `_resolve_paragraph`)
- **THEN** the exception class is still `ValueError`
- **AND** it is NOT a `BatchOperationError`

### Requirement: Public Export Of Structured Error Classes

The `docx_editor` package SHALL re-export `HashMismatchError`, `TextNotFoundError`, `ParagraphIndexError`, and `BatchOperationError` from its top-level `__init__.py` so that LLM-facing consumers can `from docx_editor import ...` them without knowing the internal module layout.

#### Scenario: All structured error classes import from the top level

- **WHEN** a consumer runs `from docx_editor import HashMismatchError, TextNotFoundError, ParagraphIndexError, BatchOperationError`
- **THEN** the import succeeds
- **AND** each name refers to the canonical class (no shadow re-definition)

### Requirement: Documented Recovery Contract For LLM Consumers

The `skills/docx/SKILL.md` file SHALL document every structured error's fields and the recovery pattern a caller should follow, so that LLM consumers can self-correct from the error without external coaching.

The documentation SHALL, at minimum, list for each of `HashMismatchError`, `TextNotFoundError`, `ParagraphIndexError`, and `BatchOperationError`:

- The exception class name
- Its structured field names (instance attributes)
- The recovery action a caller should take (e.g. "use `P{index}#{actual_hash}` to retry")

#### Scenario: SKILL.md lists structured error fields

- **WHEN** a reader opens `skills/docx/SKILL.md`
- **THEN** the file contains a section documenting each of the four structured errors
- **AND** each error's entry names its instance attributes and a recovery pattern
