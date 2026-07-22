"""Custom exceptions for docx_editor library."""

from pathlib import Path


class DocxEditError(Exception):
    """Base exception for all docx_editor errors."""

    pass


class DocumentNotFoundError(DocxEditError):
    """Raised when the document file cannot be found.

    Attributes:
        path: The path that did not exist, or None if unknown.
    """

    def __init__(self, message: str, *, path: Path | None = None):
        self.path = path
        super().__init__(message)


class InvalidDocumentError(DocxEditError):
    """Raised when the document is not a valid .docx file.

    Attributes:
        path: The path that failed validation, or None if unknown.
    """

    def __init__(self, message: str, *, path: Path | None = None):
        self.path = path
        super().__init__(message)


class WorkspaceError(DocxEditError):
    """Raised when there's an error with workspace operations."""

    pass


class WorkspaceExistsError(WorkspaceError):
    """Raised when trying to open a document with an existing workspace."""

    pass


class WorkspaceSyncError(WorkspaceError):
    """Raised when the workspace is out of sync with the source document.

    Two triggers: the source changed on disk since the workspace was created,
    or the workspace holds unsaved changes from a previous session that the
    source never received.

    Attributes:
        workspace_path: The workspace that is out of sync, or None if unknown.
        source_path: The source document it belongs to, or None if unknown.
    """

    def __init__(
        self,
        message: str,
        *,
        workspace_path: Path | None = None,
        source_path: Path | None = None,
    ):
        self.workspace_path = workspace_path
        self.source_path = source_path
        super().__init__(message)


class WorkspaceLockedError(WorkspaceError):
    """Raised when a live session already holds a document's workspace lock.

    Two sessions sharing one workspace silently overwrite each other's edits
    (last save wins), so opening a document whose advisory lock file names a
    still-running process is refused — whether that process is another one or
    the caller's own (two Document objects on one workspace clobber each other
    either way). A lock left behind by a dead process is reclaimed silently
    and never raises. ``force_recreate=True`` discards the workspace and its
    lock regardless.

    Attributes:
        pid: PID recorded in the lock file, or None if it could not be read.
        lock_path: The lock file that blocked the open.
    """

    def __init__(
        self,
        message: str,
        *,
        pid: int | None = None,
        lock_path: Path | None = None,
    ):
        self.pid = pid
        self.lock_path = lock_path
        super().__init__(message)


class SessionError(DocxEditError):
    """Raised when there's an error with a persistent session kernel."""

    pass


class SessionDeadError(SessionError):
    """Raised when session files exist but the kernel process is gone or unreachable.

    The kernel crashed or was killed: its in-memory state (open documents,
    variables) is lost. Recover with 'docx-session stop' to clean up the stale
    files, then 'docx-session start'.
    """

    pass


class DocumentOpenError(DocxEditError):
    """Raised when saving to a destination that appears open in another program.

    Word writes a ``~$`` owner (lock) file next to any document it has open. If
    that stub exists at save time, saving would race Word's own writes and risk
    corruption, so ``save()`` refuses unless ``force=True``. Also raised when the
    OS denies the final replace with a ``PermissionError`` — on Windows, Word
    holding the file open is exactly this case.

    Attributes:
        path: The destination that could not be written.
        owner_file: The ``~$`` owner file that triggered the guard, or None when
            the error came from the OS denying the replace instead.
    """

    def __init__(
        self,
        message: str,
        *,
        path: Path | None = None,
        owner_file: Path | None = None,
    ):
        self.path = path
        self.owner_file = owner_file
        super().__init__(message)


class DocumentClosedError(DocxEditError):
    """Raised when an operation is attempted on a closed Document.

    ``close()`` discards the workspace (unless ``cleanup=False``), so the
    object cannot keep serving reads or edits. Reopen the source with
    ``Document.open(e.path)`` to continue.

    Attributes:
        path: Source path of the closed document, or None if unknown.
    """

    def __init__(self, message: str, *, path: Path | None = None):
        self.path = path
        super().__init__(message)


class XMLError(DocxEditError):
    """Raised when there's an error parsing or manipulating XML."""

    pass


class NodeNotFoundError(XMLError):
    """Raised when a requested XML node cannot be found."""

    pass


class MultipleNodesFoundError(XMLError):
    """Raised when multiple nodes match when only one was expected."""

    pass


class RevisionError(DocxEditError):
    """Raised when there's an error with revision operations.

    Attributes:
        revision_id: The revision id the operation targeted, or None if the
            error is not about a specific revision.
        group_id: The revision group id the operation targeted (e.g. an
            unknown group passed to ``accept_group``/``reject_group``), or
            None if the error is not about a group.
    """

    def __init__(
        self,
        message: str,
        *,
        revision_id: int | None = None,
        group_id: int | None = None,
    ):
        self.revision_id = revision_id
        self.group_id = group_id
        super().__init__(message)


class CommentError(DocxEditError):
    """Raised when there's an error with comment operations.

    Attributes:
        comment_id: The comment id the operation targeted (e.g. the parent id
            of a failed reply), or None if no comment id applies.
    """

    def __init__(self, message: str, *, comment_id: int | None = None):
        self.comment_id = comment_id
        super().__init__(message)


def _truncate_preview(text: str, limit: int = 80) -> str:
    """Cap a paragraph preview at ``limit`` characters, adding '...' if truncated."""
    if len(text) <= limit:
        return text
    return text[:limit] + "..."


def _append_preview(msg: str, preview: str) -> str:
    """Append the 'Current content' suffix with a period-safe separator."""
    sep = " " if msg.endswith(".") else ". "
    return f'{msg}{sep}Current content: "{preview}"'


class TextNotFoundError(DocxEditError):
    """Raised when the target text cannot be found in the document.

    Also raised when the text exists but an explicit ``occurrence`` is out of
    range — the message then reports the actual count instead of claiming the
    text is absent, and ``occurrence``/``total_occurrences`` are set.

    Attributes:
        search_text: The text that was being searched for.
        paragraph_ref: Paragraph reference the search was scoped to, or None
            if the search was document-wide.
        paragraph_preview: Current visible text of the scoped paragraph
            (truncated to 80 chars with "..." suffix), or None if unscoped.
        occurrence: 0-based occurrence index requested (for nth-match
            lookups), or None otherwise.
        total_occurrences: Actual number of occurrences found, or None for
            non-occurrence lookups.
    """

    def __init__(
        self,
        search_text: str,
        *,
        paragraph_ref: str | None = None,
        paragraph_preview: str | None = None,
        occurrence: int | None = None,
        total_occurrences: int | None = None,
    ):
        self.search_text = search_text
        self.paragraph_ref = paragraph_ref
        self.paragraph_preview = _truncate_preview(paragraph_preview) if paragraph_preview is not None else None
        self.occurrence = occurrence
        self.total_occurrences = total_occurrences

        if occurrence is not None and total_occurrences is not None:
            msg = (
                f"Only {total_occurrences} occurrence(s) of '{search_text}' found, "
                f"but occurrence={occurrence} requested."
            )
            if paragraph_ref is not None:
                msg += f" Paragraph: {paragraph_ref}."
        elif paragraph_ref is not None:
            msg = f"Text not found in paragraph {paragraph_ref}: '{search_text}'"
        else:
            msg = f"Text not found: '{search_text}'"

        if self.paragraph_preview is not None:
            msg = _append_preview(msg, self.paragraph_preview)

        super().__init__(msg)


class AmbiguousTextError(DocxEditError):
    """Raised when an edit target matches more than once and no occurrence was chosen.

    Edit methods called without an explicit ``occurrence`` require the target
    text to be unique within the search scope; otherwise the intended match is
    ambiguous and editing the first one silently risks changing the wrong text.
    Pass ``occurrence=`` (0-based) to pick a match, or enumerate every match
    with ``find_all()``.

    Attributes:
        search_text: The text that matched multiple times.
        paragraph_ref: Paragraph reference the search was scoped to, or None
            if the search was document-wide.
        paragraph_preview: Current visible text of the scoped paragraph
            (truncated to 80 chars with "..." suffix), or None if unscoped.
        total_occurrences: Number of matches found in the search scope.
    """

    def __init__(
        self,
        search_text: str,
        *,
        paragraph_ref: str | None = None,
        paragraph_preview: str | None = None,
        total_occurrences: int,
    ):
        self.search_text = search_text
        self.paragraph_ref = paragraph_ref
        self.paragraph_preview = _truncate_preview(paragraph_preview) if paragraph_preview is not None else None
        self.total_occurrences = total_occurrences

        scope = f"paragraph {paragraph_ref}" if paragraph_ref is not None else "the document"
        msg = (
            f"'{search_text}' matches {total_occurrences} times in {scope}. "
            f"Pass occurrence= (0-based) to pick one, or use find_all() to list every match."
        )
        if self.paragraph_preview is not None:
            msg = _append_preview(msg, self.paragraph_preview)
        super().__init__(msg)


class ParagraphIndexError(DocxEditError):
    """Raised when a paragraph reference's 1-indexed number is out of range.

    Attributes:
        index: The paragraph index supplied by the caller (1-indexed).
        total_paragraphs: Number of paragraphs the document actually has.
    """

    def __init__(self, index: int, total_paragraphs: int):
        self.index = index
        self.total_paragraphs = total_paragraphs
        if total_paragraphs == 0:
            msg = f"Paragraph index {index} out of range. Document has no paragraphs."
        else:
            msg = (
                f"Paragraph index {index} out of range. "
                f"Document has {total_paragraphs} paragraphs "
                f"(1-indexed, valid: P1-P{total_paragraphs})."
            )
        super().__init__(msg)


class BatchOperationError(DocxEditError):
    """Raised when a single operation in a batch fails (validation or apply).

    Attributes:
        operation_index: 0-indexed position of the failing operation in the
            input list.
        reason: Human-readable description of why the operation was rejected.
        original: The underlying exception that caused the failure (also set
            as ``__cause__``), or None. Typed recovery data stays reachable
            through it, e.g. ``e.original.actual_hash`` for a stale hash.
    """

    def __init__(self, operation_index: int, reason: str, *, original: Exception | None = None):
        self.operation_index = operation_index
        self.reason = reason
        self.original = original
        super().__init__(f"Operation {operation_index}: {reason}")


class HashMismatchError(DocxEditError):
    """Raised when a paragraph's content hash doesn't match the expected hash."""

    def __init__(self, paragraph_index: int, expected_hash: str, actual_hash: str, paragraph_preview: str):
        self.paragraph_index = paragraph_index
        self.expected_hash = expected_hash
        self.actual_hash = actual_hash
        self.paragraph_preview = paragraph_preview
        super().__init__(
            f"Paragraph P{paragraph_index} content has changed. "
            f"Expected hash '{expected_hash}', got '{actual_hash}'. "
            f'Current content: "{paragraph_preview}". '
            f"Use P{paragraph_index}#{actual_hash} to target current content."
        )
