"""Custom exceptions for docx_editor library."""


class DocxEditError(Exception):
    """Base exception for all docx_editor errors."""

    pass


class DocumentNotFoundError(DocxEditError):
    """Raised when the document file cannot be found."""

    pass


class InvalidDocumentError(DocxEditError):
    """Raised when the document is not a valid .docx file."""

    pass


class WorkspaceError(DocxEditError):
    """Raised when there's an error with workspace operations."""

    pass


class WorkspaceExistsError(WorkspaceError):
    """Raised when trying to open a document with an existing workspace."""

    pass


class WorkspaceSyncError(WorkspaceError):
    """Raised when the source document has changed since workspace creation."""

    pass


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
    """Raised when there's an error with revision operations."""

    pass


class CommentError(DocxEditError):
    """Raised when there's an error with comment operations."""

    pass


class TextNotFoundError(DocxEditError):
    """Raised when the target text cannot be found in the document.

    Attributes:
        search_text: The text that was being searched for.
        paragraph_ref: Paragraph reference the search was scoped to, or None
            if the search was document-wide.
        paragraph_preview: Current visible text of the scoped paragraph
            (truncated to 80 chars with "..." suffix), or None if unscoped.
        occurrence: 1-indexed occurrence number requested (for nth-match
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
        elif paragraph_ref is not None:
            msg = f"Text not found in paragraph {paragraph_ref}: '{search_text}'"
            if self.paragraph_preview is not None:
                msg += f'. Current content: "{self.paragraph_preview}"'
        else:
            msg = f"Text not found: '{search_text}'"

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
        super().__init__(
            f"Paragraph index {index} out of range. "
            f"Document has {total_paragraphs} paragraphs "
            f"(1-indexed, valid: P1-P{total_paragraphs})."
        )


class BatchOperationError(DocxEditError):
    """Raised when a single operation in a batch fails validation.

    Attributes:
        operation_index: 0-indexed position of the failing operation in the
            input list.
        reason: Human-readable description of why the operation was rejected.
    """

    def __init__(self, operation_index: int, reason: str):
        self.operation_index = operation_index
        self.reason = reason
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


def _truncate_preview(text: str, limit: int = 80) -> str:
    """Cap a paragraph preview at ``limit`` characters, adding '...' if truncated."""
    if len(text) <= limit:
        return text
    return text[:limit] + "..."
