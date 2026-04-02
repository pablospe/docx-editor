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
    """Raised when the target text cannot be found in the document."""

    pass


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
