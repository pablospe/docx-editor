"""docx_editor - Pure Python Track Changes Library for Word Documents.

A standalone library for Word document track changes and comments,
without requiring Microsoft Word installed.

Example:
    from docx_editor import Document

    # Open and edit with hash-anchored paragraph references
    doc = Document.open("contract.docx")
    refs = doc.list_paragraphs()                               # Snapshot paragraphs
    doc.replace("30 days", "60 days", paragraph="P2#f3c1")     # Tracked replacement
    doc.insert_after("Section 5", "New clause", paragraph="P3#a7b2")  # Tracked insertion
    doc.delete("obsolete text", paragraph="P5#c4d8")           # Tracked deletion

    # Comments
    doc.add_comment("Section 5", "Please review")
    doc.reply_to_comment(comment_id=0, "Approved")

    # Revision management
    revisions = doc.list_revisions()
    doc.accept_revision(revision_id=1)
    doc.reject_all(author="OtherUser")

    # Save and close
    doc.save()
    doc.close()
"""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("docx-editor")
except PackageNotFoundError:  # pragma: no cover - source tree without installation metadata
    __version__ = "0.2.2"

from .comments import Comment
from .document import Document
from .exceptions import (
    BatchOperationError,
    CommentError,
    DocumentNotFoundError,
    DocumentOpenError,
    DocxEditError,
    HashMismatchError,
    InvalidDocumentError,
    MultipleNodesFoundError,
    NodeNotFoundError,
    ParagraphIndexError,
    RevisionError,
    TextNotFoundError,
    WorkspaceError,
    WorkspaceExistsError,
    WorkspaceLockedError,
    WorkspaceSyncError,
    XMLError,
)
from .track_changes import EditOperation, EditValidationResult, Revision, SearchResult
from .xml_editor import (
    ListItem,
    ParagraphInfo,
    ParagraphLocation,
    ParagraphRef,
    TableCell,
    compute_paragraph_hash,
)

__all__ = [
    # Main classes
    "Document",
    "EditOperation",
    "EditValidationResult",
    "Revision",
    "SearchResult",
    "Comment",
    # Exceptions
    "DocxEditError",
    "DocumentNotFoundError",
    "DocumentOpenError",
    "InvalidDocumentError",
    "WorkspaceError",
    "WorkspaceExistsError",
    "WorkspaceLockedError",
    "WorkspaceSyncError",
    "XMLError",
    "NodeNotFoundError",
    "MultipleNodesFoundError",
    "RevisionError",
    "CommentError",
    "TextNotFoundError",
    "HashMismatchError",
    "ParagraphIndexError",
    "BatchOperationError",
    # Paragraph refs
    "ParagraphInfo",
    "ParagraphRef",
    "compute_paragraph_hash",
    # Paragraph location
    "ListItem",
    "ParagraphLocation",
    "TableCell",
]

# Internal text-map machinery removed from the public API (see SearchResult /
# Document.find_text instead). Accessing these via the top-level package warns
# for one release before removal; the real objects remain importable from
# docx_editor.xml_editor.
_DEPRECATED_INTERNALS = frozenset({"TextMap", "TextMapMatch", "TextPosition", "build_text_map", "find_in_text_map"})


def __getattr__(name: str):
    if name in _DEPRECATED_INTERNALS:
        import warnings

        from . import xml_editor

        warnings.warn(
            f"docx_editor.{name} is internal and will be removed from the public API "
            f"in the next release; use Document.find_text()/SearchResult instead, or "
            f"import from docx_editor.xml_editor if you need the internals.",
            DeprecationWarning,
            stacklevel=2,
        )
        return getattr(xml_editor, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
