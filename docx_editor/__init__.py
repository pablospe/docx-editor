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

__version__ = "0.0.1"

from .comments import Comment
from .document import Document
from .exceptions import (
    CommentError,
    DocumentNotFoundError,
    DocxEditError,
    HashMismatchError,
    InvalidDocumentError,
    MultipleNodesFoundError,
    NodeNotFoundError,
    RevisionError,
    TextNotFoundError,
    WorkspaceError,
    WorkspaceExistsError,
    WorkspaceSyncError,
    XMLError,
)
from .track_changes import EditOperation, Revision
from .xml_editor import (
    ParagraphRef,
    TextMap,
    TextMapMatch,
    TextPosition,
    build_text_map,
    compute_paragraph_hash,
    find_in_text_map,
)

__all__ = [
    # Main classes
    "Document",
    "EditOperation",
    "Revision",
    "Comment",
    # Exceptions
    "DocxEditError",
    "DocumentNotFoundError",
    "InvalidDocumentError",
    "WorkspaceError",
    "WorkspaceExistsError",
    "WorkspaceSyncError",
    "XMLError",
    "NodeNotFoundError",
    "MultipleNodesFoundError",
    "RevisionError",
    "CommentError",
    "TextNotFoundError",
    "HashMismatchError",
    # Text map & paragraph refs
    "TextPosition",
    "TextMap",
    "TextMapMatch",
    "ParagraphRef",
    "build_text_map",
    "compute_paragraph_hash",
    "find_in_text_map",
]
