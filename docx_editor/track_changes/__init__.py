"""Track changes management for docx_editor.

Provides RevisionManager for creating and managing tracked changes
(insertions, deletions, moves, paragraph-property changes).
"""

from .manager import RevisionManager
from .models import (
    ALL_REVISION_TAGS,
    CHANGE_RECORD_TAGS,
    HANDLED_REVISION_TAGS,
    MOVE_RANGE_TAGS,
    UNHANDLED_REVISION_TAGS,
    EditOperation,
    EditResult,
    EditValidationResult,
    GroupSource,
    ResolveResult,
    Revision,
    RevisionCensus,
    RevisionType,
    SearchResult,
    UnhandledRevision,
    count_revision_elements,
    iter_revision_elements,
)

__all__ = [
    "ALL_REVISION_TAGS",
    "CHANGE_RECORD_TAGS",
    "HANDLED_REVISION_TAGS",
    "MOVE_RANGE_TAGS",
    "UNHANDLED_REVISION_TAGS",
    "EditOperation",
    "EditResult",
    "EditValidationResult",
    "GroupSource",
    "ResolveResult",
    "Revision",
    "RevisionCensus",
    "RevisionManager",
    "RevisionType",
    "SearchResult",
    "UnhandledRevision",
    "count_revision_elements",
    "iter_revision_elements",
]
