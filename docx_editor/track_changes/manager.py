"""``RevisionManager``: creates and resolves tracked changes."""

from ..xml_editor import DocxXMLEditor
from .batch import _BatchMixin
from .delete import _DeleteMixin
from .insert import _InsertMixin
from .listing import _ListingMixin
from .locate import _LocateMixin
from .models import GroupSource
from .registry import _RegistryMixin
from .replace import _ReplaceMixin
from .resolution import _ResolutionMixin


class RevisionManager(
    _RegistryMixin,
    _LocateMixin,
    _BatchMixin,
    _ReplaceMixin,
    _DeleteMixin,
    _InsertMixin,
    _ListingMixin,
    _ResolutionMixin,
):
    """Manages track changes in a Word document.

    Provides methods for creating tracked insertions, deletions, replacements,
    and for accepting/rejecting revisions.
    """

    def __init__(self, editor: DocxXMLEditor):
        """Initialize with a DocxXMLEditor for the document.xml file.

        Args:
            editor: DocxXMLEditor instance for word/document.xml
        """
        self.editor = editor
        # Revision groups: every revision created by one logical operation
        # shares a group id, so callers can accept/reject the operation as a
        # unit. The registry is in-memory and rebuilt on every open — nothing
        # is written into the .docx. Revisions already present in the file
        # get inferred groups from _reconstruct_groups below; edits made
        # through this manager record theirs and continue the numbering.
        self._groups: dict[int, tuple[int, ...]] = {}
        # Maps revision id -> group id. A None value means explicitly
        # ungrouped: a split-off tail of an ungroupable own insertion,
        # registered so the active _grouped capture cannot claim it
        # (membership is key-based) while group_id_of/list_revisions still
        # report None.
        self._revision_groups: dict[int, int | None] = {}
        # Maps group id -> provenance ("recorded" | "inferred").
        self._group_sources: dict[int, GroupSource] = {}
        self._group_counter = 1
        # Changeset tier (one whole call ⊇ ≥1 group). A changeset is the
        # (author, date) equivalence class over groups; these registries
        # mirror the group ones exactly, one level up. Recorded changesets
        # continue this counter past the inferred ones, just as groups do.
        self._changesets: dict[int, tuple[int, ...]] = {}
        self._group_changesets: dict[int, int] = {}
        self._changeset_sources: dict[int, GroupSource] = {}
        self._changeset_counter = 1
        # Reentrancy flag for _changeset(): only the outermost boundary
        # bundles, so batch_rewrite -> rewrite_paragraph merges into one
        # changeset (mirrors editor.frozen_timestamp's reuse guard).
        self._in_changeset = False
        # Ids of paragraph-mark insertions (tracked splits), recorded and
        # inferred alike. Lets split_count() answer without a DOM walk, so
        # result-ref building stays cheap for the common no-split edit.
        self._paragraph_mark_ids: set[int] = set()
        # Bulk resolution (``_resolve_all``/``_resolve_ids``) defers the move
        # range-mark sweep to one walk at the end instead of one per move half.
        self._defer_range_sweep = False
        self._range_sweep_pending = False
        self._reconstruct_groups()
