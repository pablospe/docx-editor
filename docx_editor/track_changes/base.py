"""Typed base for the ``RevisionManager`` mixins.

``RevisionManager`` is assembled from mixins that each own one cluster of
methods but share the instance state set up in ``RevisionManager.__init__``.
This class declares that state (annotations only, no values) and stubs the
few cross-cluster callees, so ty resolves ``self.<name>`` inside a mixin
without a ``[[tool.ty.overrides]]`` suppression. Nothing here runs: the
annotations create no class attributes and every stub is overridden by the
real method on ``RevisionManager``.

``scripts/check_pure_move.py`` (check 3) requires every non-blank line in
this file to be a verbatim copy of a line elsewhere in the package, an
attribute annotation, or ``raise NotImplementedError`` -- so keep the
explanations in this docstring and the class body bare.
"""

from xml.dom.minidom import Element

from ..xml_editor import DocxXMLEditor
from .models import GroupSource


class _RevisionManagerBase:
    editor: DocxXMLEditor
    _groups: dict[int, tuple[int, ...]]
    _revision_groups: dict[int, int | None]
    _group_sources: dict[int, GroupSource]
    _group_counter: int
    _changesets: dict[int, tuple[int, ...]]
    _group_changesets: dict[int, int]
    _changeset_sources: dict[int, GroupSource]
    _changeset_counter: int
    _in_changeset: bool
    _paragraph_mark_ids: set[int]
    _defer_range_sweep: bool
    _range_sweep_pending: bool

    def _revision_element_index(self) -> dict[str, list[Element]]:
        raise NotImplementedError

    def _is_in_document(self, elem) -> bool:
        raise NotImplementedError
