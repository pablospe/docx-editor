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

from collections.abc import Iterator
from contextlib import contextmanager
from typing import Literal
from xml.dom.minidom import Element

from ..xml_editor import DocxXMLEditor, ParagraphRef, TextMap, TextMapMatch, TextPosition
from .models import GroupSource, _GroupCapture, _RegistrySnapshot


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

    @contextmanager
    def _grouped(self) -> Iterator[_GroupCapture]:
        raise NotImplementedError

    @contextmanager
    def _changeset(self) -> Iterator[None]:
        raise NotImplementedError

    def _registry_snapshot(self) -> _RegistrySnapshot:
        raise NotImplementedError

    def _restore_registry(self, snapshot: _RegistrySnapshot) -> None:
        raise NotImplementedError

    def _resolve_paragraph(self, ref: ParagraphRef, paragraphs: list[Element] | None = None):
        raise NotImplementedError

    def _locate_in_paragraph(self, paragraph, paragraph_ref: str, text: str, occurrence: int | None) -> TextMapMatch:
        raise NotImplementedError

    def _get_run_info(self, node) -> tuple[Element | None, str]:
        raise NotImplementedError

    def _get_node_text(self, node) -> str:
        raise NotImplementedError

    def _set_node_text(self, node, text: str) -> None:
        raise NotImplementedError

    def _replace_across_nodes(self, match: TextMapMatch, replace_with: str) -> int:
        raise NotImplementedError

    def _find_ancestor(self, node, tag_name: str) -> Element | None:
        raise NotImplementedError

    def _owns_ins(self, ins_elem) -> bool:
        raise NotImplementedError

    def _insert_own_ins_within_foreign_ins(self, ins_elem, edge_node, offset: int, text: str, rPr_xml: str) -> int:
        raise NotImplementedError

    def _delete_across_nodes(self, match: TextMapMatch) -> int:
        raise NotImplementedError

    def _insert_near_match(self, match: TextMapMatch, text: str, position: Literal["before", "after"]) -> int:
        raise NotImplementedError

    @staticmethod
    def _plain_run_xml(rPr_xml: str, text: str) -> str:
        raise NotImplementedError

    def _insert_into_run(self, run, rPr_xml: str, node, offset: int, fragment: str) -> list:
        raise NotImplementedError

    def _ensure_splittable(self, p1) -> None:
        raise NotImplementedError

    def _reject_unsplittable_boundary(self, paragraph, text_map: TextMap, pos: int) -> None:
        raise NotImplementedError

    def _apply_paragraph_splits(self, p1, split_pos: int, segments: list[str]) -> int:
        raise NotImplementedError

    def _locate_document_wide(self, text: str, occurrence: int | None = None) -> TextMapMatch:
        raise NotImplementedError

    def _group_positions_by_ins(self, positions: list) -> list[tuple[Element | None, list[TextPosition]]]:
        raise NotImplementedError

    def _delete_from_ins_positions(self, positions: list) -> tuple[int, Element | None]:
        raise NotImplementedError

    def _split_ins_after_child(self, ins_elem, child) -> None:
        raise NotImplementedError

    def _delete_regular_segment(self, positions: list) -> tuple[int, Element | None]:
        raise NotImplementedError

    def _split_replace(self, match: TextMapMatch, replace_with: str) -> int:
        raise NotImplementedError

    def _build_cross_boundary_parts(self, match: TextMapMatch) -> list[tuple[Element, str, str, str, str, int]]:
        raise NotImplementedError

    def _classify_segments(self, match: TextMapMatch) -> list[tuple[bool | None, list[TextPosition]]]:
        raise NotImplementedError

    def _ins_identity_attrs(self, ins_elem) -> str:
        raise NotImplementedError

    def _adopt_split_tail(self, original_ins, new_nodes) -> None:
        raise NotImplementedError
