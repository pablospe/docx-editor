"""Track changes management for docx_editor.

Provides RevisionManager for creating and managing tracked changes (insertions/deletions).
"""

import difflib
import re
from collections import OrderedDict
from collections.abc import Callable, Iterable, Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from typing import Literal
from xml.dom.minidom import Element

from .exceptions import (
    AmbiguousTextError,
    BatchOperationError,
    DocxEditError,
    HashMismatchError,
    ParagraphIndexError,
    RevisionError,
    TextNotFoundError,
)
from .xml_editor import (
    DocxXMLEditor,
    ParagraphRef,
    TextMap,
    TextMapMatch,
    TextPosition,
    _escape_xml,
    _reject_control_chars,
    _require_valid_occurrence,
    build_text_map,
    compute_paragraph_hash,
    compute_text_hash,
    count_in_text_map,
    find_in_text_map,
    get_rPr_xml,
    get_text_node_data,
    rebuild_run_fragments,
    render_plain_wt,
)

# Provenance of a revision group: created by an edit through this open
# Document ("recorded") vs reconstructed at parse time ("inferred").
GroupSource = Literal["recorded", "inferred"]


@dataclass
class _RegistrySnapshot:
    """Copy of the group + changeset registry, for rollback with a DOM snapshot."""

    counter: int
    groups: dict[int, tuple[int, ...]]
    revision_groups: dict[int, int | None]
    group_sources: dict[int, GroupSource]
    changeset_counter: int
    changesets: dict[int, tuple[int, ...]]
    group_changesets: dict[int, int]
    changeset_sources: dict[int, GroupSource]


@dataclass
class EditOperation:
    """A single edit operation for batch processing.

    Prefer the typed constructors (:meth:`replace`, :meth:`delete`,
    :meth:`insert_after`, :meth:`insert_before`) — they validate arguments at
    construction time with the same rules ``batch_edit`` applies, so mistakes
    surface immediately instead of at apply time. The raw
    ``EditOperation(action=..., ...)`` form remains supported.
    """

    action: Literal["replace", "delete", "insert_after", "insert_before"]
    paragraph: str  # Required: hash-anchored ref like "P3#a7b2"
    find: str | None = None  # For replace
    replace_with: str | None = None  # For replace
    text: str | None = None  # For delete (text to delete) or insert (text to insert)
    anchor: str | None = None  # For insert_after/insert_before
    occurrence: int | None = None  # None = target must be unique in the paragraph

    @staticmethod
    def _validate_common(constructor: str, paragraph: str, occurrence: int | None) -> None:
        """Construction-time checks shared by all typed constructors."""
        ParagraphRef.parse(paragraph)
        _require_valid_occurrence(occurrence, f"EditOperation.{constructor}(): ")

    @classmethod
    def replace(cls, find: str, replace_with: str, *, paragraph: str, occurrence: int | None = None) -> "EditOperation":
        """Build a validated replace operation (mirrors ``Document.replace``).

        Args:
            find: Text to find and replace (must be non-empty)
            replace_with: Replacement text (empty string allowed — replacing
                with nothing is a valid tracked deletion)
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence within the paragraph (0 = first).
                Omitted → ``find`` must be unique in the paragraph, else the
                batch fails with a wrapped AmbiguousTextError at apply time.

        Raises:
            ValueError: If the paragraph ref is malformed, ``occurrence`` is
                not a non-negative integer, ``find`` is not a non-empty
                string, or ``replace_with`` is not a string.
        """
        cls._validate_common("replace", paragraph, occurrence)
        if not isinstance(find, str) or not find:
            raise ValueError(
                f"EditOperation.replace(): 'find' must be a non-empty string — the text to search for, got {find!r}"
            )
        if not isinstance(replace_with, str):
            raise ValueError(
                f"EditOperation.replace(): 'replace_with' must be a string (empty string is allowed), "
                f"got {replace_with!r}"
            )
        _reject_control_chars(find, field="'find'", ctx="EditOperation.replace(): ", allow_newline=False)
        _reject_control_chars(replace_with, field="'replace_with'", ctx="EditOperation.replace(): ", allow_newline=True)
        return cls(
            action="replace",
            paragraph=paragraph,
            find=find,
            replace_with=replace_with,
            occurrence=occurrence,
        )

    @classmethod
    def delete(cls, text: str, *, paragraph: str, occurrence: int | None = None) -> "EditOperation":
        """Build a validated delete operation (mirrors ``Document.delete``).

        Args:
            text: Text to mark as deleted (must be non-empty)
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence within the paragraph (0 = first).
                Omitted → ``text`` must be unique in the paragraph, else the
                batch fails with a wrapped AmbiguousTextError at apply time.

        Raises:
            ValueError: If the paragraph ref is malformed, ``occurrence`` is
                not a non-negative integer, or ``text`` is not a non-empty
                string.
        """
        cls._validate_common("delete", paragraph, occurrence)
        if not isinstance(text, str) or not text:
            raise ValueError(
                f"EditOperation.delete(): 'text' must be a non-empty string — the text to mark as deleted, got {text!r}"
            )
        _reject_control_chars(text, field="'text'", ctx="EditOperation.delete(): ", allow_newline=False)
        return cls(action="delete", paragraph=paragraph, text=text, occurrence=occurrence)

    @classmethod
    def _insert(
        cls,
        action: Literal["insert_after", "insert_before"],
        anchor: str,
        text: str,
        paragraph: str,
        occurrence: int | None,
    ) -> "EditOperation":
        cls._validate_common(action, paragraph, occurrence)
        if not isinstance(anchor, str) or not anchor:
            raise ValueError(
                f"EditOperation.{action}(): 'anchor' must be a non-empty string — the text to insert near, "
                f"got {anchor!r}"
            )
        if not isinstance(text, str):
            raise ValueError(
                f"EditOperation.{action}(): 'text' must be a string (empty string is allowed), got {text!r}"
            )
        _reject_control_chars(anchor, field="'anchor'", ctx=f"EditOperation.{action}(): ", allow_newline=False)
        _reject_control_chars(text, field="'text'", ctx=f"EditOperation.{action}(): ", allow_newline=True)
        return cls(action=action, paragraph=paragraph, anchor=anchor, text=text, occurrence=occurrence)

    @classmethod
    def insert_after(cls, anchor: str, text: str, *, paragraph: str, occurrence: int | None = None) -> "EditOperation":
        """Build a validated insert_after operation (mirrors ``Document.insert_after``).

        Args:
            anchor: Text to find as insertion point (must be non-empty)
            text: Text to insert after the anchor
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence of anchor within the paragraph
                (0 = first). Omitted → ``anchor`` must be unique in the
                paragraph, else the batch fails with a wrapped
                AmbiguousTextError at apply time.

        Raises:
            ValueError: If the paragraph ref is malformed, ``occurrence`` is
                not a non-negative integer, ``anchor`` is not a non-empty
                string, or ``text`` is not a string.
        """
        return cls._insert("insert_after", anchor, text, paragraph, occurrence)

    @classmethod
    def insert_before(cls, anchor: str, text: str, *, paragraph: str, occurrence: int | None = None) -> "EditOperation":
        """Build a validated insert_before operation (mirrors ``Document.insert_before``).

        Args:
            anchor: Text to find as insertion point (must be non-empty)
            text: Text to insert before the anchor
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence of anchor within the paragraph
                (0 = first). Omitted → ``anchor`` must be unique in the
                paragraph, else the batch fails with a wrapped
                AmbiguousTextError at apply time.

        Raises:
            ValueError: If the paragraph ref is malformed, ``occurrence`` is
                not a non-negative integer, ``anchor`` is not a non-empty
                string, or ``text`` is not a string.
        """
        return cls._insert("insert_before", anchor, text, paragraph, occurrence)


def _not_an_edit_operation_message(op: object) -> str:
    """Shared batch_edit/validate_batch message for a non-EditOperation element,
    so the raising and never-raises paths cannot drift apart."""
    return (
        f"expected EditOperation, got {type(op).__name__} — build operations with "
        "EditOperation.replace()/.delete()/.insert_after()/.insert_before()"
    )


@dataclass
class EditValidationResult:
    """Outcome of validating one EditOperation in a dry-run batch."""

    index: int  # 0-based position in the input operations list
    paragraph: str | None  # the operation's paragraph ref (None if it was missing)
    valid: bool  # True if the op would apply cleanly
    error: str | None = None  # human-readable reason when not valid


class EditResult(str):
    """Result of a tracked edit: the new paragraph ref plus revision-group info.

    Subclasses ``str`` — the string value *is* the new hash-anchored
    paragraph reference (e.g. ``"P2#c3d4"``), so an EditResult works
    unchanged anywhere a ref string is expected (``paragraph=`` of
    follow-up edits, equality with plain strings, dict keys).

    Extra attributes:

    - ``group_id``: id of the revision group holding every revision this
      operation created, usable with ``accept_group``/``reject_group``.
      None when the operation created no new revisions — e.g. text spliced
      into one of your own pending insertions (physically merged, so it is
      inseparable from the earlier operation at the XML level), a rewrite
      that found no differences, or a rewrite whose changes all landed
      inside your own pending insertions. Group ids are per-open-Document
      and renumbered on each open — after close()/reopen the same revisions
      belong to a freshly inferred group with a new id, so never carry a
      group_id across sessions (see ``Document.accept_group``).
    - ``changeset_id``: id of the changeset (one whole call: this single
      edit, or the entire ``batch_edit``/``batch_rewrite``) that this
      operation's group belongs to, usable with
      ``accept_changeset``/``reject_changeset``. One changeset contains ≥1
      group; a single edit is a one-group changeset. None whenever
      ``group_id`` is None. Per-open-Document and renumbered on each open,
      exactly like ``group_id``.
    - ``revision_ids``: the w:ids of the group's member revisions, in
      creation order, as of this edit's return; ``()`` when ``group_id`` is
      None. A later edit that splits one of these insertions adds the
      split-off half to the group — ``Document.list_revisions`` reflects
      live membership.
    - ``refs``: every resulting paragraph ref, in document order. A normal
      edit stays inside one paragraph, so ``refs`` is ``(str(self),)``. A
      ``\\n`` edit is a tracked paragraph split, so ``refs`` carries the first
      paragraph (== the string value) plus one ref per new paragraph the split
      created (``("P2#…", "P3#…", …)``). Like all refs these are valid until
      the next structural edit; after a split, later paragraphs' indexes have
      shifted, so re-resolve (``list_paragraphs``/``find_text``) before reusing
      stale refs.
    """

    group_id: int | None
    changeset_id: int | None
    revision_ids: tuple[int, ...]
    refs: tuple[str, ...]

    def __new__(
        cls,
        ref: str,
        group_id: int | None = None,
        revision_ids: tuple[int, ...] = (),
        changeset_id: int | None = None,
        refs: tuple[str, ...] | None = None,
    ) -> "EditResult":
        result = super().__new__(cls, ref)
        result.group_id = group_id
        result.changeset_id = changeset_id
        result.revision_ids = revision_ids
        result.refs = refs if refs is not None else (str(ref),)
        return result


@dataclass(frozen=True)
class SearchResult:
    """Public result of ``Document.find_text`` / ``find_all`` — no DOM internals.

    ``start``/``end`` are character offsets in the *containing paragraph's*
    visible text (text maps are per-paragraph), not document-wide offsets.

    ``paragraph_ref`` is computed at search time and is directly usable as the
    ``paragraph=`` argument of follow-up edits; like refs from
    ``list_paragraphs()``, it is valid until that paragraph is edited.

    ``find_text``'s ``occurrence`` counts matches document-wide, while edit
    methods count within one paragraph — ``paragraph_occurrence`` bridges the
    two: pass it as the ``occurrence=`` of a follow-up edit to target exactly
    the match ``find_text`` located.

    ``paragraph_index`` is the 1-based document-order index of the containing
    paragraph — the same integer embedded in ``paragraph_ref`` — so consumers
    never need to string-parse the ref to compare or sort by position.

    ``repr()``/``str()`` are compact one-liners
    (``SearchResult(P3#a7b2 occ=0 '30 days')``) so printing a list of results
    stays cheap; matched text longer than 60 characters is elided with
    ``"..."`` in the display only. Every field — including the full ``text``
    — remains accessible as an attribute.
    """

    start: int  # Start offset in the paragraph's visible text
    end: int  # Exclusive end offset, same coordinate space
    text: str  # The matched text
    paragraph_ref: str  # Hash-anchored ref like "P3#a7b2"
    paragraph_occurrence: int  # Occurrence index of this match within its paragraph
    spans_revision: bool  # True if the match crosses a tracked-revision boundary
    paragraph_index: int  # 1-based index of the containing paragraph (same as in paragraph_ref)

    def __repr__(self) -> str:
        text = self.text if len(self.text) <= 60 else self.text[:57] + "..."
        spans = " spans_rev" if self.spans_revision else ""
        return f"SearchResult({self.paragraph_ref} occ={self.paragraph_occurrence} {text!r}{spans})"


@dataclass(frozen=True)
class _LocatedMatch:
    """Internal: a text-map match plus the paragraph identity needed for refs."""

    match: TextMapMatch
    paragraph_index: int  # 1-based document-order index of the containing w:p
    paragraph: Element  # The containing w:p element
    paragraph_occurrence: int  # Occurrence index of the match within that paragraph


@dataclass
class Revision:
    """Represents a tracked change (insertion or deletion).

    Location and nesting fields are populated by ``list_revisions``:

    - ``paragraph_ref``: hash-anchored reference (``"P{i}#{hash}"``) of the
      containing paragraph, or None when the revision sits outside any
      ``<w:p>`` (e.g. ``<w:trPr>`` row markers).
    - ``occurrence``: 0-based occurrence index of ``text`` within the
      containing paragraph, counted in the view where the revision's text
      lives — the accepted (visible) view for insertions, the original
      (pre-revision) view for deletions. For insertions it plugs directly
      into the ``occurrence=`` parameter of replace()/delete()/
      add_comment(). None whenever targeting-by-text does not apply: empty
      text, a host insertion whose original text no longer matches its
      visible span (a nested deletion consumed part of it), or a nested
      deletion (its text never existed in the original document).
    - ``nested_under``: id of the nearest enclosing revision (e.g. a foreign
      deletion inside another author's pending insertion), else None.
    - ``contains_ids``: ids of revisions nested inside this one, in document
      order.
    - ``group_id``: id of the revision group this revision belongs to (all
      revisions from one logical edit share it — see
      ``Document.accept_group``). Edits made through the open Document
      record their group directly; revisions already in the file get an
      *inferred* group reconstructed at parse time (see ``group_source``).
      None only when the revision is ungroupable: it sits outside any
      paragraph, lacks an author or date, shares a duplicated id with
      another revision (id-keyed lookup cannot tell the occurrences
      apart), or is a mid-session split half of a foreign insertion.
      Revisions with non-numeric ids are omitted from ``list_revisions()``
      entirely — no id-keyed operation could target them.
    - ``group_source``: provenance of ``group_id`` — ``"recorded"`` for
      groups created by edits through this open Document, ``"inferred"``
      for groups reconstructed at parse time by the heuristic (contiguous
      same-paragraph revisions sharing identical ``w:author`` + ``w:date``
      are one group). Our own writes stamp collision-bumped dates (within
      one open session, two changesets by one author never share a
      second), so inferred groups match the original changesets;
      same-paragraph ops of one ``batch_edit``/``batch_rewrite`` call
      share their date and merge by design. The counter is per-session,
      so own writes from a previous session merge like foreign revisions:
      identical author + date can still over-merge (``w:date`` has second
      precision). None iff ``group_id`` is None.
    - ``changeset_id``: id of the changeset (one whole call — see
      ``EditResult.changeset_id``) this revision's group belongs to,
      resolvable with ``Document.accept_changeset``/``reject_changeset``.
      A changeset is the ``(author, date)`` equivalence class over groups:
      every group sharing this revision's ``w:author`` + ``w:date`` is in
      the same changeset (recorded edits bundle one call; inferred
      changesets partition reconstructed groups by that key). None iff
      ``group_id`` is None.
    - ``changeset_source``: provenance of ``changeset_id`` — ``"recorded"``
      for changesets bundled by a call through this open Document,
      ``"inferred"`` for changesets partitioned at parse time. None iff
      ``changeset_id`` is None.
    """

    id: int
    type: Literal["insertion", "deletion"]
    author: str
    date: datetime | None
    text: str
    paragraph_ref: str | None = None
    occurrence: int | None = None
    nested_under: int | None = None
    contains_ids: tuple[int, ...] = ()
    group_id: int | None = None
    group_source: GroupSource | None = None
    changeset_id: int | None = None
    changeset_source: GroupSource | None = None

    def __repr__(self) -> str:
        kind = "ins" if self.type == "insertion" else "del"
        location = f" @{self.paragraph_ref}" if self.paragraph_ref else ""
        preview = self.text[:30] + ("..." if len(self.text) > 30 else "")
        nested = f", nested_under={self.nested_under}" if self.nested_under is not None else ""
        contains = f", contains={list(self.contains_ids)}" if self.contains_ids else ""
        inferred = "(inferred)" if self.group_source == "inferred" else ""
        group = f", group={self.group_id}{inferred}" if self.group_id is not None else ""
        changeset = f", cs={self.changeset_id}" if self.changeset_id is not None else ""
        return f"Revision({kind} {self.id}{location}: '{preview}' by {self.author}{nested}{contains}{group}{changeset})"


def _ancestor_paragraph(elem) -> Element | None:
    """Nearest <w:p> ancestor of ``elem``, or None if outside any paragraph.

    Not replaceable by ``xml_editor._innermost_ancestor``: this loop also
    stops at the first non-element ancestor, so it terminates even on node
    chains whose ``parentNode`` never yields None (mock DOMs in tests).
    """
    node = elem.parentNode
    while node is not None and node.nodeType == node.ELEMENT_NODE:
        if node.tagName == "w:p":
            return node
        node = node.parentNode
    return None


def _nearest_revision_ancestor_id(elem) -> int | None:
    """id of the closest <w:ins>/<w:del> ancestor carrying a w:id, else None."""
    node = elem.parentNode
    while node is not None and node.nodeType == node.ELEMENT_NODE:
        if node.tagName in ("w:ins", "w:del"):
            rev_id = node.getAttribute("w:id")
            if rev_id:
                return int(rev_id)
        node = node.parentNode
    return None


def _descendant_revision_ids(elem) -> tuple[int, ...]:
    """ids of all <w:ins>/<w:del> descendants of ``elem``, in document order."""
    ids: list[int] = []

    def walk(node) -> None:
        for child in node.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            if child.tagName in ("w:ins", "w:del"):
                rev_id = child.getAttribute("w:id")
                if rev_id:
                    ids.append(int(rev_id))
            walk(child)

    walk(elem)
    return tuple(ids)


def _revision_elements(root) -> list[Element]:
    """All <w:ins>/<w:del> elements under ``root``, in document order.

    Same unconditional-recursion pattern as ``_descendant_revision_ids``:
    nested revisions (e.g. a w:del inside a w:ins) are included, each
    appearing right after its host.
    """
    elems: list[Element] = []

    def walk(node) -> None:
        for child in node.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            if child.tagName in ("w:ins", "w:del"):
                elems.append(child)
            walk(child)

    walk(root)
    return elems


def _insertion_text_nodes(elem) -> list:
    """All <w:t>/<w:delText> descendants of a <w:ins>, in document order.

    Including <w:delText> means a host insertion whose content was later
    deleted by a nested <w:del> still reports the full text it originally
    inserted (plain <w:delText> never appears under <w:ins> otherwise).
    """
    nodes: list = []

    def walk(node) -> None:
        for child in node.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            if child.tagName in ("w:t", "w:delText"):
                nodes.append(child)
            else:
                walk(child)

    walk(elem)
    return nodes


def _has_ancestor(node, ancestor) -> bool:
    """True if ``ancestor`` is ``node`` itself or one of its ancestors."""
    current = node
    while current is not None:
        if current is ancestor:
            return True
        current = current.parentNode
    return False


def _first_child_element(parent, tag: str) -> Element | None:
    """The first *direct* child of ``parent`` with tag name ``tag``, or None."""
    for child in parent.childNodes:
        if child.nodeType == child.ELEMENT_NODE and getattr(child, "tagName", "") == tag:
            return child
    return None


def _next_element_sibling(node):
    """The next sibling that is an element (skipping text/whitespace nodes)."""
    while node is not None and node.nodeType != node.ELEMENT_NODE:
        node = node.nextSibling
    return node


def _paragraph_mark_ins(paragraph) -> Element | None:
    """The paragraph-mark insertion of ``paragraph``: the ``<w:ins>`` marker
    inside ``<w:pPr><w:rPr>`` that flags this paragraph's mark as an inserted
    revision (a tracked paragraph split), or None when the mark is not tracked.
    """
    pPr = _first_child_element(paragraph, "w:pPr")
    if pPr is None:
        return None
    rPr = _first_child_element(pPr, "w:rPr")
    if rPr is None:
        return None
    return _first_child_element(rPr, "w:ins")


def _is_paragraph_mark_ins(ins) -> bool:
    """True if ``ins`` is a paragraph-mark insertion (child of ``w:pPr/w:rPr``)."""
    parent = ins.parentNode
    if parent is None or getattr(parent, "tagName", "") != "w:rPr":
        return False
    grandparent = parent.parentNode
    return grandparent is not None and getattr(grandparent, "tagName", "") == "w:pPr"


@dataclass
class _GroupCapture:
    """Filled in by ``RevisionManager._grouped`` when its with-block exits."""

    group_id: int | None = None


def _occurrence_in_text_map(tm: TextMap, elem, text: str) -> int | None:
    """0-based occurrence index of ``text`` at ``elem``'s own span in ``tm``.

    Mirrors ``find_in_text_map``'s stepping (``idx + 1`` between matches) so
    the result plugs directly into the ``occurrence=`` parameter of the
    anchor APIs. Returns None when the revision's span cannot be equated
    with ``text``: no position in the map belongs to ``elem``, the map
    text at the span's start doesn't spell ``text``, or the spelled-out
    span extends beyond ``elem`` (e.g. a partially consumed host insertion
    whose missing suffix happens to be spelled by the following text —
    anchoring there would silently cross the revision boundary).
    """
    if not text:
        return None
    start = next((i for i, pos in enumerate(tm.positions) if _has_ancestor(pos.node, elem)), None)
    if start is None:
        return None
    if not tm.text.startswith(text, start):
        return None
    if not all(_has_ancestor(pos.node, elem) for pos in tm.positions[start : start + len(text)]):
        return None
    count = 0
    idx = tm.text.find(text)
    while idx != -1 and idx < start:
        count += 1
        idx = tm.text.find(text, idx + 1)
    return count


class _RevisionLocationContext:
    """Per-``list_revisions``-call cache of paragraph indexes, refs, and text maps."""

    def __init__(self, dom):
        self._p_index = {id(p): i for i, p in enumerate(dom.getElementsByTagName("w:p"), start=1)}
        self._refs: dict[int, str] = {}
        self._maps: dict[tuple[int, str], TextMap] = {}

    def paragraph_ref(self, p) -> str | None:
        """Hash-anchored ref ("P{i}#{hash}") of ``p``; None if not indexed."""
        key = id(p)
        index = self._p_index.get(key)
        if index is None:
            return None
        if key not in self._refs:
            # Hash from the cached accepted map, shared with the occurrence
            # path, instead of compute_paragraph_hash (which builds its own).
            self._refs[key] = f"P{index}#{compute_text_hash(self.text_map(p, 'accepted').text)}"
        return self._refs[key]

    def text_map(self, p, view: Literal["accepted", "original"]) -> TextMap:
        """Cached text map of ``p`` for ``view``."""
        key = (id(p), view)
        if key not in self._maps:
            self._maps[key] = build_text_map(p, view=view)
        return self._maps[key]


class RevisionManager:
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
        self._reconstruct_groups()

    def _reconstruct_groups(self) -> None:
        """Infer revision groups for the revisions already in the document.

        Word offers no grouping concept, so nothing about the original
        logical edits survives in the file; this reconstructs them with a
        heuristic: maximal runs of consecutive revisions (document order,
        nested revisions included) in the *same paragraph* sharing identical
        raw ``w:author`` + ``w:date`` strings become one group. Same-paragraph
        contiguity is load-bearing — (author, date) alone would merge every
        edit an author made in the same second across the whole document.
        Singletons get groups too, matching live behavior where a
        one-revision edit gets a one-member group.

        Known, accepted imprecisions:

        - Same-paragraph over-merge is confined to one batch call for our
          own writes within one open session: each changeset stamps a
          collision-bumped date (never reused across an author's
          changesets in that session), so only ops of a single
          batch_edit/batch_rewrite call — which share their date by
          design — merge. The counter is not seeded from dates already in
          the file, so a previous session's own writes behave like
          foreign ones: identical author + date (w:date has second
          precision) can still over-merge.
        - A revision inside a nested paragraph (e.g. a text box's
          w:txbxContent) interrupts the outer paragraph's run in document
          order — conservative over-split.
        - Foreign revisions group by the same heuristic; provenance is
          always honest ("inferred").
        - When Word already resolved part of a former edit, the remaining
          revisions reconstruct as a rump group — accept_group/reject_group
          are rump-tolerant.

        A revision stays unregistered (group_id None, breaking the current
        run) when it has no ancestor <w:p> (e.g. w:trPr row markers), an
        empty author or date, or a non-numeric id (nonconforming producers;
        list_revisions() omits non-numeric ids entirely — no id-keyed
        operation could target them). A duplicated id is wholly ungrouped —
        every occurrence, whether or not individually groupable — because
        id-keyed lookup cannot tell the occurrences apart, so no group may
        contain an ambiguous member.
        """
        run_key: tuple[int, str, str] | None = None
        run_para: Element | None = None
        run_members: list[int] = []
        # (author, date) -> changeset id, for the inferred changeset tier.
        # Groups anywhere in the document sharing an (author, date) join one
        # changeset — a global equivalence class, not a contiguous run.
        changeset_by_key: dict[tuple[str, str], int] = {}

        def close_run() -> None:
            nonlocal run_members
            if run_members:
                group_id = self._group_counter
                self._group_counter += 1
                self._groups[group_id] = tuple(run_members)
                for rev_id in run_members:
                    self._revision_groups[rev_id] = group_id
                self._group_sources[group_id] = "inferred"
                # run_members non-empty guarantees run_key is set (members are
                # only appended after run_key becomes non-None). Its author+date
                # are the group's; drop the paragraph identity so groups in
                # different paragraphs still share one changeset.
                assert run_key is not None
                _, author, date = run_key
                cs_key = (author, date)
                cs_id = changeset_by_key.get(cs_key)
                if cs_id is None:
                    cs_id = self._changeset_counter
                    self._changeset_counter += 1
                    changeset_by_key[cs_key] = cs_id
                    self._changeset_sources[cs_id] = "inferred"
                self._changesets[cs_id] = (*self._changesets.get(cs_id, ()), group_id)
                self._group_changesets[group_id] = cs_id
                run_members = []

        elements = _revision_elements(self.editor.dom.documentElement)

        # Pre-scan for duplicated ids, independent of groupability: an
        # ungroupable occurrence (e.g. missing its date) must still bar its
        # groupable twin from winning a group that id-keyed lookup would
        # then report for both elements.
        seen_ids: set[int] = set()
        duplicate_ids: set[int] = set()
        for elem in elements:
            try:
                elem_id = int(elem.getAttribute("w:id"))
            except ValueError:
                continue
            if elem_id in seen_ids:
                duplicate_ids.add(elem_id)
            seen_ids.add(elem_id)

        for elem in elements:
            paragraph = _ancestor_paragraph(elem)
            author = elem.getAttribute("w:author")
            date = elem.getAttribute("w:date")
            try:
                rev_id = int(elem.getAttribute("w:id"))
            except ValueError:
                rev_id = None
            if rev_id is not None and _is_paragraph_mark_ins(elem):
                self._paragraph_mark_ids.add(rev_id)
            if paragraph is None or not author or not date or rev_id is None or rev_id in duplicate_ids:
                if rev_id is not None and rev_id in duplicate_ids:
                    self._revision_groups[rev_id] = None  # explicitly ungrouped
                close_run()
                run_key = None
                run_para = None
                continue
            # A paragraph boundary whose mark is an inserted revision by the
            # same author+date is NOT a group boundary: the two paragraphs were
            # one before a tracked split, so their revisions stay one group.
            same_run = (
                run_key is not None
                and author == run_key[1]
                and date == run_key[2]
                and (paragraph is run_para or self._is_split_continuation(run_para, paragraph, author, date))
            )
            if not same_run:
                close_run()
            run_key = (id(paragraph), author, date)
            run_para = paragraph
            run_members.append(rev_id)
        close_run()

    def _is_split_continuation(self, prev_para, new_para, author: str, date: str) -> bool:
        """True if ``new_para`` is the tail half of a tracked split of ``prev_para``.

        The signal is durable and Word-preserved: ``prev_para`` carries a
        paragraph-mark insertion (``<w:pPr><w:rPr><w:ins>``) by this same
        author+date and ``new_para`` is its immediate next paragraph. Without
        this, a reopened split — whose revisions span two paragraphs — would
        reconstruct as two separate inferred groups.
        """
        if prev_para is None:
            return False
        mark = _paragraph_mark_ins(prev_para)
        if mark is None:
            return False
        if mark.getAttribute("w:author") != author or mark.getAttribute("w:date") != date:
            return False
        return _next_element_sibling(prev_para.nextSibling) is new_para

    @contextmanager
    def _grouped(self) -> Iterator[_GroupCapture]:
        """Register every revision the wrapped operation creates as one group.

        Captures the <w:ins>/<w:del> elements newly created during the
        with-block (freshly assigned ``w:id``; pre-existing revisions
        re-serialized by insertion splits are never collected), then keeps
        those that are (i)
        authored by us — a split half of a *foreign* insertion gets a fresh
        id but keeps the foreign author, and must not join our group —
        (ii) still attached to the DOM, excluding create-then-remove churn,
        and (iii) not already registered — by ``_adopt_split_tail``, which
        adopts a split-off tail of one of our *own* insertions into its
        origin group (or marks it ungrouped when the origin has no group),
        never letting the splitting operation claim it; or by
        ``_reconstruct_groups`` (redundantly — the collector only reports
        freshly assigned ids, which pre-existing revisions never have).
        A group id is allocated only when that filtered set is non-empty;
        it is exposed on the yielded _GroupCapture after the block exits.
        If the operation raises, nothing is registered.
        """
        capture = _GroupCapture()
        with self.editor.collect_tracked_changes() as collected, self.editor.frozen_timestamp():
            yield capture
        members: list[int] = []
        seen: set[int] = set()
        for elem in collected:
            if elem.getAttribute("w:author") != self.editor.author:
                continue
            if not _has_ancestor(elem, self.editor.dom):
                continue
            try:
                # Ids we assign are always numeric; a non-numeric one was
                # copied from a nonconforming producer's element — leave it
                # ungrouped rather than fail the edit (same tolerance as
                # _get_next_change_id).
                rev_id = int(elem.getAttribute("w:id"))
            except ValueError:
                continue
            if rev_id in self._revision_groups:
                continue
            if rev_id not in seen:
                seen.add(rev_id)
                members.append(rev_id)
        if members:
            group_id = self._group_counter
            self._group_counter += 1
            self._groups[group_id] = tuple(members)
            for rev_id in members:
                self._revision_groups[rev_id] = group_id
            self._group_sources[group_id] = "recorded"
            capture.group_id = group_id

    @contextmanager
    def _changeset(self) -> Iterator[None]:
        """Bundle every group one public call creates as one changeset.

        The changeset is the intent tier: one whole ``batch_edit``/
        ``batch_rewrite`` call, or one single edit, contains ≥1 group.

        Reentrant by reuse (mirrors ``editor.frozen_timestamp``): a nested
        entry defers to the enclosing boundary, so ``batch_rewrite`` ->
        ``rewrite_paragraph`` merges into ONE changeset while a standalone
        ``rewrite_paragraph`` gets its own. Only the outermost entry bundles,
        and only on clean exit — an exception propagates out of the generator
        at the ``yield`` before the bundling code runs, so a failed batch never
        bundles a ghost changeset (its registry is restored by
        ``_restore_registry`` anyway).

        Yields nothing: the bundled id is read back through
        ``changeset_id_of(group_id)`` (via ``_edit_result``/``list_revisions``),
        never off the context manager itself.

        Members are the *recorded* groups whose ids fall in
        ``[start, group_counter)`` and are not yet assigned to a changeset. A
        split tail adopted into an existing group's id (``_adopt_split_tail``
        allocates no new id in range) stays with that group's changeset and is
        never re-bundled.
        """
        if self._in_changeset:
            yield
            return
        self._in_changeset = True
        start = self._group_counter
        try:
            yield
        finally:
            self._in_changeset = False
        members = [
            gid
            for gid in range(start, self._group_counter)
            if gid in self._groups and self._group_sources.get(gid) == "recorded" and gid not in self._group_changesets
        ]
        if members:
            changeset_id = self._changeset_counter
            self._changeset_counter += 1
            self._changesets[changeset_id] = tuple(members)
            for gid in members:
                self._group_changesets[gid] = changeset_id
            self._changeset_sources[changeset_id] = "recorded"

    def _adopt_split_tail(self, original_ins, new_nodes) -> None:
        """Keep a split-off tail of one of our own insertions in its origin group.

        Editing the middle of our own pending insertion physically splits it:
        the trailing half is re-created as a fresh <w:ins> with a fresh w:id.
        That tail is leftover content of the *original* insertion's operation,
        not of the operation doing the splitting — so it joins the original
        insertion's group — recorded or inferred alike, keeping
        reject_group/accept_group of that earlier operation complete. When
        the origin has no group (an ungroupable insertion by this author,
        e.g. one missing its w:date), the tail is registered as explicitly
        ungrouped (None) instead. Either registration stops the active
        ``_grouped`` capture from claiming the tail for the splitting
        operation — otherwise rejecting that operation's group would rip a
        leftover piece out of a pre-existing insertion.
        """
        try:
            origin_group = self._revision_groups.get(int(original_ins.getAttribute("w:id")))
        except ValueError:  # pragma: no cover - our own ins always has a numeric id
            origin_group = None
        for node in new_nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":  # pragma: no branch
                tail_id = int(node.getAttribute("w:id"))
                self._revision_groups[tail_id] = origin_group
                if origin_group is not None:
                    self._groups[origin_group] = (*self._groups[origin_group], tail_id)

    def _registry_snapshot(self) -> _RegistrySnapshot:
        """Snapshot the group + changeset registry for rollback with a DOM snapshot."""
        return _RegistrySnapshot(
            counter=self._group_counter,
            groups=dict(self._groups),
            revision_groups=dict(self._revision_groups),
            group_sources=dict(self._group_sources),
            changeset_counter=self._changeset_counter,
            changesets=dict(self._changesets),
            group_changesets=dict(self._group_changesets),
            changeset_sources=dict(self._changeset_sources),
        )

    def _restore_registry(self, snapshot: _RegistrySnapshot) -> None:
        """Restore the group + changeset registry captured by ``_registry_snapshot``."""
        self._group_counter = snapshot.counter
        self._groups = snapshot.groups
        self._revision_groups = snapshot.revision_groups
        self._group_sources = snapshot.group_sources
        self._changeset_counter = snapshot.changeset_counter
        self._changesets = snapshot.changesets
        self._group_changesets = snapshot.group_changesets
        self._changeset_sources = snapshot.changeset_sources

    def group_id_of(self, revision_id: int) -> int | None:
        """Group id of a revision (recorded or inferred), or None if ungrouped."""
        return self._revision_groups.get(revision_id)

    def group_revisions(self, group_id: int) -> tuple[int, ...]:
        """Member revision ids of ``group_id``, in creation order (recorded
        groups) or document order (inferred groups).

        Raises:
            RevisionError: If the group id is unknown to this manager (group
                ids are per-open-Document and renumbered on each open).
        """
        members = self._groups.get(group_id)
        if members is None:
            raise RevisionError(
                f"Unknown revision group: {group_id}. Group ids are per-open-Document and "
                f"renumbered on each open (recorded for this session's edits, inferred by "
                f"reconstruction for revisions already in the file); use a group_id from "
                f"this session's EditResult or list_revisions().",
                group_id=group_id,
            )
        return members

    def changeset_id_of(self, group_id: int) -> int | None:
        """Changeset id of a group (recorded or inferred), or None if unassigned."""
        return self._group_changesets.get(group_id)

    def split_count(self, group_id: int) -> int:
        """Number of tracked paragraph splits a group made.

        Counts the group's paragraph-mark insertions; the split spans
        ``split_count + 1`` consecutive paragraphs. Zero for a normal edit.
        Answered from ``_paragraph_mark_ids`` — no DOM walk — so building an
        EditResult stays cheap for the common no-split edit.
        """
        return sum(1 for rev_id in self._groups.get(group_id, ()) if rev_id in self._paragraph_mark_ids)

    def changeset_groups(self, changeset_id: int) -> tuple[int, ...]:
        """Member group ids of ``changeset_id``, in group-creation order
        (recorded changesets) or document order (inferred changesets).

        Raises:
            RevisionError: If the changeset id is unknown to this manager
                (changeset ids are per-open-Document and renumbered on each
                open, exactly like group ids).
        """
        members = self._changesets.get(changeset_id)
        if members is None:
            raise RevisionError(
                f"Unknown changeset: {changeset_id}. Changeset ids are per-open-Document and "
                f"renumbered on each open (recorded for this session's calls, inferred by "
                f"reconstruction for revisions already in the file); use a changeset_id from "
                f"this session's EditResult or list_revisions().",
                changeset_id=changeset_id,
            )
        return members

    def _resolve_paragraph(self, ref: ParagraphRef, paragraphs: list[Element] | None = None):
        """Resolve a ParagraphRef to its <w:p> element, validating the hash.

        Args:
            ref: Parsed paragraph reference
            paragraphs: Optional pre-fetched list of every <w:p> element, so
                batch callers pay for one full-DOM walk per batch instead of
                one per operation. Default None fetches fresh.

        Returns:
            The <w:p> DOM element

        Raises:
            ParagraphIndexError: If paragraph index is out of range
            HashMismatchError: If the hash doesn't match current content
        """
        if paragraphs is None:
            paragraphs = self.editor.dom.getElementsByTagName("w:p")
        if ref.index < 1 or ref.index > len(paragraphs):
            raise ParagraphIndexError(ref.index, len(paragraphs))
        p = paragraphs[ref.index - 1]
        actual_hash = compute_paragraph_hash(p)
        if actual_hash != ref.hash:
            tm = build_text_map(p)
            preview = tm.text[:80]
            if len(tm.text) > 80:
                preview += "..."
            raise HashMismatchError(ref.index, ref.hash, actual_hash, preview)
        return p

    def _locate_in_paragraph(self, paragraph, paragraph_ref: str, text: str, occurrence: int | None) -> TextMapMatch:
        """The single scoped find-or-raise path shared by every paragraph-scoped edit.

        Args:
            paragraph: The resolved <w:p> element.
            paragraph_ref: The caller's ref string (for error messages).
            text: Text to locate.
            occurrence: Which occurrence (0 = first). None means the text must
                be unique within the paragraph.

        Raises:
            ValueError: If ``occurrence`` is negative or not an integer, or ``text`` is not a
                non-empty string.
            TextNotFoundError: If the text is absent, or ``occurrence`` is out
                of range (then with ``occurrence``/``total_occurrences`` set).
            AmbiguousTextError: If ``occurrence`` is None and the text matches
                more than once in the paragraph.
        """
        _require_valid_occurrence(occurrence)
        if not isinstance(text, str) or not text:
            raise ValueError(f"search text must be a non-empty string, got {text!r}")
        text_map = build_text_map(paragraph)
        total = count_in_text_map(text_map, text)

        occ = occurrence if occurrence is not None else 0
        match = find_in_text_map(text_map, text, occ)
        if match is None:
            if total > 0:
                raise TextNotFoundError(
                    text,
                    paragraph_ref=paragraph_ref,
                    paragraph_preview=text_map.text,
                    occurrence=occ,
                    total_occurrences=total,
                )
            raise TextNotFoundError(
                text,
                paragraph_ref=paragraph_ref,
                paragraph_preview=text_map.text,
            )
        if occurrence is None and total > 1:
            raise AmbiguousTextError(
                text,
                paragraph_ref=paragraph_ref,
                paragraph_preview=text_map.text,
                total_occurrences=total,
            )
        return match

    def batch_edit(self, operations: list[EditOperation]) -> list[int]:
        """Apply multiple edits atomically with upfront hash validation.

        Validates all paragraph hashes before applying any edits.
        Applies edits in reverse paragraph order so earlier paragraphs'
        hashes remain valid throughout. The whole call is one changeset:
        every op's revisions share one ``w:date``, while each op still
        records its own revision group in-session.

        Args:
            operations: List of EditOperation objects (each must have paragraph set)

        Returns:
            List of change IDs, one per operation (in original input order)

        Raises:
            BatchOperationError: If any operation fails — validation (element
                is not an EditOperation, malformed ref, stale hash, bad index)
                or apply (missing text, ambiguous target). Carries
                ``operation_index`` so the caller knows which op failed, and
                ``original`` (also ``__cause__``) with the underlying typed
                exception. No edits are applied on failure.
        """
        if not operations:
            return []

        # One full-DOM <w:p> walk shared by the whole batch. Safe because
        # batch ops never add, remove, or replace <w:p> elements (they only
        # rewrite runs inside a paragraph) and minidom returns a plain
        # non-live list. After a rollback the DOM is replaced, but the
        # exception propagates immediately and this list is never used again.
        paragraphs = self.editor.dom.getElementsByTagName("w:p")

        # Parse and validate all refs upfront
        parsed: list[tuple[int, ParagraphRef, EditOperation]] = []
        for i, op in enumerate(operations):
            if not isinstance(op, EditOperation):
                raise BatchOperationError(i, _not_an_edit_operation_message(op))
            if not op.paragraph:
                raise BatchOperationError(i, "paragraph reference is required for batch mode")
            try:
                ref = ParagraphRef.parse(op.paragraph)
                self._resolve_paragraph(ref, paragraphs)  # Raises HashMismatchError if stale
            except (ValueError, DocxEditError) as e:
                raise BatchOperationError(i, str(e), original=e) from e
            parsed.append((i, ref, op))

        # Sort by paragraph index descending (reverse order) for application
        # Stable sort preserves original order for same-paragraph edits
        parsed.sort(key=lambda x: x[1].index, reverse=True)

        # Snapshot DOM and group registry before any mutation so we can roll
        # back on partial failure — without the registry snapshot, rollback
        # would leave ghost groups pointing at reverted revision ids.
        snapshot = self.editor.dom.toxml(encoding=self.editor.encoding)
        registry_snapshot = self._registry_snapshot()

        try:
            results = [0] * len(operations)
            # One batch call = one changeset: every op's revisions share one
            # w:date (the per-op _grouped() freezes join this outer scope) and
            # every op's group bundles into that one changeset. Inside the try
            # so a failed op never bundles a changeset (the generator raises at
            # its yield before bundling) and rollback restores the registry.
            with self._changeset(), self.editor.frozen_timestamp():
                for original_idx, _ref, op in parsed:
                    try:
                        # Each op is its own logical edit: one group per op, so
                        # callers can accept one op and reject another.
                        with self._grouped():
                            change_id = self._apply_single_edit(op, paragraphs)
                    except (ValueError, DocxEditError) as e:
                        raise BatchOperationError(original_idx, str(e), original=e) from e
                    results[original_idx] = change_id
            return results
        except Exception:
            # Restore via the line-tracking parser so parse_position is preserved.
            # If rollback itself fails, surface the original edit error — it is
            # the actionable one; a rollback failure is a secondary symptom.
            try:
                self.editor._reload_dom_from_bytes(snapshot)
            except Exception:
                pass
            self._restore_registry(registry_snapshot)
            raise

    def _resolve_action_target(self, op: EditOperation) -> str:
        """Validate op's required args and return the text this op must locate.

        Shared by ``_apply_single_edit`` and ``_validate_single`` so the two
        paths cannot drift out of sync. Rejects a negative ``occurrence`` up
        front (the one non-well-formed input the text-map search chokes on)
        so both paths fail cleanly before the search.

        Raises:
            ValueError: If ``occurrence`` is negative or not an integer, required arguments for
                op.action are missing or not strings, or the action is
                unrecognized.
        """
        _require_valid_occurrence(op.occurrence)

        if op.action == "replace":
            if not op.find or not isinstance(op.replace_with, str):
                raise ValueError("replace requires 'find' and a string 'replace_with'")
            _reject_control_chars(op.find, field="'find'", ctx="replace(): ", allow_newline=False)
            _reject_control_chars(op.replace_with, field="'replace_with'", ctx="replace(): ", allow_newline=True)
            return op.find
        elif op.action == "delete":
            if not op.text:
                raise ValueError("delete requires 'text'")
            _reject_control_chars(op.text, field="'text'", ctx="delete(): ", allow_newline=False)
            return op.text
        elif op.action in ("insert_after", "insert_before"):
            if not op.anchor or not isinstance(op.text, str):
                raise ValueError(f"{op.action} requires 'anchor' and a string 'text'")
            _reject_control_chars(op.anchor, field="'anchor'", ctx=f"{op.action}(): ", allow_newline=False)
            _reject_control_chars(op.text, field="'text'", ctx=f"{op.action}(): ", allow_newline=True)
            return op.anchor
        else:
            raise ValueError(f"Unknown action: {op.action}")

    def _apply_single_edit(self, op: EditOperation, paragraphs: list[Element] | None = None) -> int:
        """Apply a single edit operation. Paragraph hash was already validated.

        ``paragraphs`` is the batch's shared <w:p> snapshot (see batch_edit);
        None fetches fresh.
        """
        ref = ParagraphRef.parse(op.paragraph)
        if paragraphs is None:
            paragraphs = self.editor.dom.getElementsByTagName("w:p")
        p = paragraphs[ref.index - 1]

        target = self._resolve_action_target(op)
        match = self._locate_in_paragraph(p, op.paragraph, target, op.occurrence)

        if op.action == "replace":
            assert op.replace_with is not None  # guaranteed by _resolve_action_target
            return self._replace_across_nodes(match, op.replace_with)
        elif op.action == "delete":
            return self._delete_across_nodes(match)
        else:  # insert_after / insert_before
            assert op.text is not None  # guaranteed by _resolve_action_target
            position = "after" if op.action == "insert_after" else "before"
            return self._insert_near_match(match, op.text, position)

    def validate_batch(self, operations: list[EditOperation]) -> list[EditValidationResult]:
        """Validate a batch of edits without applying any of them.

        Mirrors the checks in ``batch_edit`` / ``_apply_single_edit`` (paragraph
        ref format, hash freshness, per-action argument requirements, and target
        text existence) but never raises and never mutates the document. Each
        operation gets its own result so the caller sees the full picture even
        when some ops are valid and others are not.

        Limitation: each operation is validated independently against the
        *current* document state; sequential effects are not simulated. A batch
        with multiple operations on the same paragraph (where one op's edit
        changes what a later op would see) may validate differently than it
        applies. Cross-paragraph batches are unaffected, since edits never
        change the paragraph count.

        Args:
            operations: List of EditOperation objects (each should have paragraph set)

        Returns:
            One EditValidationResult per operation, in input order. An element
            that is not an EditOperation at all comes back as an invalid
            result (``paragraph=None``), never as an exception.
        """
        if not operations:
            return []
        # One <w:p> walk for the whole dry run — validation is read-only, so
        # the snapshot is trivially stable (same sharing as batch_edit).
        paragraphs = self.editor.dom.getElementsByTagName("w:p")
        results = []
        for i, op in enumerate(operations):
            if not isinstance(op, EditOperation):
                results.append(
                    EditValidationResult(index=i, paragraph=None, valid=False, error=_not_an_edit_operation_message(op))
                )
                continue
            error = self._validate_single(op, paragraphs)
            results.append(
                EditValidationResult(
                    index=i,
                    paragraph=op.paragraph,
                    valid=error is None,
                    error=error,
                )
            )
        return results

    def _validate_single(self, op: EditOperation, paragraphs: list[Element] | None = None) -> str | None:
        """Return an error message if ``op`` would fail, or None if it is valid.

        Reuses ``_resolve_paragraph``, ``_resolve_action_target``, and
        ``_locate_in_paragraph`` — the same helpers ``_apply_single_edit`` uses —
        so dry-run validation cannot drift from real application semantics
        (out-of-range and ambiguous targets produce the same error text).
        Reads only. ``paragraphs`` is the dry run's shared <w:p> snapshot
        (see validate_batch); None fetches fresh.
        """
        if not op.paragraph:
            return "paragraph reference is required for batch mode"

        try:
            ref = ParagraphRef.parse(op.paragraph)
        except ValueError as e:
            return str(e)

        try:
            p = self._resolve_paragraph(ref, paragraphs)
        except (ParagraphIndexError, HashMismatchError) as e:
            return str(e)

        # Resolve required args + the text this op must locate via the same
        # helper _apply_single_edit uses (which also rejects a negative
        # occurrence), so validation cannot drift from application semantics and
        # the locate below only raises the same errors application would.
        try:
            target = self._resolve_action_target(op)
        except ValueError as e:
            return str(e)

        try:
            self._locate_in_paragraph(p, op.paragraph, target, op.occurrence)
        except (ValueError, DocxEditError) as e:
            return str(e)

        return None

    def batch_rewrite(self, rewrites: list[tuple[str, str]]) -> list[int | None]:
        """Rewrite multiple paragraphs with upfront hash validation.

        The whole call is one changeset: all rewrites share one ``w:date``,
        while each rewrite still records its own revision group in-session.

        Returns:
            One revision group id per rewrite, in input order (None for
            rewrites that created no revisions) — each rewrite gets its own
            group via :meth:`rewrite_paragraph`.

        Raises:
            BatchOperationError: If any rewrite fails — validation (malformed
                ref, duplicate paragraph, non-string new_text, stale hash,
                bad index) or apply.
                Carries ``operation_index`` and ``original`` (also
                ``__cause__``) with the underlying typed exception.
        """
        if not rewrites:
            return []

        # Parse and validate all refs upfront
        parsed: list[tuple[int, ParagraphRef, str]] = []
        seen_indices: set[int] = set()
        for i, (ref_str, new_text) in enumerate(rewrites):
            try:
                ref = ParagraphRef.parse(ref_str)
                self._resolve_paragraph(ref)  # Raises HashMismatchError if stale
            except (ValueError, DocxEditError) as e:
                raise BatchOperationError(i, str(e), original=e) from e
            if ref.index in seen_indices:
                raise BatchOperationError(
                    i,
                    f"duplicate paragraph P{ref.index}. Each paragraph can appear at most once in a batch rewrite.",
                )
            if not isinstance(new_text, str):
                raise BatchOperationError(
                    i,
                    f"'new_text' must be a string (empty string is allowed — it deletes all text), got {new_text!r}",
                )
            seen_indices.add(ref.index)
            parsed.append((i, ref, new_text))

        # Sort by paragraph index descending
        parsed.sort(key=lambda x: x[1].index, reverse=True)

        # Snapshot DOM and group registry before any mutation so we can roll
        # back on partial failure — same atomicity contract as batch_edit.
        snapshot = self.editor.dom.toxml(encoding=self.editor.encoding)
        registry_snapshot = self._registry_snapshot()

        try:
            # Apply rewrites in reverse paragraph order. One batch call = one
            # changeset: all rewrites share one w:date (each rewrite still
            # gets its own group; duplicate paragraphs are rejected above, so
            # reconstruction keeps one group per paragraph despite the
            # shared date).
            group_ids: list[int | None] = [None] * len(rewrites)
            with self._changeset(), self.editor.frozen_timestamp():
                for original_idx, ref, new_text in parsed:
                    try:
                        # The inner rewrite_paragraph's own _changeset() becomes
                        # a no-op (reentrancy guard), so the whole call is one
                        # changeset with one group per rewrite.
                        group_ids[original_idx] = self.rewrite_paragraph(f"P{ref.index}#{ref.hash}", new_text)
                    except (ValueError, DocxEditError) as e:
                        raise BatchOperationError(original_idx, str(e), original=e) from e
            return group_ids
        except Exception:
            # Restore via the line-tracking parser so parse_position is preserved.
            # If rollback itself fails, surface the original edit error — it is
            # the actionable one; a rollback failure is a secondary symptom.
            try:
                self.editor._reload_dom_from_bytes(snapshot)
            except Exception:
                pass
            self._restore_registry(registry_snapshot)
            raise

    def rewrite_paragraph(self, ref_str: str, new_text: str) -> int | None:
        """Rewrite a paragraph's text, generating fine-grained tracked changes.

        Diffs old vs new text at word level and applies minimal tracked changes
        (insertions, deletions, replacements) to transform the paragraph. All
        revisions created by one call are registered as a single revision
        group, so ``accept_group``/``reject_group`` can resolve the rewrite
        as a unit.

        Args:
            ref_str: Paragraph reference string (e.g., "P3#a7b2")
            new_text: Desired new text for the paragraph

        Returns:
            The rewrite's revision group id, or None when no revisions were
            created (old text already equals ``new_text``, or every change
            was absorbed into this author's own pending insertions).

        Raises:
            ValueError: If ``new_text`` is not a string
            HashMismatchError: If the paragraph hash doesn't match
            IndexError: If paragraph index is out of range
        """
        if not isinstance(new_text, str):
            raise ValueError(
                f"rewrite_paragraph(): 'new_text' must be a string "
                f"(empty string is allowed — it deletes all text), got {new_text!r}"
            )
        _reject_control_chars(new_text, field="'new_text'", ctx="rewrite_paragraph(): ", allow_newline=True)
        with self._changeset(), self._grouped() as capture:
            self._rewrite_paragraph_inner(ref_str, new_text)
        return capture.group_id

    def _rewrite_paragraph_inner(self, ref_str: str, new_text: str) -> None:
        """Diff-and-apply body of ``rewrite_paragraph`` (runs inside _grouped)."""
        ref = ParagraphRef.parse(ref_str)
        p = self._resolve_paragraph(ref)
        text_map = build_text_map(p)
        old_text = text_map.text

        if old_text == new_text:
            return

        old_tokens = _tokenize_words(old_text)
        new_tokens = _tokenize_words(new_text)

        sm = difflib.SequenceMatcher(None, old_tokens, new_tokens)
        opcodes = sm.get_opcodes()

        # Convert token indices to character offsets in old_text
        old_token_offsets = []
        pos = 0
        for tok in old_tokens:
            old_token_offsets.append(pos)
            pos += len(tok)

        # Build hunks: list of (tag, old_char_start, old_char_end, new_fragment)
        hunks = []
        for tag, i1, i2, j1, j2 in opcodes:
            if tag == "equal":
                continue

            # Character range in old_text
            if i1 < len(old_tokens):
                old_char_start = old_token_offsets[i1]
            else:
                old_char_start = len(old_text)

            if i2 > 0:
                last_tok = old_tokens[i2 - 1]
                old_char_end = old_token_offsets[i2 - 1] + len(last_tok)
            else:
                old_char_end = old_char_start

            new_fragment = "".join(new_tokens[j1:j2])

            hunks.append((tag, old_char_start, old_char_end, new_fragment))

        # Process hunks in reverse order for position stability
        for tag, old_char_start, old_char_end, new_fragment in reversed(hunks):
            # Rebuild text_map each iteration since DOM changes
            text_map = build_text_map(p)

            if tag == "replace":
                match_text = old_text[old_char_start:old_char_end]
                match = self._find_match_at_position(text_map, match_text, old_char_start)
                self._replace_across_nodes(match, new_fragment)

            elif tag == "delete":
                match_text = old_text[old_char_start:old_char_end]
                match = self._find_match_at_position(text_map, match_text, old_char_start)
                self._delete_across_nodes(match)

            elif tag == "insert":
                self._rewrite_insert_at(p, text_map, old_char_start, new_fragment)

    def _find_match_at_position(self, text_map: TextMap, search: str, expected_pos: int) -> TextMapMatch:
        """Find text at an expected character position in the text map.

        Unlike find_in_text_map which finds the first occurrence, this
        verifies the match is at the expected position. Used by
        rewrite_paragraph() to avoid matching the wrong occurrence when
        the same text appears multiple times in a paragraph.

        Raises RevisionError if the text is not found at the expected position.
        """
        idx = text_map.find(search, expected_pos)
        if idx == -1 or idx != expected_pos:
            raise RevisionError(f"Rewrite failed: could not locate '{search}' at position {expected_pos}")
        end = idx + len(search)
        positions = text_map.get_nodes_for_range(idx, end)
        if positions:
            first_ins = positions[0].is_inside_ins
            spans = any(p.is_inside_ins != first_ins for p in positions)
        else:
            spans = False
        return TextMapMatch(
            start=idx,
            end=end,
            text=search,
            positions=positions,
            spans_boundary=spans,
        )

    def _rewrite_insert_at(self, paragraph, text_map: TextMap, char_pos: int, text: str) -> None:
        """Insert text at a character position within a paragraph.

        Used by rewrite_paragraph() for 'insert' opcodes.

        Args:
            paragraph: The <w:p> DOM element
            text_map: Current text map for the paragraph
            char_pos: Character position in visible text to insert at
            text: Text to insert
        """
        if "\n" in text:
            self._ensure_splittable(paragraph)
            segments = text.split("\n")
            if segments[0]:
                self._rewrite_insert_at(paragraph, text_map, char_pos, segments[0])
            self._apply_paragraph_splits(paragraph, char_pos + len(segments[0]), segments[1:])
            return
        if not text_map.positions:
            # Empty paragraph — append insertion
            # Get rPr from any existing run, or use empty
            runs = paragraph.getElementsByTagName("w:r")
            rPr_xml = get_rPr_xml(runs[0]) if runs else ""

            ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"
            # Insert before w:sectPr if present, else append
            sect_prs = paragraph.getElementsByTagName("w:sectPr")
            if sect_prs:
                self.editor.insert_before(sect_prs[0], ins_xml)
            else:
                self.editor.append_to(paragraph, ins_xml)
            return

        if char_pos >= len(text_map.positions):
            # Insert at end — after last character's run
            last_pos = text_map.positions[-1]
            run, rPr_xml = self._get_run_info(last_pos.node)
            if not run:
                return

            # Inside our own <w:ins>: splice text directly; a foreign
            # author's insertion gets our own sibling <w:ins> instead
            ins_ancestor = self._find_ancestor(run, "w:ins")
            if ins_ancestor:
                if self._owns_ins(ins_ancestor):
                    node_text = self._get_node_text(last_pos.node)
                    self._set_node_text(last_pos.node, node_text + text)
                    _set_xml_space_preserve(last_pos.node)
                else:
                    node_text = self._get_node_text(last_pos.node)
                    self._insert_own_ins_within_foreign_ins(ins_ancestor, last_pos.node, len(node_text), text, rPr_xml)
                return

            ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"
            self.editor.insert_after(run, ins_xml)
            return

        # Insert at a position within the text
        pos = text_map.positions[char_pos]
        run, rPr_xml = self._get_run_info(pos.node)
        if not run:
            return

        # Inside our own <w:ins>: splice text directly; a foreign author's
        # insertion gets our own sibling <w:ins> (splitting theirs mid-content)
        ins_ancestor = self._find_ancestor(run, "w:ins")
        if ins_ancestor:
            if self._owns_ins(ins_ancestor):
                node_text = self._get_node_text(pos.node)
                offset = pos.offset_in_node
                self._set_node_text(pos.node, node_text[:offset] + text + node_text[offset:])
                _set_xml_space_preserve(pos.node)
            else:
                self._insert_own_ins_within_foreign_ins(ins_ancestor, pos.node, pos.offset_in_node, text, rPr_xml)
            return

        # Split the run at the offset and insert <w:ins> between
        node_text = self._get_node_text(pos.node)
        offset = pos.offset_in_node
        before_text = node_text[:offset]
        after_text = node_text[offset:]

        xml_parts = []
        if before_text:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before_text)}</w:t></w:r>")
        xml_parts.append(f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>")
        if after_text:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after_text)}</w:t></w:r>")

        new_xml = "".join(xml_parts)
        self.editor.replace_node(run, new_xml)

    def count_matches(self, text: str) -> int:
        """Count how many times a text string appears in the document.

        Uses text maps for accurate counting across element boundaries.

        Args:
            text: Text to search for

        Returns:
            Number of occurrences found
        """
        count = 0
        for paragraph in self.editor.dom.getElementsByTagName("w:p"):
            count += count_in_text_map(build_text_map(paragraph), text)
        return count

    def _locate_document_wide(self, text: str, occurrence: int | None = None) -> TextMapMatch:
        """Document-wide nth-occurrence lookup via text maps.

        Totals come from :meth:`count_matches` rather than a single
        ``count_in_text_map`` call (as ``_locate_in_paragraph`` uses) because
        text maps are per-paragraph — there is no one document-wide map — and
        the exact total is part of the error contract.

        Raises:
            ValueError: If ``occurrence`` is negative or not an integer, or ``text`` is not a
                non-empty string.
            TextNotFoundError: If the text is not found or occurrence doesn't
                exist; ``total_occurrences`` matches :meth:`count_matches`.
            AmbiguousTextError: If ``occurrence`` is None and the text matches
                more than once in the document.
        """
        _require_valid_occurrence(occurrence)
        if not isinstance(text, str) or not text:
            raise ValueError(f"search text must be a non-empty string, got {text!r}")
        occ = occurrence if occurrence is not None else 0
        match = self._find_across_boundaries(text, occ)
        if match is None:
            total = self.count_matches(text)
            if total:
                raise TextNotFoundError(text, occurrence=occ, total_occurrences=total)
            raise TextNotFoundError(text)
        if occurrence is None:
            total = self.count_matches(text)
            if total > 1:
                raise AmbiguousTextError(text, total_occurrences=total)
        return match

    def _find_across_boundaries_located(self, text: str, occurrence: int = 0) -> _LocatedMatch | None:
        """Find the nth occurrence of text across element boundaries.

        Searches across all paragraphs using text maps, keeping paragraph
        identity so callers can build hash-anchored refs.

        Returns:
            A _LocatedMatch, or None if not found.
        """
        current_occurrence = 0
        for idx, paragraph in enumerate(self.editor.dom.getElementsByTagName("w:p"), start=1):
            text_map = build_text_map(paragraph)
            local_occ = 0
            while True:
                match = find_in_text_map(text_map, text, local_occ)
                if match is None:
                    break
                if current_occurrence == occurrence:
                    return _LocatedMatch(
                        match=match,
                        paragraph_index=idx,
                        paragraph=paragraph,
                        paragraph_occurrence=local_occ,
                    )
                current_occurrence += 1
                local_occ += 1
        return None

    def _find_across_boundaries(self, text: str, occurrence: int = 0) -> TextMapMatch | None:
        """Find the nth occurrence of text across element boundaries.

        Searches across all paragraphs using text maps.
        Returns TextMapMatch or None.
        """
        located = self._find_across_boundaries_located(text, occurrence)
        return located.match if located is not None else None

    def find_text(self, text: str, occurrence: int = 0, paragraph: str | None = None) -> SearchResult | None:
        """Find the nth occurrence of text, as a public SearchResult.

        Searches across element boundaries. With ``paragraph=None``,
        ``occurrence`` counts matches document-wide (0 = first); with a
        paragraph reference, the search is scoped to that paragraph and
        ``occurrence`` counts within it. Returns None if not found.

        Raises:
            ValueError: If ``text`` is not a non-empty string, ``occurrence``
                is not a non-negative integer (None included — the default is
                0, not None), or ``paragraph`` is malformed.
            ParagraphIndexError: If ``paragraph``'s index is out of range.
            HashMismatchError: If ``paragraph``'s hash is stale.
        """
        if not isinstance(text, str) or not text:
            raise ValueError(f"find_text(): search text must be a non-empty string, got {text!r}")
        _require_valid_occurrence(occurrence, "find_text(): ", allow_none=False)

        if paragraph is not None:
            results = self.find_all(text, paragraph=paragraph)
            # 0 <= guard: a bare results[occurrence] would let a negative
            # index silently return a match from the end.
            return results[occurrence] if 0 <= occurrence < len(results) else None

        located = self._find_across_boundaries_located(text, occurrence)
        if located is None:
            return None
        return SearchResult(
            start=located.match.start,
            end=located.match.end,
            text=located.match.text,
            paragraph_ref=f"P{located.paragraph_index}#{compute_paragraph_hash(located.paragraph)}",
            paragraph_occurrence=located.paragraph_occurrence,
            spans_revision=located.match.spans_boundary,
            paragraph_index=located.paragraph_index,
        )

    def find_all(self, text: str, paragraph: str | None = None) -> list[SearchResult]:
        """Enumerate every match of ``text`` as a list of SearchResults.

        One call replaces the N+1 ``find_text`` probes needed to enumerate N
        hits. Each result's ``paragraph_ref``/``paragraph_occurrence`` plug
        directly into a follow-up edit's ``paragraph=``/``occurrence=``.

        Args:
            text: Text to search for (must be non-empty).
            paragraph: Optional paragraph reference (e.g., "P2#f3c1") to scope
                the search. None searches the whole document.

        Returns:
            SearchResults in document order; ``[]`` when nothing matches (it
            is an enumeration API, not a lookup — no-match is not an error).

        Raises:
            ValueError: If ``text`` is not a non-empty string, or
                ``paragraph`` is malformed.
            ParagraphIndexError: If ``paragraph``'s index is out of range.
            HashMismatchError: If ``paragraph``'s hash is stale.
        """
        if not isinstance(text, str) or not text:
            raise ValueError(f"find_all(): search text must be a non-empty string, got {text!r}")

        if paragraph is not None:
            ref = ParagraphRef.parse(paragraph)
            paragraphs = [(ref.index, self._resolve_paragraph(ref))]
        else:
            paragraphs = list(enumerate(self.editor.dom.getElementsByTagName("w:p"), start=1))

        results: list[SearchResult] = []
        for idx, p in paragraphs:
            text_map = build_text_map(p)
            paragraph_ref: str | None = None
            local_occ = 0
            while (match := find_in_text_map(text_map, text, local_occ)) is not None:
                if paragraph_ref is None:
                    paragraph_ref = f"P{idx}#{compute_paragraph_hash(p)}"
                results.append(
                    SearchResult(
                        start=match.start,
                        end=match.end,
                        text=match.text,
                        paragraph_ref=paragraph_ref,
                        paragraph_occurrence=local_occ,
                        spans_revision=match.spans_boundary,
                        paragraph_index=idx,
                    )
                )
                local_occ += 1
        return results

    def replace_text(
        self, find: str, replace_with: str, occurrence: int | None = None, paragraph: str | None = None
    ) -> int:
        """Replace text with tracked changes (deletion + insertion).

        Finds the specified occurrence of `find` text and replaces it with `replace_with`,
        creating a tracked deletion for the old text and insertion for the new text.

        Words shared by ``find`` and ``replace_with`` at either end are
        trimmed before revisions are written, so only the changed words become
        a deletion/insertion pair. A replace that only adds or only removes
        words degenerates into a pure insertion or deletion; when
        ``replace_with`` equals the found text, nothing is written and -1 is
        returned (no-op). The insertion carries the formatting (rPr) that
        covers the most characters of the trimmed span — runs sharing
        identical formatting tally together, ties breaking to the
        earliest-seen formatting.

        Args:
            find: Text to find and replace
            replace_with: Replacement text
            occurrence: Which occurrence to replace (0 = first, 1 = second,
                etc.). Omitted → ``find`` must be unique in the search scope,
                else AmbiguousTextError.
            paragraph: Optional paragraph reference (e.g., "P2#f3c1") to scope the search

        Returns:
            The change ID of the insertion (or of the deletion when the
            replace degenerates to a pure deletion; -1 for a no-op)

        Raises:
            ValueError: If ``find`` is not a non-empty string, ``replace_with``
                is not a string, or ``occurrence`` is negative or not an integer
            TextNotFoundError: If the text is not found or occurrence doesn't exist
            AmbiguousTextError: If ``occurrence`` is omitted and ``find``
                matches more than once in the search scope
            HashMismatchError: If the paragraph hash doesn't match
        """
        if not isinstance(replace_with, str):
            raise ValueError(f"'replace_with' must be a string (empty string is allowed), got {replace_with!r}")
        _reject_control_chars(find, field="'find'", ctx="replace(): ", allow_newline=False)
        _reject_control_chars(replace_with, field="'replace_with'", ctx="replace(): ", allow_newline=True)
        with self._changeset(), self._grouped():
            if paragraph is not None:
                ref = ParagraphRef.parse(paragraph)
                p = self._resolve_paragraph(ref)
                match = self._locate_in_paragraph(p, paragraph, find, occurrence)
                return self._replace_across_nodes(match, replace_with)

            match = self._locate_document_wide(find, occurrence)
            return self._replace_across_nodes(match, replace_with)

    def suggest_deletion(self, text: str, occurrence: int | None = None, paragraph: str | None = None) -> int:
        """Mark text as deleted with tracked changes.

        Args:
            text: Text to mark as deleted
            occurrence: Which occurrence to delete (0 = first, 1 = second,
                etc.). Omitted → ``text`` must be unique in the search scope,
                else AmbiguousTextError.
            paragraph: Optional paragraph reference (e.g., "P2#f3c1") to scope the search

        Returns:
            The change ID of the deletion

        Raises:
            ValueError: If ``text`` is not a non-empty string, or
                ``occurrence`` is negative or not an integer
            TextNotFoundError: If the text is not found or occurrence doesn't exist
            AmbiguousTextError: If ``occurrence`` is omitted and ``text``
                matches more than once in the search scope
            HashMismatchError: If the paragraph hash doesn't match
        """
        _reject_control_chars(text, field="'text'", ctx="delete(): ", allow_newline=False)
        with self._changeset(), self._grouped():
            if paragraph is not None:
                ref = ParagraphRef.parse(paragraph)
                p = self._resolve_paragraph(ref)
                match = self._locate_in_paragraph(p, paragraph, text, occurrence)
                return self._delete_across_nodes(match)

            match = self._locate_document_wide(text, occurrence)
            return self._delete_across_nodes(match)

    def _get_run_info(self, node) -> tuple[Element | None, str]:
        """Get the parent w:r element and its rPr XML for a w:t node."""
        run = node.parentNode
        while run and run.nodeName != "w:r":
            run = run.parentNode
        if not run:
            return None, ""
        return run, get_rPr_xml(run)

    def _get_node_text(self, node) -> str:
        """Get text content of a w:t node by concatenating ALL child text nodes.

        Thin wrapper around :func:`get_text_node_data` — kept as a method so
        existing call sites (and external subclasses) don't have to change.
        """
        return get_text_node_data(node)

    def _set_node_text(self, node, text: str) -> None:
        """Replace all text content of a w:t/w:delText element with ``text``.

        Removes every existing TEXT_NODE child and appends a single new one
        carrying the full content. Necessary because assigning to
        ``firstChild.data`` would leave any sibling text nodes behind,
        corrupting the document when the element holds split text (issue #9).
        """
        for child in list(node.childNodes):
            if child.nodeType == child.TEXT_NODE:
                node.removeChild(child)
        node.appendChild(node.ownerDocument.createTextNode(text))

    def _build_cross_boundary_parts(self, match: TextMapMatch) -> list[tuple[Element, str, str, str, str, int]]:
        """Build per-node data for a cross-boundary match.

        Returns list of (run, rPr_xml, before_text, matched_part, after_text, node_id) tuples,
        one per unique w:t node involved in the match. Nodes are in document order.
        """
        # Group positions by their w:t node (not run — a run can have multiple w:t nodes)
        node_data = OrderedDict()
        for pos in match.positions:
            run, rPr_xml = self._get_run_info(pos.node)
            if run is None:
                continue
            nid = id(pos.node)
            if nid not in node_data:
                node_data[nid] = {
                    "run": run,
                    "rPr_xml": rPr_xml,
                    "node": pos.node,
                    "first_offset": pos.offset_in_node,
                    "last_offset": pos.offset_in_node,
                }
            else:
                node_data[nid]["last_offset"] = pos.offset_in_node

        result = []
        for nid, info in node_data.items():
            node_text = self._get_node_text(info["node"])
            first = info["first_offset"]
            last = info["last_offset"]
            before = node_text[:first]
            matched = node_text[first : last + 1]
            after = node_text[last + 1 :]
            result.append((info["run"], info["rPr_xml"], before, matched, after, nid))
        return result

    def _majority_rPr(self, parts: list[tuple[Element, str, str, str, str, int]]) -> str:
        """rPr of the run(s) contributing the most characters to the match.

        Tallies ``len(matched_part)`` per distinct serialized rPr string, in
        first-seen order; ties break to the earliest-seen rPr. Grouping by
        serialized string means semantically equal but differently-ordered
        rPr children tally separately — deterministic, and runs from the same
        source serialize identically.
        """
        tally: dict[str, int] = {}
        for _run, rPr_xml, _before, matched, _after, _nid in parts:
            tally[rPr_xml] = tally.get(rPr_xml, 0) + len(matched)
        if not tally:
            return ""
        return max(tally, key=lambda k: tally[k])

    def _classify_segments(self, match: TextMapMatch) -> list[tuple[bool | None, list[TextPosition]]]:
        """Group match positions into contiguous segments by revision context.

        Returns list of (is_inside_ins, positions_list) tuples.
        """
        segments = []
        current_ins = None
        current_positions = []
        for pos in match.positions:
            if pos.is_inside_ins != current_ins:
                if current_positions:
                    segments.append((current_ins, current_positions))
                current_ins = pos.is_inside_ins
                current_positions = [pos]
            else:
                current_positions.append(pos)
        if current_positions:
            segments.append((current_ins, current_positions))
        return segments

    def _replace_across_nodes(self, match: TextMapMatch, replace_with: str) -> int:
        """Replace text spanning multiple w:t elements, handling mixed revision contexts.

        Words shared by the matched text and ``replace_with`` at either end
        are trimmed first, so only the changed words become revisions. A
        replace whose span trims to nothing on one side degenerates into a
        pure insertion or deletion; one that trims away entirely (replacement
        equals the match) is a no-op returning -1.

        A ``\\n`` in ``replace_with`` means a tracked paragraph split — routed
        to :meth:`_split_replace` (no affix-trimming).
        """
        if "\n" in replace_with:
            return self._split_replace(match, replace_with)
        prefix, suffix = _trim_replace_affixes(match.text, replace_with)
        del_text = match.text[prefix : len(match.text) - suffix]
        ins_text = replace_with[prefix : len(replace_with) - suffix]

        if not del_text and not ins_text:
            return -1
        if not ins_text:
            return self._delete_across_nodes(match.narrowed(prefix, suffix))
        if not del_text:
            if prefix:
                return self._insert_near_match(match.narrowed(0, len(match.text) - prefix), ins_text, "after")
            return self._insert_near_match(match.narrowed(len(match.text) - suffix, 0), ins_text, "before")

        trimmed = match.narrowed(prefix, suffix)
        if trimmed.spans_boundary:
            return self._replace_mixed_state(trimmed, ins_text)
        return self._replace_same_context(trimmed, ins_text)

    def _replace_same_context(self, match: TextMapMatch, replace_with: str) -> int:
        """Replace text spanning multiple runs in the same revision context.

        Groups the match by parent run, then for each run:
        - Keeps text before the match as an unchanged run
        - Puts matched text into w:del
        - Keeps text after the match as an unchanged run
        - Inserts w:ins with replacement text after the last deletion
        """
        parts = self._build_cross_boundary_parts(match)
        if not parts:
            return -1

        # Site D: all positions inside <w:ins> — dispatch on insertion ownership
        if all(p.is_inside_ins for p in match.positions):
            first_node = match.positions[0].node
            ins_elem = self._find_ancestor(first_node, "w:ins")
            ins_groups = self._group_positions_by_ins(match.positions)

            if all(g_ins is None or self._owns_ins(g_ins) for g_ins, _ in ins_groups):
                # All our own — edit in place (historical behavior)
                # Save parent/next sibling before removal may detach ins_elem
                ins_parent = ins_elem.parentNode if ins_elem else None
                ins_next = ins_elem.nextSibling if ins_elem else None

                self._remove_from_insertion(match.positions)

                ins_rPr = self._majority_rPr(parts)
                new_run_xml = f"<w:r>{ins_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r>"

                if ins_elem and ins_elem.parentNode:
                    # ins_elem still in DOM — insert replacement inside it
                    first_child = ins_elem.firstChild
                    if first_child:
                        self.editor.insert_before(first_child, new_run_xml)
                    else:  # pragma: no cover - removing all content deletes the ins outright
                        self.editor.append_to(ins_elem, new_run_xml)
                elif ins_parent:  # pragma: no branch - an ins removed from the DOM had a parent
                    # ins_elem was fully removed — wrap replacement in a new <w:ins>
                    ins_wrapper_xml = f"<w:ins>{new_run_xml}</w:ins>"
                    if ins_next:
                        new_nodes = self.editor.insert_before(ins_next, ins_wrapper_xml)
                    else:
                        new_nodes = self.editor.append_to(ins_parent, ins_wrapper_xml)
                    # The wrapper is a new revision holding this operation's
                    # replacement text — return its id so the caller's
                    # EditResult reaches the group that contains it.
                    for node in new_nodes:
                        if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":  # pragma: no branch
                            return int(node.getAttribute("w:id"))
                # Reached by the splice-in-place branch: no new revision.
                return -1

            # Foreign insertion(s) involved — preserve them: nest our deletion
            # inside, then place our replacement <w:ins> right after it,
            # splitting the foreign ins when trailing content follows.
            first_id, last_del = self._delete_from_ins_positions(match.positions)

            ins_rPr = self._majority_rPr(parts)
            replacement_xml = f"<w:ins><w:r>{ins_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r></w:ins>"
            if last_del is None:  # pragma: no cover - a foreign group always creates a del
                return first_id
            del_ins = self._find_ancestor(last_del, "w:ins")
            if del_ins is not None:
                self._split_ins_after_child(del_ins, last_del)
                new_nodes = self.editor.insert_after(del_ins, replacement_xml)
            else:  # pragma: no cover - Site D positions are inside ins, so the del is nested
                new_nodes = self.editor.insert_after(last_del, replacement_xml)
            for node in new_nodes:
                if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":  # pragma: no branch
                    return int(node.getAttribute("w:id"))
            return first_id  # pragma: no cover - the fragment always yields a w:ins

        # The insertion carries the rPr covering the most characters of the
        # match (same-rPr runs tally together; ties → earliest seen)
        ins_rPr = self._majority_rPr(parts)

        # Group parts by run for multi-w:t preservation
        run_order: list[int] = []
        run_map: dict[int, dict] = {}
        for run, rPr_xml, before, matched, after, nid in parts:
            rid = id(run)
            if rid not in run_map:
                run_order.append(rid)
                run_map[rid] = {"run": run, "rPr_xml": rPr_xml, "parts": []}
            run_map[rid]["parts"].append((before, matched, after, nid))

        xml_parts = []
        part_idx = 0
        total_parts = len(parts)
        for rid in run_order:
            info = run_map[rid]
            run = info["run"]
            rPr_xml = info["rPr_xml"]

            # Build deterministic node-to-part mapping using node ids from parts
            node_to_part = {nid: (before, matched, after) for before, matched, after, nid in info["parts"]}
            parts_emitted = 0

            # Keyword-only defaults bind this iteration's state (B023)
            def render_wt(wt, *, node_to_part=node_to_part, run_rPr=rPr_xml, base_idx=part_idx) -> list[str]:
                nonlocal parts_emitted
                fragments: list[str] = []
                if id(wt) in node_to_part:
                    before, matched, after = node_to_part[id(wt)]
                    parts_emitted += 1

                    if before:
                        fragments.append(f"<w:r>{run_rPr}<w:t>{_escape_xml(before)}</w:t></w:r>")
                    fragments.append(
                        f"<w:del><w:r>{run_rPr}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>"
                    )

                    # Insert replacement after the last deletion
                    if base_idx + parts_emitted == total_parts:
                        fragments.append(f"<w:ins><w:r>{ins_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r></w:ins>")

                    if after:
                        fragments.append(f"<w:r>{run_rPr}<w:t>{_escape_xml(after)}</w:t></w:r>")
                else:
                    # Unmatched sibling — preserve
                    fragments.extend(render_plain_wt(wt, run_rPr))
                return fragments

            # Emit the run's children in document order (w:t split around the
            # match; w:tab/w:br/w:drawing/… preserved in place)
            xml_parts.extend(rebuild_run_fragments(run, rPr_xml, render_wt))
            part_idx += len(node_to_part)

        # Replace all affected runs: insert new XML before first run, remove all runs
        first_run = parts[0][0]
        new_xml = "".join(xml_parts)
        nodes = self.editor.insert_before(first_run, new_xml)

        seen = set()
        for run, _, _, _, _, _ in parts:
            if id(run) in seen:
                continue
            seen.add(id(run))
            parent = run.parentNode
            if parent:
                parent.removeChild(run)

        # Find insertion node ID
        for node in nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":
                return int(node.getAttribute("w:id"))

        return -1

    def _replace_mixed_state(self, match: TextMapMatch, replace_with: str) -> int:
        """Replace text spanning revision boundaries via atomic decomposition.

        For each segment:
        - Regular text: wrap in <w:del> (standard deletion)
        - Inside <w:ins>: remove the matched portion (undo partial insertion)

        Then insert new text as <w:ins>.
        """
        segments = self._classify_segments(match)

        # First position's run anchors the insertion point; the insertion's
        # rPr follows the majority-by-characters rule across the whole match
        first_run, first_rPr = self._get_run_info(match.positions[0].node)
        parts = self._build_cross_boundary_parts(match)
        ins_rPr = self._majority_rPr(parts) if parts else first_rPr

        # Find the first affected element to use as insertion reference point.
        # For regular text, it's the run; for ins text, it's the w:ins element.
        first_pos = match.positions[0]
        if first_pos.is_inside_ins:
            ref_node = self._find_ancestor(first_pos.node, "w:ins")
        else:
            ref_node = first_run

        if ref_node is None:
            return -1

        # Place a marker before ref_node so we can find the insertion point
        # after deletion processing (which may remove ref_node).
        marker = self.editor.dom.createComment("replace-marker")
        ref_node.parentNode.insertBefore(marker, ref_node)  # type: ignore[union-attr]

        # Process each segment to delete/remove the matched text
        # (author-aware: foreign insertions get a nested <w:del>, our own are
        # edited in place)
        for is_inside_ins, positions in segments:
            if is_inside_ins:
                self._delete_from_ins_positions(positions)
            else:
                self._delete_regular_segment(positions)

        # Insert replacement after the last <w:del> sibling following the marker,
        # so it appears after any preserved prefix text.
        ins_xml = f"<w:ins><w:r>{ins_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r></w:ins>"
        last_del = None
        sibling = marker.nextSibling
        while sibling:
            if sibling.nodeType == sibling.ELEMENT_NODE:
                if sibling.tagName == "w:del":
                    last_del = sibling
                elif last_del is not None:
                    # Stop at first non-del element after we found a del
                    break
            sibling = sibling.nextSibling

        if last_del:
            new_nodes = self.editor.insert_after(last_del, ins_xml)
        else:
            # No deletions found — insert after marker
            new_nodes = self.editor.insert_after(marker, ins_xml)

        # Remove marker
        if marker.parentNode:
            marker.parentNode.removeChild(marker)

        # Return the change ID of the new insertion
        for node in new_nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":
                return int(node.getAttribute("w:id"))
        return -1

    def _find_ancestor(self, node, tag_name: str) -> Element | None:
        """Find the nearest ancestor with the given tag name."""
        parent = node.parentNode
        while parent:
            if parent.nodeType == parent.ELEMENT_NODE and parent.tagName == tag_name:
                return parent
            parent = parent.parentNode
        return None

    def _owns_ins(self, ins_elem) -> bool:
        """Whether ``ins_elem`` is the current author's own pending insertion.

        A missing or empty ``w:author`` reads as foreign: we must never
        destructively edit an insertion we cannot attribute to ourselves.
        Comparison is exact string equality — differing Unicode normalization
        or case reads as foreign, which fails safe (we nest instead of
        destroy).
        """
        author = ins_elem.getAttribute("w:author")
        return bool(author) and author == self.editor.author

    def _ins_identity_attrs(self, ins_elem) -> str:
        """Serialize the identity attributes of an insertion for re-creation.

        Returns ``w:author``/``w:date``/``w16du:dateUtc`` (those present) as an
        XML attribute string, so a re-created half of a split ``w:ins`` keeps
        the original author's identity; attribute injection only adds a
        fresh ``w:id``.
        """
        parts = []
        for attr in ("w:author", "w:date", "w16du:dateUtc"):
            value = ins_elem.getAttribute(attr)
            if value:
                parts.append(f' {attr}="{_escape_xml(value)}"')
        return "".join(parts)

    def _group_positions_by_ins(self, positions: list) -> list[tuple[Element | None, list[TextPosition]]]:
        """Group contiguous positions by their ancestor <w:ins> element.

        Positions are in document order, so positions sharing an ins element
        are contiguous; adjacent distinct ins elements form separate groups.
        A group's element is None for positions outside any insertion.
        """
        groups: list[tuple[Element | None, list[TextPosition]]] = []
        current_ins = None
        for pos in positions:
            ins_elem = self._find_ancestor(pos.node, "w:ins")
            if not groups or ins_elem is not current_ins:
                groups.append((ins_elem, [pos]))
                current_ins = ins_elem
            else:
                groups[-1][1].append(pos)
        return groups

    def _delete_from_ins_positions(self, positions: list) -> tuple[int, Element | None]:
        """Author-aware deletion of match positions that sit inside <w:ins>.

        Our own insertions are edited in place (text physically removed, as
        before). A foreign author's insertion is preserved: the matched text
        is wrapped in a nested <w:del> carrying our authorship — Word's own
        representation for deleting another reviewer's pending insertion.

        Returns (first created del id or -1, last created del element or None).
        """
        first_del_id = -1
        last_del: Element | None = None
        for ins_elem, group in self._group_positions_by_ins(positions):
            if ins_elem is None or self._owns_ins(ins_elem):
                self._remove_from_insertion(group)
            else:
                del_id, group_last_del = self._delete_regular_segment(group)
                if first_del_id == -1:
                    first_del_id = del_id
                if group_last_del is not None:  # pragma: no branch - a foreign group always creates a del
                    last_del = group_last_del
        return first_del_id, last_del

    def _split_ins_after_child(self, ins_elem, child) -> None:
        """Split ``ins_elem`` after ``child``, keeping the author's identity.

        Everything following ``child`` moves into a fresh sibling <w:ins>
        that copies this insertion's w:author/w:date (fresh w:id via
        attribute injection). ``ins_elem`` is typically another author's
        insertion — the copied identity keeps both halves attributed to
        them. No-op when nothing follows ``child``. ``child`` may be a
        descendant; the split happens after the direct child containing it.
        """
        while child.parentNode is not ins_elem:
            child = child.parentNode
        trailing = []
        node = child.nextSibling
        while node is not None:
            trailing.append(node)
            node = node.nextSibling
        if not any(n.nodeType == n.ELEMENT_NODE for n in trailing):
            return
        children_xml = "".join(n.toxml() for n in trailing)
        for n in trailing:
            ins_elem.removeChild(n)
        identity_xml = self._ins_identity_attrs(ins_elem)
        self.editor.insert_after(ins_elem, f"<w:ins{identity_xml}>{children_xml}</w:ins>")

    def _split_foreign_ins_at(self, edge_node, offset: int) -> Element | None:
        """Make (edge_node, offset) fall on a child boundary of its <w:ins>.

        Splits the run containing ``edge_node`` at ``offset`` when the split
        point falls mid-run. Returns the last element that belongs to the
        left side of the split point (None when the split point is at the
        very start of the insertion's content).

        Group caveat: when the enclosing *foreign* insertion is later split
        into fresh-id halves, those halves are not adopted into the origin's
        inferred group (adoption is deliberately limited to our own
        insertions) — resolving that group then affects only part of the
        visual insertion. Foreign grouping is best-effort by design.
        """
        run, rPr_xml = self._get_run_info(edge_node)
        if not run:  # pragma: no cover - a w:t node always sits inside a run
            return None
        node_text = self._get_node_text(edge_node)

        # This site splits a run into left/right halves rather than rendering
        # per-w:t, and must know which side each child lands on, so it keeps
        # its own direct-children walk instead of rebuild_run_fragments.
        # Non-text children (w:tab, w:br, w:drawing, …) stay in document
        # order on whichever side of the split point they fall.
        left_parts: list[str] = []
        right_parts: list[str] = []
        side = left_parts
        for child in run.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            tag = getattr(child, "tagName", "")
            if tag == "w:rPr":
                continue
            if child is edge_node:
                if node_text[:offset]:
                    left_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(node_text[:offset])}</w:t></w:r>")
                if node_text[offset:]:
                    right_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(node_text[offset:])}</w:t></w:r>")
                side = right_parts
                continue
            if tag == "w:t":
                wt_text = self._get_node_text(child)
                if wt_text:
                    side.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(wt_text)}</w:t></w:r>")
                continue
            side.append(f"<w:r>{rPr_xml}{child.toxml()}</w:r>")

        if not right_parts:
            # Split point is at the end of this run — run boundary already
            return run
        if not left_parts:
            # Split point is immediately before this run
            # (isinstance so ty narrows minidom's sibling union — the usual
            # nodeType comparison doesn't)
            prev = run.previousSibling
            while prev is not None:
                if isinstance(prev, Element):
                    return prev
                prev = prev.previousSibling
            return None

        new_nodes = self.editor.replace_node(run, "".join(left_parts + right_parts))
        elements = [n for n in new_nodes if n.nodeType == n.ELEMENT_NODE]
        return elements[len(left_parts) - 1]

    def _insert_own_ins_within_foreign_ins(self, ins_elem, edge_node, offset: int, text: str, rPr_xml: str) -> int:
        """Insert our own <w:ins> at (edge_node, offset) inside a foreign ins.

        Never splices into the foreign insertion (that would credit the other
        author) and never nests <w:ins> in <w:ins> (invalid OOXML). Boundary
        offsets produce a plain sibling; mid-insertion offsets split the
        foreign ins into two identity-preserving halves with our ins between.

        Returns the new insertion's change id.
        """
        own_ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"
        wt_nodes = self._get_wt_nodes_in_ancestor(ins_elem)
        node_text = self._get_node_text(edge_node)

        if edge_node is wt_nodes[0] and offset == 0:
            new_nodes = self.editor.insert_before(ins_elem, own_ins_xml)
        elif edge_node is wt_nodes[-1] and offset == len(node_text):
            new_nodes = self.editor.insert_after(ins_elem, own_ins_xml)
        else:
            boundary = self._split_foreign_ins_at(edge_node, offset)
            if boundary is None:  # pragma: no cover - non-boundary offsets never split at the start
                new_nodes = self.editor.insert_before(ins_elem, own_ins_xml)
            else:
                self._split_ins_after_child(ins_elem, boundary)
                new_nodes = self.editor.insert_after(ins_elem, own_ins_xml)

        for node in new_nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":  # pragma: no branch
                return int(node.getAttribute("w:id"))
        return -1  # pragma: no cover - the fragment always yields a w:ins

    def _remove_from_insertion(self, positions: list) -> None:
        """Remove matched text from inside a <w:ins> element.

        Handles segments spanning multiple w:t nodes within the insertion.
        If the entire insertion text is matched, removes the <w:ins> element.
        If partial, truncates or splits.
        """
        # Group positions by w:t node to handle multi-node segments
        node_groups = OrderedDict()
        for pos in positions:
            nid = id(pos.node)
            if nid not in node_groups:
                node_groups[nid] = {"node": pos.node, "first": pos.offset_in_node, "last": pos.offset_in_node}
            else:
                node_groups[nid]["last"] = pos.offset_in_node

        groups = list(node_groups.values())
        first_group = groups[0]
        last_group = groups[-1]

        first_node = first_group["node"]
        last_node = last_group["node"]
        first_offset = first_group["first"]
        last_offset = last_group["last"]

        before = self._get_node_text(first_node)[:first_offset]
        after = self._get_node_text(last_node)[last_offset + 1 :]

        ins_elem = self._find_ancestor(first_node, "w:ins")

        if not before and not after and len(groups) == len(self._get_wt_nodes_in_ancestor(ins_elem)):
            # Entire insertion matched -- remove the <w:ins> element
            if ins_elem and ins_elem.parentNode:
                ins_elem.parentNode.removeChild(ins_elem)
        elif len(groups) == 1 and first_node is last_node:
            # Single node — use simple truncate/split logic
            node_text = self._get_node_text(first_node)
            before_text = node_text[:first_offset]
            after_text = node_text[last_offset + 1 :]

            if not before_text and not after_text:
                # Entire single node matched
                if ins_elem and ins_elem.parentNode:
                    if len(self._get_wt_nodes_in_ancestor(ins_elem)) == 1:
                        # Sole w:t — remove entire <w:ins>
                        ins_elem.parentNode.removeChild(ins_elem)
                    else:
                        # Other w:t nodes exist — remove just this w:t (and run if empty)
                        self._remove_wt_and_maybe_run(first_node)
            elif not before_text:
                self._set_node_text(first_node, after_text)
                _set_xml_space_preserve(first_node)
            elif not after_text:
                self._set_node_text(first_node, before_text)
                _set_xml_space_preserve(first_node)
            else:
                # Middle split
                self._set_node_text(first_node, before_text)
                _set_xml_space_preserve(first_node)
                run = self._find_ancestor(first_node, "w:r")
                if ins_elem and run:
                    rPr_xml = get_rPr_xml(run)
                    after_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(after_text)}</w:t></w:r></w:ins>"
                    new_nodes = self.editor.insert_after(ins_elem, after_xml)
                    self._adopt_split_tail(ins_elem, new_nodes)
        else:
            # Multi-node: truncate first node to before, last node to after,
            # remove intermediate nodes entirely.
            # Only remove the w:t node; remove the run only if no w:t children remain.
            if before:
                self._set_node_text(first_node, before)
                _set_xml_space_preserve(first_node)
            else:
                self._remove_wt_and_maybe_run(first_node)

            if after:
                self._set_node_text(last_node, after)
                _set_xml_space_preserve(last_node)
            else:
                self._remove_wt_and_maybe_run(last_node)

            # Remove intermediate nodes unconditionally (entire text is matched)
            for group in groups[1:-1]:
                self._remove_wt_and_maybe_run(group["node"])

    def _remove_wt_and_maybe_run(self, wt_node) -> None:
        """Remove a w:t node, and its parent w:r if no meaningful children remain.

        Preserves the run if it still contains non-text content children
        like w:tab, w:br, w:drawing, etc.
        """
        run = self._find_ancestor(wt_node, "w:r")
        if wt_node.parentNode:
            wt_node.parentNode.removeChild(wt_node)
        if run and not run.getElementsByTagName("w:t") and run.parentNode:
            has_content_children = any(
                n.nodeType == n.ELEMENT_NODE and getattr(n, "tagName", None) not in ("w:t", "w:rPr")
                for n in run.childNodes
            )
            if not has_content_children:
                run.parentNode.removeChild(run)

    def _get_wt_nodes_in_ancestor(self, ancestor) -> list:
        """Get all w:t nodes inside an ancestor element."""
        if ancestor is None:
            return []
        return ancestor.getElementsByTagName("w:t")

    def _delete_regular_segment(self, positions: list) -> tuple[int, Element | None]:
        """Wrap matched text in <w:del> in place, run by run.

        Groups positions by run first, then by w:t node within each run,
        so that each run is removed exactly once even when it contains
        multiple w:t nodes involved in the match. The rebuilt runs stay at
        the original location, so this serves both regular top-level text
        and nesting a deletion inside a foreign author's <w:ins> (the new
        <w:del> is stamped with the current author either way).

        Returns (first created del id or -1, last created del element or None).
        """
        # Group positions by run, then by node within each run
        run_groups: OrderedDict[int, dict] = OrderedDict()
        for pos in positions:
            run, rPr_xml = self._get_run_info(pos.node)
            if not run:
                continue
            rid = id(run)
            if rid not in run_groups:
                run_groups[rid] = {"run": run, "rPr_xml": rPr_xml, "nodes": OrderedDict()}
            nid = id(pos.node)
            node_map = run_groups[rid]["nodes"]
            if nid not in node_map:
                node_map[nid] = {"node": pos.node, "first": pos.offset_in_node, "last": pos.offset_in_node}
            else:
                node_map[nid]["last"] = pos.offset_in_node

        # Flatten to a list of (run_info, node_group) for global indexing
        all_node_groups: list[tuple[dict, dict]] = []
        for run_info in run_groups.values():
            for ng in run_info["nodes"].values():
                all_node_groups.append((run_info, ng))

        total = len(all_node_groups)
        first_del_id = -1
        last_del: Element | None = None
        processed_runs: set[int] = set()

        for _global_idx, (run_info, _) in enumerate(all_node_groups):
            run = run_info["run"]
            rPr_xml = run_info["rPr_xml"]
            rid = id(run)

            if rid in processed_runs:
                continue

            node_items = list(run_info["nodes"].values())

            # Render ALL w:t nodes in this run, preserving unmatched ones.
            # Keyword-only defaults bind this iteration's state (B023).
            def render_wt(wt, *, run_info=run_info, run_rPr=rPr_xml, node_items=node_items, rid=rid) -> list[str]:
                fragments: list[str] = []
                if id(wt) not in run_info["nodes"]:
                    # Unmatched sibling — preserve as-is
                    return render_plain_wt(wt, run_rPr)

                ng = run_info["nodes"][id(wt)]
                node_text = self._get_node_text(ng["node"])
                first_offset = ng["first"]
                last_offset = ng["last"]

                # Determine this node group's position in the global sequence
                run_keys = list(run_groups.keys())
                local_idx = node_items.index(ng)
                preceding_nodes = sum(len(run_groups[k]["nodes"]) for k in run_keys[: run_keys.index(rid)])
                global_pos = preceding_nodes + local_idx
                is_first_overall = global_pos == 0
                is_last_overall = global_pos == total - 1

                before = node_text[:first_offset] if is_first_overall else ""
                after = node_text[last_offset + 1 :] if is_last_overall else ""

                # For intermediate nodes, the entire text is matched
                if not is_first_overall and not is_last_overall:
                    matched = node_text
                else:
                    matched = node_text[first_offset : last_offset + 1]

                if before:
                    fragments.append(f"<w:r>{run_rPr}<w:t>{_escape_xml(before)}</w:t></w:r>")
                fragments.append(f"<w:del><w:r>{run_rPr}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>")
                if after:
                    fragments.append(f"<w:r>{run_rPr}<w:t>{_escape_xml(after)}</w:t></w:r>")
                return fragments

            # Emit the run's children in document order (w:tab/w:br/w:drawing/…
            # preserved in place)
            new_xml = "".join(rebuild_run_fragments(run, rPr_xml, render_wt))
            nodes = self.editor.insert_before(run, new_xml)
            if run.parentNode:
                run.parentNode.removeChild(run)
            processed_runs.add(rid)

            for n in nodes:
                if n.nodeType == n.ELEMENT_NODE and n.tagName == "w:del":
                    if first_del_id == -1:
                        first_del_id = int(n.getAttribute("w:id"))
                    last_del = n

        return first_del_id, last_del

    def _delete_across_nodes(self, match: TextMapMatch) -> int:
        """Delete text spanning multiple w:t elements."""
        if match.spans_boundary:
            return self._delete_mixed_state(match)
        return self._delete_same_context(match)

    def _delete_same_context(self, match: TextMapMatch) -> int:
        """Delete text spanning multiple runs in the same revision context."""
        parts = self._build_cross_boundary_parts(match)
        if not parts:
            return -1

        # Site F: all positions inside <w:ins> — author-aware dispatch (our own
        # insertions edited in place, foreign ones get a nested <w:del>)
        if all(p.is_inside_ins for p in match.positions):
            first_id, _ = self._delete_from_ins_positions(match.positions)
            return first_id

        # Group parts by run, using node ids from _build_cross_boundary_parts
        run_parts: OrderedDict[int, list] = OrderedDict()
        for part in parts:
            rid = id(part[0])
            if rid not in run_parts:
                run_parts[rid] = []
            run_parts[rid].append(part)

        xml_parts = []
        for _rid, rparts in run_parts.items():
            run = rparts[0][0]
            rPr_xml = rparts[0][1]

            # Build deterministic node-to-part mapping using node ids from parts
            node_to_part = {nid: (rp_xml, before, matched, after) for _, rp_xml, before, matched, after, nid in rparts}

            # Keyword-only defaults bind this iteration's state (B023)
            def render_wt(wt, *, node_to_part=node_to_part, run_rPr=rPr_xml) -> list[str]:
                fragments: list[str] = []
                if id(wt) in node_to_part:
                    rp_xml, before, matched, after = node_to_part[id(wt)]
                    if before:
                        fragments.append(f"<w:r>{rp_xml}<w:t>{_escape_xml(before)}</w:t></w:r>")
                    fragments.append(f"<w:del><w:r>{rp_xml}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>")
                    if after:
                        fragments.append(f"<w:r>{rp_xml}<w:t>{_escape_xml(after)}</w:t></w:r>")
                else:
                    # Unmatched sibling — preserve
                    fragments.extend(render_plain_wt(wt, run_rPr))
                return fragments

            # Emit the run's children in document order (matched w:t as
            # <w:del>; w:tab/w:br/w:drawing/… preserved in place)
            xml_parts.extend(rebuild_run_fragments(run, rPr_xml, render_wt))

        first_run = parts[0][0]
        new_xml = "".join(xml_parts)
        nodes = self.editor.insert_before(first_run, new_xml)

        seen = set()
        for run, _, _, _, _, _ in parts:
            if id(run) in seen:
                continue
            seen.add(id(run))
            parent = run.parentNode
            if parent:
                parent.removeChild(run)

        # Find deletion node ID
        for node in nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:del":
                return int(node.getAttribute("w:id"))

        return -1

    def _delete_mixed_state(self, match: TextMapMatch) -> int:
        """Delete text spanning revision boundaries.

        Regular text segments are wrapped in <w:del>.
        Insertion text segments are removed (undoing partial insertion).
        """
        segments = self._classify_segments(match)

        first_del_id = -1
        for is_inside_ins, positions in segments:
            if is_inside_ins:
                del_id, _ = self._delete_from_ins_positions(positions)
            else:
                del_id, _ = self._delete_regular_segment(positions)
            if first_del_id == -1:
                first_del_id = del_id

        return first_del_id

    def insert_text_after(
        self, anchor: str, text: str, occurrence: int | None = None, paragraph: str | None = None
    ) -> int:
        """Insert text after anchor with tracked changes.

        Args:
            anchor: Text to find as the anchor point
            text: Text to insert after the anchor
            occurrence: Which occurrence of anchor to use (0 = first,
                1 = second, etc.). Omitted → ``anchor`` must be unique in the
                search scope, else AmbiguousTextError.
            paragraph: Optional paragraph reference (e.g., "P2#f3c1") to scope the search

        Returns:
            The change ID of the insertion

        Raises:
            ValueError: If ``anchor`` is not a non-empty string, ``text`` is
                not a string, or ``occurrence`` is negative or not an integer
            TextNotFoundError: If the anchor text is not found or occurrence doesn't exist
            AmbiguousTextError: If ``occurrence`` is omitted and ``anchor``
                matches more than once in the search scope
            HashMismatchError: If the paragraph hash doesn't match
        """
        return self._insert_text(anchor, text, position="after", occurrence=occurrence, paragraph=paragraph)

    def insert_text_before(
        self, anchor: str, text: str, occurrence: int | None = None, paragraph: str | None = None
    ) -> int:
        """Insert text before anchor with tracked changes.

        Args:
            anchor: Text to find as the anchor point
            text: Text to insert before the anchor
            occurrence: Which occurrence of anchor to use (0 = first,
                1 = second, etc.). Omitted → ``anchor`` must be unique in the
                search scope, else AmbiguousTextError.
            paragraph: Optional paragraph reference (e.g., "P2#f3c1") to scope the search

        Returns:
            The change ID of the insertion

        Raises:
            ValueError: If ``anchor`` is not a non-empty string, ``text`` is
                not a string, or ``occurrence`` is negative or not an integer
            TextNotFoundError: If the anchor text is not found or occurrence doesn't exist
            AmbiguousTextError: If ``occurrence`` is omitted and ``anchor``
                matches more than once in the search scope
            HashMismatchError: If the paragraph hash doesn't match
        """
        return self._insert_text(anchor, text, position="before", occurrence=occurrence, paragraph=paragraph)

    def _insert_text(
        self,
        anchor: str,
        text: str,
        position: Literal["before", "after"],
        occurrence: int | None = None,
        paragraph: str | None = None,
    ) -> int:
        """Insert text before or after anchor with tracked changes."""
        if not isinstance(text, str):
            raise ValueError(f"'text' must be a string (empty string is allowed), got {text!r}")
        _reject_control_chars(anchor, field="'anchor'", ctx=f"insert_{position}(): ", allow_newline=False)
        _reject_control_chars(text, field="'text'", ctx=f"insert_{position}(): ", allow_newline=True)
        with self._changeset(), self._grouped():
            if paragraph is not None:
                ref = ParagraphRef.parse(paragraph)
                p = self._resolve_paragraph(ref)
                match = self._locate_in_paragraph(p, paragraph, anchor, occurrence)
                return self._insert_near_match(match, text, position)

            match = self._locate_document_wide(anchor, occurrence)
            return self._insert_near_match(match, text, position)

    def _insert_near_match(self, match: TextMapMatch, text: str, position: Literal["before", "after"]) -> int:
        """Insert text before/after a match, splitting the edge w:t at the match boundary.

        A ``\\n`` in ``text`` means a tracked paragraph split — routed to
        :meth:`_split_insert`.
        """
        if "\n" in text:
            return self._split_insert(match, text, position)
        positions = match.positions
        if not positions:
            return -1

        if position == "after":
            edge = positions[-1]
            offset = edge.offset_in_node + 1
        else:
            edge = positions[0]
            offset = edge.offset_in_node

        run, rPr_xml = self._get_run_info(edge.node)
        if not run:
            return -1

        # Edge run inside <w:ins>: splice into our own insertion (no wrapper,
        # so no nested <w:ins>); a foreign author's insertion gets our own
        # sibling <w:ins>, splitting theirs when the anchor falls mid-content.
        ins_ancestor = self._find_ancestor(run, "w:ins")
        if ins_ancestor:
            if self._owns_ins(ins_ancestor):
                node_text = self._get_node_text(edge.node)
                self._set_node_text(edge.node, node_text[:offset] + text + node_text[offset:])
                _set_xml_space_preserve(edge.node)
                return -1
            return self._insert_own_ins_within_foreign_ins(ins_ancestor, edge.node, offset, text, rPr_xml)

        # Rebuild the edge run: split its w:t at the offset and wrap text in
        # <w:ins>; non-text children (w:tab, w:br, w:drawing, …) stay in place
        ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"

        def render_wt(wt) -> list[str]:
            fragments: list[str] = []
            if wt is edge.node:
                node_text = self._get_node_text(wt)
                before_text = node_text[:offset]
                after_text = node_text[offset:]
                if before_text:
                    fragments.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before_text)}</w:t></w:r>")
                fragments.append(ins_xml)
                if after_text:
                    fragments.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after_text)}</w:t></w:r>")
            else:
                # Preserve unmatched sibling w:t
                fragments.extend(render_plain_wt(wt, rPr_xml))
            return fragments

        nodes = self.editor.replace_node(run, "".join(rebuild_run_fragments(run, rPr_xml, render_wt)))

        for node in nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":
                return int(node.getAttribute("w:id"))
        return -1

    # ==================== Paragraph splits (\n) ====================

    def _ensure_splittable(self, p1) -> None:
        """Refuse to split a paragraph that ends a document section.

        A paragraph-level ``w:sectPr`` marks a section boundary; moving it (or
        duplicating it) across a split would silently corrupt the section
        structure. Uncommon — the document-final section mark lives on
        ``w:body``, not in a paragraph — so refuse cleanly rather than guess.
        """
        pPr = _first_child_element(p1, "w:pPr")
        if pPr is not None and _first_child_element(pPr, "w:sectPr") is not None:
            raise RevisionError(
                "Cannot split a paragraph that carries a section mark (w:sectPr) — the section "
                "boundary would be ambiguous. Edit around the section break instead."
            )

    def _reject_unsplittable_boundary(self, paragraph, text_map: TextMap, pos: int) -> None:
        """Refuse a split whose boundary would cut inside an existing revision.

        A split at visible position ``pos`` must fall on a run that is a *direct*
        child of ``paragraph``; a boundary inside a pre-existing
        ``<w:ins>``/``<w:del>``, hyperlink, or other inline container is not yet
        supported. Called up front — before any delete/insert — by the split
        dispatchers so a refused split never leaves a partial mutation (single
        edits have no DOM rollback), and again by ``_collect_tail_nodes`` as the
        backstop for the multi-op rewrite path. End-of-paragraph splits (empty
        tail) are always fine.
        """
        if pos >= len(text_map.text):
            return
        edge = text_map.positions[pos]
        run = self._find_ancestor(edge.node, "w:r")
        if run is None or run.parentNode is not paragraph:
            raise RevisionError(
                "Cannot split a paragraph at a point inside an existing revision, "
                "hyperlink, or other inline container (not yet supported)."
            )

    def _split_replace(self, match: TextMapMatch, replace_with: str) -> int:
        """Replace ``match`` with text containing ``\\n`` — a tracked split.

        Deletes the whole matched text, inserts the first segment where it was,
        then splits the paragraph once per ``\\n``, inserting each following
        segment at the start of the new paragraph. All revisions (del, the
        per-paragraph insertions, and each inserted paragraph mark) are created
        inside the caller's active ``_grouped``/``_changeset`` scope, so the
        whole split is one revision group and one changeset. Affix-trimming is
        deliberately skipped: a structural change reads clearer as a whole-find
        deletion plus segmented insertions.
        """
        segments = replace_with.split("\n")
        p1 = _ancestor_paragraph(match.positions[0].node)
        self._ensure_splittable(p1)
        # The first split lands where the match's tail begins (match.end, since
        # the match is deleted and segment 0 reinserted where it was). Reject an
        # unsplittable boundary now, before mutating (no single-op rollback).
        self._reject_unsplittable_boundary(p1, build_text_map(p1), match.end)
        start = match.start
        self._delete_across_nodes(match)
        if segments[0]:
            self._rewrite_insert_at(p1, build_text_map(p1), start, segments[0])
        return self._apply_paragraph_splits(p1, start + len(segments[0]), segments[1:])

    def _split_insert(self, match: TextMapMatch, text: str, position: Literal["before", "after"]) -> int:
        """Insert ``text`` (containing ``\\n``) near ``match`` — a tracked split.

        The first segment is inserted at the anchor boundary; each subsequent
        ``\\n`` splits the paragraph, its segment landing at the start of the
        new paragraph. One group, one changeset (see :meth:`_split_replace`).
        """
        segments = text.split("\n")
        p1 = _ancestor_paragraph(match.positions[0].node)
        self._ensure_splittable(p1)
        base = match.end if position == "after" else match.start
        # The first split lands at the anchor boundary (segment 0 is inserted
        # there, pushing the original content right). Reject an unsplittable
        # boundary now, before mutating (no single-op rollback).
        self._reject_unsplittable_boundary(p1, build_text_map(p1), base)
        if segments[0]:
            self._rewrite_insert_at(p1, build_text_map(p1), base, segments[0])
        return self._apply_paragraph_splits(p1, base + len(segments[0]), segments[1:])

    def _apply_paragraph_splits(self, p1, split_pos: int, segments: list[str]) -> int:
        """Split ``p1`` once per entry in ``segments``, threading the tail.

        ``split_pos`` is the visible-text position in the current paragraph
        where the first split falls; each segment is inserted at the start of
        the paragraph it opens. Returns the id of the last inserted paragraph
        mark (a member of the operation's group), so the caller's EditResult
        reaches the group.
        """
        member_id = -1
        current_p = p1
        pos = split_pos
        for segment in segments:
            new_p, mark_id = self._split_paragraph_at_position(current_p, pos)
            member_id = mark_id
            if segment:
                self._insert_ins_at_paragraph_start(new_p, segment)
            current_p = new_p
            pos = len(segment)
        return member_id

    def _split_paragraph_at_position(self, p1, pos: int):
        """Split ``p1`` at visible position ``pos``.

        Everything from ``pos`` onward moves into a new following paragraph
        (a copy of ``p1``'s properties, minus any section mark); ``p1``'s
        paragraph mark is flagged as an inserted revision. Returns the new
        paragraph element and the mark insertion's id.
        """
        tail = self._collect_tail_nodes(p1, pos)
        new_p = self._new_tail_paragraph(p1)
        for node in tail:
            new_p.appendChild(node)
        mark_id = self._flag_paragraph_mark_inserted(p1)
        return new_p, mark_id

    def _collect_tail_nodes(self, p1, pos: int) -> list:
        """Direct children of ``p1`` (in order) holding visible text from ``pos`` on.

        Splits the run at the boundary when ``pos`` falls mid-run. The
        paragraph properties (``w:pPr``) are never included. Raises
        RevisionError when the boundary sits inside an existing revision or
        other inline container (deferred — our own split flows always cut on a
        run that is a direct child of the paragraph).
        """
        text_map = build_text_map(p1)
        if pos >= len(text_map.text):
            return []
        self._reject_unsplittable_boundary(p1, text_map, pos)
        edge = text_map.positions[pos]
        run = self._find_ancestor(edge.node, "w:r")
        if run is None:  # pragma: no cover - guarded by _reject_unsplittable_boundary
            return []
        if edge.offset_in_node == 0:
            first_tail = run
        else:
            # Split the run at the offset; the tail starts at the right half.
            boundary = self._split_foreign_ins_at(edge.node, edge.offset_in_node)
            first_tail = _next_element_sibling(boundary.nextSibling) if boundary is not None else None
        tail: list = []
        node = first_tail
        while node is not None:
            nxt = node.nextSibling
            tail.append(node)
            node = nxt
        return tail

    def _new_tail_paragraph(self, p1):
        """Create and insert an empty following paragraph copying ``p1``'s pPr.

        The copy drops any section mark (``w:sectPr`` stays on the last
        paragraph only), pPr-change tracking, and any inherited paragraph-mark
        revision. The new ``w:p`` is stamped (paraId/rsids) via injection.
        """
        doc = self.editor.dom
        new_p = doc.createElement("w:p")
        orig_pPr = _first_child_element(p1, "w:pPr")
        if orig_pPr is not None:
            pPr_copy = orig_pPr.cloneNode(True)
            assert pPr_copy is not None  # cloneNode of an element returns an element
            # A section mark stays on the last paragraph only; pPr-change
            # tracking and any inherited mark revision do not belong on the copy.
            for tag in ("w:sectPr", "w:pPrChange"):
                child = _first_child_element(pPr_copy, tag)
                if child is not None:
                    pPr_copy.removeChild(child)
            rPr = _first_child_element(pPr_copy, "w:rPr")
            if rPr is not None:
                for tag in ("w:ins", "w:del"):
                    mark = _first_child_element(rPr, tag)
                    if mark is not None:
                        rPr.removeChild(mark)
            new_p.appendChild(pPr_copy)
        p1.parentNode.insertBefore(new_p, p1.nextSibling)
        self.editor._inject_attributes_to_nodes([new_p])
        return new_p

    def _flag_paragraph_mark_inserted(self, p1) -> int:
        """Flag ``p1``'s paragraph mark as an inserted revision.

        Adds an empty ``<w:ins>`` as the first child of the paragraph-mark
        ``<w:pPr><w:rPr>`` (created in schema order when absent). Injection
        stamps id/author/date and, inside an active ``_grouped`` scope, records
        it as a group member. Returns the mark insertion's id.
        """
        doc = self.editor.dom
        pPr = _first_child_element(p1, "w:pPr")
        if pPr is None:
            pPr = doc.createElement("w:pPr")
            p1.insertBefore(pPr, p1.firstChild)
        rPr = _first_child_element(pPr, "w:rPr")
        if rPr is None:
            rPr = doc.createElement("w:rPr")
            anchor = _first_child_element(pPr, "w:sectPr") or _first_child_element(pPr, "w:pPrChange")
            if anchor is not None:
                pPr.insertBefore(rPr, anchor)
            else:
                pPr.appendChild(rPr)
        ins = doc.createElement("w:ins")
        rPr.insertBefore(ins, rPr.firstChild)
        self.editor._inject_attributes_to_nodes([ins])
        mark_id = int(ins.getAttribute("w:id"))
        self._paragraph_mark_ids.add(mark_id)
        return mark_id

    def _insert_ins_at_paragraph_start(self, paragraph, text: str) -> None:
        """Insert ``text`` as a tracked insertion at the start of ``paragraph``.

        Lands right after ``w:pPr`` (before any moved tail content).
        """
        ins_xml = f"<w:ins><w:r><w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"
        pPr = _first_child_element(paragraph, "w:pPr")
        if pPr is not None:
            self.editor.insert_after(pPr, ins_xml)
            return
        first = _next_element_sibling(paragraph.firstChild)
        if first is not None:
            self.editor.insert_before(first, ins_xml)
        else:
            self.editor.append_to(paragraph, ins_xml)

    def list_revisions(
        self,
        author: str | None = None,
        paragraph: str | None = None,
        *,
        with_location: bool = True,
    ) -> list[Revision]:
        """List all tracked changes in the document.

        Args:
            author: If provided, filter by author name
            paragraph: If provided, a paragraph reference (e.g. "P3#a7b2")
                from list_paragraphs(); only revisions inside that paragraph
                are returned.
            with_location: If False, skip computing ``paragraph_ref`` and
                ``occurrence`` (they stay None) — the location work builds a
                text map and hash per revision-bearing paragraph, wasted on
                callers that only need ids (accept_all/reject_all re-list on
                every pass). Forced True when ``paragraph`` is given, since
                the filter matches on ``paragraph_ref``.

        Returns:
            List of Revision objects sorted by id — see :class:`Revision`.
            Nesting fields are always populated; the location fields
            (``paragraph_ref``/``occurrence``) unless ``with_location=False``.

        Raises:
            ValueError: If ``paragraph`` is malformed
            ParagraphIndexError: If the paragraph index is out of range
            HashMismatchError: If the paragraph hash doesn't match current content
        """
        paragraph_filter = None
        if paragraph is not None:
            ref = ParagraphRef.parse(paragraph)
            self._resolve_paragraph(ref)  # validates index and hash
            paragraph_filter = f"P{ref.index}#{ref.hash}"

        ctx = None
        if with_location or paragraph_filter is not None:
            ctx = _RevisionLocationContext(self.editor.dom)

        def matches(rev: Revision | None) -> bool:
            if rev is None:
                return False
            if author is not None and rev.author != author:
                return False
            return paragraph_filter is None or rev.paragraph_ref == paragraph_filter

        revisions = []

        # Find all insertions
        for ins_elem in self.editor.dom.getElementsByTagName("w:ins"):
            rev = self._parse_revision(ins_elem, "insertion", ctx)
            if matches(rev):
                revisions.append(rev)

        # Find all deletions
        for del_elem in self.editor.dom.getElementsByTagName("w:del"):
            rev = self._parse_revision(del_elem, "deletion", ctx)
            if matches(rev):
                revisions.append(rev)

        # Sort by ID
        revisions.sort(key=lambda r: r.id)
        return revisions

    def get_markup_text(self) -> str:
        """Render document text with inline revision markup.

        Each paragraph is one line; tracked changes wrap their content as
        ``[ins#{id}:{author}]...[/ins]`` / ``[del#{id}:{author}]...[/del]``,
        nesting included (e.g. ``[ins#1:A]kept [del#9:B]gone[/del][/ins]``).

        A human/agent verification view, not a parseable format: author
        names are not escaped, tabs/breaks are not rendered, and text inside
        a drawing's text box appears both inline in the host paragraph's
        line and again as its own line (same as get_text()).
        """

        def render(node) -> str:
            parts: list[str] = []
            for child in node.childNodes:
                if child.nodeType != child.ELEMENT_NODE:
                    continue
                if child.tagName in ("w:ins", "w:del"):
                    kind = "ins" if child.tagName == "w:ins" else "del"
                    rev_id = child.getAttribute("w:id") or "?"
                    rev_author = child.getAttribute("w:author") or "Unknown"
                    parts.append(f"[{kind}#{rev_id}:{rev_author}]{render(child)}[/{kind}]")
                elif child.tagName in ("w:t", "w:delText"):
                    parts.append(get_text_node_data(child))
                else:
                    parts.append(render(child))
            return "".join(parts)

        return "\n".join(render(p) for p in self.editor.dom.getElementsByTagName("w:p"))

    def _parse_revision(
        self,
        elem,
        rev_type: Literal["insertion", "deletion"],
        ctx: _RevisionLocationContext | None = None,
    ) -> Revision | None:
        """Parse a w:ins or w:del element into a Revision object.

        Args:
            elem: The <w:ins>/<w:del> element
            rev_type: Which kind of revision ``elem`` is
            ctx: Per-call location cache from list_revisions. None (detached
                elements, unit tests) leaves paragraph_ref/occurrence unset.
        """
        rev_id = elem.getAttribute("w:id")
        if not rev_id:
            return None
        try:
            rev_id_int = int(rev_id)
        except ValueError:
            # Nonconforming producer: Revision.id is an int and every
            # id-keyed operation targets ints, so a non-numeric w:id is
            # unrepresentable — omit it rather than crash the listing.
            return None

        author = elem.getAttribute("w:author") or "Unknown"
        date_str = elem.getAttribute("w:date")

        try:
            date = datetime.fromisoformat(date_str.replace("Z", "+00:00")) if date_str else None
        except ValueError:
            date = None

        # Extract text content
        if rev_type == "insertion":
            text_elems = _insertion_text_nodes(elem)
        else:
            text_elems = elem.getElementsByTagName("w:delText")
            if not text_elems:
                # Deliberate interop fallback: nonconforming producers may
                # leave plain w:t inside w:del. Fires only when the w:del has
                # no w:delText at all; mixed content reads only w:delText.
                text_elems = elem.getElementsByTagName("w:t")

        text = "".join(self._get_node_text(t_elem) for t_elem in text_elems)

        paragraph_ref = None
        occurrence = None
        if ctx is not None:
            paragraph = _ancestor_paragraph(elem)
            if paragraph is not None:
                paragraph_ref = ctx.paragraph_ref(paragraph)
                if text:
                    # Insertions live in the visible text; deletions in the
                    # original (pre-revision) text.
                    view: Literal["accepted", "original"] = "accepted" if rev_type == "insertion" else "original"
                    occurrence = _occurrence_in_text_map(ctx.text_map(paragraph, view), elem, text)

        group_id = self._revision_groups.get(rev_id_int)
        changeset_id = self.changeset_id_of(group_id) if group_id is not None else None
        return Revision(
            id=rev_id_int,
            type=rev_type,
            author=author,
            date=date,
            text=text,
            paragraph_ref=paragraph_ref,
            occurrence=occurrence,
            nested_under=_nearest_revision_ancestor_id(elem),
            contains_ids=_descendant_revision_ids(elem),
            group_id=group_id,
            group_source=self._group_sources.get(group_id) if group_id is not None else None,
            changeset_id=changeset_id,
            changeset_source=self._changeset_sources.get(changeset_id) if changeset_id is not None else None,
        )

    def _revision_element_index(self) -> dict[str, Element]:
        """Map ``w:id`` -> its <w:ins>/<w:del> element, one full-DOM walk per tag.

        Built once per group/changeset resolution and threaded through
        ``accept_revision``/``reject_revision`` so locating a member is an O(1)
        lookup instead of a fresh full-document scan (ISSUES.md #57).

        One element per id is exact for this path: a duplicate w:id never
        becomes a group member — ``_reconstruct_groups`` bars every duplicated
        id from every inferred group, and our own allocator keeps recorded ids
        unique — so a member id always maps to a single element. The
        w:ins-before-w:del insertion order only settles a tie-break group
        members never hit (it mirrors the fresh scan, which checks insertions
        first).
        """
        element_index: dict[str, Element] = {}
        for tag in ("w:ins", "w:del"):
            for elem in self.editor.dom.getElementsByTagName(tag):
                element_index.setdefault(elem.getAttribute("w:id"), elem)
        return element_index

    def _is_in_document(self, elem) -> bool:
        """True if ``elem`` is still attached to the live document tree.

        Accepting/rejecting a revision detaches its element (unwrap or
        removeChild), and a member nested inside an already-resolved member
        detaches together with its host. Either way the element is no longer
        reachable from the document root, so a snapshot lookup must treat it as
        gone — reproducing the "not found by getElementsByTagName" signal the
        fresh scan relied on for rump tolerance and no-double-count.
        """
        node = elem
        root = self.editor.dom
        while node is not None:
            if node is root:
                return True
            node = node.parentNode
        return False

    def _find_revision_element(self, revision_id: int, element_index: dict[str, Element] | None) -> Element | None:
        """Locate the live <w:ins>/<w:del> element for ``revision_id``.

        ``element_index is None`` scans the document fresh (insertions before
        deletions), matching the historical lookup exactly. Otherwise the id is
        resolved through the pre-built ``element_index`` and confirmed
        still-attached via ``_is_in_document``.
        """
        if element_index is None:
            for ins_elem in self.editor.dom.getElementsByTagName("w:ins"):
                if ins_elem.getAttribute("w:id") == str(revision_id):
                    return ins_elem
            for del_elem in self.editor.dom.getElementsByTagName("w:del"):
                if del_elem.getAttribute("w:id") == str(revision_id):
                    return del_elem
            return None
        elem = element_index.get(str(revision_id))
        if elem is None or not self._is_in_document(elem):
            return None
        return elem

    def accept_revision(self, revision_id: int, element_index: dict[str, Element] | None = None) -> bool:
        """Accept a revision by ID.

        For insertions: removes the w:ins wrapper, keeping the content.
        For deletions: removes the w:del element entirely.

        Args:
            revision_id: The w:id of the revision to accept
            element_index: Optional pre-built w:id -> element map (see
                ``_revision_element_index``) that lets group/changeset
                resolution skip a full-DOM scan per member. ``None`` scans
                fresh (standalone calls, accept_all/reject_all).

        Returns:
            True if revision was accepted, False if not found
        """
        elem = self._find_revision_element(revision_id, element_index)
        if elem is None:
            return False
        if elem.tagName == "w:ins":
            # Accept insertion: unwrap the content
            self._unwrap_element(elem)
        else:  # w:del
            # Accept deletion: remove the element entirely
            self._remove_element(elem)
        return True

    def reject_revision(self, revision_id: int, element_index: dict[str, Element] | None = None) -> bool:
        """Reject a revision by ID.

        For insertions: removes the w:ins element and its content entirely.
        For deletions: removes the w:del wrapper and converts w:delText back to w:t.

        Args:
            revision_id: The w:id of the revision to reject
            element_index: Optional pre-built w:id -> element map (see
                ``_revision_element_index``) that lets group/changeset
                resolution skip a full-DOM scan per member. ``None`` scans
                fresh (standalone calls, accept_all/reject_all).

        Returns:
            True if revision was rejected, False if not found
        """
        elem = self._find_revision_element(revision_id, element_index)
        if elem is None:
            return False
        if elem.tagName == "w:ins":
            if _is_paragraph_mark_ins(elem):
                # Reject a paragraph-mark insertion: remove the mark and rejoin
                # the tail paragraph (the inverse of the tracked split).
                self._rejoin_paragraph(elem)
            else:
                # Reject insertion: remove entirely
                self._remove_element(elem)
        else:  # w:del
            # Reject deletion: restore the deleted text
            self._restore_deletion(elem)
        return True

    def _resolve_ids(self, members: Iterable[int], resolve: Callable[[int, dict[str, Element] | None], bool]) -> int:
        """Apply ``resolve`` (accept/reject_revision) to every id in ``members``.

        Same reverse-id, loop-until-no-progress pattern as accept_all/
        reject_all, restricted to ``members``: nested members become
        resolvable once their host is processed, and members already resolved
        individually are simply skipped. Shared by group and changeset
        resolution (a changeset passes the union of its groups' revisions).

        The w:id -> element index is built once here and threaded through every
        ``resolve`` call, so resolution costs two full-DOM walks total instead
        of one scan per member per pass (ISSUES.md #57). The index stays valid
        across passes: accept/reject only ever *detach* elements, and
        ``_is_in_document`` (inside ``resolve``) treats a detached member as
        already gone.
        """
        members = list(members)
        element_index = self._revision_element_index()
        count = 0
        while True:
            progressed = False
            for rev_id in sorted(members, reverse=True):
                if resolve(rev_id, element_index):
                    count += 1
                    progressed = True
            if not progressed:
                return count

    def _resolve_group(self, group_id: int, resolve: Callable[[int, dict[str, Element] | None], bool]) -> int:
        """Apply ``resolve`` to every member revision of a group."""
        return self._resolve_ids(self.group_revisions(group_id), resolve)

    def _resolve_changeset(self, changeset_id: int, resolve: Callable[[int, dict[str, Element] | None], bool]) -> int:
        """Apply ``resolve`` to every revision across all groups of a changeset.

        A revision belongs to exactly one group, so the changeset's groups
        never share a member — the flattened list has no duplicates, and
        ``_resolve_ids`` would tolerate any anyway (a re-resolve returns False).
        """
        revision_ids = [rev_id for group_id in self.changeset_groups(changeset_id) for rev_id in self._groups[group_id]]
        return self._resolve_ids(revision_ids, resolve)

    def accept_group(self, group_id: int) -> int:
        """Accept every revision in a revision group.

        Args:
            group_id: Group id from an edit's :class:`EditResult` (or a
                Revision's ``group_id``).

        Returns:
            Number of revisions accepted. Members already resolved
            individually are skipped (and not counted).

        Raises:
            RevisionError: If the group id is unknown to this manager.
        """
        return self._resolve_group(group_id, self.accept_revision)

    def reject_group(self, group_id: int) -> int:
        """Reject every revision in a revision group.

        Args:
            group_id: Group id from an edit's :class:`EditResult` (or a
                Revision's ``group_id``).

        Returns:
            Number of revisions rejected. Members already resolved
            individually are skipped (and not counted).

        Raises:
            RevisionError: If the group id is unknown to this manager.
        """
        return self._resolve_group(group_id, self.reject_revision)

    def accept_changeset(self, changeset_id: int) -> int:
        """Accept every revision in a changeset (one whole call's groups).

        Args:
            changeset_id: Changeset id from an edit's :class:`EditResult` (or
                a Revision's ``changeset_id``).

        Returns:
            Number of revisions accepted across the changeset's groups.
            Members already resolved individually are skipped (rump-tolerant).

        Raises:
            RevisionError: If the changeset id is unknown to this manager.
        """
        return self._resolve_changeset(changeset_id, self.accept_revision)

    def reject_changeset(self, changeset_id: int) -> int:
        """Reject every revision in a changeset (one whole call's groups).

        Args:
            changeset_id: Changeset id from an edit's :class:`EditResult` (or
                a Revision's ``changeset_id``).

        Returns:
            Number of revisions rejected across the changeset's groups.
            Members already resolved individually are skipped (rump-tolerant).

        Raises:
            RevisionError: If the changeset id is unknown to this manager.
        """
        return self._resolve_changeset(changeset_id, self.reject_revision)

    def accept_all(self, author: str | None = None) -> int:
        """Accept all revisions, optionally filtered by author.

        Repeats passes until no listed revision can be processed, fully
        resolving nested revisions in Word-authored files (e.g. a w:del inside
        a w:ins) and terminating even when an author filter leaves other
        authors' revisions in the document. Revisions are matched by w:id, so
        if Word emits duplicate ids across authors, a filtered call may also
        process a same-id revision by another author.

        Args:
            author: If provided, only accept revisions by this author

        Returns:
            Number of revisions accepted
        """
        count = 0
        while True:
            progressed = False
            # Process in reverse order by ID to avoid index issues
            revisions = self.list_revisions(author=author, with_location=False)
            for rev in sorted(revisions, key=lambda r: r.id, reverse=True):
                if self.accept_revision(rev.id):
                    count += 1
                    progressed = True
            if not progressed:
                return count

    def reject_all(self, author: str | None = None) -> int:
        """Reject all revisions, optionally filtered by author.

        Repeats passes until no listed revision can be processed, fully
        resolving nested revisions in Word-authored files (e.g. a w:del inside
        a w:ins) and terminating even when an author filter leaves other
        authors' revisions in the document. Revisions are matched by w:id, so
        if Word emits duplicate ids across authors, a filtered call may also
        process a same-id revision by another author.

        Args:
            author: If provided, only reject revisions by this author

        Returns:
            Number of revisions rejected
        """
        count = 0
        while True:
            progressed = False
            # Process in reverse order by ID to avoid index issues
            revisions = self.list_revisions(author=author, with_location=False)
            for rev in sorted(revisions, key=lambda r: r.id, reverse=True):
                if self.reject_revision(rev.id):
                    count += 1
                    progressed = True
            if not progressed:
                return count

    def _unwrap_element(self, elem) -> None:
        """Remove an element's wrapper, keeping its children in place."""
        parent = elem.parentNode
        while elem.firstChild:
            child = elem.firstChild
            parent.insertBefore(child, elem)
        parent.removeChild(elem)

    def _remove_element(self, elem) -> None:
        """Detach an element (and its whole subtree) from its parent."""
        elem.parentNode.removeChild(elem)

    def _restore_deletion(self, del_elem) -> None:
        """Restore deleted content by converting w:delText back to w:t."""
        # Convert all w:delText to w:t
        for del_text in list(del_elem.getElementsByTagName("w:delText")):
            t_elem = self.editor.dom.createElement("w:t")
            # Copy content
            while del_text.firstChild:
                t_elem.appendChild(del_text.firstChild)
            # Copy attributes
            for i in range(del_text.attributes.length):
                attr = del_text.attributes.item(i)
                t_elem.setAttribute(attr.name, attr.value)
            del_text.parentNode.replaceChild(t_elem, del_text)

        # Update run attributes: w:rsidDel back to w:rsidR
        for run in del_elem.getElementsByTagName("w:r"):
            if run.hasAttribute("w:rsidDel"):
                run.setAttribute("w:rsidR", run.getAttribute("w:rsidDel"))
                run.removeAttribute("w:rsidDel")

        # Unwrap the w:del element
        self._unwrap_element(del_elem)

    def _rejoin_paragraph(self, mark_ins) -> None:
        """Reject a paragraph-mark insertion: drop the mark and merge the next
        paragraph's content back into this one — the inverse of a tracked split.

        The paragraph owning the mark survives (keeping its original
        properties, including any section mark); the following paragraph's
        content is appended and that paragraph is removed. Order-independent
        across a multi-split group: ``_resolve_ids`` dissolves the later (higher
        id) marks first, so an intermediate paragraph is never removed while it
        still owns a pending mark.
        """
        p1 = _ancestor_paragraph(mark_ins)
        rPr = mark_ins.parentNode
        rPr.removeChild(mark_ins)
        # Tidy the empty property wrappers the split created (keep any that
        # carried other properties).
        if not _next_element_sibling(rPr.firstChild):
            pPr = rPr.parentNode
            pPr.removeChild(rPr)
            if not _next_element_sibling(pPr.firstChild):
                pPr.parentNode.removeChild(pPr)
        if p1 is None:  # pragma: no cover - a mark-ins always sits in a paragraph
            return
        p2 = _next_element_sibling(p1.nextSibling)
        if p2 is None or getattr(p2, "tagName", "") != "w:p":
            return  # best effort: no following paragraph to rejoin into
        p2_pPr = _first_child_element(p2, "w:pPr")
        for child in list(p2.childNodes):
            if child is p2_pPr:
                continue
            p1.appendChild(child)
        p2.parentNode.removeChild(p2)


def _tokenize_words(text: str) -> list[str]:
    """Split text into alternating word and whitespace tokens."""
    return re.findall(r"\S+|\s+", text)


def _trim_replace_affixes(find: str, replace_with: str) -> tuple[int, int]:
    """Compute the character lengths of the word-level common prefix and
    suffix shared by ``find`` and ``replace_with``.

    Trimming is word-granular (tokens from :func:`_tokenize_words`) so a
    replace only revises whole changed words, matching the diff granularity
    of ``rewrite_paragraph``. The suffix scan is bounded by each side's
    remainder after the prefix, so a token shared at both ends is never
    consumed twice.

    Returns:
        ``(prefix_len, suffix_len)`` in characters.
    """
    f_toks = _tokenize_words(find)
    r_toks = _tokenize_words(replace_with)

    i = 0
    while i < len(f_toks) and i < len(r_toks) and f_toks[i] == r_toks[i]:
        i += 1
    j = 0
    while j < len(f_toks) - i and j < len(r_toks) - i and f_toks[-(j + 1)] == r_toks[-(j + 1)]:
        j += 1

    prefix_len = sum(len(tok) for tok in f_toks[:i])
    suffix_len = sum(len(tok) for tok in f_toks[len(f_toks) - j :])
    return prefix_len, suffix_len


def _set_xml_space_preserve(wt_elem) -> None:
    """Set xml:space='preserve' on a w:t element to preserve whitespace."""
    wt_elem.setAttribute("xml:space", "preserve")
