"""Revision listing: the ``RevisionManager`` mixin that reports revisions and renders markup text."""

from typing import Literal
from xml.dom.minidom import Element

from ..xml_editor import (
    _TEXTBOX_CONTENT,
    ParagraphRef,
    TextMap,
    body_paragraphs,
    build_text_map,
    compute_text_hash,
    get_text_node_data,
)
from .base import _RevisionManagerBase
from .dom import (
    _addressable_paragraph,
    _adjudicable_id,
    _deletion_text_nodes,
    _descendant_revision_ids,
    _insertion_text_nodes,
    _nearest_revision_ancestor_id,
    _occurrence_in_text_map,
    _parse_w_date,
)
from .models import (
    _MARKUP_KIND_BY_TAG,
    _REVISION_TYPE_BY_TAG,
    HANDLED_REVISION_TAGS,
    UNHANDLED_REVISION_TAGS,
    Revision,
    RevisionType,
    UnhandledRevision,
    iter_revision_elements,
)


class _RevisionLocationContext:
    """Per-``list_revisions``-call cache of paragraph indexes, refs, and text maps."""

    def __init__(self, dom):
        self._p_index = {id(p): i for i, p in enumerate(body_paragraphs(dom), start=1)}
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


class _ListingMixin(_RevisionManagerBase):
    def has_own_revisions(self) -> bool:
        """Whether the document currently holds a revision authored by us.

        Counts *pending* revisions only — accepting or rejecting one removes
        it from the DOM, so a document we edited and then fully accepted
        answers False, which is the honest answer: it carries no redline.
        Ownership is the same test the rest of this class uses, ``w:author``
        equal to the editor's author, so a foreign redline by someone whose
        author string happens to match ours reads as ours.

        Returns:
            True on the first ``w:ins``/``w:del`` we authored, False if there
            is none.
        """
        for tag in ("w:ins", "w:del"):
            for elem in self.editor.dom.getElementsByTagName(tag):
                if elem.getAttribute("w:author") == self.editor.author:
                    return True
        return False

    def list_revisions(
        self,
        author: str | None = None,
        paragraph: str | None = None,
        *,
        with_location: bool = True,
    ) -> list[Revision]:
        """List the document's revisions: insertions, deletions, move halves
        and paragraph-property changes.

        Walks ``HANDLED_REVISION_TAGS`` (``w:ins``, ``w:del``, ``w:moveFrom``,
        ``w:moveTo``, ``w:pPrChange``) in one recursive pass — every row is
        adjudicable by ``accept_revision``/``reject_revision``. Every other
        revision type in the OOXML schema (run/section/table property
        changes, table-structure revisions, custom-XML range marks) is
        invisible here and is listed instead by ``list_unhandled_revisions()``.
        A move's range marks are scaffolding: never listed, swept with their
        content.

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
        for elem in iter_revision_elements(self.editor.dom, HANDLED_REVISION_TAGS):
            rev = self._parse_revision(elem, _REVISION_TYPE_BY_TAG[elem.tagName], ctx)
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
        The two halves of a content move render the same way as
        ``[moveFrom#{id}:{author}]...[/moveFrom]`` /
        ``[moveTo#{id}:{author}]...[/moveTo]``. A ``w:pPrChange`` has no text
        and does not appear.

        A human/agent verification view, not a parseable format: author
        names are not escaped and tabs/breaks are not rendered (unlike
        ``get_visible_text``, where a tab mark is a ``\\t``). Text inside a
        drawing's text box does not appear at all — box content is excluded
        from every text view and from paragraph enumeration (same as
        get_visible_text()); see
        :func:`~docx_editor.xml_editor.body_paragraphs`. A revision whose only
        content is unrendered therefore shows as an empty pair of brackets
        (``[ins#11:R][/ins]`` for an insertion carrying nothing but a box) —
        the marker says a revision is there, not that it inserted nothing.
        """

        def render(node) -> str:
            parts: list[str] = []
            for child in node.childNodes:
                if child.nodeType != child.ELEMENT_NODE:
                    continue
                if child.tagName in _MARKUP_KIND_BY_TAG:
                    kind = _MARKUP_KIND_BY_TAG[child.tagName]
                    rev_id = child.getAttribute("w:id") or "?"
                    rev_author = child.getAttribute("w:author") or "Unknown"
                    parts.append(f"[{kind}#{rev_id}:{rev_author}]{render(child)}[/{kind}]")
                elif child.tagName in ("w:t", "w:delText"):
                    parts.append(get_text_node_data(child))
                elif child.tagName == _TEXTBOX_CONTENT:
                    continue  # box content belongs to the box, not this paragraph
                else:
                    parts.append(render(child))
            return "".join(parts)

        return "\n".join(render(p) for p in body_paragraphs(self.editor.dom))

    def _parse_revision(
        self,
        elem,
        rev_type: RevisionType,
        ctx: _RevisionLocationContext | None = None,
    ) -> Revision | None:
        """Parse a handled revision element into a Revision object.

        Only ``HANDLED_REVISION_TAGS`` are representable as a
        :class:`Revision`; the rest of the revision schema surfaces as
        :class:`UnhandledRevision` instead.

        Args:
            elem: The <w:ins>/<w:del>/<w:moveFrom>/<w:moveTo>/<w:pPrChange> element
            rev_type: Which kind of revision ``elem`` is
            ctx: Per-call location cache from list_revisions. None (detached
                elements, unit tests) leaves paragraph_ref/occurrence unset.
        """
        rev_id_int = _adjudicable_id(elem)
        if rev_id_int is None:
            # Nonconforming producer: unrepresentable here, reported by
            # ``list_unhandled_revisions`` instead (see ``_unhandled_elements``).
            return None

        author = elem.getAttribute("w:author") or "Unknown"
        date = _parse_w_date(elem)

        # Extract text content
        if rev_type == "property_change":
            # A change record holds the previous properties, never text.
            text_elems = []
        elif rev_type != "deletion":
            # Insertions and both move halves: Word writes plain w:t inside a
            # w:moveFrom (a hand-authored one may use w:delText); the shared
            # walk reads both, box-excluded.
            text_elems = _insertion_text_nodes(elem)
        else:
            text_elems = _deletion_text_nodes(elem)

        text = "".join(self._get_node_text(t_elem) for t_elem in text_elems)

        paragraph_ref = None
        occurrence = None
        if ctx is not None:
            paragraph = _addressable_paragraph(elem)
            if paragraph is not None:
                paragraph_ref = ctx.paragraph_ref(paragraph)
                if text and paragraph_ref is not None:
                    # An occurrence with no ref cannot be acted on (the
                    # paragraph is not addressable — e.g. it lives inside a
                    # text box), so half a location is worse than none.
                    # Insertions and moved-to text live in the visible text;
                    # deletions and moved-from text in the original
                    # (pre-revision) text.
                    view: Literal["accepted", "original"] = (
                        "accepted" if rev_type in ("insertion", "move_to") else "original"
                    )
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

    def _revision_element_index(self) -> dict[str, list[Element]]:
        """Map ``w:id`` -> its handled revision elements, in one recursive walk.

        ``HANDLED_REVISION_TAGS`` only: these are the revision elements
        ``accept_revision``/``reject_revision`` can act on, so indexing the
        rest of the schema would build lookups nothing could consume (the
        honesty floor reports them separately — see ``accept_all``). One
        ``iter_revision_elements`` pass rather than a ``getElementsByTagName``
        per tag: five tags would otherwise cost five full-DOM walks per
        group/changeset call (the ISSUES.md #57 pin).

        Built once per resolution call and threaded through
        ``accept_revision``/``reject_revision`` so locating a member is a dict
        lookup instead of a fresh full-document scan (ISSUES.md #57).

        Each id maps to a *list* because Word does not guarantee unique w:id:
        one reviewer's <w:ins> and another's <w:del> can share an id. Group and
        changeset members never collide (``_reconstruct_groups`` bars every
        duplicated id from every inferred group, and our own allocator keeps
        recorded ids unique), so for those callers every list holds exactly one
        element. Whole-document resolution (``_resolve_all``) is where the
        duplicates live: keeping them all lets one index serve every same-id
        element, so they resolve within a single pass instead of costing a
        rebuilt index and a whole extra pass each.

        Each id's list is in document order — the same order the fresh scan
        in ``_find_revision_element`` uses, so a duplicated id resolves the
        same element whichever path finds it.
        """
        element_index: dict[str, list[Element]] = {}
        for elem in iter_revision_elements(self.editor.dom, HANDLED_REVISION_TAGS):
            element_index.setdefault(elem.getAttribute("w:id"), []).append(elem)
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

    def _find_revision_element(
        self, revision_id: int, element_index: dict[str, list[Element]] | None
    ) -> Element | None:
        """Locate the live handled revision element for ``revision_id``.

        ``element_index is None`` scans the document fresh, returning the
        first match in document order. Otherwise the id is resolved through
        the pre-built ``element_index``, returning the first candidate still
        attached to the document.

        Returning the *first still-attached* candidate (rather than a single
        remembered element) is what lets duplicate ids resolve from one index:
        once an element is detached its successor becomes the answer, so N
        same-id revisions resolve in one pass rather than N.
        """
        if element_index is None:
            wanted = str(revision_id)
            for elem in iter_revision_elements(self.editor.dom, HANDLED_REVISION_TAGS):
                if elem.getAttribute("w:id") == wanted:
                    return elem
            return None
        for elem in element_index.get(str(revision_id), ()):
            if self._is_in_document(elem):
                return elem
        return None

    def _unhandled_elements(self, author: str | None = None) -> list[Element]:
        """Pending elements this library cannot resolve, in document order.

        Every ``UNHANDLED_REVISION_TAGS`` element, plus any
        ``HANDLED_REVISION_TAGS`` element with no numeric ``w:id``
        (``_adjudicable_id``): ``list_revisions`` omits such a mark because
        nothing id-keyed could target it, so without this it would vanish
        from both listings and ``accept_all`` would claim a clean document
        while the mark — and, for a ``w:moveFrom``, its hidden text — stays.

        Pending, not merely present: the recorded subtree of a change record is
        skipped (``skip_change_records``), so a ``w:cellIns`` inside a
        ``w:tcPrChange``'s historical ``w:tcPr`` is not counted as a second
        revision alongside the change itself.

        Author filtering follows ``list_revisions``: a missing ``w:author``
        reads as ``"Unknown"``, so an unattributed mark is matched only by
        ``author="Unknown"`` and excluded from every other filtered scan.
        """
        unhandled = frozenset(UNHANDLED_REVISION_TAGS)
        elems = [
            e
            for e in iter_revision_elements(
                self.editor.dom, UNHANDLED_REVISION_TAGS + HANDLED_REVISION_TAGS, skip_change_records=True
            )
            if e.tagName in unhandled or _adjudicable_id(e) is None
        ]
        if author is None:
            return elems
        return [e for e in elems if (e.getAttribute("w:author") or "Unknown") == author]

    def list_unhandled_revisions(self, author: str | None = None) -> list[UnhandledRevision]:
        """List the revision elements this library does not accept or reject.

        The complement of ``list_revisions``: everything in the OOXML revision
        schema except ``HANDLED_REVISION_TAGS`` and a move's range marks —
        run/section/table property changes, table-structure revisions,
        ``w:numberingChange``, custom-XML range marks (see
        ``UNHANDLED_REVISION_TAGS``). They survive open/edit/save unchanged and
        are left pending by ``accept_all``/``reject_all``. A handled-type mark
        with no numeric ``w:id`` is listed here too: ``list_revisions`` cannot
        represent it, and a mark that appears in neither listing would let
        ``accept_all`` claim a clean document it did not deliver.

        A separate method rather than extra rows from ``list_revisions``: these
        carry nothing ``accept_revision(rev.id)`` could act on, so mixing them
        in would break the loop idiom where every listed row is adjudicable.

        Args:
            author: If provided, filter by author name. Marks with no
                ``w:author`` read as ``"Unknown"``, so they match only
                ``author="Unknown"``.

        Returns:
            List of UnhandledRevision in document order — see
            :class:`UnhandledRevision`.
        """
        elems = self._unhandled_elements(author)
        if not elems:
            # The common case on an ins/del-only document. Building the
            # location context walks the whole DOM, and SKILL.md tells callers
            # to check this routinely, so do not pay for it to return [].
            return []
        ctx = _RevisionLocationContext(self.editor.dom)
        rows: list[UnhandledRevision] = []
        for elem in elems:
            paragraph = _addressable_paragraph(elem)
            rows.append(
                UnhandledRevision(
                    tag=elem.tagName,
                    id=_adjudicable_id(elem),
                    author=elem.getAttribute("w:author") or "Unknown",
                    date=_parse_w_date(elem),
                    paragraph_ref=ctx.paragraph_ref(paragraph) if paragraph is not None else None,
                )
            )
        return rows
