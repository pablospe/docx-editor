"""Free minidom helpers for revision-bearing elements."""

from datetime import datetime
from xml.dom.minidom import Element

from ..xml_editor import _TEXTBOX_CONTENT, TextMap, _inside_textbox, is_tab_node
from .models import _RUN_TRACK_CHANGE_TAGS, HANDLED_REVISION_TAGS, iter_revision_elements


def _parse_w_date(elem) -> datetime | None:
    """Parse an element's ``w:date``, or None when absent or unparseable.

    Word writes UTC as a trailing ``Z``, which ``fromisoformat`` only accepts
    from 3.11; the replacement keeps the 3.10 floor working. A nonconforming
    producer's unparseable stamp reads as None rather than failing the listing.
    """
    date_str = elem.getAttribute("w:date")
    if not date_str:
        return None
    try:
        return datetime.fromisoformat(date_str.replace("Z", "+00:00"))
    except ValueError:
        return None


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


def _addressable_paragraph(elem) -> Element | None:
    """The ``<w:p>`` a location may name for ``elem``, or None.

    ``_ancestor_paragraph`` climbs straight past ``w:txbxContent``, so a mark
    with no ``<w:p>`` of its own inside a text box — a ``w:trPr`` row marker
    or a ``w:tblPrChange`` in a box's table — would otherwise report the
    *host* paragraph's ref and be returned by a ``paragraph=`` filter on it,
    attributing the box's content to a paragraph whose text excludes it.
    """
    paragraph = _ancestor_paragraph(elem)
    if paragraph is None or _inside_textbox(elem, paragraph):
        return None
    return paragraph


def _adjudicable_id(elem) -> int | None:
    """``elem``'s ``w:id`` as an int, or None when it is missing or non-numeric.

    ``Revision.id`` is an int and every id-keyed operation targets ints, so a
    mark without a numeric id is unreachable by ``accept_revision``/
    ``reject_revision`` however handled its tag is. Nonconforming producers
    only; the one definition is shared by the listing (which omits such a
    mark) and the honesty floor (which reports it) so the two cannot drift.
    """
    raw_id = elem.getAttribute("w:id")
    if not raw_id:
        return None
    try:
        return int(raw_id)
    except ValueError:
        return None


def _nearest_revision_ancestor_id(elem) -> int | None:
    """id of the closest handled-revision ancestor with a numeric w:id, else None.

    An ancestor without one (``_adjudicable_id`` is None) is skipped, not
    fatal: it is reported by ``list_unhandled_revisions`` instead.
    """
    node = elem.parentNode
    while node is not None and node.nodeType == node.ELEMENT_NODE:
        if node.tagName in HANDLED_REVISION_TAGS:
            rev_id = _adjudicable_id(node)
            if rev_id is not None:
                return rev_id
        node = node.parentNode
    return None


def _descendant_revision_ids(elem) -> tuple[int, ...]:
    """ids of all handled-revision descendants of ``elem``, in document order.

    Descendants without a numeric w:id are skipped (see
    ``_nearest_revision_ancestor_id``).
    """
    ids: list[int] = []
    for child in iter_revision_elements(elem, HANDLED_REVISION_TAGS):
        rev_id = _adjudicable_id(child)
        if rev_id is not None:
            ids.append(rev_id)
    return tuple(ids)


def _revision_elements(root) -> list[Element]:
    """All handled-revision elements under ``root``, in document order.

    Unconditional recursion (``iter_revision_elements`` without
    ``skip_change_records``): nested revisions (e.g. a w:del inside a w:ins)
    are included, each appearing right after its host.
    """
    return list(iter_revision_elements(root, HANDLED_REVISION_TAGS))


def _insertion_text_nodes(elem) -> list:
    """All <w:t>/<w:delText> descendants and run-level <w:tab/> marks of a
    <w:ins>, <w:moveFrom> or <w:moveTo>, in document order.

    Including <w:delText> means a host insertion whose content was later
    deleted by a nested <w:del> still reports the full text it originally
    inserted (plain <w:delText> never appears under <w:ins> otherwise).
    Tabs are included so ``Revision.text`` spells the same ``"\\t"`` the text
    map does and the revision's ``occurrence`` keeps resolving. The same walk
    serves both move halves: Word writes plain <w:t> inside a <w:moveFrom>, a
    hand-authored one may use <w:delText>, and either way the moved text is
    what the row should report.

    Text inside a drawing's text box is skipped: an insertion wrapping a run
    that carries a box reports the text it inserted, not the box's content,
    which is excluded from every text view (see ``body_paragraphs``).
    """
    nodes: list = []

    def walk(node) -> None:
        for child in node.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            if child.tagName == _TEXTBOX_CONTENT:
                continue
            if child.tagName in ("w:t", "w:delText") or is_tab_node(child):
                nodes.append(child)
            else:
                walk(child)

    walk(elem)
    return nodes


def _deletion_text_nodes(elem) -> list:
    """All <w:delText> descendants and run-level <w:tab/> marks of a <w:del>,
    in document order.

    Box content is excluded the same way the insertion walk excludes it: a
    w:del wrapping a run that carries a drawing reports the body text it
    deleted, not the box's. Bounded by ``elem``, so a w:del living *inside* a
    box still reports its own text.

    Deliberate interop fallback: nonconforming producers may leave plain w:t
    inside w:del. Fires only when the w:del has no w:delText at all; mixed
    content reads only w:delText.
    """

    def collect(text_tag: str) -> list:
        return [
            n
            for n in elem.getElementsByTagName("*")
            if (n.tagName == text_tag or is_tab_node(n)) and not _inside_textbox(n, elem)
        ]

    nodes = collect("w:delText")
    if not any(n.tagName == "w:delText" for n in nodes):
        nodes = collect("w:t")
    return nodes


def _has_ancestor(node, ancestor) -> bool:
    """True if ``ancestor`` is ``node`` itself or one of its ancestors."""
    current = node
    while current is not None:
        if current is ancestor:
            return True
        current = current.parentNode
    return False


def _outermost_revision(elem: Element) -> Element:
    """``elem``, or the outermost run-level revision wrapper it is nested inside.

    Where a comment marker may be anchored. A marker placed *inside* another
    author's pending insertion (or the destination half of their move) is
    carried away when that revision is rejected, stranding its twin outside
    as an unpaired range marker; hoisting
    to the outermost revision keeps both markers in run-level content whatever
    anyone later does to the host. Our own group's members are never above
    ``elem`` (``group_spans`` spans outermost members only), so every revision
    ancestor found here is somebody else's.
    """
    outermost = elem
    node = elem.parentNode
    while isinstance(node, Element) and node.tagName in _RUN_TRACK_CHANGE_TAGS:
        outermost = node
        node = node.parentNode
    return outermost


def _node_depth(node) -> int:
    """Number of ancestors above ``node``, root included."""
    depth = 0
    current = node.parentNode
    while current is not None:
        depth += 1
        current = current.parentNode
    return depth


def _first_child_element(parent, tag: str) -> Element | None:
    """The first *direct* child of ``parent`` with tag name ``tag``, or None."""
    for child in parent.childNodes:
        if child.nodeType == child.ELEMENT_NODE and getattr(child, "tagName", "") == tag:
            return child
    return None


def _first_content_child(run) -> Element | None:
    """First element child of ``run`` that is not its ``w:rPr``."""
    for child in run.childNodes:
        if child.nodeType == child.ELEMENT_NODE and child.tagName != "w:rPr":
            return child
    return None  # pragma: no cover - callers pass the run holding a text-map node, so a content child exists


def _next_element_sibling(node) -> Element | None:
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


def _is_paragraph_mark_marker(elem) -> bool:
    """True if ``elem`` marks a paragraph mark: a child of ``w:pPr/w:rPr``.

    Word records a revision of the paragraph mark itself — an inserted
    (split), deleted (merge) or moved paragraph boundary — as an empty
    ``w:ins``/``w:del``/``w:moveFrom``/``w:moveTo`` in the paragraph's own
    run properties, rather than around any content.
    """
    parent = elem.parentNode
    if parent is None or getattr(parent, "tagName", "") != "w:rPr":
        return False
    grandparent = parent.parentNode
    return grandparent is not None and getattr(grandparent, "tagName", "") == "w:pPr"


def _is_paragraph_mark_ins(ins) -> bool:
    """True if ``ins`` is a paragraph-mark insertion (child of ``w:pPr/w:rPr``)."""
    return ins.tagName == "w:ins" and _is_paragraph_mark_marker(ins)


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


def _set_xml_space_preserve(wt_elem) -> None:
    """Set xml:space='preserve' on a w:t element to preserve whitespace."""
    wt_elem.setAttribute("xml:space", "preserve")
