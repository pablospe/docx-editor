"""Insert sites: the ``RevisionManager`` mixin that turns text insertions and paragraph splits into tracked changes."""

from typing import Literal
from xml.dom.minidom import Element

from ..exceptions import RevisionError
from ..xml_editor import (
    ParagraphRef,
    TextMap,
    TextMapMatch,
    _escape_xml,
    _reject_control_chars,
    build_text_map,
    get_rPr_xml,
    rebuild_run_fragments,
    render_plain_wt,
)
from .base import _RevisionManagerBase
from .dom import (
    _ancestor_paragraph,
    _first_child_element,
    _first_content_child,
    _next_element_sibling,
    _paragraph_mark_ins,
    _set_xml_space_preserve,
)
from .models import _RUN_TRACK_CHANGE_TAGS


class _InsertMixin(_RevisionManagerBase):
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
        _reject_control_chars(anchor, field="'anchor'", ctx=f"insert_{position}(): ", allow_tab=True)
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
            if not self._owns_ins(ins_ancestor):
                return self._insert_own_ins_within_foreign_ins(ins_ancestor, edge.node, offset, text, rPr_xml)
            if edge.is_tab:
                # A tab holds no text to splice into: a plain sibling run,
                # still inside our own <w:ins>, lands the text beside it.
                self._insert_into_run(run, rPr_xml, edge.node, offset, self._plain_run_xml(rPr_xml, text))
                return -1
            node_text = self._get_node_text(edge.node)
            self._set_node_text(edge.node, node_text[:offset] + text + node_text[offset:])
            _set_xml_space_preserve(edge.node)
            return -1

        # Rebuild the edge run: split its w:t at the offset (or land beside a
        # w:tab edge) and wrap text in <w:ins>; every other child stays in place
        ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"
        nodes = self._insert_into_run(run, rPr_xml, edge.node, offset, ins_xml)

        for node in nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":  # pragma: no branch
                return int(node.getAttribute("w:id"))
        return -1  # pragma: no cover - the fragment always yields a w:ins

    @staticmethod
    def _plain_run_xml(rPr_xml: str, text: str) -> str:
        """A bare run carrying ``text`` — for splicing beside a tab inside our
        own pending insertion, where a nested ``<w:ins>`` would be invalid."""
        return f"<w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r>"

    def _insert_into_run(self, run, rPr_xml: str, node, offset: int, fragment: str) -> list:
        """Rebuild ``run`` with ``fragment`` spliced in at (``node``, ``offset``).

        ``node`` is a direct ``w:t`` or ``w:tab`` child of ``run``. A ``w:t``
        is split around ``offset``; a tab is one character, so the fragment
        lands before it (``offset == 0``) or after it (``offset == 1``) and the
        ``<w:tab/>`` is re-emitted intact. Every other child keeps its place
        in document order. Returns the nodes that replaced ``run``.
        """

        def render_wt(wt) -> list[str]:
            if wt is not node:
                return render_plain_wt(wt, rPr_xml)
            node_text = self._get_node_text(wt)
            fragments: list[str] = []
            if node_text[:offset]:
                fragments.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(node_text[:offset])}</w:t></w:r>")
            fragments.append(fragment)
            if node_text[offset:]:
                fragments.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(node_text[offset:])}</w:t></w:r>")
            return fragments

        def render_other(child) -> list[str]:
            own = f"<w:r>{rPr_xml}{child.toxml()}</w:r>"
            if child is not node:
                return [own]
            return [fragment, own] if offset == 0 else [own, fragment]

        return self.editor.replace_node(run, "".join(rebuild_run_fragments(run, rPr_xml, render_wt, render_other)))

    # ==================== Paragraph splits (\n) ====================

    def _ensure_splittable(self, p1) -> None:
        """Refuse to split a paragraph that can't take a fresh tracked mark.

        Two cases refuse cleanly (before any mutation):

        - A paragraph-level ``w:sectPr`` marks a section boundary; moving it (or
          duplicating it) across a split would silently corrupt the section
          structure. Uncommon — the document-final section mark lives on
          ``w:body``, not in a paragraph.
        - The paragraph already carries a pending inserted paragraph mark. A
          second split would add a second ``<w:ins>`` to one ``<w:pPr><w:rPr>``
          (invalid OOXML — ``CT_ParaRPr`` allows one ``ins``) and mis-attribute
          the marks to the wrong breaks. A single multi-``\\n`` op never trips
          this (each split lands on a fresh tail paragraph); only re-splitting an
          already-split half in a separate op does.
        """
        pPr = _first_child_element(p1, "w:pPr")
        if pPr is not None and _first_child_element(pPr, "w:sectPr") is not None:
            raise RevisionError(
                "Cannot split a paragraph that carries a section mark (w:sectPr) — the section "
                "boundary would be ambiguous. Edit around the section break instead."
            )
        if _paragraph_mark_ins(p1) is not None:
            raise RevisionError(
                "Cannot split a paragraph that already has a pending inserted paragraph mark — "
                "accept or reject that split before splitting the same paragraph again."
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
        tail) are always fine. A boundary on a ``<w:tab/>`` needs no extra
        check here: ``_collect_tail_nodes`` splits the run around the tab.
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
        # Formatting to fall back on when the new paragraph's tail is empty
        # (a split at the end of the current paragraph): the trailing run of the
        # paragraph being split carries the boundary formatting. A paragraph left
        # runless by an empty segment carries no formatting of its own, so the
        # last known boundary rPr is kept rather than reset — that propagates
        # through a run of empty tails (e.g. appending "A\n\nC" past the last
        # word), so every segment keeps the surrounding format instead of only
        # the first. A paragraph that *has* runs is a genuine boundary and
        # replaces the fallback, even when its last run carries no rPr.
        fallback_rPr = ""
        for segment in segments:
            boundary_runs = current_p.getElementsByTagName("w:r")
            if boundary_runs:
                fallback_rPr = get_rPr_xml(boundary_runs[-1])
            new_p, mark_id = self._split_paragraph_at_position(current_p, pos)
            member_id = mark_id
            if segment:
                self._insert_ins_at_paragraph_start(new_p, segment, fallback_rPr)
            current_p = new_p
            pos = len(segment)
        return member_id

    def _split_paragraph_at_position(self, p1, pos: int) -> tuple[Element, int]:
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

        Splits the run at the boundary when ``pos`` falls mid-run — mid-node,
        or at the start of a node that is not the run's first content child
        (``<w:t>foo</w:t><w:tab/><w:t>bar</w:t>`` split at ``bar`` keeps
        ``foo`` and the tab in the head; any leading child — a second
        ``w:t``, a ``w:tab``, a ``w:br``, a drawing — stays on the side it
        precedes). The paragraph properties (``w:pPr``)
        are never included. Raises RevisionError when the boundary sits inside
        an existing revision or other inline container (deferred — our own
        split flows always cut on a run that is a direct child of the
        paragraph).
        """
        text_map = build_text_map(p1)
        if pos >= len(text_map.text):
            return []
        self._reject_unsplittable_boundary(p1, text_map, pos)
        edge = text_map.positions[pos]
        run = self._find_ancestor(edge.node, "w:r")
        if run is None:  # pragma: no cover - guarded by _reject_unsplittable_boundary
            return []
        if edge.offset_in_node == 0 and _first_content_child(run) is edge.node:
            first_tail = run
        else:
            # Split the run at the offset; the tail starts at the right half.
            # A None boundary means nothing preceded the split point, so the
            # (unchanged) run itself opens the tail.
            boundary = self._split_foreign_ins_at(edge.node, edge.offset_in_node)
            first_tail = _next_element_sibling(boundary.nextSibling) if boundary is not None else run
        tail: list = []
        node = first_tail
        while node is not None:
            nxt = node.nextSibling
            tail.append(node)
            node = nxt
        return tail

    def _new_tail_paragraph(self, p1) -> Element:
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
                for tag in _RUN_TRACK_CHANGE_TAGS:
                    mark = _first_child_element(rPr, tag)
                    if mark is not None:
                        rPr.removeChild(mark)
            new_p.appendChild(pPr_copy)
        p1.parentNode.insertBefore(new_p, p1.nextSibling)
        self.editor._inject_attributes_to_nodes([new_p])
        return new_p

    def _paragraph_mark_rPr(self, p1) -> Element:
        """Return ``p1``'s paragraph-mark ``<w:pPr><w:rPr>``, creating both in
        schema order (``w:rPr`` before any ``w:sectPr``/``w:pPrChange``) if absent.
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
        return rPr

    def _flag_paragraph_mark_inserted(self, p1) -> int:
        """Flag ``p1``'s paragraph mark as an inserted revision.

        Adds an empty ``<w:ins>`` as the first child of the paragraph-mark
        ``<w:pPr><w:rPr>`` (created in schema order when absent). Injection
        stamps id/author/date and, inside an active ``_grouped`` scope, records
        it as a group member. Returns the mark insertion's id.
        """
        rPr = self._paragraph_mark_rPr(p1)
        ins = self.editor.dom.createElement("w:ins")
        rPr.insertBefore(ins, rPr.firstChild)
        self.editor._inject_attributes_to_nodes([ins])
        mark_id = int(ins.getAttribute("w:id"))
        self._paragraph_mark_ids.add(mark_id)
        return mark_id

    def _insert_ins_at_paragraph_start(self, paragraph, text: str, fallback_rPr_xml: str = "") -> None:
        """Insert ``text`` as a tracked insertion at the start of ``paragraph``.

        Lands right after ``w:pPr`` (before any moved tail content). The run
        inherits the formatting (rPr) of the tail it sits directly before — the
        moved boundary run — so a split-inserted segment matches the surrounding
        text instead of dropping to document default. When the tail is empty (a
        split at the paragraph's end), ``fallback_rPr_xml`` (the last boundary
        formatting the split loop saw — several paragraphs back when empty
        segments intervene) is used instead.
        """
        runs = paragraph.getElementsByTagName("w:r")
        rPr_xml = get_rPr_xml(runs[0]) if runs else fallback_rPr_xml
        ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"
        pPr = _first_child_element(paragraph, "w:pPr")
        if pPr is not None:
            self.editor.insert_after(pPr, ins_xml)
            return
        first = _next_element_sibling(paragraph.firstChild)
        if first is not None:
            self.editor.insert_before(first, ins_xml)
        else:
            self.editor.append_to(paragraph, ins_xml)
