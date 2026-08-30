"""Revision resolution: the ``RevisionManager`` mixin that accepts and rejects revisions."""

import warnings
from collections.abc import Callable, Iterable, Iterator
from contextlib import contextmanager
from xml.dom.minidom import Element

from ..exceptions import UnhandledRevisionWarning
from .base import _RevisionManagerBase
from .dom import (
    _ancestor_paragraph,
    _first_child_element,
    _is_paragraph_mark_ins,
    _is_paragraph_mark_marker,
    _next_element_sibling,
    _paragraph_mark_ins,
)
from .models import (
    MOVE_RANGE_TAGS,
    ResolveResult,
    iter_revision_elements,
)


class _ResolutionMixin(_RevisionManagerBase):
    def accept_revision(self, revision_id: int, element_index: dict[str, list[Element]] | None = None) -> bool:
        """Accept a revision by ID.

        - insertion (``w:ins``): removes the wrapper, keeping the content.
        - deletion (``w:del``): removes the element entirely.
        - move_to (``w:moveTo``): removes the wrapper, keeping the moved text
          at its destination.
        - move_from (``w:moveFrom``): removes the element — the text leaves
          its source. The two halves of a move are independent rows; resolve
          them together (``accept_all``, or the inferred changeset both halves
          share) so the text neither doubles nor vanishes.
        - property_change (``w:pPrChange``): removes the record, keeping the
          paragraph's current properties.

        A paragraph-mark move marker (``w:pPr/w:rPr/w:moveFrom``/``w:moveTo``)
        is dropped, the same approximate treatment a deleted paragraph mark
        gets (see ``accept_all``). After every resolution, range marks whose
        content is all gone are swept (``_sweep_move_range_marks``) — move
        content can leave inside any resolved host, not only a move half.

        Args:
            revision_id: The w:id of the revision to accept
            element_index: Optional pre-built w:id -> element map (see
                ``_revision_element_index``) that lets group/changeset
                resolution skip a full-DOM scan per member. ``None`` scans
                fresh (standalone calls); accept_all/reject_all pass one too.

        Returns:
            True if revision was accepted, False if not found
        """
        elem = self._find_revision_element(revision_id, element_index)
        if elem is None:
            return False
        tag = elem.tagName
        if tag == "w:ins":
            # Accept insertion: unwrap the content
            self._unwrap_element(elem)
        elif tag in ("w:del", "w:pPrChange"):
            # Accept deletion: remove the element entirely. Accept a property
            # change: drop the record of the previous properties.
            self._remove_element(elem)
        elif tag == "w:moveTo" and not _is_paragraph_mark_marker(elem):
            # Accept the destination half: the moved text stays.
            self._unwrap_element(elem)
        else:  # w:moveFrom, or a paragraph-mark marker of either half
            # Accept the source half: the moved-away text goes.
            self._remove_element(elem)
        # Unconditional: move content can leave with any resolved host (a
        # w:del wrapping a w:moveFrom, a rejoined paragraph's mark marker),
        # not only when a move half is the element resolved.
        self._sweep_after_move()
        return True

    def reject_revision(self, revision_id: int, element_index: dict[str, list[Element]] | None = None) -> bool:
        """Reject a revision by ID.

        - insertion (``w:ins``): removes the element and its content entirely
          (a paragraph-mark insertion rejoins the split paragraph).
        - deletion (``w:del``): removes the wrapper and converts ``w:delText``
          back to ``w:t``.
        - move_to (``w:moveTo``): removes the element — the text leaves its
          destination, and with it any revision nested inside. Unlike a
          foreign ``w:ins``, a foreign ``w:moveTo`` is not split around this
          session's own edits (its text is plain visible text to the editor),
          so an own edit made inside one is carried away by its rejection —
          a known gap.
        - move_from (``w:moveFrom``): restores the text at its source exactly
          as a rejected deletion is (``w:delText`` -> ``w:t``, ``w:rsidDel``
          -> ``w:rsidR``, unwrap). Word writes plain ``w:t`` inside a
          ``w:moveFrom``, which the same path leaves as is. Resolve both
          halves together — see ``accept_revision``.
        - property_change (``w:pPrChange``): puts the recorded previous
          properties back (``_restore_paragraph_properties``).

        A paragraph-mark move marker is dropped, as on accept.

        Args:
            revision_id: The w:id of the revision to reject
            element_index: Optional pre-built w:id -> element map (see
                ``_revision_element_index``) that lets group/changeset
                resolution skip a full-DOM scan per member. ``None`` scans
                fresh (standalone calls); accept_all/reject_all pass one too.

        Returns:
            True if revision was rejected, False if not found
        """
        elem = self._find_revision_element(revision_id, element_index)
        if elem is None:
            return False
        tag = elem.tagName
        if tag == "w:ins":
            if _is_paragraph_mark_ins(elem):
                # Reject a paragraph-mark insertion: remove the mark and rejoin
                # the tail paragraph (the inverse of the tracked split).
                self._rejoin_paragraph(elem)
            else:
                # Reject insertion: remove entirely
                self._remove_element(elem)
        elif tag == "w:del":
            # Reject deletion: restore the deleted text
            self._restore_deletion(elem)
        elif tag == "w:pPrChange":
            self._restore_paragraph_properties(elem)
        elif tag == "w:moveFrom" and not _is_paragraph_mark_marker(elem):
            # Reject the source half: the text stays where it was.
            self._restore_deletion(elem)
        else:  # w:moveTo, or a paragraph-mark marker of either half
            # Reject the destination half: the moved-in text goes.
            self._remove_element(elem)
        self._sweep_after_move()  # unconditional — see accept_revision
        return True

    def _resolve_ids(
        self, members: Iterable[int], resolve: Callable[[int, dict[str, list[Element]] | None], bool]
    ) -> int:
        """Apply ``resolve`` (accept/reject_revision) to every id in ``members``.

        Reverse-id, loop-until-no-progress pattern: nested members become
        resolvable once their host is processed, and members already resolved
        individually are simply skipped. Shared by group resolution and
        changeset resolution (a changeset passes the union of its groups'
        revisions).

        Members are a fixed set with unique ids (the allocator guarantees it),
        so the w:id -> element index is built once here and threaded through
        every ``resolve`` call: resolution costs two full-DOM walks instead of
        one scan per member per pass (ISSUES.md #57). The index stays valid
        across passes because accept/reject only ever *detach* elements, and
        ``_is_in_document`` (inside ``resolve``) treats a detached member as
        already gone.

        Termination rests on ``resolve`` returning False once an id has no live
        element left. accept_all/reject_all cannot use this loop — they resolve
        every id in the document, where Word may repeat a w:id across authors;
        see ``_resolve_all``.
        """
        members = list(members)
        element_index = self._revision_element_index()
        count = 0
        with self._deferred_range_sweep():
            while True:
                progressed = False
                for rev_id in sorted(members, reverse=True):
                    if resolve(rev_id, element_index):
                        count += 1
                        progressed = True
                if not progressed:
                    return count

    def _resolve_group(self, group_id: int, resolve: Callable[[int, dict[str, list[Element]] | None], bool]) -> int:
        """Apply ``resolve`` to every member revision of a group."""
        return self._resolve_ids(self.group_revisions(group_id), resolve)

    def _resolve_changeset(
        self, changeset_id: int, resolve: Callable[[int, dict[str, list[Element]] | None], bool]
    ) -> int:
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

    def _resolve_all(self, author: str | None, resolve: Callable[[int, dict[str, list[Element]] | None], bool]) -> int:
        """Apply ``resolve`` to every listed revision, re-listing on each pass.

        "Every listed revision" is every ``HANDLED_REVISION_TAGS`` element,
        because that is what ``list_revisions`` walks — so the loop terminates
        as soon as none remains, regardless of what other revision types are
        still pending. ``_resolve_all_reporting`` is what turns that into an
        honest count for callers.

        The whole-document counterpart to ``_resolve_ids``. Two differences,
        both forced by resolving *every* revision rather than a known member
        set:

        * **Re-listing each pass is what terminates the loop.** The listing
          shrinks as revisions are resolved, so an empty (or no-progress) pass
          ends it. ``_resolve_ids`` can instead lean on ``resolve`` returning
          False, because its members are a fixed set of unique ids. Relying on
          that here would not terminate: a whole-document listing can hand the
          same id back repeatedly.
        * **The index maps each id to every element carrying it.** Word may
          reuse one w:id across authors (A's w:ins and B's w:del both id=7);
          ``_find_revision_element`` returns the first still-attached
          candidate, so same-id revisions resolve within one pass.

        The index is built once for the whole call, not per pass: resolution
        only ever detaches elements, so a single build stays a valid superset
        and ``_is_in_document`` filters what is gone. Each revision then costs
        a dict lookup instead of a fresh full-document scan (ISSUES.md #56).
        Per pass this is a constant number of full-DOM walks — the work
        itself is still linear in the number of revisions listed.
        """
        count = 0
        element_index = self._revision_element_index()
        with self._deferred_range_sweep():
            while True:
                revisions = self.list_revisions(author=author, with_location=False)
                if not revisions:
                    return count
                progressed = False
                for rev in sorted(revisions, key=lambda r: r.id, reverse=True):
                    if resolve(rev.id, element_index):
                        count += 1
                        progressed = True
                if not progressed:
                    return count

    def _resolve_all_reporting(
        self,
        author: str | None,
        resolve: Callable[[int, dict[str, list[Element]] | None], bool],
        verb: str,
    ) -> ResolveResult:
        """``_resolve_all`` plus the honesty floor: what is still pending after it.

        The census is taken *after* resolution, so a foreign mark carried away
        inside a rejected insertion's subtree correctly does not appear, and
        the number always answers "what did this call leave behind".

        Move range marks are swept afterwards whatever was resolved, so a
        document holding only stray marks (no move content at all) still
        comes out clean.
        """
        count = self._resolve_all(author, resolve)
        self._sweep_after_move()
        unhandled_types: dict[str, int] = {}
        for elem in self._unhandled_elements(author):
            unhandled_types[elem.tagName] = unhandled_types.get(elem.tagName, 0) + 1
        if unhandled_types:
            total = sum(unhandled_types.values())
            listing = ", ".join(f"{tag} x{n}" for tag, n in sorted(unhandled_types.items()))
            scope = f" (author={author!r})" if author is not None else ""
            warnings.warn(
                f"{verb}{scope} resolved {count} revision(s) but left {total} unresolved: {listing}. "
                f"This library resolves insertions, deletions, moves and paragraph-property "
                f"changes only; inspect the rest with list_unhandled_revisions() before "
                f"reporting the document as fully adjudicated.",
                UnhandledRevisionWarning,
                # 4 frames out of warnings.warn: _resolve_all_reporting ->
                # RevisionManager.accept_all -> Document.accept_all -> caller.
                # Tuned for the supported Document path; calling RevisionManager
                # directly makes the chain one frame shorter.
                stacklevel=4,
            )
        return ResolveResult(count, unhandled_types)

    def accept_all(self, author: str | None = None) -> ResolveResult:
        """Accept every listed revision, optionally filtered by author.

        Walks ``HANDLED_REVISION_TAGS``: insertions, deletions, both halves of
        a content move, and paragraph-property changes. Delegates to
        ``_resolve_all``, which re-lists on each pass: that fully resolves
        *nested* revisions in Word-authored files (e.g. a w:del inside a
        w:ins) and terminates even when an author filter leaves other authors'
        revisions in the document. Revisions are matched by w:id, so if Word
        emits duplicate ids across authors, a filtered call may also process a
        same-id revision by another author.

        A move is resolved as a unit: its ``w:moveFrom`` content is removed and
        its ``w:moveTo`` content unwrapped in the same call, and the range
        marks bracketing them are swept once empty — the text ends up at its
        destination exactly once. A ``w:pPrChange`` record is dropped, keeping
        the paragraph's current properties.

        Every other revision type in the OOXML schema — run/section/table
        property changes, table-structure revisions, ``w:numberingChange``,
        custom-XML range marks — is left untouched and counted on the result
        (``.unhandled`` / ``.unhandled_types``, listed by
        ``list_unhandled_revisions()``), with an ``UnhandledRevisionWarning``
        when the count is nonzero.

        Paragraph-mark and row markers resolve *approximately* and are not
        part of that count, because the marker itself is consumed: a deleted
        or moved paragraph mark (``w:pPr/w:rPr/w:del``, ``.../w:moveFrom``,
        ``.../w:moveTo``) should merge or split paragraphs, and ``w:trPr`` row
        markers should add or drop the table row — today only the marker is
        removed. So accepting a moved *table* leaves its source behind as a
        table of empty cells (the moved text is at the destination exactly
        once; the rows are not). A document's exposure is visible via
        ``docx_editor.track_changes.count_revision_elements`` (``ins_del_contexts``).

        A listed revision can also leave the document inside another one's
        resolution rather than through its own — a ``w:pPrChange`` on the tail
        paragraph of a rejected tracked split goes with that paragraph's
        ``w:pPr`` (see ``_rejoin_paragraph``) — and is then counted neither
        as resolved nor as unhandled.

        Args:
            author: If provided, only accept revisions by this author

        Returns:
            :class:`ResolveResult` — an int carrying the number of revisions
            accepted, plus ``.unhandled``/``.unhandled_types``.
        """
        return self._resolve_all_reporting(author, self.accept_revision, "accept_all()")

    def reject_all(self, author: str | None = None) -> ResolveResult:
        """Reject every listed revision, optionally filtered by author.

        Walks ``HANDLED_REVISION_TAGS``: insertions, deletions, both halves of
        a content move, and paragraph-property changes. Delegates to
        ``_resolve_all``, which re-lists on each pass: that fully resolves
        *nested* revisions in Word-authored files (e.g. a w:del inside a
        w:ins) and terminates even when an author filter leaves other authors'
        revisions in the document. Revisions are matched by w:id, so if Word
        emits duplicate ids across authors, a filtered call may also process a
        same-id revision by another author.

        A move is undone as a unit: its ``w:moveTo`` content is removed and
        its ``w:moveFrom`` content restored in place, range marks swept — the
        text is back at its source exactly once. A ``w:pPrChange`` puts the
        recorded previous paragraph properties back.

        Every other revision type is left untouched and reported exactly as in
        ``accept_all`` — see that docstring for the full scope, including the
        structural marker cases that resolve approximately.

        Args:
            author: If provided, only reject revisions by this author

        Returns:
            :class:`ResolveResult` — an int carrying the number of revisions
            rejected, plus ``.unhandled``/``.unhandled_types``.
        """
        return self._resolve_all_reporting(author, self.reject_revision, "reject_all()")

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

    def _restore_paragraph_properties(self, ppr_change) -> None:
        """Reject a ``w:pPrChange``: put the recorded previous properties back.

        The record's ``w:pPr`` child (CT_PPrBase) holds the properties the
        paragraph had before the change. The live ``w:pPr``'s base children —
        everything except ``w:rPr``, ``w:sectPr`` and the record itself — are
        replaced by the recorded ones, inserted ahead of ``w:rPr``/``w:sectPr``
        so schema order holds even for producers that wrote the record before
        ``w:rPr`` (LibreOffice does). A record with no ``w:pPr`` child, which
        LibreOffice writes for "previously no properties", restores exactly
        that: the base children are cleared. The recorded style id is put back
        verbatim even when ``styles.xml`` lacks it — that is what the file says
        the paragraph had (Word falls back to Normal for an unknown id).
        """
        ppr = ppr_change.parentNode
        if getattr(ppr, "tagName", "") != "w:pPr":
            # Schema-invalid placement: nothing to restore into. Drop the record
            # rather than rewrite whatever parent this is.
            self._remove_element(ppr_change)
            return
        # The live w:pPr's tail — its paragraph-mark w:rPr (which may itself
        # carry a pending mark revision), section mark and the record — is
        # kept; everything before it is the base being replaced. The same
        # tags are skipped on the recorded side because CT_PPrBase cannot
        # hold them: a nonconforming record's copies are dropped, not
        # restored over the live ones.
        tail_tags = ("w:rPr", "w:sectPr", "w:pPrChange")
        for child in list(ppr.childNodes):
            if child.nodeType == child.ELEMENT_NODE and child.tagName not in tail_tags:
                ppr.removeChild(child)
        anchor = next(
            (c for c in ppr.childNodes if c.nodeType == c.ELEMENT_NODE and c.tagName in tail_tags),
            None,
        )
        recorded = _first_child_element(ppr_change, "w:pPr")
        if recorded is not None:
            # Same document, so moving the nodes is enough — insertBefore
            # detaches them from the record first.
            for child in list(recorded.childNodes):
                if child.nodeType == child.ELEMENT_NODE and getattr(child, "tagName", "") not in tail_tags:
                    ppr.insertBefore(child, anchor)
        self._remove_element(ppr_change)

    def _sweep_after_move(self) -> None:
        """Sweep range marks now, or note that a deferred bulk sweep is owed.

        Nothing to do — no walk at all — unless the document holds a move
        range mark (``DocxXMLEditor.holds_move_range_marks``): the sweep only
        ever removes those, and nearly every document has none.
        """
        if not self.editor.holds_move_range_marks:
            return
        if self._defer_range_sweep:
            self._range_sweep_pending = True
        else:
            self._sweep_move_range_marks()

    @contextmanager
    def _deferred_range_sweep(self) -> Iterator[None]:
        """Defer ``_sweep_after_move`` to one sweep when the block exits.

        The sweep is a full-document walk; ``_resolve_all``/``_resolve_ids``
        resolve many move halves per call and would otherwise walk the
        document once per half — quadratic in the number of moves. Range
        marks are only ever *removed* by the sweep, so sweeping once at the
        end sees exactly the same content state as sweeping after each half.
        """
        previous = (self._defer_range_sweep, self._range_sweep_pending)
        self._defer_range_sweep, self._range_sweep_pending = True, False
        try:
            yield
        finally:
            pending = self._range_sweep_pending
            self._defer_range_sweep, self._range_sweep_pending = previous
            if pending:
                self._sweep_after_move()

    def _sweep_move_range_marks(self) -> None:
        """Remove move range marks whose content has all been resolved.

        One recursive walk in document order over the four range-mark tags
        plus ``w:moveFrom``/``w:moveTo``. Per family (From/To), a
        ``*RangeStart`` is paired with the ``*RangeEnd`` carrying the same
        ``w:id``; the pair is removed when no pending content of its family
        lies between them. Marks left unpaired (a Start with no End or the
        reverse — a damaged file, or an End that precedes its Start), marks
        with no ``w:id`` at all, and a Start whose id a later Start reuses are
        removed once no pending content of their family remains anywhere in
        the document. Never destroys content: only the empty marks go.

        Called after every move resolution (once per bulk call — see
        ``_deferred_range_sweep``) and at the end of ``accept_all``/
        ``reject_all``. Word writes the From-range's End at body
        level for a moved table (outside any paragraph), which is why this is
        a document walk rather than a paragraph-local one.
        """
        # family -> open Start marks by w:id -> [element, content seen inside]
        open_starts: dict[str, dict[str, list]] = {"From": {}, "To": {}}
        unpaired: dict[str, list[Element]] = {"From": [], "To": []}
        content_seen: dict[str, bool] = {"From": False, "To": False}
        tags = MOVE_RANGE_TAGS + ("w:moveFrom", "w:moveTo")
        marks_seen = 0
        removed: list[Element] = []
        for elem in list(iter_revision_elements(self.editor.dom, tags)):
            tag = elem.tagName
            family = "From" if tag.startswith("w:moveFrom") else "To"
            if tag in MOVE_RANGE_TAGS:
                marks_seen += 1
            if tag in ("w:moveFrom", "w:moveTo"):
                content_seen[family] = True
                for entry in open_starts[family].values():
                    entry[1] = True
            elif not elem.getAttribute("w:id"):
                # Unpairable: an id-less mark could otherwise pair a Start of
                # one range with the End of another.
                unpaired[family].append(elem)
            elif tag.endswith("RangeStart"):
                shadowed = open_starts[family].pop(elem.getAttribute("w:id"), None)
                if shadowed is not None:
                    unpaired[family].append(shadowed[0])
                open_starts[family][elem.getAttribute("w:id")] = [elem, False]
            else:  # *RangeEnd
                entry = open_starts[family].pop(elem.getAttribute("w:id"), None)
                if entry is None:
                    unpaired[family].append(elem)
                elif not entry[1]:
                    removed += [entry[0], elem]
        for family in ("From", "To"):
            if content_seen[family]:
                continue
            removed += [entry[0] for entry in open_starts[family].values()]
            removed += unpaired[family]
        for elem in removed:
            self._remove_element(elem)
        # No marks left: later resolutions skip the walk entirely.
        self.editor.holds_move_range_marks = marks_seen > len(removed)

    def _restore_deletion(self, del_elem) -> None:
        """Restore deleted content by converting w:delText back to w:t.

        ``del_elem`` is a ``w:del`` or a ``w:moveFrom``. A ``w:del`` or
        ``w:moveFrom`` nested inside it is a separate, still-pending revision:
        it is left wrapped for its own accept/reject, its runs keep
        ``w:rsidDel``, and under a nested ``w:del`` they keep ``w:delText``.
        """

        def inside_nested(node, tags: tuple[str, ...]) -> bool:
            parent = node.parentNode
            while parent is not None and parent is not del_elem:
                if getattr(parent, "tagName", "") in tags:
                    return True
                parent = parent.parentNode
            return False

        # Convert all w:delText to w:t — except under a nested w:del. (A
        # nested w:moveFrom gets the conversion: plain w:t is the form Word
        # writes inside one anyway.)
        for del_text in list(del_elem.getElementsByTagName("w:delText")):
            if inside_nested(del_text, ("w:del",)):
                continue
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
        # Update run attributes: w:rsidDel back to w:rsidR — a run inside a
        # nested, still-pending w:del or w:moveFrom keeps its rsidDel.
        for run in del_elem.getElementsByTagName("w:r"):
            if run.hasAttribute("w:rsidDel") and not inside_nested(run, ("w:del", "w:moveFrom")):
                run.setAttribute("w:rsidR", run.getAttribute("w:rsidDel"))
                run.removeAttribute("w:rsidDel")

        # Unwrap the w:del element
        self._unwrap_element(del_elem)

    def _rejoin_paragraph(self, mark_ins) -> None:
        """Reject a paragraph-mark insertion: drop the mark and merge the next
        paragraph's content back into this one — the inverse of a tracked split.

        The paragraph owning the mark survives (keeping its original
        properties, including any section mark); the following paragraph's
        content is appended and that paragraph is removed.

        If that following paragraph is itself an intermediate half of a
        multi-split, its ``w:pPr`` carries its *own* pending mark (the break to
        the paragraph after it). That mark is migrated onto the surviving
        paragraph so the later break stays tracked — otherwise rejecting a
        non-terminal mark individually (a public ``reject_revision``) would drop
        it with ``p2``'s ``pPr``, leaving a permanent, untracked break.
        ``reject_group``/``reject_changeset`` dissolve later (higher-id) marks
        first, so ``p2`` has no mark by the time it is merged and this migration
        is a no-op.
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
        if p2 is None or getattr(p2, "tagName", "") != "w:p":  # pragma: no cover - a mark implies a following paragraph
            return  # best effort: no following paragraph to rejoin into
        downstream_mark = _paragraph_mark_ins(p2)
        p2_pPr = _first_child_element(p2, "w:pPr")
        for child in list(p2.childNodes):
            if child is p2_pPr:
                continue
            p1.appendChild(child)
        p2_parent = p2.parentNode
        assert p2_parent is not None  # a live sibling always has a parent (narrows for ty)
        p2_parent.removeChild(p2)
        if downstream_mark is not None:
            # p1 now ends where p2 did, so it inherits p2's break-to-successor
            # mark (same id/author/date, still a live group member). insertBefore
            # re-parents it out of p2's detached rPr automatically.
            new_rPr = self._paragraph_mark_rPr(p1)
            new_rPr.insertBefore(downstream_mark, new_rPr.firstChild)
