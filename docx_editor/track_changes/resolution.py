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
    _first_child_not_in,
    _is_paragraph_mark_ins,
    _is_paragraph_mark_marker,
    _next_element_sibling,
    _num_pr_holds_revision,
    _ppr_base_anchor,
    _ppr_base_children,
)
from .models import (
    _RUN_TRACK_CHANGE_TAGS,
    MOVE_RANGE_TAGS,
    ResolveResult,
    iter_revision_elements,
)


class _ResolutionMixin(_RevisionManagerBase):
    def accept_revision(
        self,
        revision_id: int,
        element_index: dict[str, list[Element]] | None = None,
        *,
        author: str | None = None,
    ) -> bool:
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

        A w:id must name one revision here. Word does not guarantee w:id is
        unique, so a standalone call counts every pending element carrying the
        id — handled revisions and the marks ``list_unhandled_revisions()``
        reports alike — and raises ``RevisionError`` when there is more than
        one and any of them is resolvable, rather than resolve whichever
        comes first in the document (possibly another author's). ``author=``
        narrows the count to one reviewer's; same-author duplicates stay
        refused. ``accept_all``/``reject_all`` (optionally by author) take
        them, and a revision whose only twin is an unhandled mark is still
        reachable through its group when it has one, which the registry
        builds from handled revisions alone.

        Args:
            revision_id: The w:id of the revision to accept
            element_index: Optional pre-built w:id -> element map (see
                ``_revision_element_index``) that lets group/changeset
                resolution skip a full-DOM scan per member. ``None`` scans
                fresh (standalone calls); accept_all/reject_all pass one too.
            author: Standalone calls only — narrow the id to this author's
                revision when several pending revisions share it. Ignored
                with an ``element_index``, whose caller has already
                identified the members.

        Returns:
            True if revision was accepted, False if not found — including an
            id carried only by marks this library cannot resolve.

        Raises:
            RevisionError: Standalone call on an id that more than one
                pending revision carries (after ``author`` narrowing), at
                least one of them resolvable. The document is untouched.
        """
        elem = self._find_revision_element(revision_id, element_index, author)
        if elem is None:
            return False
        tag = elem.tagName
        if tag == "w:ins":
            # Accept insertion: unwrap the content
            self._unwrap_element(elem)
        elif tag == "w:del":
            # Accept deletion: remove the element entirely (its comment
            # markers stay, see _remove_content).
            self._remove_content(elem)
        elif tag == "w:pPrChange":
            # Accept a property change: drop the record of the previous
            # properties.
            self._remove_element(elem)
        elif tag == "w:moveTo" and not _is_paragraph_mark_marker(elem):
            # Accept the destination half: the moved text stays.
            self._unwrap_element(elem)
        else:  # w:moveFrom, or a paragraph-mark marker of either half
            # Accept the source half: the moved-away text goes.
            self._remove_content(elem)
        # Unconditional: move content can leave with any resolved host (a
        # w:del wrapping a w:moveFrom, a rejoined paragraph's mark marker),
        # not only when a move half is the element resolved.
        self._sweep_after_move()
        return True

    def reject_revision(
        self,
        revision_id: int,
        element_index: dict[str, list[Element]] | None = None,
        *,
        author: str | None = None,
    ) -> bool:
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

        Comment markers inside removed content are kept (``_remove_content``):
        a comment wholly on rejected text becomes a point comment where the
        text was, and its body in ``comments.xml`` is never touched.

        A w:id shared by more than one pending revision is refused exactly as
        ``accept_revision`` describes; ``author=`` narrows it.

        Args:
            revision_id: The w:id of the revision to reject
            element_index: Optional pre-built w:id -> element map (see
                ``_revision_element_index``) that lets group/changeset
                resolution skip a full-DOM scan per member. ``None`` scans
                fresh (standalone calls); accept_all/reject_all pass one too.
            author: Standalone calls only — narrow the id to this author's
                revision when several pending revisions share it. Ignored
                with an ``element_index``, whose caller has already
                identified the members.

        Returns:
            True if revision was rejected, False if not found — including an
            id carried only by marks this library cannot resolve.

        Raises:
            RevisionError: Standalone call on an id that more than one
                pending revision carries (after ``author`` narrowing), at
                least one of them resolvable. The document is untouched.
        """
        elem = self._find_revision_element(revision_id, element_index, author)
        if elem is None:
            return False
        tag = elem.tagName
        if tag == "w:ins":
            if _is_paragraph_mark_ins(elem):
                # Reject a paragraph-mark insertion: remove the mark and rejoin
                # the tail paragraph (the inverse of the tracked split).
                self._rejoin_paragraph(elem)
            else:
                # Reject insertion: remove entirely (its comment markers
                # stay, see _remove_content).
                self._remove_content(elem)
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
            self._remove_content(elem)
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

        The index needs no author scoping: a member id is never a duplicated
        one. ``_reconstruct_groups`` leaves every occurrence of a repeated w:id
        wholly ungrouped, and a freshly allocated id is above every id already
        in the file (``_seed_max_change_id``/``_fold_revision_ids`` fold them
        all into a high-water mark), so no member can share its id with a
        second element — of this author or any other.

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

        "Every listed revision" is every pending ``HANDLED_REVISION_TAGS``
        element, because that is what ``list_revisions`` walks — so the loop
        terminates as soon as none remains, regardless of what other revision
        types are still pending. ``_resolve_all_reporting`` is what turns that
        into an honest count for callers.

        The whole-document counterpart to ``_resolve_ids``. Two differences,
        both forced by resolving *every* revision rather than a known member
        set:

        * **Re-listing each pass is what terminates the loop.** The listing
          shrinks as revisions are resolved, so an empty pass ends it, and so
          does a no-progress pass on a freshly built index (see the last
          paragraph). ``_resolve_ids`` can instead lean on ``resolve`` returning
          False, because its members are a fixed set of unique ids. Relying on
          that here would not terminate: a whole-document listing can hand the
          same id back repeatedly.
        * **The index maps each id to every element carrying it.** Word may
          reuse one w:id across authors (A's w:ins and B's w:del both id=7);
          unfiltered, ``_find_revision_element`` returns the first
          still-attached candidate, so same-id revisions resolve within one
          pass. A filtered call instead indexes only ``author``'s elements
          (``_revision_element_index(author)``), so another author's element
          sharing that id is never a candidate: the listing and the index
          agree on ownership, and the call cannot resolve a row it did not
          list. It may resolve fewer — a listed row nested inside a rejected
          insertion goes with its host — but never counts someone else's.
          Targeting, not survival: another author's revision nested inside a
          rejected insertion still goes with it, uncounted.

        The index is built once up front, not per pass: resolution detaches
        elements, so a build stays a valid superset and ``_is_in_document``
        filters what is gone. Each revision then costs a dict lookup instead
        of a fresh full-document scan (ISSUES.md #56). Per pass this is a
        constant number of full-DOM walks — the work itself is still linear
        in the number of revisions listed.

        One resolution adds a pending revision instead: rejecting a
        ``w:pPrChange`` moves its recorded ``w:pPr`` back into the paragraph,
        and a ``w:numPr/w:ins`` recorded there — skipped by the index as
        history — is live from then on. The next pass lists it, finds no index
        entry and makes no progress. So a pass without progress rebuilds the
        index and retries — unless nothing has been resolved since the index
        was built, when a rebuild could not differ and the call ends. Every
        rebuild is therefore followed by a resolution (which detaches an
        element) or by the return, and a run in which no pass stalls never
        rebuilds.
        """
        count = 0
        element_index = self._revision_element_index(author)
        index_is_fresh = True
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
                if progressed:
                    index_is_fresh = False
                elif index_is_fresh:
                    return count
                else:
                    element_index = self._revision_element_index(author)
                    index_is_fresh = True

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
        revisions in the document. An ``author=`` call *targets* only that
        author's own elements even when a producer repeats a w:id across
        authors: the id lookup is scoped to the author being resolved. That is
        about which elements are targeted, not about what survives — rejecting
        an insertion still removes whatever another author nested inside it.

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
        resolution rather than through its own, and is then counted neither
        as resolved nor as unhandled. Rejecting an insertion removes whatever
        is nested inside it. Rejecting a tracked paragraph split (see
        ``_rejoin_paragraph``) removes only what the merged paragraph has no
        room for. The tail's ``w:pPrChange`` goes when the head carries a
        record of its own, or a ``w:numPr`` holding a revision that adopting
        the tail's properties would replace. The tail's ``w:sectPr`` (with
        any ``w:sectPrChange`` in it) goes when both paragraphs carry one. A
        paragraph-mark revision next to the rejected ``w:ins`` (``w:rPrChange``
        included) goes when the tail's mark carries one of the same tag. The
        tail's base properties survive only when its ``w:pPrChange`` moves
        with them (otherwise the head's are kept), so a revision embedded in
        them — a ``w:ins`` or ``w:numberingChange`` in its ``w:numPr`` — is
        kept or dropped with them. Every other revision on either paragraph stays pending on
        the merged paragraph.

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
        revisions in the document. An ``author=`` call *targets* only that
        author's own elements even when a producer repeats a w:id across
        authors: the id lookup is scoped to the author being resolved. That is
        about which elements are targeted, not about what survives — rejecting
        an insertion still removes whatever another author nested inside it.

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

    def _remove_content(self, elem) -> None:
        """Remove revision content (a rejected insertion, an accepted
        deletion) without losing the comment markers inside it.

        Every ``w:commentRangeStart``/``w:commentRangeEnd``/``w:commentReference``
        in ``elem``'s subtree is first moved, in document order, to just before
        ``elem``: the nearest surviving position. Order is kept, so a range
        wholly inside becomes a point comment where the text was and a range
        straddling the edge is clipped to the text that survives. The comment
        itself (``comments.xml``) is never touched by resolution.
        """
        self._hoist_comment_markers(elem)
        self._remove_element(elem)

    def _hoist_comment_markers(self, elem) -> None:
        """Move ``elem``'s comment markers to just before it (see ``_remove_content``).

        A reference moves with its run when that run holds nothing else but
        ``w:rPr`` (the shape Word and this library write); otherwise it is
        detached into a new bare ``w:r``, so the run's text is not resurrected.
        """
        parent = elem.parentNode
        marks = [
            node
            for node in elem.getElementsByTagName("*")  # document order
            if node.tagName in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference")
        ]
        for mark in marks:
            if mark.tagName == "w:commentReference":
                run = mark.parentNode
                if getattr(run, "tagName", "") == "w:r" and all(
                    child is mark or child.tagName == "w:rPr"
                    for child in run.childNodes
                    if child.nodeType == child.ELEMENT_NODE
                ):
                    mark = run
                else:
                    bare = self.editor.dom.createElement("w:r")
                    bare.appendChild(mark)  # detaches it from its old run
                    mark = bare
            parent.insertBefore(mark, elem)

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
        recorded = _first_child_element(ppr_change, "w:pPr")
        self._replace_base_properties(ppr, recorded)
        self._remove_element(ppr_change)

    def _replace_base_properties(self, ppr, source) -> None:
        """Replace ``ppr``'s base properties with those of ``source`` (another
        ``w:pPr``, or None for "no properties"), moving the nodes over ahead
        of ``ppr``'s ``w:rPr``/``w:sectPr``/``w:pPrChange``.

        ``ppr``'s own tail — its paragraph-mark ``w:rPr`` (which may carry a
        pending mark revision), section mark and record — is kept. The same
        tags are skipped on the ``source`` side: a record's ``w:pPr``
        (CT_PPrBase) cannot hold them, so a nonconforming record's copies are
        dropped, not restored over the live ones.
        """
        for child in _ppr_base_children(ppr):
            ppr.removeChild(child)
        if source is None:
            return
        anchor = _ppr_base_anchor(ppr)
        # Same document, so moving the nodes is enough — insertBefore
        # detaches them from their old parent first.
        for child in _ppr_base_children(source):
            ppr.insertBefore(child, anchor)

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
        # A w:moveFrom/w:moveTo recorded inside a change record is previous
        # state, not pending content: it must not keep stray marks alive.
        for elem in list(iter_revision_elements(self.editor.dom, tags, skip_change_records=True)):
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

        The paragraph owning the mark (``p1``) survives, keeping its base
        properties and identity; the following paragraph's (``p2``) content
        is appended and ``p2`` is removed. ``p2``'s base properties, and any
        revision inside them such as a ``w:numPr``'s ``w:ins``, are dropped
        unless ``p2``'s ``w:pPrChange`` brings them over (below).

        The merged paragraph ends where ``p2`` ended, so its paragraph mark
        *is* ``p2``'s mark: ``p2``'s mark ``w:rPr`` becomes the merged mark's.
        Every revision on ``p2``'s mark therefore stays pending on the merged
        paragraph — another author's ``w:del``/``w:moveFrom``/``w:moveTo``,
        an unhandled ``w:rPrChange``, and the ``w:ins`` of the next break when
        ``p2`` is an intermediate half of a multi-split. The nodes are moved,
        not copied, so a bulk resolution's element index still finds them.

        ``p1``'s mark formatting is dropped with the rejected mark, but a
        revision sitting next to the rejected ``w:ins`` (a ``w:del``,
        ``w:moveFrom`` or ``w:moveTo`` marker, or an unhandled
        ``w:rPrChange``) is not: it moves onto the merged mark, in schema
        order, unless ``p2``'s mark already carries one of the same tag (a
        mark holds one of each; ``p2``'s wins and ``p1``'s is dropped
        uncounted). A split written by an earlier release left a
        pre-existing revision of the original mark there, and it cannot be
        told apart from a revision of the new break, so both are kept: a
        kept revision can still be rejected, a dropped one is gone.

        A ``w:pPrChange`` on ``p2`` moves to the merged paragraph together
        with ``p2``'s base properties (the ones the record is a change *to*),
        replacing ``p1``'s, so accepting it keeps them and rejecting it
        restores the recorded ones. It does not move — ``p1``'s properties
        stay and ``p2``'s record is dropped — in two cases: ``p1`` carries a
        record of its own (a ``w:pPr`` holds one), or ``p1``'s ``w:numPr``
        holds a revision (a ``w:ins`` or ``w:numberingChange``), which the
        replacement, or a later rejection of the moved record, would remove.

        A section mark (``w:sectPr``) on ``p2`` moves over the same way, with
        any ``w:sectPrChange`` inside it, so the section still ends at the
        merged paragraph. When ``p1`` carries its own, ``p1``'s wins and
        ``p2``'s is dropped.
        """
        p1 = _ancestor_paragraph(mark_ins)
        p2 = _next_element_sibling(p1.nextSibling) if p1 is not None else None
        rPr = mark_ins.parentNode
        on_live_mark = rPr.parentNode.parentNode is p1
        rPr.removeChild(mark_ins)
        if not on_live_mark:
            # A w:pPr/w:rPr/w:ins somewhere other than directly under p1 (a
            # nonconforming w:pPr inside a run, say) is not p1's mark: there
            # is no break to undo, so only the marker goes.
            if _next_element_sibling(rPr.firstChild) is None:
                rPr.parentNode.removeChild(rPr)
            return
        if p1 is None or p2 is None or p2.tagName != "w:p":  # pragma: no cover - no paragraph follows directly
            # Best effort: nothing to rejoin, so only the mark goes.
            self._tidy_mark_properties(rPr.parentNode)
            return
        p2_pPr = _first_child_element(p2, "w:pPr")
        for child in list(p2.childNodes):
            if child is not p2_pPr:
                p1.appendChild(child)
        p2_parent = p2.parentNode
        assert p2_parent is not None  # a live sibling always has a parent (narrows for ty)
        p2_parent.removeChild(p2)

        # Get-or-create puts the slot in schema order; p2's w:rPr takes it.
        slot = self._paragraph_mark_rPr(p1)
        p1_pPr = slot.parentNode
        assert p1_pPr is not None  # _paragraph_mark_rPr returns an attached w:rPr (narrows for ty)
        p2_rPr = _first_child_element(p2_pPr, "w:rPr") if p2_pPr is not None else None
        if p2_rPr is None:
            p2_rPr = self.editor.dom.createElement("w:rPr")
        p1_pPr.replaceChild(p2_rPr, slot)
        # p1's own mark revisions survive unless p2's mark has the same tag.
        # CT_ParaRPr order: the markers lead, in _RUN_TRACK_CHANGE_TAGS order;
        # the formatting follows; w:rPrChange is last.
        for index, tag in enumerate(_RUN_TRACK_CHANGE_TAGS):
            marker = _first_child_element(slot, tag)
            if marker is not None and _first_child_element(p2_rPr, tag) is None:
                p2_rPr.insertBefore(marker, _first_child_not_in(p2_rPr, _RUN_TRACK_CHANGE_TAGS[: index + 1]))
        rPr_change = _first_child_element(slot, "w:rPrChange")
        if rPr_change is not None and _first_child_element(p2_rPr, "w:rPrChange") is None:
            p2_rPr.appendChild(rPr_change)
        p2_sect = _first_child_element(p2_pPr, "w:sectPr") if p2_pPr is not None else None
        if p2_sect is not None and _first_child_element(p1_pPr, "w:sectPr") is None:
            # Schema order: after the w:rPr just placed, before any record.
            p1_pPr.insertBefore(p2_sect, _first_child_element(p1_pPr, "w:pPrChange"))
        p2_change = _first_child_element(p2_pPr, "w:pPrChange") if p2_pPr is not None else None
        if (
            p2_change is not None
            and _first_child_element(p1_pPr, "w:pPrChange") is None
            and not _num_pr_holds_revision(p1_pPr)
        ):
            self._replace_base_properties(p1_pPr, p2_pPr)
            p1_pPr.appendChild(p2_change)  # schema order: the record is last
        self._tidy_mark_properties(p1_pPr)

    def _tidy_mark_properties(self, pPr) -> None:
        """Drop the empty ``w:rPr``, then the empty ``w:pPr``, a tracked split
        leaves behind (either is kept when it carries anything else)."""
        rPr = _first_child_element(pPr, "w:rPr")
        if rPr is not None and _next_element_sibling(rPr.firstChild) is None:
            pPr.removeChild(rPr)
        if _next_element_sibling(pPr.firstChild) is None:
            pPr.parentNode.removeChild(pPr)
