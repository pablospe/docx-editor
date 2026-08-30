"""Group and changeset registry: the ``RevisionManager`` mixin that records and queries revision grouping."""

from collections.abc import Iterable, Iterator
from contextlib import contextmanager
from xml.dom.minidom import Element

from ..exceptions import RevisionError
from .base import _RevisionManagerBase
from .dom import (
    _ancestor_paragraph,
    _has_ancestor,
    _is_paragraph_mark_ins,
    _is_paragraph_mark_marker,
    _next_element_sibling,
    _outermost_revision,
    _paragraph_mark_ins,
    _revision_elements,
)
from .models import _GroupCapture, _RegistrySnapshot


class _RegistryMixin(_RevisionManagerBase):
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
        # (author, date) of the current run. Paragraph identity is tracked
        # separately in run_para, so it is not part of the key.
        run_key: tuple[str, str] | None = None
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
                # are the group's changeset key, so groups in different
                # paragraphs sharing them still join one changeset.
                assert run_key is not None
                cs_key = run_key
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
                and author == run_key[0]
                and date == run_key[1]
                and (paragraph is run_para or self._is_split_continuation(run_para, paragraph, author, date))
            )
            if not same_run:
                close_run()
            run_key = (author, date)
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
        if prev_para is None:  # pragma: no cover - run_para is always set alongside run_key
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
        """Group id of a revision (recorded or inferred), or None if ungrouped.

        A negative id is the "no revision was written" sentinel the edit
        methods return (a no-op, or an amendment to our own pending
        insertion), never a lookup key — resolving it must not collide with
        a document that happens to carry a negative ``w:id``.
        """
        if revision_id < 0:
            return None
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

    def group_spans(self, group_ids: Iterable[int]) -> dict[int, tuple[Element, Element]]:
        """First and last *content-level* revision element of each group, document order.

        The anchor lookup behind ``note=``: a comment bracketing a whole edit
        needs the outermost elements the edit created, not a text position —
        a deletion's text is not in the accepted text map at all, and an
        insertion's text may repeat in its paragraph.

        Paragraph-mark insertions are excluded: a comment marker cannot live in
        ``w:pPr/w:rPr``. A group whose only revision is one — a pure tracked
        split — is therefore absent from the result, as is a group with no live
        revision at all; callers must handle a missing key rather than assume
        one span per requested group.

        An endpoint nested inside *another author's* revision is reported as
        that outermost revision instead, so a marker placed beside it cannot be
        carried away when that author's proposal is rejected.

        One full-DOM walk regardless of how many groups are asked about, so a
        whole batch's notes cost one pass.
        """
        wanted = {rev_id: gid for gid in group_ids for rev_id in self._groups.get(gid, ())}
        if not wanted:
            return {}
        members: dict[int, list[Element]] = {}
        for elem in _revision_elements(self.editor.dom):
            if _is_paragraph_mark_marker(elem):
                continue
            try:
                rev_id = int(elem.getAttribute("w:id"))
            except ValueError:  # pragma: no cover - our own ids are always numeric
                continue
            gid = wanted.get(rev_id)
            if gid is not None:
                members.setdefault(gid, []).append(elem)
        spans: dict[int, tuple[Element, Element]] = {}
        for gid, elems in members.items():
            # A member nested inside another member of the same group would put
            # a marker *inside* the revision it explains, where rejecting would
            # carry it away and orphan its twin. Span the outermost members only.
            outer = [e for e in elems if not any(other is not e and _has_ancestor(e, other) for other in elems)]
            if outer:  # pragma: no branch - a member is outermost or nested in one
                # Hoisting keeps both markers out of a foreign revision that a
                # later reject could carry away; it only ever widens the span
                # outwards, so first still precedes last.
                spans[gid] = (_outermost_revision(outer[0]), _outermost_revision(outer[-1]))
        return spans

    def groups_are_dead(self, group_ids: Iterable[int]) -> set[int]:
        """The subset of ``group_ids`` with no revision element left in the document.

        A group goes dead when the last revision it holds is accepted, rejected,
        or carried away inside a resolved host — which is exactly when anything
        keyed to it (a ``note=`` comment) has nothing left to explain. One
        ``w:id`` index for the whole call, so asking about every registered
        group costs the same walk as asking about one.

        A group id this manager does not know reads as dead: it has no live
        revision either.
        """
        group_ids = list(group_ids)
        if not group_ids:
            return set()
        element_index = self._revision_element_index()
        return {
            gid
            for gid in group_ids
            if not any(
                self._is_in_document(elem)
                for rev_id in self._groups.get(gid, ())
                for elem in element_index.get(str(rev_id), ())
            )
        }
