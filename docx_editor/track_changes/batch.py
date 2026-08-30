"""Batch and rewrite: the ``RevisionManager`` mixin that applies edit batches and paragraph rewrites."""

from xml.dom.minidom import Element

from ..exceptions import (
    BatchOperationError,
    DocxEditError,
    HashMismatchError,
    ParagraphIndexError,
    RevisionError,
)
from ..xml_editor import (
    ParagraphRef,
    TextMap,
    TextMapMatch,
    _escape_xml,
    _reject_control_chars,
    _require_valid_occurrence,
    body_paragraphs,
    build_text_map,
    get_rPr_xml,
)
from .base import _RevisionManagerBase
from .diff import _diff_hunks
from .dom import _set_xml_space_preserve
from .models import (
    EditOperation,
    EditValidationResult,
    _not_an_edit_operation_message,
    _validate_edit_target,
    _validate_note,
    _ValidationOutcome,
)


class _BatchMixin(_RevisionManagerBase):
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

        # One full-DOM <w:p> walk shared by the whole batch. Ops never remove
        # or reorder <w:p> elements, and a tracked split only inserts directly
        # after the paragraph being edited; descending-index application means
        # such an insertion always lands after the ops still to be processed,
        # so no pending index shifts. minidom returns a plain non-live list.
        # After a rollback the DOM is replaced, but the exception propagates
        # immediately and this list is never used again.
        paragraphs = body_paragraphs(self.editor.dom)

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
            ValueError: If ``occurrence`` is negative or not an integer, ``note``
                is neither None nor a valid note string, required arguments for
                op.action are missing or not strings, or the action is
                unrecognized.
        """
        _require_valid_occurrence(op.occurrence)
        # Validated on both paths, so a raw EditOperation(action=..., note=<bad>)
        # that skipped the typed constructors fails the dry run *and* fails the
        # batch atomically, before any mutation.
        _validate_note(op.note, ctx=f"{op.action}(): ")

        if op.action == "replace":
            if not op.find or not isinstance(op.replace_with, str):
                raise ValueError("replace requires 'find' and a string 'replace_with'")
            _validate_edit_target(op.find, field="'find'", ctx="replace(): ")
            _reject_control_chars(op.replace_with, field="'replace_with'", ctx="replace(): ", allow_newline=True)
            return op.find
        elif op.action == "delete":
            if not op.text:
                raise ValueError("delete requires 'text'")
            _validate_edit_target(op.text, field="'text'", ctx="delete(): ")
            return op.text
        elif op.action in ("insert_after", "insert_before"):
            if not op.anchor or not isinstance(op.text, str):
                raise ValueError(f"{op.action} requires 'anchor' and a string 'text'")
            _reject_control_chars(op.anchor, field="'anchor'", ctx=f"{op.action}(): ", allow_tab=True)
            _reject_control_chars(op.text, field="'text'", ctx=f"{op.action}(): ", allow_newline=True)
            return op.anchor
        else:
            raise ValueError(f"Unknown action: {op.action}")

    def _apply_single_edit(self, op: EditOperation, paragraphs: list[Element]) -> int:
        """Apply a single edit operation. Paragraph hash was already validated.

        ``paragraphs`` is the batch's shared body-paragraph snapshot (see
        batch_edit).
        """
        ref = ParagraphRef.parse(op.paragraph)
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
            result (``paragraph=None``), never as an exception. Rows that failed
            on a stale hash also carry ``current_ref`` — the ref for that
            paragraph's current content — so the caller can retry without
            parsing the message.
        """
        if not operations:
            return []
        # One <w:p> walk for the whole dry run — validation is read-only, so
        # the snapshot is trivially stable (same sharing as batch_edit).
        paragraphs = body_paragraphs(self.editor.dom)
        results = []
        for i, op in enumerate(operations):
            if not isinstance(op, EditOperation):
                results.append(
                    EditValidationResult(index=i, paragraph=None, valid=False, error=_not_an_edit_operation_message(op))
                )
                continue
            outcome = self._validate_single(op, paragraphs)
            results.append(
                EditValidationResult(
                    index=i,
                    paragraph=op.paragraph,
                    valid=outcome.error is None,
                    error=outcome.error,
                    current_ref=outcome.current_ref,
                )
            )
        return results

    def _validate_single(self, op: EditOperation, paragraphs: list[Element] | None = None) -> "_ValidationOutcome":
        """Return why ``op`` would fail, or an all-None outcome if it is valid.

        Reuses ``_resolve_paragraph``, ``_resolve_action_target``, and
        ``_locate_in_paragraph`` — the same helpers ``_apply_single_edit`` uses —
        so dry-run validation cannot drift from real application semantics
        (out-of-range and ambiguous targets produce the same error text).
        Reads only. ``paragraphs`` is the dry run's shared <w:p> snapshot
        (see validate_batch); None fetches fresh.
        """
        if not op.paragraph:
            return _ValidationOutcome(error="paragraph reference is required for batch mode")

        try:
            ref = ParagraphRef.parse(op.paragraph)
        except ValueError as e:
            return _ValidationOutcome(error=str(e))

        try:
            p = self._resolve_paragraph(ref, paragraphs)
        except HashMismatchError as e:
            # The one mechanically fixable failure: hand back the ref that
            # targets this paragraph's current content (EditValidationResult
            # .current_ref) so callers never regex the hash out of the prose.
            return _ValidationOutcome(error=str(e), current_ref=f"P{e.paragraph_index}#{e.actual_hash}")
        except ParagraphIndexError as e:
            return _ValidationOutcome(error=str(e))

        # Resolve required args + the text this op must locate via the same
        # helper _apply_single_edit uses (which also rejects a negative
        # occurrence), so validation cannot drift from application semantics and
        # the locate below only raises the same errors application would.
        try:
            target = self._resolve_action_target(op)
        except ValueError as e:
            return _ValidationOutcome(error=str(e))

        try:
            self._locate_in_paragraph(p, op.paragraph, target, op.occurrence)
        except (ValueError, DocxEditError) as e:
            return _ValidationOutcome(error=str(e))

        return _ValidationOutcome(error=None)

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

        # One full-DOM <w:p> walk shared by the whole batch. Rewrites never
        # remove or reorder <w:p> elements, and a tracked split (a "\n" in
        # new_text) only *inserts* the new paragraph directly after the one
        # being rewritten. Because application runs in descending index order
        # (see the loop below), every insertion lands after the paragraphs
        # still to be processed, so no pending ref's snapshot index shifts.
        paragraphs = body_paragraphs(self.editor.dom)

        # Parse and validate all refs upfront
        parsed: list[tuple[int, ParagraphRef, str]] = []
        seen_indices: set[int] = set()
        for i, (ref_str, new_text) in enumerate(rewrites):
            try:
                ref = ParagraphRef.parse(ref_str)
                self._resolve_paragraph(ref, paragraphs)  # Raises HashMismatchError if stale
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
                    f"'new_text' must be a string (empty string deletes all text of a tab-free paragraph), "
                    f"got {new_text!r}",
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
                        group_ids[original_idx] = self.rewrite_paragraph(
                            f"P{ref.index}#{ref.hash}", new_text, paragraphs=paragraphs
                        )
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

    def rewrite_paragraph(self, ref_str: str, new_text: str, *, paragraphs: list[Element] | None = None) -> int | None:
        """Rewrite a paragraph's text, generating fine-grained tracked changes.

        Diffs old vs new text at word level and applies minimal tracked changes
        (insertions, deletions, replacements) to transform the paragraph. All
        revisions created by one call are registered as a single revision
        group, so ``accept_group``/``reject_group`` can resolve the rewrite
        as a unit.

        Args:
            ref_str: Paragraph reference string (e.g., "P3#a7b2")
            new_text: Desired new text for the paragraph
            paragraphs: Optional pre-fetched ``<w:p>`` list, threaded in by
                ``batch_rewrite`` so a whole batch shares one full-DOM walk
                (ISSUES.md #51). None fetches fresh — the default for
                standalone calls.

        Returns:
            The rewrite's revision group id, or None when no revisions were
            created (old text already equals ``new_text``, or every change
            was absorbed into this author's own pending insertions —
            amending one of those is undone by rejecting *its* group).

        Raises:
            ValueError: If ``new_text`` is not a string, or does not hold the
                same number of tab marks (``\\t``) as the paragraph — a rewrite
                keeps every tab and changes the text between them (ISSUES.md #6)
            HashMismatchError: If the paragraph hash doesn't match
            IndexError: If paragraph index is out of range
        """
        if not isinstance(new_text, str):
            raise ValueError(
                f"rewrite_paragraph(): 'new_text' must be a string "
                f"(empty string deletes all text of a tab-free paragraph), got {new_text!r}"
            )
        # new_text may hold tabs: _rewrite_paragraph_inner requires the same
        # count as the paragraph and diffs the text between them segment by
        # segment (ISSUES.md #6).
        _reject_control_chars(
            new_text, field="'new_text'", ctx="rewrite_paragraph(): ", allow_newline=True, allow_tab=True
        )
        with self._changeset(), self._grouped() as capture:
            self._rewrite_paragraph_inner(ref_str, new_text, paragraphs)
        return capture.group_id

    def _rewrite_paragraph_inner(self, ref_str: str, new_text: str, paragraphs: list[Element] | None = None) -> None:
        """Diff-and-apply body of ``rewrite_paragraph`` (runs inside _grouped).

        ``paragraphs`` is the batch's shared <w:p> snapshot (see batch_rewrite);
        None fetches fresh.
        """
        ref = ParagraphRef.parse(ref_str)
        p = self._resolve_paragraph(ref, paragraphs)
        text_map = build_text_map(p)
        old_text = text_map.text

        if old_text == new_text:
            return

        # A rewrite keeps the paragraph's tab marks: nothing writes or removes
        # a tracked tab (ISSUES.md #6), so new_text must hold the same number
        # of "\t", and the text between consecutive tabs is diffed segment by
        # segment. Exact by construction — no hunk ever spans a tab — and
        # refused before any mutation, like the split preflight below.
        old_tabs, new_tabs = old_text.count("\t"), new_text.count("\t")
        if old_tabs != new_tabs:
            raise ValueError(
                f"rewrite_paragraph(): 'new_text' must keep the paragraph's tab marks — it has {new_tabs} "
                f"tab(s) where the paragraph has {old_tabs}, and a rewrite can neither add nor remove a "
                f"'\\t' (nothing writes a tracked tab). Keep every tab and rewrite the text between them, "
                f"or use replace()/delete() on the text beside a tab (ISSUES.md #6)."
            )
        hunks = _diff_hunks(old_text, new_text)

        # Preflight the split (\n) hunks against the pre-mutation state: the
        # reversed hunk loop below has no rollback, so anything that can't split
        # cleanly must refuse before any hunk mutates.
        split_hunks = [(tag, s, e, frag) for (tag, s, e, frag) in hunks if tag != "delete" and "\n" in frag]
        if len(split_hunks) > 1:
            # Every split hunk targets this one paragraph, so a second would
            # re-mark it (invalid — see _ensure_splittable). A single hunk with
            # several \n is fine (it threads through fresh tail paragraphs).
            raise RevisionError(
                "Cannot rewrite a paragraph with newlines that fall in separate edits — that "
                "would split one paragraph in more than one place. Use separate split_paragraph()/"
                "replace() calls, or keep the rewrite's newlines within one contiguous change."
            )
        # A replace splits at its match end, an insert at its start; higher-position
        # hunks never shift these lower/equal boundaries, so the pre-mutation
        # positions stay valid.
        for tag, old_char_start, old_char_end, _frag in split_hunks:
            self._ensure_splittable(p)
            boundary = old_char_end if tag == "replace" else old_char_start
            self._reject_unsplittable_boundary(p, text_map, boundary)

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
                node_text = self._get_node_text(last_pos.node)
                if not self._owns_ins(ins_ancestor):
                    self._insert_own_ins_within_foreign_ins(ins_ancestor, last_pos.node, len(node_text), text, rPr_xml)
                elif last_pos.is_tab:
                    self._insert_into_run(run, rPr_xml, last_pos.node, 1, self._plain_run_xml(rPr_xml, text))
                else:
                    self._set_node_text(last_pos.node, node_text + text)
                    _set_xml_space_preserve(last_pos.node)
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
            if not self._owns_ins(ins_ancestor):
                self._insert_own_ins_within_foreign_ins(ins_ancestor, pos.node, pos.offset_in_node, text, rPr_xml)
            elif pos.is_tab:
                self._insert_into_run(run, rPr_xml, pos.node, 0, self._plain_run_xml(rPr_xml, text))
            else:
                node_text = self._get_node_text(pos.node)
                offset = pos.offset_in_node
                self._set_node_text(pos.node, node_text[:offset] + text + node_text[offset:])
                _set_xml_space_preserve(pos.node)
            return

        # Split the run at the offset and insert <w:ins> between; the run's
        # other children (sibling w:t, w:tab, w:br, …) keep their places.
        ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"
        self._insert_into_run(run, rPr_xml, pos.node, pos.offset_in_node, ins_xml)
