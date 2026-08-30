"""Replace sites: the ``RevisionManager`` mixin that turns a text replacement into tracked changes."""

from collections import OrderedDict
from xml.dom.minidom import Element

from ..exceptions import RevisionError
from ..xml_editor import (
    ParagraphRef,
    TextMapMatch,
    TextPosition,
    _escape_xml,
    _reject_control_chars,
    get_rPr_xml,
    get_text_node_data,
    is_tab_node,
    rebuild_run_fragments,
    render_plain_wt,
)
from .base import _RevisionManagerBase
from .diff import _trim_replace_affixes
from .dom import _node_depth
from .models import _validate_edit_target


class _ReplaceMixin(_RevisionManagerBase):
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
            replace degenerates to a pure deletion). -1 whenever no new
            revision is written: a no-op, or a replace landing wholly inside
            your own pending insertion — that amends the insertion in place
            (whole-insertion and partial matches alike), so undoing it means
            rejecting the group of the insertion it amended.

        Raises:
            ValueError: If ``find`` is not a non-empty string or contains a
                tab (``\\t`` — a tab mark can be matched but not replaced
                yet, ISSUES.md #6), ``replace_with`` is not a string, or ``occurrence`` is
                negative or not an integer
            TextNotFoundError: If the text is not found or occurrence doesn't exist
            AmbiguousTextError: If ``occurrence`` is omitted and ``find``
                matches more than once in the search scope
            HashMismatchError: If the paragraph hash doesn't match
        """
        if not isinstance(replace_with, str):
            raise ValueError(f"'replace_with' must be a string (empty string is allowed), got {replace_with!r}")
        _validate_edit_target(find, field="'find'", ctx="replace(): ")
        _reject_control_chars(replace_with, field="'replace_with'", ctx="replace(): ", allow_newline=True)
        with self._changeset(), self._grouped():
            if paragraph is not None:
                ref = ParagraphRef.parse(paragraph)
                p = self._resolve_paragraph(ref)
                match = self._locate_in_paragraph(p, paragraph, find, occurrence)
                return self._replace_across_nodes(match, replace_with)

            match = self._locate_document_wide(find, occurrence)
            return self._replace_across_nodes(match, replace_with)

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
        A ``<w:tab/>`` mark reads as ``"\\t"``, the one character it occupies
        in the text map, so offset arithmetic and boundary checks
        (``offset == len(node_text)``) work unchanged on a tab position.
        """
        if is_tab_node(node):
            return "\t"
        return get_text_node_data(node)

    def _set_node_text(self, node, text: str) -> None:
        """Replace all text content of a w:t/w:delText element with ``text``.

        Removes every existing TEXT_NODE child and appends a single new one
        carrying the full content. Necessary because assigning to
        ``firstChild.data`` would leave any sibling text nodes behind,
        corrupting the document when the element holds split text (issue #9).

        A ``<w:tab/>`` holds no text: every tab-adjacent edit must route
        through the run rebuilders instead, so splicing into one is a bug
        (it would write ``<w:tab>text</w:tab>``) and fails loudly here.
        """
        if is_tab_node(node):
            raise RevisionError("Internal error: attempted to splice text into a <w:tab/> mark (ISSUES.md #6)")
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
            ins_groups = self._group_positions_by_ins(match.positions)

            if all(g_ins is None or self._owns_ins(g_ins) for g_ins, _ in ins_groups):
                # All our own — splice the replacement in at the match position
                return self._replace_within_own_ins(
                    parts, [g_ins for g_ins, _ in ins_groups if g_ins is not None], replace_with
                )

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

    def _replace_within_own_ins(
        self,
        parts: list[tuple[Element, str, str, str, str, int]],
        ins_elems: list[Element],
        replace_with: str,
    ) -> int:
        """Splice ``replace_with`` into our own pending insertion, in place.

        Text we inserted ourselves was never in the original document, so
        replacing it writes no <w:del>/<w:ins> pair: the matched characters
        are physically removed and a plain run carrying the replacement takes
        their place *at the match position* inside the same <w:ins>.

        Each affected run is rebuilt where it stands rather than relocated to
        the insertion's start (the historical behavior — it reordered every
        match that did not begin at the insertion's first character), so text
        surviving on either side of the match — and any sibling w:t, w:tab,
        w:br … in the same run — keeps its document order. An insertion left
        with no content is removed.

        Returns -1: editing our own pending text creates no new revision.
        """
        ins_rPr = self._majority_rPr(parts)
        replacement_xml = f"<w:r>{ins_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r>"
        last_nid = parts[-1][5]

        run_parts: OrderedDict[int, list] = OrderedDict()
        for part in parts:
            run_parts.setdefault(id(part[0]), []).append(part)

        # Deepest run first: a run nested in another matched run's w:drawing
        # must be rebuilt while still attached, or its ancestor re-serializes a
        # stale copy of the drawing and the edit inside it — the replacement
        # included — is dropped with the detached subtree. sorted() is stable,
        # so runs at the same depth keep document order.
        for rparts in sorted(run_parts.values(), key=lambda rp: -_node_depth(rp[0][0])):
            run = rparts[0][0]
            rPr_xml = rparts[0][1]
            node_to_part = {nid: (before, after) for _run, _rp_xml, before, _matched, after, nid in rparts}

            # Keyword-only defaults bind this iteration's state (B023)
            def render_wt(wt, *, node_to_part=node_to_part, run_rPr=rPr_xml) -> list[str]:
                if id(wt) not in node_to_part:
                    # Unmatched sibling — preserve
                    return render_plain_wt(wt, run_rPr)
                before, after = node_to_part[id(wt)]
                fragments: list[str] = []
                if before:
                    fragments.append(f"<w:r>{run_rPr}<w:t>{_escape_xml(before)}</w:t></w:r>")
                # The matched text itself is dropped; the replacement takes
                # the place of the match's last node.
                if id(wt) == last_nid:
                    fragments.append(replacement_xml)
                if after:
                    fragments.append(f"<w:r>{run_rPr}<w:t>{_escape_xml(after)}</w:t></w:r>")
                return fragments

            # Emit the run's children in document order (w:tab/w:br/w:drawing/…
            # preserved in place), then swap it for the rebuilt fragments
            new_xml = "".join(rebuild_run_fragments(run, rPr_xml, render_wt))
            if new_xml:
                self.editor.insert_before(run, new_xml)
            if run.parentNode:  # pragma: no branch - a matched run is always attached
                run.parentNode.removeChild(run)

        # An insertion whose whole content was matched keeps nothing but the
        # replacement; one that did not receive the replacement is now empty.
        for ins_elem in ins_elems:
            if ins_elem.parentNode and not any(child.nodeType == child.ELEMENT_NODE for child in ins_elem.childNodes):
                ins_elem.parentNode.removeChild(ins_elem)
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
