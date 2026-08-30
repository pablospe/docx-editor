"""Delete sites: the ``RevisionManager`` mixin that turns a text deletion into tracked changes."""

from collections import OrderedDict
from xml.dom.minidom import Element

from ..xml_editor import (
    ParagraphRef,
    TextMapMatch,
    TextPosition,
    _escape_xml,
    get_rPr_xml,
    is_tab_node,
    rebuild_run_fragments,
    render_plain_wt,
)
from .base import _RevisionManagerBase
from .dom import _set_xml_space_preserve
from .models import _validate_edit_target


class _DeleteMixin(_RevisionManagerBase):
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
            ValueError: If ``text`` is not a non-empty string or contains a
                tab (``\\t`` — a tab mark can be matched but not deleted yet, ISSUES.md #6),
                or ``occurrence`` is negative or not an integer
            TextNotFoundError: If the text is not found or occurrence doesn't exist
            AmbiguousTextError: If ``occurrence`` is omitted and ``text``
                matches more than once in the search scope
            HashMismatchError: If the paragraph hash doesn't match
        """
        _validate_edit_target(text, field="'text'", ctx="delete(): ")
        with self._changeset(), self._grouped():
            if paragraph is not None:
                ref = ParagraphRef.parse(paragraph)
                p = self._resolve_paragraph(ref)
                match = self._locate_in_paragraph(p, paragraph, text, occurrence)
                return self._delete_across_nodes(match)

            match = self._locate_document_wide(text, occurrence)
            return self._delete_across_nodes(match)

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

        Despite the name this is the general run splitter — the paragraph
        split path (:meth:`_collect_tail_nodes`) uses it on plain runs too.

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
        # order on whichever side of the split point they fall. A w:tab edge
        # is one character, never sliced: offset 0 puts it on the right,
        # offset 1 on the left.
        left_parts: list[str] = []
        right_parts: list[str] = []
        side = left_parts
        for child in run.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            tag = getattr(child, "tagName", "")
            if tag == "w:rPr":
                continue
            if child is edge_node and is_tab_node(child):
                (right_parts if offset == 0 else left_parts).append(f"<w:r>{rPr_xml}{child.toxml()}</w:r>")
                side = right_parts
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
        wt_nodes = self._content_nodes_in_ancestor(ins_elem)
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

        if not before and not after and len(groups) == len(self._content_nodes_in_ancestor(ins_elem)):
            # Entire insertion matched -- remove the <w:ins> element
            if ins_elem and ins_elem.parentNode:
                ins_elem.parentNode.removeChild(ins_elem)
        elif len(groups) == 1 and first_node is last_node:
            # Single node — use simple truncate/split logic
            node_text = self._get_node_text(first_node)
            before_text = node_text[:first_offset]
            after_text = node_text[last_offset + 1 :]

            if not before_text and not after_text:
                # Entire single node matched. A sole content node was taken by
                # the whole-insertion branch above, so other content (another
                # w:t, or a tab) remains: remove just this node, and its run
                # if that empties it.
                if ins_elem and ins_elem.parentNode:
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

    def _content_nodes_in_ancestor(self, ancestor) -> list:
        """Every character-bearing node inside ``ancestor``, in document order:
        ``w:t`` elements and run-level ``<w:tab/>`` marks.

        Counting tabs keeps the "entire insertion matched → drop the
        ``<w:ins>``" fast paths honest: a match can never include a tab, so
        an insertion holding one is never wholly matched.
        """
        if ancestor is None:
            return []
        return [n for n in ancestor.getElementsByTagName("*") if n.tagName == "w:t" or is_tab_node(n)]

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
