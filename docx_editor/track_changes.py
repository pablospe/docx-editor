"""Track changes management for docx_editor.

Provides RevisionManager for creating and managing tracked changes (insertions/deletions).
"""

from collections import OrderedDict
from dataclasses import dataclass
from datetime import datetime
from typing import Literal

from .exceptions import RevisionError, TextNotFoundError
from .xml_editor import DocxXMLEditor, TextMapMatch, build_text_map, find_in_text_map


@dataclass
class Revision:
    """Represents a tracked change (insertion or deletion)."""

    id: int
    type: Literal["insertion", "deletion"]
    author: str
    date: datetime | None
    text: str

    def __repr__(self) -> str:
        type_symbol = "+" if self.type == "insertion" else "-"
        return f"Revision({type_symbol}{self.id}: '{self.text[:30]}...' by {self.author})"


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
            text_map = build_text_map(paragraph)
            local_occ = 0
            while find_in_text_map(text_map, text, local_occ) is not None:
                count += 1
                local_occ += 1
        return count

    def _get_nth_match(self, text: str, occurrence: int):
        """Get the nth occurrence of text (0-indexed).

        Args:
            text: Text to search for
            occurrence: Which occurrence to get (0 = first, 1 = second, etc.)

        Returns:
            The matching w:t element

        Raises:
            TextNotFoundError: If not enough occurrences exist
        """
        matches = self.editor.find_all_nodes(tag="w:t", contains=text)
        if not matches:
            raise TextNotFoundError(f"Text not found: '{text}'")
        if occurrence >= len(matches):
            raise TextNotFoundError(
                f"Only {len(matches)} occurrence(s) of '{text}' found, "
                f"but occurrence={occurrence} requested (0-indexed)"
            )
        return matches[occurrence]

    def _find_across_boundaries(self, text: str, occurrence: int = 0) -> TextMapMatch | None:
        """Find the nth occurrence of text across element boundaries.

        Searches across all paragraphs using text maps.
        Returns TextMapMatch or None.
        """
        current_occurrence = 0
        for paragraph in self.editor.dom.getElementsByTagName("w:p"):
            text_map = build_text_map(paragraph)
            local_occ = 0
            while True:
                match = find_in_text_map(text_map, text, local_occ)
                if match is None:
                    break
                if current_occurrence == occurrence:
                    return match
                current_occurrence += 1
                local_occ += 1
        return None

    def replace_text(self, find: str, replace_with: str, occurrence: int = 0) -> int:
        """Replace text with tracked changes (deletion + insertion).

        Finds the specified occurrence of `find` text and replaces it with `replace_with`,
        creating a tracked deletion for the old text and insertion for the new text.

        Args:
            find: Text to find and replace
            replace_with: Replacement text
            occurrence: Which occurrence to replace (0 = first, 1 = second, etc.)

        Returns:
            The change ID of the insertion

        Raises:
            TextNotFoundError: If the text is not found or occurrence doesn't exist
        """
        # Find the text element containing the search text
        try:
            elem = self._get_nth_match(find, occurrence)
        except TextNotFoundError:
            # Fall back to cross-boundary search
            match = self._find_across_boundaries(find, occurrence)
            if match is None:
                raise
            return self._replace_across_nodes(match, replace_with)

        # Get the parent run
        run = elem.parentNode
        while run and run.nodeName != "w:r":
            run = run.parentNode

        if not run:
            raise RevisionError("Could not find parent w:r element")

        # Get the full text content
        full_text = elem.firstChild.data if elem.firstChild else ""
        start_idx = full_text.find(find)

        if start_idx == -1:
            raise TextNotFoundError(f"Text not found: '{find}'")

        # If run has multiple w:t children, delegate to cross-boundary path
        # to avoid losing sibling w:t nodes when replacing the whole run.
        if len(run.getElementsByTagName("w:t")) > 1:
            match = self._find_across_boundaries(find, occurrence)
            if match is not None:
                return self._replace_across_nodes(match, replace_with)

        # Build replacement XML
        before_text = full_text[:start_idx]
        after_text = full_text[start_idx + len(find) :]

        # Site A: If inside <w:ins>, edit text in-place (no del/ins wrappers)
        if self._find_ancestor(run, "w:ins"):
            elem.firstChild.data = before_text + replace_with + after_text
            _set_xml_space_preserve(elem)
            return -1

        # Preserve run properties if present
        rPr_xml = ""
        rPr_elems = run.getElementsByTagName("w:rPr")
        if rPr_elems:
            rPr_xml = rPr_elems[0].toxml()

        # Build the replacement runs
        xml_parts = []

        # Text before the match (unchanged)
        if before_text:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before_text)}</w:t></w:r>")

        # Deletion of old text
        xml_parts.append(f"<w:del><w:r>{rPr_xml}<w:delText>{_escape_xml(find)}</w:delText></w:r></w:del>")

        # Insertion of new text
        xml_parts.append(f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(replace_with)}</w:t></w:r></w:ins>")

        # Text after the match (unchanged)
        if after_text:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after_text)}</w:t></w:r>")

        # Replace the original run
        new_xml = "".join(xml_parts)
        nodes = self.editor.replace_node(run, new_xml)

        # Find the insertion node to get its ID
        for node in nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":
                return int(node.getAttribute("w:id"))

        return -1

    def suggest_deletion(self, text: str, occurrence: int = 0) -> int:
        """Mark text as deleted with tracked changes.

        Args:
            text: Text to mark as deleted
            occurrence: Which occurrence to delete (0 = first, 1 = second, etc.)

        Returns:
            The change ID of the deletion

        Raises:
            TextNotFoundError: If the text is not found or occurrence doesn't exist
        """
        # Find the text element containing the search text
        try:
            elem = self._get_nth_match(text, occurrence)
        except TextNotFoundError:
            # Fall back to cross-boundary search
            match = self._find_across_boundaries(text, occurrence)
            if match is None:
                raise
            return self._delete_across_nodes(match)

        # Get the parent run
        run = elem.parentNode
        while run and run.nodeName != "w:r":
            run = run.parentNode

        if not run:
            raise RevisionError("Could not find parent w:r element")

        # Get the full text content
        full_text = elem.firstChild.data if elem.firstChild else ""
        start_idx = full_text.find(text)

        if start_idx == -1:
            raise TextNotFoundError(f"Text not found: '{text}'")

        # If run has multiple w:t children, delegate to cross-boundary path
        if len(run.getElementsByTagName("w:t")) > 1:
            match = self._find_across_boundaries(text, occurrence)
            if match is not None:
                return self._delete_across_nodes(match)

        # Preserve run properties if present
        rPr_xml = ""
        rPr_elems = run.getElementsByTagName("w:rPr")
        if rPr_elems:
            rPr_xml = rPr_elems[0].toxml()

        before_text = full_text[:start_idx]
        after_text = full_text[start_idx + len(text) :]

        # Site B: If inside <w:ins>, shrink/remove the insertion (no <w:del>)
        ins_ancestor = self._find_ancestor(run, "w:ins")
        if ins_ancestor:
            remaining = before_text + after_text
            if remaining:
                elem.firstChild.data = remaining
                _set_xml_space_preserve(elem)
            else:
                # Entire w:t text removed — check if other w:t nodes exist
                if len(self._get_wt_nodes_in_ancestor(ins_ancestor)) == 1:
                    # Sole w:t — remove entire <w:ins>
                    if ins_ancestor.parentNode:
                        ins_ancestor.parentNode.removeChild(ins_ancestor)
                else:
                    # Other w:t nodes exist — remove just this w:t node
                    if elem.parentNode:
                        elem.parentNode.removeChild(elem)
                    # If the run has no more w:t children, remove it too
                    if not run.getElementsByTagName("w:t") and run.parentNode:
                        run.parentNode.removeChild(run)
            return -1

        # Build the replacement runs
        xml_parts = []

        # Text before the match (unchanged)
        if before_text:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before_text)}</w:t></w:r>")

        # Deletion of the target text
        xml_parts.append(f"<w:del><w:r>{rPr_xml}<w:delText>{_escape_xml(text)}</w:delText></w:r></w:del>")

        # Text after the match (unchanged)
        if after_text:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after_text)}</w:t></w:r>")

        # Replace the original run
        new_xml = "".join(xml_parts)
        nodes = self.editor.replace_node(run, new_xml)

        # Find the deletion node to get its ID
        for node in nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:del":
                return int(node.getAttribute("w:id"))

        return -1

    def _get_run_info(self, node):
        """Get the parent w:r element and its rPr XML for a w:t node."""
        run = node.parentNode
        while run and run.nodeName != "w:r":
            run = run.parentNode
        if not run:
            return None, ""
        rPr_xml = ""
        rPr_elems = run.getElementsByTagName("w:rPr")
        if rPr_elems:
            rPr_xml = rPr_elems[0].toxml()
        return run, rPr_xml

    def _get_node_text(self, node) -> str:
        """Get text content of a w:t node."""
        text = ""
        for child in node.childNodes:
            if child.nodeType == child.TEXT_NODE:
                text += child.data
        return text

    def _build_cross_boundary_parts(self, match: TextMapMatch):
        """Build per-node data for a cross-boundary match.

        Returns list of (run, rPr_xml, before_text, matched_part, after_text) tuples,
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
        for info in node_data.values():
            node_text = self._get_node_text(info["node"])
            first = info["first_offset"]
            last = info["last_offset"]
            before = node_text[:first]
            matched = node_text[first : last + 1]
            after = node_text[last + 1 :]
            result.append((info["run"], info["rPr_xml"], before, matched, after))
        return result

    def _classify_segments(self, match: TextMapMatch):
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
        """Replace text spanning multiple w:t elements, handling mixed revision contexts."""
        if match.spans_boundary:
            return self._replace_mixed_state(match, replace_with)
        return self._replace_same_context(match, replace_with)

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

        # Site D: If all positions inside <w:ins>, edit in-place
        if all(p.is_inside_ins for p in match.positions):
            first_node = match.positions[0].node
            ins_elem = self._find_ancestor(first_node, "w:ins")
            # Save parent/next sibling before removal may detach ins_elem
            ins_parent = ins_elem.parentNode if ins_elem else None
            ins_next = ins_elem.nextSibling if ins_elem else None

            self._remove_from_insertion(match.positions)

            first_rPr = parts[0][1]
            new_run_xml = f"<w:r>{first_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r>"

            if ins_elem and ins_elem.parentNode:
                # ins_elem still in DOM — insert replacement inside it
                nodes = self.editor._parse_fragment(new_run_xml)
                first_child = ins_elem.firstChild
                if first_child:
                    for node in nodes:
                        ins_elem.insertBefore(node, first_child)
                else:
                    for node in nodes:
                        ins_elem.appendChild(node)
            elif ins_parent:
                # ins_elem was fully removed — wrap replacement in a new <w:ins>
                ins_wrapper_xml = f"<w:ins>{new_run_xml}</w:ins>"
                nodes = self.editor._parse_fragment(ins_wrapper_xml)
                for node in nodes:
                    ins_parent.insertBefore(node, ins_next)
            return -1

        # Use first run's rPr for the insertion
        first_rPr = parts[0][1]

        # Collect matched node ids from match positions
        matched_node_ids = {id(pos.node) for pos in match.positions}

        # Group parts by run for multi-w:t preservation
        run_order: list[int] = []
        run_map: dict[int, dict] = {}
        for run, rPr_xml, before, matched, after in parts:
            rid = id(run)
            if rid not in run_map:
                run_order.append(rid)
                run_map[rid] = {"run": run, "rPr_xml": rPr_xml, "parts": []}
            run_map[rid]["parts"].append((before, matched, after))

        xml_parts = []
        part_idx = 0
        total_parts = len(parts)
        for rid in run_order:
            info = run_map[rid]
            run = info["run"]
            rPr_xml = info["rPr_xml"]

            # Iterate all w:t in document order, emitting unmatched as preserved
            all_wt = list(run.getElementsByTagName("w:t"))
            part_sub_idx = 0
            for wt in all_wt:
                if id(wt) in matched_node_ids and part_sub_idx < len(info["parts"]):
                    before, matched, after = info["parts"][part_sub_idx]
                    part_sub_idx += 1

                    if before:
                        xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before)}</w:t></w:r>")
                    xml_parts.append(f"<w:del><w:r>{rPr_xml}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>")

                    # Insert replacement after the last deletion
                    if part_idx + part_sub_idx == total_parts:
                        xml_parts.append(f"<w:ins><w:r>{first_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r></w:ins>")

                    if after:
                        xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after)}</w:t></w:r>")
                else:
                    # Unmatched sibling — preserve
                    wt_text = self._get_node_text(wt)
                    if wt_text:
                        xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(wt_text)}</w:t></w:r>")

            part_idx += len(info["parts"])

        # Replace all affected runs: insert new XML before first run, remove all runs
        first_run = parts[0][0]
        new_xml = "".join(xml_parts)
        nodes = self.editor.insert_before(first_run, new_xml)

        seen = set()
        for run, _, _, _, _ in parts:
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

        # Get rPr from first position's run for the new insertion
        first_run, first_rPr = self._get_run_info(match.positions[0].node)

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
        marker = self.editor.dom.createElement("w:_marker")
        ref_node.parentNode.insertBefore(marker, ref_node)

        # Process each segment to delete/remove the matched text
        for is_inside_ins, positions in segments:
            if is_inside_ins:
                self._remove_from_insertion(positions)
            else:
                self._delete_regular_segment(positions)

        # Insert replacement after the last <w:del> sibling following the marker,
        # so it appears after any preserved prefix text.
        ins_xml = f"<w:ins><w:r>{first_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r></w:ins>"
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

    def _find_ancestor(self, node, tag_name: str):
        """Find the nearest ancestor with the given tag name."""
        parent = node.parentNode
        while parent:
            if parent.nodeType == parent.ELEMENT_NODE and parent.tagName == tag_name:
                return parent
            parent = parent.parentNode
        return None

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
                        # Other w:t nodes exist — remove just this w:t
                        run = self._find_ancestor(first_node, "w:r")
                        if first_node.parentNode:
                            first_node.parentNode.removeChild(first_node)
                        # If the run has no more w:t children, remove the run
                        if run and not run.getElementsByTagName("w:t") and run.parentNode:
                            run.parentNode.removeChild(run)
            elif not before_text:
                first_node.firstChild.data = after_text
            elif not after_text:
                first_node.firstChild.data = before_text
            else:
                # Middle split
                first_node.firstChild.data = before_text
                run = self._find_ancestor(first_node, "w:r")
                if ins_elem and run:
                    rPr_xml = ""
                    rPr_elems = run.getElementsByTagName("w:rPr")
                    if rPr_elems:
                        rPr_xml = rPr_elems[0].toxml()
                    after_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(after_text)}</w:t></w:r></w:ins>"
                    self.editor.insert_after(ins_elem, after_xml)
        else:
            # Multi-node: truncate first node to before, last node to after,
            # remove intermediate nodes entirely.
            # Only remove the w:t node; remove the run only if no w:t children remain.
            if before:
                first_node.firstChild.data = before
                _set_xml_space_preserve(first_node)
            else:
                self._remove_wt_and_maybe_run(first_node)

            if after:
                last_node.firstChild.data = after
                _set_xml_space_preserve(last_node)
            else:
                self._remove_wt_and_maybe_run(last_node)

            # Remove intermediate nodes
            for group in groups[1:-1]:
                node = group["node"]
                node_text = self._get_node_text(node)
                if not node_text or node_text == self._get_node_text(node):
                    self._remove_wt_and_maybe_run(node)

    def _remove_wt_and_maybe_run(self, wt_node) -> None:
        """Remove a w:t node, and its parent w:r if no w:t children remain."""
        run = self._find_ancestor(wt_node, "w:r")
        if wt_node.parentNode:
            wt_node.parentNode.removeChild(wt_node)
        if run and not run.getElementsByTagName("w:t") and run.parentNode:
            run.parentNode.removeChild(run)

    def _get_wt_nodes_in_ancestor(self, ancestor) -> list:
        """Get all w:t nodes inside an ancestor element."""
        if ancestor is None:
            return []
        return ancestor.getElementsByTagName("w:t")

    def _delete_regular_segment(self, positions: list) -> int:
        """Delete regular (non-insertion) text by wrapping in <w:del>.

        Groups positions by run first, then by w:t node within each run,
        so that each run is removed exactly once even when it contains
        multiple w:t nodes involved in the match.

        Returns the w:id of the first <w:del> element created, or -1.
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
        processed_runs: set[int] = set()

        for _global_idx, (run_info, _) in enumerate(all_node_groups):
            run = run_info["run"]
            rPr_xml = run_info["rPr_xml"]
            rid = id(run)

            if rid in processed_runs:
                continue

            # Build xml_parts for ALL w:t nodes in this run, preserving unmatched ones
            matched_node_ids = {id(ng["node"]) for ng in run_info["nodes"].values()}
            node_items = list(run_info["nodes"].values())
            all_wt_nodes = list(run.getElementsByTagName("w:t"))
            xml_parts: list[str] = []

            for wt_node in all_wt_nodes:
                if id(wt_node) not in matched_node_ids:
                    # Unmatched sibling — preserve as-is
                    wt_text = self._get_node_text(wt_node)
                    if wt_text:
                        xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(wt_text)}</w:t></w:r>")
                    continue

                ng = run_info["nodes"][id(wt_node)]
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
                    xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before)}</w:t></w:r>")
                xml_parts.append(f"<w:del><w:r>{rPr_xml}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>")
                if after:
                    xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after)}</w:t></w:r>")

            new_xml = "".join(xml_parts)
            nodes = self.editor.insert_before(run, new_xml)
            if run.parentNode:
                run.parentNode.removeChild(run)
            processed_runs.add(rid)

            if first_del_id == -1:
                for n in nodes:
                    if n.nodeType == n.ELEMENT_NODE and n.tagName == "w:del":
                        first_del_id = int(n.getAttribute("w:id"))
                        break

        return first_del_id

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

        # Site F: If all positions inside <w:ins>, remove from insertion directly
        if all(p.is_inside_ins for p in match.positions):
            self._remove_from_insertion(match.positions)
            return -1

        # Group parts by run to handle multi-w:t runs
        matched_nodes = set()
        run_parts: OrderedDict[int, list] = OrderedDict()
        for part in parts:
            run = part[0]
            rid = id(run)
            if rid not in run_parts:
                run_parts[rid] = []
            run_parts[rid].append(part)
            # Track which w:t nodes are matched (from _build_cross_boundary_parts,
            # the node is identified by before/matched/after text)

        # Collect matched node ids from match positions
        for pos in match.positions:
            matched_nodes.add(id(pos.node))

        xml_parts = []
        for _rid, rparts in run_parts.items():
            run = rparts[0][0]
            rPr_xml = rparts[0][1]

            # Emit unmatched w:t siblings before and matched parts in document order
            all_wt = list(run.getElementsByTagName("w:t"))
            matched_parts_by_node = {}
            for rp in rparts:
                # Find which w:t node this part corresponds to
                for wt in all_wt:
                    if id(wt) in matched_nodes and id(wt) not in matched_parts_by_node:
                        matched_parts_by_node[id(wt)] = rp
                        break

            for wt in all_wt:
                if id(wt) in matched_parts_by_node:
                    _, rp_xml, before, matched, after = matched_parts_by_node[id(wt)]
                    if before:
                        xml_parts.append(f"<w:r>{rp_xml}<w:t>{_escape_xml(before)}</w:t></w:r>")
                    xml_parts.append(f"<w:del><w:r>{rp_xml}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>")
                    if after:
                        xml_parts.append(f"<w:r>{rp_xml}<w:t>{_escape_xml(after)}</w:t></w:r>")
                elif id(wt) not in matched_nodes:
                    # Unmatched sibling — preserve
                    wt_text = self._get_node_text(wt)
                    if wt_text:
                        xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(wt_text)}</w:t></w:r>")

        first_run = parts[0][0]
        new_xml = "".join(xml_parts)
        nodes = self.editor.insert_before(first_run, new_xml)

        seen = set()
        for run, _, _, _, _ in parts:
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
                self._remove_from_insertion(positions)
            else:
                del_id = self._delete_regular_segment(positions)
                if first_del_id == -1:
                    first_del_id = del_id

        return first_del_id

    def insert_text_after(self, anchor: str, text: str, occurrence: int = 0) -> int:
        """Insert text after anchor with tracked changes.

        Args:
            anchor: Text to find as the anchor point
            text: Text to insert after the anchor
            occurrence: Which occurrence of anchor to use (0 = first, 1 = second, etc.)

        Returns:
            The change ID of the insertion

        Raises:
            TextNotFoundError: If the anchor text is not found or occurrence doesn't exist
        """
        return self._insert_text(anchor, text, position="after", occurrence=occurrence)

    def insert_text_before(self, anchor: str, text: str, occurrence: int = 0) -> int:
        """Insert text before anchor with tracked changes.

        Args:
            anchor: Text to find as the anchor point
            text: Text to insert before the anchor
            occurrence: Which occurrence of anchor to use (0 = first, 1 = second, etc.)

        Returns:
            The change ID of the insertion

        Raises:
            TextNotFoundError: If the anchor text is not found or occurrence doesn't exist
        """
        return self._insert_text(anchor, text, position="before", occurrence=occurrence)

    def _insert_text(self, anchor: str, text: str, position: Literal["before", "after"], occurrence: int = 0) -> int:
        """Insert text before or after anchor with tracked changes."""
        # Find the text element containing the anchor text
        try:
            elem = self._get_nth_match(anchor, occurrence)
        except TextNotFoundError:
            # Fall back to cross-boundary search
            match = self._find_across_boundaries(anchor, occurrence)
            if match is None:
                raise TextNotFoundError(f"Anchor text not found: '{anchor}'") from None
            return self._insert_near_match(match, text, position)

        # Get the parent run
        run = elem.parentNode
        while run and run.nodeName != "w:r":
            run = run.parentNode

        if not run:
            raise RevisionError("Could not find parent w:r element")

        # Preserve run properties if present
        rPr_xml = ""
        rPr_elems = run.getElementsByTagName("w:rPr")
        if rPr_elems:
            rPr_xml = rPr_elems[0].toxml()

        # Site C: If inside <w:ins>, edit text in-place (safe for multi-w:t)
        full_text = elem.firstChild.data if elem.firstChild else ""
        anchor_idx = full_text.find(anchor)

        if anchor_idx == -1:
            raise TextNotFoundError(f"Anchor text not found: '{anchor}'")

        before_text = full_text[:anchor_idx]
        after_text = full_text[anchor_idx + len(anchor) :]

        if self._find_ancestor(run, "w:ins"):
            if position == "before":
                elem.firstChild.data = before_text + text + anchor + after_text
            else:
                elem.firstChild.data = before_text + anchor + text + after_text
            _set_xml_space_preserve(elem)
            return -1

        # If run has multiple w:t children, delegate to cross-boundary path
        # to avoid losing sibling w:t nodes when replacing the whole run.
        if len(run.getElementsByTagName("w:t")) > 1:
            match = self._find_across_boundaries(anchor, occurrence)
            if match is not None:
                return self._insert_near_match(match, text, position)

        # Build split runs + insertion
        xml_parts = []
        if before_text:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before_text)}</w:t></w:r>")

        ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"

        if position == "before":
            xml_parts.append(ins_xml)
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(anchor)}</w:t></w:r>")
        else:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(anchor)}</w:t></w:r>")
            xml_parts.append(ins_xml)

        if after_text:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after_text)}</w:t></w:r>")

        # Replace the original run with the split + insertion
        new_xml = "".join(xml_parts)
        nodes = self.editor.replace_node(run, new_xml)

        # Find the insertion node to get its ID
        for node in nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":
                return int(node.getAttribute("w:id"))

        return -1

    def _insert_near_match(self, match: TextMapMatch, text: str, position: str) -> int:
        """Insert text before/after a cross-boundary match."""
        positions = match.positions
        if not positions:
            return -1

        # Get rPr from first run
        first_run, rPr_xml = self._get_run_info(positions[0].node)

        # Sites H/I: If ref run is inside <w:ins>, insert bare <w:r> (no wrapper)
        if position == "after":
            last_run, _ = self._get_run_info(positions[-1].node)
            if not last_run:
                return -1
            ref_run = last_run
        else:
            if not first_run:
                return -1
            ref_run = first_run

        all_inside_ins = all(p.is_inside_ins for p in positions)
        if all_inside_ins and self._find_ancestor(ref_run, "w:ins"):
            bare_xml = f"<w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r>"
            if position == "after":
                self.editor.insert_after(ref_run, bare_xml)
            else:
                self.editor.insert_before(ref_run, bare_xml)
            return -1

        ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"

        if position == "after":
            nodes = self.editor.insert_after(ref_run, ins_xml)
        else:
            nodes = self.editor.insert_before(ref_run, ins_xml)

        for node in nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":
                return int(node.getAttribute("w:id"))
        return -1

    def list_revisions(self, author: str | None = None) -> list[Revision]:
        """List all tracked changes in the document.

        Args:
            author: If provided, filter by author name

        Returns:
            List of Revision objects
        """
        revisions = []

        # Find all insertions
        for ins_elem in self.editor.dom.getElementsByTagName("w:ins"):
            rev = self._parse_revision(ins_elem, "insertion")
            if rev and (author is None or rev.author == author):
                revisions.append(rev)

        # Find all deletions
        for del_elem in self.editor.dom.getElementsByTagName("w:del"):
            rev = self._parse_revision(del_elem, "deletion")
            if rev and (author is None or rev.author == author):
                revisions.append(rev)

        # Sort by ID
        revisions.sort(key=lambda r: r.id)
        return revisions

    def _parse_revision(self, elem, rev_type: Literal["insertion", "deletion"]) -> Revision | None:
        """Parse a w:ins or w:del element into a Revision object."""
        rev_id = elem.getAttribute("w:id")
        if not rev_id:
            return None

        author = elem.getAttribute("w:author") or "Unknown"
        date_str = elem.getAttribute("w:date")

        try:
            date = datetime.fromisoformat(date_str.replace("Z", "+00:00")) if date_str else None
        except ValueError:
            date = None

        # Extract text content
        if rev_type == "insertion":
            text_elems = elem.getElementsByTagName("w:t")
        else:
            text_elems = elem.getElementsByTagName("w:delText")

        text_parts = []
        for t_elem in text_elems:
            if t_elem.firstChild:
                text_parts.append(t_elem.firstChild.data)

        return Revision(
            id=int(rev_id),
            type=rev_type,
            author=author,
            date=date,
            text="".join(text_parts),
        )

    def accept_revision(self, revision_id: int) -> bool:
        """Accept a revision by ID.

        For insertions: removes the w:ins wrapper, keeping the content.
        For deletions: removes the w:del element entirely.

        Args:
            revision_id: The w:id of the revision to accept

        Returns:
            True if revision was accepted, False if not found
        """
        # Try to find as insertion
        for ins_elem in self.editor.dom.getElementsByTagName("w:ins"):
            if ins_elem.getAttribute("w:id") == str(revision_id):
                # Accept insertion: unwrap the content
                self._unwrap_element(ins_elem)
                return True

        # Try to find as deletion
        for del_elem in self.editor.dom.getElementsByTagName("w:del"):
            if del_elem.getAttribute("w:id") == str(revision_id):
                # Accept deletion: remove the element entirely
                parent = del_elem.parentNode
                parent.removeChild(del_elem)
                return True

        return False

    def reject_revision(self, revision_id: int) -> bool:
        """Reject a revision by ID.

        For insertions: removes the w:ins element and its content entirely.
        For deletions: removes the w:del wrapper and converts w:delText back to w:t.

        Args:
            revision_id: The w:id of the revision to reject

        Returns:
            True if revision was rejected, False if not found
        """
        # Try to find as insertion
        for ins_elem in self.editor.dom.getElementsByTagName("w:ins"):
            if ins_elem.getAttribute("w:id") == str(revision_id):
                # Reject insertion: remove entirely
                parent = ins_elem.parentNode
                parent.removeChild(ins_elem)
                return True

        # Try to find as deletion
        for del_elem in self.editor.dom.getElementsByTagName("w:del"):
            if del_elem.getAttribute("w:id") == str(revision_id):
                # Reject deletion: restore the deleted text
                self._restore_deletion(del_elem)
                return True

        return False

    def accept_all(self, author: str | None = None) -> int:
        """Accept all revisions, optionally filtered by author.

        Args:
            author: If provided, only accept revisions by this author

        Returns:
            Number of revisions accepted
        """
        count = 0
        revisions = self.list_revisions(author=author)
        # Process in reverse order by ID to avoid index issues
        for rev in sorted(revisions, key=lambda r: r.id, reverse=True):
            if self.accept_revision(rev.id):
                count += 1
        return count

    def reject_all(self, author: str | None = None) -> int:
        """Reject all revisions, optionally filtered by author.

        Args:
            author: If provided, only reject revisions by this author

        Returns:
            Number of revisions rejected
        """
        count = 0
        revisions = self.list_revisions(author=author)
        # Process in reverse order by ID to avoid index issues
        for rev in sorted(revisions, key=lambda r: r.id, reverse=True):
            if self.reject_revision(rev.id):
                count += 1
        return count

    def _unwrap_element(self, elem) -> None:
        """Remove an element's wrapper, keeping its children in place."""
        parent = elem.parentNode
        while elem.firstChild:
            child = elem.firstChild
            parent.insertBefore(child, elem)
        parent.removeChild(elem)

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


def _set_xml_space_preserve(wt_elem) -> None:
    """Set xml:space='preserve' on a w:t element to preserve whitespace."""
    wt_elem.setAttribute("xml:space", "preserve")


def _escape_xml(text: str) -> str:
    """Escape text for safe XML inclusion."""
    return (
        text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )
