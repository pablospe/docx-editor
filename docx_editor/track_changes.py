"""Track changes management for docx_editor.

Provides RevisionManager for creating and managing tracked changes (insertions/deletions).
"""

from dataclasses import dataclass
from datetime import datetime
from typing import Literal

from .exceptions import RevisionError, TextNotFoundError
from .xml_editor import DocxXMLEditor, TextMapMatch, build_text_map


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

        Args:
            text: Text to search for

        Returns:
            Number of occurrences found
        """
        matches = self.editor.find_all_nodes(tag="w:t", contains=text)
        return len(matches)

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
            start = 0
            while True:
                idx = text_map.find(text, start)
                if idx == -1:
                    break
                if current_occurrence == occurrence:
                    end = idx + len(text)
                    positions = text_map.get_nodes_for_range(idx, end)
                    if positions:
                        first_ins = positions[0].is_inside_ins
                        spans = any(p.is_inside_ins != first_ins for p in positions)
                    else:
                        spans = False
                    return TextMapMatch(
                        start=idx,
                        end=end,
                        text=text,
                        positions=positions,
                        spans_boundary=spans,
                    )
                current_occurrence += 1
                start = idx + 1
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

        # Build replacement XML
        before_text = full_text[:start_idx]
        after_text = full_text[start_idx + len(find) :]

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

        # Preserve run properties if present
        rPr_xml = ""
        rPr_elems = run.getElementsByTagName("w:rPr")
        if rPr_elems:
            rPr_xml = rPr_elems[0].toxml()

        before_text = full_text[:start_idx]
        after_text = full_text[start_idx + len(text) :]

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
        """Build per-run data for a cross-boundary match.

        Returns list of (run, rPr_xml, before_text, matched_part, after_text) tuples,
        one per unique run involved in the match. Runs are in document order.
        """
        from collections import OrderedDict

        # Group positions by their w:t node, then by parent run
        run_data = OrderedDict()  # id(run) -> (run, rPr, first_offset, last_offset, node)
        for pos in match.positions:
            run, rPr_xml = self._get_run_info(pos.node)
            if run is None:
                continue
            rid = id(run)
            if rid not in run_data:
                run_data[rid] = {
                    "run": run,
                    "rPr_xml": rPr_xml,
                    "node": pos.node,
                    "first_offset": pos.offset_in_node,
                    "last_offset": pos.offset_in_node,
                }
            else:
                run_data[rid]["last_offset"] = pos.offset_in_node

        result = []
        for info in run_data.values():
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

        # Use first run's rPr for the insertion
        first_rPr = parts[0][1]

        xml_parts = []
        for idx, (_run, rPr_xml, before, matched, after) in enumerate(parts):
            if before:
                xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before)}</w:t></w:r>")

            xml_parts.append(f"<w:del><w:r>{rPr_xml}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>")

            # Insert replacement after the last deletion
            if idx == len(parts) - 1:
                xml_parts.append(f"<w:ins><w:r>{first_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r></w:ins>")

            if after:
                xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after)}</w:t></w:r>")

        # Replace all affected runs: insert new XML before first run, remove all runs
        first_run = parts[0][0]
        new_xml = "".join(xml_parts)
        nodes = self.editor.insert_before(first_run, new_xml)

        for run, _, _, _, _ in parts:
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

        # Insert the new text before the first affected element
        ins_xml = f"<w:ins><w:r>{first_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r></w:ins>"
        new_nodes = self.editor.insert_before(ref_node, ins_xml)

        # Process each segment to delete/remove the matched text
        for is_inside_ins, positions in segments:
            if is_inside_ins:
                self._remove_from_insertion(positions)
            else:
                self._delete_regular_segment(positions)

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

        Modifies the <w:t> text node in-place. If the entire text is matched,
        removes the <w:ins> element. If partial, truncates or splits.
        """
        node = positions[0].node
        first_offset = positions[0].offset_in_node
        last_offset = positions[-1].offset_in_node

        node_text = self._get_node_text(node)
        before = node_text[:first_offset]
        after = node_text[last_offset + 1 :]

        ins_elem = self._find_ancestor(node, "w:ins")

        if not before and not after:
            # Entire node matched -- remove the <w:ins> element
            if ins_elem and ins_elem.parentNode:
                ins_elem.parentNode.removeChild(ins_elem)
        elif not before:
            # Matched at start -- truncate to after portion
            node.firstChild.data = after
        elif not after:
            # Matched at end -- truncate to before portion
            node.firstChild.data = before
        else:
            # Matched in middle -- keep "before" in current node, create new ins for "after"
            node.firstChild.data = before

            run = self._find_ancestor(node, "w:r")
            if ins_elem and run:
                rPr_xml = ""
                rPr_elems = run.getElementsByTagName("w:rPr")
                if rPr_elems:
                    rPr_xml = rPr_elems[0].toxml()

                after_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(after)}</w:t></w:r></w:ins>"
                self.editor.insert_after(ins_elem, after_xml)

    def _delete_regular_segment(self, positions: list) -> None:
        """Delete regular (non-insertion) text by wrapping in <w:del>.

        Splits the run if needed to isolate the matched portion.
        """
        node = positions[0].node
        first_offset = positions[0].offset_in_node
        last_offset = positions[-1].offset_in_node

        run, rPr_xml = self._get_run_info(node)
        if not run:
            return

        node_text = self._get_node_text(node)
        before = node_text[:first_offset]
        matched = node_text[first_offset : last_offset + 1]
        after = node_text[last_offset + 1 :]

        xml_parts = []
        if before:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before)}</w:t></w:r>")
        xml_parts.append(f"<w:del><w:r>{rPr_xml}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>")
        if after:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after)}</w:t></w:r>")

        new_xml = "".join(xml_parts)
        self.editor.insert_before(run, new_xml)
        run.parentNode.removeChild(run)

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

        xml_parts = []
        for _run, rPr_xml, before, matched, after in parts:
            if before:
                xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before)}</w:t></w:r>")

            xml_parts.append(f"<w:del><w:r>{rPr_xml}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>")

            if after:
                xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after)}</w:t></w:r>")

        first_run = parts[0][0]
        new_xml = "".join(xml_parts)
        nodes = self.editor.insert_before(first_run, new_xml)

        for run, _, _, _, _ in parts:
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
                node = positions[0].node
                first_offset = positions[0].offset_in_node
                last_offset = positions[-1].offset_in_node

                run, rPr_xml = self._get_run_info(node)
                if not run:
                    continue

                node_text = self._get_node_text(node)
                before = node_text[:first_offset]
                matched = node_text[first_offset : last_offset + 1]
                after = node_text[last_offset + 1 :]

                xml_parts = []
                if before:
                    xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before)}</w:t></w:r>")
                xml_parts.append(f"<w:del><w:r>{rPr_xml}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>")
                if after:
                    xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after)}</w:t></w:r>")

                new_xml = "".join(xml_parts)
                nodes = self.editor.insert_before(run, new_xml)
                run.parentNode.removeChild(run)

                if first_del_id == -1:
                    for n in nodes:
                        if n.nodeType == n.ELEMENT_NODE and n.tagName == "w:del":
                            first_del_id = int(n.getAttribute("w:id"))
                            break

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

        # Split the run at the anchor boundary so insertion goes inline
        full_text = elem.firstChild.data if elem.firstChild else ""
        anchor_idx = full_text.find(anchor)

        if anchor_idx == -1:
            raise TextNotFoundError(f"Anchor text not found: '{anchor}'")

        before_text = full_text[:anchor_idx]
        after_text = full_text[anchor_idx + len(anchor) :]

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

        ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"

        if position == "after":
            last_run, _ = self._get_run_info(positions[-1].node)
            if not last_run:
                return -1
            nodes = self.editor.insert_after(last_run, ins_xml)
        else:
            if not first_run:
                return -1
            nodes = self.editor.insert_before(first_run, ins_xml)

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
