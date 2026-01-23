"""Track changes management for docx_editor.

Provides RevisionManager for creating and managing tracked changes (insertions/deletions).
"""

from dataclasses import dataclass
from datetime import datetime
from typing import Literal

from .exceptions import RevisionError, TextNotFoundError
from .xml_editor import DocxXMLEditor


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
            raise

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
            raise

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
            raise TextNotFoundError(f"Anchor text not found: '{anchor}'") from None

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

        # Create insertion XML
        ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"

        # Insert before or after the run
        if position == "after":
            nodes = self.editor.insert_after(run, ins_xml)
        else:
            nodes = self.editor.insert_before(run, ins_xml)

        # Find the insertion node to get its ID
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
