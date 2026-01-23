"""Comment management for docx_editor.

Provides CommentManager for creating and managing document comments.
"""

import html
import shutil
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

from .exceptions import CommentError, TextNotFoundError
from .xml_editor import DocxXMLEditor, _generate_hex_id

# Path to template files
TEMPLATE_DIR = Path(__file__).parent / "ooxml" / "templates"


@dataclass
class Comment:
    """Represents a document comment."""

    id: int
    text: str
    author: str
    date: datetime | None
    resolved: bool = False
    replies: list["Comment"] = field(default_factory=list)

    def __repr__(self) -> str:
        status = "[RESOLVED] " if self.resolved else ""
        reply_count = f" ({len(self.replies)} replies)" if self.replies else ""
        return f"Comment({self.id}: {status}'{self.text[:30]}...' by {self.author}{reply_count})"


class CommentManager:
    """Manages comments in a Word document.

    Handles the complex task of managing comments across 5 related XML files:
    - comments.xml: Main comment content
    - commentsExtended.xml: Threading information
    - commentsIds.xml: Durable IDs
    - commentsExtensible.xml: Extended properties
    - document.xml: Comment range markers
    """

    def __init__(
        self,
        workspace_path: Path,
        document_editor: DocxXMLEditor,
        author: str,
        initials: str,
    ):
        """Initialize with workspace path and document editor.

        Args:
            workspace_path: Path to the unpacked workspace folder
            document_editor: DocxXMLEditor for word/document.xml
            author: Author name for new comments
            initials: Author initials for new comments
        """
        self.workspace_path = workspace_path
        self.word_path = workspace_path / "word"
        self.document_editor = document_editor
        self.author = author
        self.initials = initials

        # Comment file paths
        self.comments_path = self.word_path / "comments.xml"
        self.comments_extended_path = self.word_path / "commentsExtended.xml"
        self.comments_ids_path = self.word_path / "commentsIds.xml"
        self.comments_extensible_path = self.word_path / "commentsExtensible.xml"

        # Cache for lazy-loaded editors
        self._editors: dict[str, DocxXMLEditor] = {}

        # Load existing comments for reply support
        self.existing_comments = self._load_existing_comments()
        self.next_comment_id = self._get_next_comment_id()

    def _get_editor(self, xml_path: Path) -> DocxXMLEditor:
        """Get or create an editor for the specified XML file."""
        path_str = str(xml_path)
        if path_str not in self._editors:
            self._editors[path_str] = DocxXMLEditor(
                xml_path,
                rsid=self.document_editor.rsid,
                author=self.author,
                initials=self.initials,
            )
        return self._editors[path_str]

    def add_comment(self, anchor_text: str, comment_text: str) -> int:
        """Add a comment anchored to specific text.

        Args:
            anchor_text: Text to attach the comment to
            comment_text: The comment content

        Returns:
            The comment ID

        Raises:
            TextNotFoundError: If the anchor text is not found
        """
        # Find the anchor element
        try:
            elem = self.document_editor.get_node(tag="w:t", contains=anchor_text)
        except Exception:
            raise TextNotFoundError(f"Anchor text not found: '{anchor_text}'") from None

        # Get the parent run and paragraph
        run = elem.parentNode
        while run and run.nodeName != "w:r":
            run = run.parentNode

        para = run
        while para and para.nodeName != "w:p":
            para = para.parentNode

        if not run or not para:
            raise CommentError("Could not find parent run/paragraph")

        comment_id = self.next_comment_id
        para_id = _generate_hex_id()
        durable_id = _generate_hex_id()
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        # Add comment range markers to document.xml
        self.document_editor.insert_before(run, self._comment_range_start_xml(comment_id))
        self.document_editor.append_to(para, self._comment_range_end_xml(comment_id))

        # Add to all comment XML files
        self._add_to_comments_xml(comment_id, para_id, comment_text, timestamp)
        self._add_to_comments_extended_xml(para_id, parent_para_id=None)
        self._add_to_comments_ids_xml(para_id, durable_id)
        self._add_to_comments_extensible_xml(durable_id)

        # Track for reply support
        self.existing_comments[comment_id] = {"para_id": para_id}
        self.next_comment_id += 1

        return comment_id

    def reply_to_comment(self, parent_comment_id: int, reply_text: str) -> int:
        """Add a reply to an existing comment.

        Args:
            parent_comment_id: The ID of the comment to reply to
            reply_text: The reply content

        Returns:
            The new comment ID for the reply

        Raises:
            CommentError: If the parent comment is not found
        """
        if parent_comment_id not in self.existing_comments:
            raise CommentError(f"Parent comment with id={parent_comment_id} not found")

        parent_info = self.existing_comments[parent_comment_id]
        comment_id = self.next_comment_id
        para_id = _generate_hex_id()
        durable_id = _generate_hex_id()
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        # Find parent comment markers in document.xml
        parent_start_elem = self.document_editor.get_node(
            tag="w:commentRangeStart", attrs={"w:id": str(parent_comment_id)}
        )
        parent_ref_elem = self.document_editor.get_node(
            tag="w:commentReference", attrs={"w:id": str(parent_comment_id)}
        )

        # Add reply markers after parent markers
        self.document_editor.insert_after(parent_start_elem, self._comment_range_start_xml(comment_id))

        parent_ref_run = parent_ref_elem.parentNode
        self.document_editor.insert_after(parent_ref_run, f'<w:commentRangeEnd w:id="{comment_id}"/>')
        self.document_editor.insert_after(parent_ref_run, self._comment_ref_run_xml(comment_id))

        # Add to all comment XML files
        self._add_to_comments_xml(comment_id, para_id, reply_text, timestamp)
        self._add_to_comments_extended_xml(para_id, parent_para_id=parent_info["para_id"])
        self._add_to_comments_ids_xml(para_id, durable_id)
        self._add_to_comments_extensible_xml(durable_id)

        # Track for further replies
        self.existing_comments[comment_id] = {"para_id": para_id}
        self.next_comment_id += 1

        return comment_id

    def list_comments(self, author: str | None = None) -> list[Comment]:
        """List all comments in the document.

        Args:
            author: If provided, filter by author name

        Returns:
            List of Comment objects (with replies nested)
        """
        if not self.comments_path.exists():
            return []

        editor = self._get_editor(self.comments_path)
        comments_dict: dict[int, Comment] = {}
        parent_map: dict[str, str] = {}  # para_id -> parent_para_id

        # Build parent map from commentsExtended.xml
        if self.comments_extended_path.exists():
            ext_editor = self._get_editor(self.comments_extended_path)
            for ex_elem in ext_editor.dom.getElementsByTagName("w15:commentEx"):
                para_id = ex_elem.getAttribute("w15:paraId")
                parent_para_id = ex_elem.getAttribute("w15:paraIdParent")
                if para_id:
                    parent_map[para_id] = parent_para_id

        # Parse all comments
        for comment_elem in editor.dom.getElementsByTagName("w:comment"):
            comment = self._parse_comment(comment_elem)
            if comment and (author is None or comment.author == author):
                # Check if resolved
                para_id = self._get_comment_para_id(comment_elem)
                if para_id and self.comments_extended_path.exists():
                    ext_editor = self._get_editor(self.comments_extended_path)
                    for ex_elem in ext_editor.dom.getElementsByTagName("w15:commentEx"):
                        if ex_elem.getAttribute("w15:paraId") == para_id:
                            comment.resolved = ex_elem.getAttribute("w15:done") == "1"
                            break

                comments_dict[comment.id] = comment

        # Build reply tree
        para_to_id: dict[str, int] = {}
        for comment_id, info in self.existing_comments.items():
            para_to_id[info["para_id"]] = comment_id

        # Nest replies
        root_comments = []
        for comment_id, comment in comments_dict.items():
            para_id = self.existing_comments.get(comment_id, {}).get("para_id")
            if para_id:
                parent_para = parent_map.get(para_id)
                if parent_para and parent_para in para_to_id:
                    parent_id = para_to_id[parent_para]
                    if parent_id in comments_dict:
                        comments_dict[parent_id].replies.append(comment)
                        continue
            root_comments.append(comment)

        return sorted(root_comments, key=lambda c: c.id)

    def resolve_comment(self, comment_id: int) -> bool:
        """Mark a comment as resolved.

        Args:
            comment_id: The comment ID to resolve

        Returns:
            True if resolved, False if not found
        """
        if comment_id not in self.existing_comments:
            return False

        para_id = self.existing_comments[comment_id]["para_id"]

        if not self.comments_extended_path.exists():
            return False

        editor = self._get_editor(self.comments_extended_path)
        for ex_elem in editor.dom.getElementsByTagName("w15:commentEx"):
            if ex_elem.getAttribute("w15:paraId") == para_id:
                ex_elem.setAttribute("w15:done", "1")
                return True

        return False

    def delete_comment(self, comment_id: int) -> bool:
        """Delete a comment from the document.

        Args:
            comment_id: The comment ID to delete

        Returns:
            True if deleted, False if not found
        """
        if comment_id not in self.existing_comments:
            return False

        para_id = self.existing_comments[comment_id]["para_id"]

        # Remove from document.xml
        try:
            range_start = self.document_editor.get_node(tag="w:commentRangeStart", attrs={"w:id": str(comment_id)})
            range_start.parentNode.removeChild(range_start)
        except Exception:
            pass

        try:
            range_end = self.document_editor.get_node(tag="w:commentRangeEnd", attrs={"w:id": str(comment_id)})
            range_end.parentNode.removeChild(range_end)
        except Exception:
            pass

        try:
            ref = self.document_editor.get_node(tag="w:commentReference", attrs={"w:id": str(comment_id)})
            # Remove the parent run containing the reference
            if ref.parentNode and ref.parentNode.nodeName == "w:r":
                ref.parentNode.parentNode.removeChild(ref.parentNode)
            else:
                ref.parentNode.removeChild(ref)
        except Exception:
            pass

        # Remove from comments.xml
        if self.comments_path.exists():
            editor = self._get_editor(self.comments_path)
            for comment_elem in editor.dom.getElementsByTagName("w:comment"):
                if comment_elem.getAttribute("w:id") == str(comment_id):
                    comment_elem.parentNode.removeChild(comment_elem)
                    break

        # Remove from commentsExtended.xml
        if self.comments_extended_path.exists():
            editor = self._get_editor(self.comments_extended_path)
            for ex_elem in editor.dom.getElementsByTagName("w15:commentEx"):
                if ex_elem.getAttribute("w15:paraId") == para_id:
                    ex_elem.parentNode.removeChild(ex_elem)
                    break

        # Remove from commentsIds.xml
        if self.comments_ids_path.exists():
            editor = self._get_editor(self.comments_ids_path)
            for id_elem in editor.dom.getElementsByTagName("w16cid:commentId"):
                if id_elem.getAttribute("w16cid:paraId") == para_id:
                    id_elem.parentNode.removeChild(id_elem)
                    break

        # Remove from commentsExtensible.xml
        if self.comments_extensible_path.exists():
            # Need durable_id, which is in commentsIds.xml - already removed
            # Just leave it, or we'd need to track durable_id
            pass

        del self.existing_comments[comment_id]
        return True

    def save_all(self) -> None:
        """Save all modified XML files."""
        for editor in self._editors.values():
            editor.save()

    # ==================== Private: Loading ====================

    def _get_next_comment_id(self) -> int:
        """Get the next available comment ID."""
        if not self.comments_path.exists():
            return 0

        editor = self._get_editor(self.comments_path)
        max_id = -1
        for comment_elem in editor.dom.getElementsByTagName("w:comment"):
            comment_id = comment_elem.getAttribute("w:id")
            if comment_id:
                try:
                    max_id = max(max_id, int(comment_id))
                except ValueError:
                    pass
        return max_id + 1

    def _load_existing_comments(self) -> dict[int, dict]:
        """Load existing comments for reply support."""
        if not self.comments_path.exists():
            return {}

        editor = self._get_editor(self.comments_path)
        existing = {}

        for comment_elem in editor.dom.getElementsByTagName("w:comment"):
            comment_id = comment_elem.getAttribute("w:id")
            if not comment_id:
                continue

            para_id = self._get_comment_para_id(comment_elem)
            if not para_id:
                continue

            existing[int(comment_id)] = {"para_id": para_id}

        return existing

    def _get_comment_para_id(self, comment_elem) -> str | None:
        """Get the para_id from a comment element."""
        for p_elem in comment_elem.getElementsByTagName("w:p"):
            para_id = p_elem.getAttribute("w14:paraId")
            if para_id:
                return para_id
        return None

    def _parse_comment(self, comment_elem) -> Comment | None:
        """Parse a w:comment element into a Comment object."""
        comment_id = comment_elem.getAttribute("w:id")
        if not comment_id:
            return None

        author = comment_elem.getAttribute("w:author") or "Unknown"
        date_str = comment_elem.getAttribute("w:date")

        try:
            date = datetime.fromisoformat(date_str.replace("Z", "+00:00")) if date_str else None
        except ValueError:
            date = None

        # Extract text content from w:t elements
        text_parts = []
        for t_elem in comment_elem.getElementsByTagName("w:t"):
            if t_elem.firstChild:
                text_parts.append(t_elem.firstChild.data)

        return Comment(
            id=int(comment_id),
            text="".join(text_parts),
            author=author,
            date=date,
        )

    # ==================== Private: XML File Creation ====================

    def _ensure_comment_file(self, path: Path, template_name: str) -> None:
        """Ensure a comment XML file exists, creating from template if needed."""
        if not path.exists():
            shutil.copy(TEMPLATE_DIR / template_name, path)

    def _add_to_comments_xml(self, comment_id: int, para_id: str, text: str, timestamp: str) -> None:
        """Add a single comment to comments.xml."""
        self._ensure_comment_file(self.comments_path, "comments.xml")

        editor = self._get_editor(self.comments_path)
        root = editor.get_node(tag="w:comments")

        escaped_text = html.escape(text)
        comment_xml = f"""<w:comment w:id="{comment_id}">
  <w:p w14:paraId="{para_id}" w14:textId="77777777">
    <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>
    <w:r><w:rPr><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>{escaped_text}</w:t></w:r>
  </w:p>
</w:comment>"""
        editor.append_to(root, comment_xml)

    def _add_to_comments_extended_xml(self, para_id: str, parent_para_id: str | None) -> None:
        """Add a single comment to commentsExtended.xml."""
        self._ensure_comment_file(self.comments_extended_path, "commentsExtended.xml")

        editor = self._get_editor(self.comments_extended_path)
        root = editor.get_node(tag="w15:commentsEx")

        if parent_para_id:
            xml = f'<w15:commentEx w15:paraId="{para_id}" w15:paraIdParent="{parent_para_id}" w15:done="0"/>'
        else:
            xml = f'<w15:commentEx w15:paraId="{para_id}" w15:done="0"/>'
        editor.append_to(root, xml)

    def _add_to_comments_ids_xml(self, para_id: str, durable_id: str) -> None:
        """Add a single comment to commentsIds.xml."""
        self._ensure_comment_file(self.comments_ids_path, "commentsIds.xml")

        editor = self._get_editor(self.comments_ids_path)
        root = editor.get_node(tag="w16cid:commentsIds")

        xml = f'<w16cid:commentId w16cid:paraId="{para_id}" w16cid:durableId="{durable_id}"/>'
        editor.append_to(root, xml)

    def _add_to_comments_extensible_xml(self, durable_id: str) -> None:
        """Add a single comment to commentsExtensible.xml."""
        self._ensure_comment_file(self.comments_extensible_path, "commentsExtensible.xml")

        editor = self._get_editor(self.comments_extensible_path)
        root = editor.get_node(tag="w16cex:commentsExtensible")

        xml = f'<w16cex:commentExtensible w16cex:durableId="{durable_id}"/>'
        editor.append_to(root, xml)

    # ==================== Private: XML Fragments ====================

    def _comment_range_start_xml(self, comment_id: int) -> str:
        """Generate XML for comment range start."""
        return f'<w:commentRangeStart w:id="{comment_id}"/>'

    def _comment_range_end_xml(self, comment_id: int) -> str:
        """Generate XML for comment range end with reference run."""
        return f"""<w:commentRangeEnd w:id="{comment_id}"/>
<w:r>
  <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
  <w:commentReference w:id="{comment_id}"/>
</w:r>"""

    def _comment_ref_run_xml(self, comment_id: int) -> str:
        """Generate XML for comment reference run."""
        return f"""<w:r>
  <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
  <w:commentReference w:id="{comment_id}"/>
</w:r>"""
