"""Main Document class for docx_editor.

Provides the primary user-facing API for editing Word documents with track changes
and comments.
"""

import html
import shutil
from pathlib import Path

from .comments import Comment, CommentManager
from .exceptions import WorkspaceSyncError
from .track_changes import EditOperation, Revision, RevisionManager
from .workspace import Workspace
from .xml_editor import DocxXMLEditor, ParagraphRef, build_text_map, compute_paragraph_hash


class Document:
    """Word document with track changes and comment support.

    This is the main entry point for docx_editor. It provides methods for:
    - Opening and saving documents
    - Making tracked changes (replace, delete, insert)
    - Managing comments (add, reply, resolve, delete)
    - Managing revisions (list, accept, reject)

    Example:
        from docx_editor import Document

        doc = Document.open("contract.docx")
        doc.replace("30 days", "60 days")
        doc.add_comment("Section 5", "Please review")
        doc.save()
        doc.close()
    """

    def __init__(self, workspace: Workspace):
        """Initialize Document with a workspace.

        Use Document.open() instead of calling this directly.

        Args:
            workspace: Workspace instance for the document
        """
        self._workspace = workspace
        self._closed = False

        # Create the document editor
        self._document_editor = DocxXMLEditor(
            workspace.document_xml_path,
            rsid=workspace.rsid,
            author=workspace.author,
            initials=workspace.initials,
        )

        # Initialize managers
        self._revision_manager = RevisionManager(self._document_editor)
        self._comment_manager = CommentManager(
            workspace.workspace_path,
            self._document_editor,
            workspace.author,
            workspace.initials,
        )

        # Setup tracking infrastructure
        self._setup_tracking()

    @classmethod
    def open(
        cls,
        path: str | Path,
        author: str | None = None,
        force_recreate: bool = False,
    ) -> "Document":
        """Open a Word document for editing.

        Creates a workspace in .docx/{document_stem}/ alongside the document.
        The workspace persists until close() is called.

        Args:
            path: Path to the .docx file
            author: Author name for tracked changes (defaults to system username)
            force_recreate: If True, delete existing workspace and create fresh

        Returns:
            Document instance ready for editing

        Example:
            doc = Document.open("contract.docx")
            doc = Document.open("contract.docx", author="Legal Team")
        """
        path = Path(path).resolve()

        if force_recreate:
            Workspace.delete(path)

        try:
            workspace = Workspace(path, author=author, create=True)
        except WorkspaceSyncError:
            # Document changed, offer to recreate
            if force_recreate:
                raise
            Workspace.delete(path)
            workspace = Workspace(path, author=author, create=True)

        return cls(workspace)

    @property
    def author(self) -> str:
        """Get the author name for tracked changes."""
        return self._workspace.author

    @property
    def source_path(self) -> Path:
        """Get the path to the source document."""
        return self._workspace.source_path

    # ==================== Track Changes API ====================

    def find_text(self, text: str, occurrence: int = 0):
        """Find text in the document, including across element boundaries.

        Returns match info or None if not found.
        """
        self._ensure_open()
        return self._revision_manager._find_across_boundaries(text, occurrence)

    def count_matches(self, text: str) -> int:
        """Count how many times a text string appears in the document.

        Use this before editing to verify your search text is unique,
        or to determine which occurrence to target.

        Args:
            text: Text to search for

        Returns:
            Number of occurrences found

        Example:
            count = doc.count_matches("Section 5")
            if count > 1:
                print(f"Warning: {count} matches found, specify occurrence")
        """
        self._ensure_open()
        return self._revision_manager.count_matches(text)

    def _compute_new_ref(self, old_ref: str) -> str:
        """Compute a fresh paragraph reference after mutation."""
        ref = ParagraphRef.parse(old_ref)
        p = self._document_editor.dom.getElementsByTagName("w:p")[ref.index - 1]
        new_hash = compute_paragraph_hash(p)
        return f"P{ref.index}#{new_hash}"

    def list_paragraphs(self, max_chars: int = 80) -> list[str]:
        """List all paragraphs with hash-anchored references.

        Returns a list of strings like "P1#a7b2| Introduction to the..."
        for use as stable paragraph references in editing operations.

        Args:
            max_chars: Maximum characters for the preview text (default 80).
                Use 0 to get only hash refs without preview text.

        Returns:
            List of hash-tagged paragraph preview strings
        """
        self._ensure_open()
        paragraphs = self._document_editor.dom.getElementsByTagName("w:p")
        result = []
        for i, p in enumerate(paragraphs, start=1):
            h = compute_paragraph_hash(p)
            tm = build_text_map(p)
            preview = tm.text[:max_chars]
            if len(tm.text) > max_chars:
                preview += "..."
            result.append(f"P{i}#{h}| {preview}")
        return result

    def get_visible_text(self) -> str:
        """Get the visible text of the document.

        Returns flattened text with paragraphs separated by newlines.
        Inserted text is included, deleted text is excluded.

        Returns:
            The visible text content
        """
        self._ensure_open()
        paragraphs = self._document_editor.dom.getElementsByTagName("w:p")
        parts = []
        for p in paragraphs:
            tm = build_text_map(p)
            parts.append(tm.text)
        return "\n".join(parts)

    def replace(self, find: str, replace_with: str, *, paragraph: str, occurrence: int = 0) -> str:
        """Replace text with tracked changes.

        Creates a tracked deletion of the old text and insertion of the new text.

        Args:
            find: Text to find and replace
            replace_with: Replacement text
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence within the paragraph (0 = first, 1 = second, etc.)

        Returns:
            New paragraph reference with updated hash (e.g., "P2#c3d4").
            Use this for follow-up edits without calling list_paragraphs().

        Example:
            new_ref = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
            doc.replace("other text", "new text", paragraph=new_ref)
        """
        self._ensure_open()
        self._revision_manager.replace_text(find, replace_with, occurrence=occurrence, paragraph=paragraph)
        return self._compute_new_ref(paragraph)

    def delete(self, text: str, *, paragraph: str, occurrence: int = 0) -> str:
        """Mark text as deleted with tracked changes.

        Args:
            text: Text to mark as deleted
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence within the paragraph (0 = first, 1 = second, etc.)

        Returns:
            New paragraph reference with updated hash (e.g., "P2#c3d4").

        Example:
            new_ref = doc.delete("obsolete clause", paragraph="P2#f3c1")
        """
        self._ensure_open()
        self._revision_manager.suggest_deletion(text, occurrence=occurrence, paragraph=paragraph)
        return self._compute_new_ref(paragraph)

    def insert_after(self, anchor: str, text: str, *, paragraph: str, occurrence: int = 0) -> str:
        """Insert text after anchor with tracked changes.

        Args:
            anchor: Text to find as insertion point
            text: Text to insert after the anchor
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence of anchor within the paragraph (0 = first)

        Returns:
            New paragraph reference with updated hash (e.g., "P2#c3d4").

        Example:
            new_ref = doc.insert_after("Section 5", " (as amended)", paragraph="P2#f3c1")
        """
        self._ensure_open()
        self._revision_manager.insert_text_after(anchor, text, occurrence=occurrence, paragraph=paragraph)
        return self._compute_new_ref(paragraph)

    def insert_before(self, anchor: str, text: str, *, paragraph: str, occurrence: int = 0) -> str:
        """Insert text before anchor with tracked changes.

        Args:
            anchor: Text to find as insertion point
            text: Text to insert before the anchor
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence of anchor within the paragraph (0 = first)

        Returns:
            New paragraph reference with updated hash (e.g., "P2#c3d4").

        Example:
            new_ref = doc.insert_before("Section 6", "New clause: ", paragraph="P2#f3c1")
        """
        self._ensure_open()
        self._revision_manager.insert_text_before(anchor, text, occurrence=occurrence, paragraph=paragraph)
        return self._compute_new_ref(paragraph)

    def batch_edit(self, operations: list[EditOperation]) -> list[str]:
        """Apply multiple edits atomically with upfront hash validation.

        All paragraph hashes are validated before any edits are applied.
        If any hash is stale, the entire batch is rejected. Edits are applied
        in reverse paragraph order so a single list_paragraphs() snapshot
        suffices for the entire batch.

        Args:
            operations: List of EditOperation objects

        Returns:
            List of new paragraph references with updated hashes, in input order.

        Example:
            new_refs = doc.batch_edit([
                EditOperation(action="replace", find="old", replace_with="new", paragraph="P20#a7b2"),
                EditOperation(action="delete", text="remove", paragraph="P15#f3c1"),
            ])
        """
        self._ensure_open()
        self._revision_manager.batch_edit(operations)
        return [self._compute_new_ref(op.paragraph) for op in operations]

    def rewrite_paragraph(self, ref: str, new_text: str) -> str:
        """Rewrite a paragraph's text with automatic fine-grained tracked changes.

        Diffs the current paragraph text against new_text at word level and
        generates minimal tracked insertions, deletions, and replacements.

        Args:
            ref: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            new_text: Desired new text for the paragraph

        Returns:
            The new paragraph reference with updated hash (e.g., "P2#c3d4").
            Use this ref for follow-up edits without calling list_paragraphs().

        Example:
            new_ref = doc.rewrite_paragraph("P2#f3c1", "The board shall approve the proposal.")
        """
        self._ensure_open()
        self._revision_manager.rewrite_paragraph(ref, new_text)
        return self._compute_new_ref(ref)

    def batch_rewrite(self, rewrites: list[tuple[str, str]]) -> list[str]:
        """Rewrite multiple paragraphs with upfront hash validation.

        All paragraph hashes are validated before any rewrites are applied.
        If any hash is stale, the entire batch is rejected before any changes
        are made. Once validation passes, rewrites are applied sequentially.

        Args:
            rewrites: List of (ref, new_text) tuples

        Returns:
            List of new paragraph references with updated hashes, in input order

        Example:
            refs = doc.list_paragraphs()
            new_refs = doc.batch_rewrite([
                ("P1#a7b2", "Updated first paragraph."),
                ("P3#c3d4", "Updated third paragraph."),
            ])
        """
        self._ensure_open()
        self._revision_manager.batch_rewrite(rewrites)
        return [self._compute_new_ref(ref) for ref, _ in rewrites]

    # ==================== Comments API ====================

    def add_comment(self, anchor_text: str, comment: str) -> int:
        """Add a comment anchored to specific text.

        Args:
            anchor_text: Text to attach the comment to
            comment: The comment content

        Returns:
            The comment ID

        Example:
            doc.add_comment("Section 5", "Please review this section")
        """
        self._ensure_open()
        return self._comment_manager.add_comment(anchor_text, comment)

    def reply_to_comment(self, comment_id: int, reply: str) -> int:
        """Add a reply to an existing comment.

        Args:
            comment_id: ID of the comment to reply to
            reply: The reply content

        Returns:
            The new comment ID for the reply

        Example:
            doc.reply_to_comment(0, "I agree with this change")
        """
        self._ensure_open()
        return self._comment_manager.reply_to_comment(comment_id, reply)

    def list_comments(self, author: str | None = None) -> list[Comment]:
        """List all comments in the document.

        Args:
            author: If provided, filter by author name

        Returns:
            List of Comment objects (with replies nested)

        Example:
            comments = doc.list_comments()
            for c in comments:
                print(f"{c.author}: {c.text}")
        """
        self._ensure_open()
        return self._comment_manager.list_comments(author=author)

    def resolve_comment(self, comment_id: int) -> bool:
        """Mark a comment as resolved.

        Args:
            comment_id: ID of the comment to resolve

        Returns:
            True if resolved, False if not found

        Example:
            doc.resolve_comment(0)
        """
        self._ensure_open()
        return self._comment_manager.resolve_comment(comment_id)

    def delete_comment(self, comment_id: int) -> bool:
        """Delete a comment from the document.

        Args:
            comment_id: ID of the comment to delete

        Returns:
            True if deleted, False if not found

        Example:
            doc.delete_comment(0)
        """
        self._ensure_open()
        return self._comment_manager.delete_comment(comment_id)

    # ==================== Revision Management API ====================

    def list_revisions(self, author: str | None = None) -> list[Revision]:
        """List all tracked changes in the document.

        Args:
            author: If provided, filter by author name

        Returns:
            List of Revision objects

        Example:
            revisions = doc.list_revisions()
            for r in revisions:
                print(f"{r.type}: {r.text} by {r.author}")
        """
        self._ensure_open()
        return self._revision_manager.list_revisions(author=author)

    def accept_revision(self, revision_id: int) -> bool:
        """Accept a revision by ID.

        For insertions: keeps the inserted content.
        For deletions: permanently removes the deleted content.

        Args:
            revision_id: ID of the revision to accept

        Returns:
            True if accepted, False if not found

        Example:
            doc.accept_revision(1)
        """
        self._ensure_open()
        return self._revision_manager.accept_revision(revision_id)

    def reject_revision(self, revision_id: int) -> bool:
        """Reject a revision by ID.

        For insertions: removes the inserted content.
        For deletions: restores the deleted content.

        Args:
            revision_id: ID of the revision to reject

        Returns:
            True if rejected, False if not found

        Example:
            doc.reject_revision(1)
        """
        self._ensure_open()
        return self._revision_manager.reject_revision(revision_id)

    def accept_all(self, author: str | None = None) -> int:
        """Accept all revisions.

        Args:
            author: If provided, only accept revisions by this author

        Returns:
            Number of revisions accepted

        Example:
            count = doc.accept_all()
            print(f"Accepted {count} revisions")
        """
        self._ensure_open()
        return self._revision_manager.accept_all(author=author)

    def reject_all(self, author: str | None = None) -> int:
        """Reject all revisions.

        Args:
            author: If provided, only reject revisions by this author

        Returns:
            Number of revisions rejected

        Example:
            count = doc.reject_all(author="OtherUser")
        """
        self._ensure_open()
        return self._revision_manager.reject_all(author=author)

    # ==================== Save/Close API ====================

    def save(self, path: str | Path | None = None, validate: bool = False) -> Path:
        """Save the document.

        Args:
            path: Output path (defaults to original source path)
            validate: If True, validate with LibreOffice before saving

        Returns:
            Path to the saved document

        Example:
            doc.save()  # Save to original path
            doc.save("contract_v2.docx")  # Save to new path
        """
        self._ensure_open()

        # Ensure comment relationships and content types
        self._ensure_comment_relationships()
        self._ensure_comment_content_types()

        # Save all editors
        self._document_editor.save()
        self._comment_manager.save_all()

        # Pack and save
        return self._workspace.save(destination=path, validate=validate)

    def close(self, cleanup: bool = True) -> None:
        """Close the document and clean up workspace.

        Args:
            cleanup: If True, delete the workspace folder

        Example:
            doc.close()  # Clean up workspace
            doc.close(cleanup=False)  # Keep workspace for inspection
        """
        if self._closed:
            return

        self._workspace.close(cleanup=cleanup)
        self._closed = True

    def __enter__(self) -> "Document":
        """Context manager entry."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Context manager exit - close without cleanup on error."""
        self.close(cleanup=exc_type is None)

    # ==================== Private Methods ====================

    def _ensure_open(self) -> None:
        """Raise error if document is closed."""
        if self._closed:
            raise ValueError("Document is closed")

    def _setup_tracking(self) -> None:
        """Set up tracked changes infrastructure in the document."""
        # Ensure people.xml exists
        people_path = self._workspace.word_path / "people.xml"
        if not people_path.exists():
            templates_dir = Path(__file__).parent / "ooxml" / "templates"
            shutil.copy(templates_dir / "people.xml", people_path)

        # Add content type for people.xml
        self._add_content_type_for_people()

        # Add relationship for people.xml
        self._add_relationship_for_people()

        # Update settings.xml with RSID
        self._update_settings()

        # Add author to people.xml
        self._add_author_to_people()

    def _add_content_type_for_people(self) -> None:
        """Add people.xml content type to [Content_Types].xml."""
        content_types_path = self._workspace.workspace_path / "[Content_Types].xml"
        editor = DocxXMLEditor(
            content_types_path,
            rsid=self._workspace.rsid,
            author=self._workspace.author,
        )

        # Check if already exists
        for override_elem in editor.dom.getElementsByTagName("Override"):
            if override_elem.getAttribute("PartName") == "/word/people.xml":
                return

        # Add Override element
        root = editor.dom.documentElement
        content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"
        override_xml = f'<Override PartName="/word/people.xml" ContentType="{content_type}"/>'
        editor.append_to(root, override_xml)
        editor.save()

    def _add_relationship_for_people(self) -> None:
        """Add people.xml relationship to document.xml.rels."""
        rels_path = self._workspace.word_path / "_rels" / "document.xml.rels"
        if not rels_path.exists():
            return

        editor = DocxXMLEditor(
            rels_path,
            rsid=self._workspace.rsid,
            author=self._workspace.author,
        )

        # Check if already exists
        for rel_elem in editor.dom.getElementsByTagName("Relationship"):
            if rel_elem.getAttribute("Target") == "people.xml":
                return

        root = editor.dom.documentElement
        root_tag = root.tagName
        prefix = root_tag.split(":")[0] + ":" if ":" in root_tag else ""
        next_rid = editor.get_next_rid()

        rel_type = "http://schemas.microsoft.com/office/2011/relationships/people"
        rel_xml = f'<{prefix}Relationship Id="{next_rid}" Type="{rel_type}" Target="people.xml"/>'
        editor.append_to(root, rel_xml)
        editor.save()

    def _update_settings(self) -> None:
        """Update settings.xml with RSID."""
        settings_path = self._workspace.word_path / "settings.xml"
        if not settings_path.exists():
            return

        editor = DocxXMLEditor(
            settings_path,
            rsid=self._workspace.rsid,
            author=self._workspace.author,
        )

        root = editor.get_node(tag="w:settings")
        prefix = root.tagName.split(":")[0] if ":" in root.tagName else "w"

        # Check if rsids section exists
        rsids_elements = editor.dom.getElementsByTagName(f"{prefix}:rsids")

        if not rsids_elements:
            # Add new rsids section
            rsids_xml = f"""<{prefix}:rsids>
  <{prefix}:rsidRoot {prefix}:val="{self._workspace.rsid}"/>
  <{prefix}:rsid {prefix}:val="{self._workspace.rsid}"/>
</{prefix}:rsids>"""

            # Try to insert after compat
            compat_elements = editor.dom.getElementsByTagName(f"{prefix}:compat")
            if compat_elements:
                editor.insert_after(compat_elements[0], rsids_xml)
            else:
                editor.append_to(root, rsids_xml)
        else:
            # Check if this rsid already exists
            rsids_elem = rsids_elements[0]
            rsid_exists = any(
                elem.getAttribute(f"{prefix}:val") == self._workspace.rsid
                for elem in rsids_elem.getElementsByTagName(f"{prefix}:rsid")
            )

            if not rsid_exists:
                rsid_xml = f'<{prefix}:rsid {prefix}:val="{self._workspace.rsid}"/>'
                editor.append_to(rsids_elem, rsid_xml)

        editor.save()

    def _add_author_to_people(self) -> None:
        """Add author to people.xml."""
        people_path = self._workspace.word_path / "people.xml"
        if not people_path.exists():
            return

        editor = DocxXMLEditor(
            people_path,
            rsid=self._workspace.rsid,
            author=self._workspace.author,
        )

        # Check if author already exists
        for person_elem in editor.dom.getElementsByTagName("w15:person"):
            if person_elem.getAttribute("w15:author") == self._workspace.author:
                return

        root = editor.get_node(tag="w15:people")

        escaped_author = html.escape(self._workspace.author, quote=True)
        person_xml = f"""<w15:person w15:author="{escaped_author}">
  <w15:presenceInfo w15:providerId="None" w15:userId="{escaped_author}"/>
</w15:person>"""
        editor.append_to(root, person_xml)
        editor.save()

    def _ensure_comment_relationships(self) -> None:
        """Ensure word/_rels/document.xml.rels has comment relationships."""
        # Only needed if comments.xml exists
        comments_path = self._workspace.word_path / "comments.xml"
        if not comments_path.exists():
            return

        rels_path = self._workspace.word_path / "_rels" / "document.xml.rels"
        editor = DocxXMLEditor(
            rels_path,
            rsid=self._workspace.rsid,
            author=self._workspace.author,
        )

        # Check if already exists
        for rel_elem in editor.dom.getElementsByTagName("Relationship"):
            if rel_elem.getAttribute("Target") == "comments.xml":
                return

        root = editor.dom.documentElement
        root_tag = root.tagName
        prefix = root_tag.split(":")[0] + ":" if ":" in root_tag else ""
        next_rid_num = int(editor.get_next_rid()[3:])

        # Add relationship elements
        rels = [
            (
                next_rid_num,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                "comments.xml",
            ),
            (
                next_rid_num + 1,
                "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
                "commentsExtended.xml",
            ),
            (
                next_rid_num + 2,
                "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
                "commentsIds.xml",
            ),
            (
                next_rid_num + 3,
                "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
                "commentsExtensible.xml",
            ),
        ]

        for rel_id, rel_type, target in rels:
            rel_xml = f'<{prefix}Relationship Id="rId{rel_id}" Type="{rel_type}" Target="{target}"/>'
            editor.append_to(root, rel_xml)

        editor.save()

    def _ensure_comment_content_types(self) -> None:
        """Ensure [Content_Types].xml has comment content types."""
        # Only needed if comments.xml exists
        comments_path = self._workspace.word_path / "comments.xml"
        if not comments_path.exists():
            return

        content_types_path = self._workspace.workspace_path / "[Content_Types].xml"
        editor = DocxXMLEditor(
            content_types_path,
            rsid=self._workspace.rsid,
            author=self._workspace.author,
        )

        # Check if already exists
        for override_elem in editor.dom.getElementsByTagName("Override"):
            if override_elem.getAttribute("PartName") == "/word/comments.xml":
                return

        root = editor.dom.documentElement

        # Add Override elements
        overrides = [
            (
                "/word/comments.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
            ),
            (
                "/word/commentsExtended.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
            ),
            (
                "/word/commentsIds.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
            ),
            (
                "/word/commentsExtensible.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml",
            ),
        ]

        for part_name, content_type in overrides:
            override_xml = f'<Override PartName="{part_name}" ContentType="{content_type}"/>'
            editor.append_to(root, override_xml)

        editor.save()
