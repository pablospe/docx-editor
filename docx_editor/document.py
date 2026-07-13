"""Main Document class for docx_editor.

Provides the primary user-facing API for editing Word documents with track changes
and comments.
"""

import html
import shutil
from collections.abc import Iterator
from pathlib import Path
from typing import Literal, overload
from xml.dom.minidom import Element

from .comments import Comment, CommentManager
from .exceptions import HashMismatchError, ParagraphIndexError
from .track_changes import EditOperation, EditValidationResult, Revision, RevisionManager
from .workspace import Workspace
from .xml_editor import (
    DocxXMLEditor,
    ParagraphInfo,
    ParagraphLocation,
    ParagraphRef,
    _build_table_index,
    _compute_paragraph_location,
    build_text_map,
    compute_paragraph_hash,
)


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
        workspace_dir: str | Path | None = None,
    ) -> "Document":
        """Open a Word document for editing.

        Creates a workspace holding the document's unpacked contents. By
        default the workspace lives under the platform user cache directory
        (Linux: ``$XDG_CACHE_HOME`` or ``~/.cache``; macOS:
        ``~/Library/Caches``; Windows: ``%LOCALAPPDATA%``), in a subfolder
        named ``docx-editor/<hash>`` where ``<hash>`` is derived from the
        document's absolute path. The workspace persists until close() is
        called.

        Args:
            path: Path to the .docx file
            author: Author name for tracked changes (defaults to system username)
            force_recreate: If True, delete any existing workspace (stale or
                in-sync) before opening, discarding whatever XML it holds, and
                re-unpack from the current source. Use this to recover from
                WorkspaceSyncError.
            workspace_dir: Base directory for the workspace. Overrides the
                DOCX_EDITOR_WORKSPACE_DIR environment variable and the platform
                cache default. Tilde-expanded; an empty value counts as unset.
                A relative path resolves against the document's directory, so
                ``workspace_dir=".docx"`` keeps the workspace next to the file
                (handy for debugging).

        Returns:
            Document instance ready for editing

        Raises:
            WorkspaceSyncError: If the source document was modified since the
                workspace was created. The message includes the workspace path.
                Pass force_recreate=True to discard the stale workspace and
                re-unpack from the current source.

        Example:
            doc = Document.open("contract.docx")
            doc = Document.open("contract.docx", author="Legal Team")
        """
        path = Path(path).resolve()

        if force_recreate:
            Workspace.delete(path, workspace_dir=workspace_dir)

        workspace = Workspace(path, author=author, create=True, workspace_dir=workspace_dir)
        return cls(workspace)

    @property
    def author(self) -> str:
        """Get the author name for tracked changes."""
        return self._workspace.author

    @property
    def source_path(self) -> Path:
        """Get the path to the source document."""
        return self._workspace.source_path

    @property
    def workspace_path(self) -> Path:
        """Get the path to this document's workspace folder.

        The workspace lives under the user cache by default, so this is the
        only way to locate the unpacked XML — e.g. after close(cleanup=False)
        or when a workspace is preserved because an exception was raised.
        """
        return self._workspace.workspace_path

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

    def paragraph_count(self) -> int:
        """Return the total number of paragraphs in the document.

        Cheap bounds check for pagination — avoids building the full
        :meth:`list_paragraphs` result just to learn the count.

        Returns:
            Total number of paragraphs (the highest valid 1-based ref index).
        """
        self._ensure_open()
        return len(self._document_editor.dom.getElementsByTagName("w:p"))

    def list_paragraphs(self, max_chars: int = 80, *, start: int = 1, limit: int | None = None) -> list[str]:
        """List paragraphs with hash-anchored references.

        Returns a list of strings like "P1#a7b2| Introduction to the..."
        for use as stable paragraph references in editing operations. Refs
        are **1-based global** indexes (P1, P2, …) and stay correct across
        pages — a slice starting at paragraph 51 emits "P51#…", not "P1#…".

        Args:
            max_chars: Maximum characters for the preview text (default 80).
                Must be >= 0. Use 0 to get only the hash refs (e.g. "P1#a7b2"),
                with no preview or "| " separator.
            start: 1-based index of the first paragraph to return (default 1).
                Must be >= 1. A ``start`` beyond the last paragraph yields an
                empty list.
            limit: Maximum number of paragraphs to return, or ``None`` for all
                paragraphs from ``start`` onward (default ``None``). Must be
                >= 0 when given; ``0`` yields an empty list.

        Returns:
            List of hash-tagged paragraph preview strings.

        Raises:
            ValueError: If ``max_chars`` < 0, ``start`` < 1, or ``limit`` < 0.

        Example:
            # Caller chooses the page size; ``start`` walks forward by it.
            count = doc.paragraph_count()
            page_size = 50
            for start in range(1, count + 1, page_size):
                for ref in doc.list_paragraphs(start=start, limit=page_size):
                    print(ref)
        """
        self._ensure_open()
        if max_chars < 0:
            raise ValueError(f"max_chars must be >= 0, got {max_chars}")
        result = []
        for i, p in self._iter_paragraph_slice(start, limit):
            h = compute_paragraph_hash(p)
            if max_chars == 0:
                result.append(f"P{i}#{h}")
                continue
            tm = build_text_map(p)
            preview = tm.text[:max_chars]
            if len(tm.text) > max_chars:
                preview += "..."
            result.append(f"P{i}#{h}| {preview}")
        return result

    def _iter_paragraph_slice(self, start: int, limit: int | None) -> Iterator[tuple[int, Element]]:
        """Return ``(index, paragraph_element)`` pairs for a 1-based slice.

        Shared pagination logic for :meth:`list_paragraphs` and
        :meth:`list_paragraphs_structured`. ``index`` is the 1-based global
        paragraph index, preserved across slices. Callers handle
        ``_ensure_open()`` themselves. Argument validation is eager so callers
        get a ``ValueError`` immediately, not on first iteration.

        Raises:
            ValueError: If ``start`` < 1, or ``limit`` < 0 when given.
        """
        if start < 1:
            raise ValueError(f"start must be >= 1, got {start}")
        if limit is not None and limit < 0:
            raise ValueError(f"limit must be >= 0, got {limit}")
        paragraphs = self._document_editor.dom.getElementsByTagName("w:p")
        begin = start - 1
        end = begin + limit if limit is not None else None
        return enumerate(paragraphs[begin:end], start=begin + 1)

    def list_paragraphs_structured(self, *, start: int = 1, limit: int | None = None) -> list[ParagraphInfo]:
        """List paragraphs as structured :class:`ParagraphInfo` records.

        Like :meth:`list_paragraphs`, but returns named records (index, ref,
        full text) instead of pipe-delimited preview strings. The ``text``
        field is always the full, untruncated paragraph text — there is no
        ``max_chars`` parameter. ``str(info)`` uses the same
        ``"P{i}#{hash}| {text}"`` delimiter format as :meth:`list_paragraphs`,
        but always with the full text (it matches :meth:`list_paragraphs`
        output only when that call's ``max_chars`` is large enough to avoid
        truncation).

        Refs are **1-based global** indexes (P1, P2, …) and stay correct
        across slices, with the same ``start``/``limit`` semantics as
        :meth:`list_paragraphs`.

        Args:
            start: 1-based index of the first paragraph to return (default 1).
                Must be >= 1. A ``start`` beyond the last paragraph yields an
                empty list.
            limit: Maximum number of paragraphs to return, or ``None`` for all
                paragraphs from ``start`` onward (default ``None``). Must be
                >= 0 when given; ``0`` yields an empty list.

        Returns:
            List of :class:`ParagraphInfo` records.

        Raises:
            ValueError: If ``start`` < 1, or ``limit`` < 0.

        Example:
            # Caller chooses the page size; ``start`` walks forward by it.
            page_size = 50
            for start in range(1, doc.paragraph_count() + 1, page_size):
                for info in doc.list_paragraphs_structured(start=start, limit=page_size):
                    print(info.ref, info.text)
        """
        self._ensure_open()
        result = []
        for i, p in self._iter_paragraph_slice(start, limit):
            h = compute_paragraph_hash(p)
            text = build_text_map(p).text
            result.append(ParagraphInfo(index=i, ref=f"P{i}#{h}", text=text))
        return result

    def get_paragraph(self, index: int) -> ParagraphInfo:
        """Return one paragraph as a structured :class:`ParagraphInfo` record.

        Single-item counterpart to :meth:`list_paragraphs_structured`. The
        returned record (index, hash-anchored ref, full untruncated text) is
        identical to the one that method would emit for the same paragraph.

        Args:
            index: 1-based paragraph index (P1 is ``index=1``). Must be in
                ``1 .. paragraph_count()``.

        Returns:
            :class:`ParagraphInfo` for the paragraph at ``index``.

        Raises:
            ParagraphIndexError: If ``index`` is out of range (``< 1`` or
                greater than :meth:`paragraph_count`).

        Example:
            info = doc.get_paragraph(1)
            print(info.ref, info.text)
        """
        self._ensure_open()
        paragraphs = self._document_editor.dom.getElementsByTagName("w:p")
        if index < 1 or index > len(paragraphs):
            raise ParagraphIndexError(index, len(paragraphs))
        p = paragraphs[index - 1]
        h = compute_paragraph_hash(p)
        return ParagraphInfo(index=index, ref=f"P{index}#{h}", text=build_text_map(p).text)

    def get_paragraph_location(self, ref: str) -> ParagraphLocation:
        """Return the structural location of the paragraph identified by ``ref``.

        Tells the caller whether a paragraph lives in the document body or
        inside a table cell, and — when in a table — gives its 1-based
        coordinates (table index, row, logical column, depth).

        ``location.table.col`` is the *logical-grid* column, accounting for
        ``w:gridSpan`` of preceding cells in the same row. A cell that
        visually sits in column 4 reports ``col=4`` even when an earlier
        cell in the row spans 2 grid columns.

        Args:
            ref: Paragraph reference from :meth:`list_paragraphs` (e.g.,
                ``"P3#a7b2"``).

        Returns:
            :class:`ParagraphLocation`. ``location.in_table`` is ``False``
            for body paragraphs; ``True`` when the paragraph is inside a
            ``<w:tc>`` cell (in which case ``location.table`` is populated).

        Raises:
            ValueError: If ``ref`` has an invalid format.
            ParagraphIndexError: If the paragraph index is out of range.
            HashMismatchError: If the hash no longer matches current
                paragraph content (paragraph was modified after the ref
                was captured).

        Example:
            for entry in doc.list_paragraphs():
                ref = entry.split("|")[0]
                loc = doc.get_paragraph_location(ref)
                if loc.in_table:
                    cell = loc.table
                    print(f"{ref}: table {cell.index} r{cell.row} c{cell.col}")
        """
        self._ensure_open()
        parsed = ParagraphRef.parse(ref)
        paragraphs = self._document_editor.dom.getElementsByTagName("w:p")
        if parsed.index < 1 or parsed.index > len(paragraphs):
            raise ParagraphIndexError(parsed.index, len(paragraphs))
        p = paragraphs[parsed.index - 1]
        actual_hash = compute_paragraph_hash(p)
        if actual_hash != parsed.hash:
            tm = build_text_map(p)
            preview = tm.text[:80]
            if len(tm.text) > 80:
                preview += "..."
            raise HashMismatchError(parsed.index, parsed.hash, actual_hash, preview)
        return _compute_paragraph_location(p)

    def list_paragraph_locations(self) -> list[tuple[str, ParagraphLocation]]:
        """List every paragraph paired with its structural location.

        Batch counterpart to :meth:`get_paragraph_location`: precomputes
        table indices once instead of re-scanning the table hierarchy per
        ref. Each entry is ``(ref, location)`` where
        ``ref`` is the same ``"P{i}#{hash}"`` token emitted by
        :meth:`list_paragraphs` (the part before ``|``) and accepted by
        :meth:`get_paragraph_location` and the editing methods.

        Returns:
            List of ``(ref, ParagraphLocation)`` tuples in document order.
            ``location.in_table`` is ``False`` for body paragraphs; ``True``
            when the paragraph is inside a ``<w:tc>`` cell.

        Example:
            for ref, loc in doc.list_paragraph_locations():
                if loc.in_table:
                    cell = loc.table
                    print(f"{ref}: table {cell.index} r{cell.row} c{cell.col}")
        """
        self._ensure_open()
        dom = self._document_editor.dom
        table_index = _build_table_index(dom)
        result = []
        for i, p in enumerate(dom.getElementsByTagName("w:p"), start=1):
            ref = f"P{i}#{compute_paragraph_hash(p)}"
            result.append((ref, _compute_paragraph_location(p, table_index)))
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

    @overload
    def batch_edit(self, operations: list[EditOperation], *, dry_run: Literal[False] = ...) -> list[str]: ...

    @overload
    def batch_edit(self, operations: list[EditOperation], *, dry_run: Literal[True]) -> list[EditValidationResult]: ...

    def batch_edit(
        self, operations: list[EditOperation], *, dry_run: bool = False
    ) -> list[str] | list[EditValidationResult]:
        """Apply multiple edits atomically with upfront hash validation.

        All paragraph hashes are validated before any edits are applied.
        If any hash is stale, the entire batch is rejected. Edits are applied
        in reverse paragraph order so a single list_paragraphs() snapshot
        suffices for the entire batch.

        Args:
            operations: List of EditOperation objects
            dry_run: If True, validate every operation without applying any
                edits and return a list of EditValidationResult (one per
                operation, in input order). The document is left unchanged.
                Each operation is validated independently against the current
                document; sequential effects between multiple operations on the
                same paragraph are not simulated (see
                RevisionManager.validate_batch).

        Returns:
            When dry_run is False: list of new paragraph references with updated
            hashes, in input order.
            When dry_run is True: list of EditValidationResult, one per operation.

        Example:
            new_refs = doc.batch_edit([
                EditOperation(action="replace", find="old", replace_with="new", paragraph="P20#a7b2"),
                EditOperation(action="delete", text="remove", paragraph="P15#f3c1"),
            ])

            # Pre-flight the batch (note: same-paragraph sequential effects are
            # not simulated — see the dry_run note above):
            results = doc.batch_edit(ops, dry_run=True)
            if all(r.valid for r in results):
                doc.batch_edit(ops)
        """
        self._ensure_open()
        if dry_run:
            return self._revision_manager.validate_batch(operations)
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

    def add_comment(
        self,
        anchor_text: str,
        comment: str,
        *,
        paragraph: str | None = None,
        occurrence: int = 0,
    ) -> int:
        """Add a comment anchored to specific text.

        Anchors are located with the same text-map search used by
        :meth:`count_matches` and the tracked-change edit methods, so anchors
        that span ``w:t`` run boundaries (formatting changes, smart-quote
        splits, ``w:ins`` wrappers) are found.

        Args:
            anchor_text: Text to attach the comment to.
            comment: The comment content.
            paragraph: Optional paragraph reference (e.g., ``"P3#a7b2"``) to
                scope the search. ``None`` searches the whole document.
            occurrence: Which occurrence to anchor to (0 = first).

        Returns:
            The comment ID.

        Example:
            doc.add_comment("Section 5", "Please review this section")
            doc.add_comment("foo", "Note", paragraph="P3#a7b2", occurrence=1)
        """
        self._ensure_open()
        return self._comment_manager.add_comment(anchor_text, comment, paragraph=paragraph, occurrence=occurrence)

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
