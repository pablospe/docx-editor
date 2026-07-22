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
from .exceptions import DocumentClosedError, HashMismatchError, ParagraphIndexError
from .track_changes import (
    EditOperation,
    EditResult,
    EditValidationResult,
    Revision,
    RevisionManager,
    SearchResult,
)
from .workspace import Workspace
from .xml_editor import (
    DocxXMLEditor,
    ListItem,
    ParagraphInfo,
    ParagraphLocation,
    ParagraphRef,
    XMLEditor,
    _build_style_numbering_map,
    _build_style_outline_map,
    _build_table_index,
    _compute_heading_paths,
    _compute_paragraph_location,
    _compute_section_indexes,
    build_text_map,
    compute_paragraph_hash,
)

# Default page size for list_paragraphs / list_paragraphs_structured. Bounds
# the output of a bare call on large documents; pass limit=None for everything.
_DEFAULT_LIST_LIMIT = 200


def _require_ref_string(paragraph: str) -> None:
    """Reject non-string paragraph refs before they can silently select the
    RevisionManager's document-wide search branch (its ``paragraph=None``
    mode is intentional at that layer, not at this one)."""
    if not isinstance(paragraph, str):
        raise ValueError(f"'paragraph' must be a paragraph ref string like 'P3#a7b2', got {type(paragraph).__name__}")


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
        ref = doc.find_text("30 days").paragraph_ref
        result = doc.replace("30 days", "60 days", paragraph=ref)
        doc.add_comment("Section 5", "Please review")

        # Every edit's revisions form a group — accept/reject them as a unit.
        # Group ids are per-open-Document (renumbered on each open; revisions
        # already in the file get inferred groups reconstructed at parse
        # time), so always use ids from this session:
        result = doc.rewrite_paragraph(result, "The board shall approve.")
        doc.reject_group(result.group_id)  # undo the whole rewrite

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

        # Create the document editor. Every post-open workspace write goes
        # through an editor (or the comment manager's template copies), so
        # routing mark_dirty through their write-ahead hooks enforces the
        # dirty-flag contract mechanically at the write layer.
        self._document_editor = DocxXMLEditor(
            workspace.document_xml_path,
            rsid=workspace.rsid,
            author=workspace.author,
            initials=workspace.initials,
            on_save=workspace.mark_dirty,
        )

        # Initialize managers
        self._revision_manager = RevisionManager(self._document_editor)
        self._comment_manager = CommentManager(
            workspace.workspace_path,
            self._document_editor,
            workspace.author,
            workspace.initials,
            on_write=workspace.mark_dirty,
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
            author: Author name for tracked changes. None (the default) uses
                the system username; otherwise it must be a non-empty string.
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
            ValueError: If ``author`` is neither None nor a non-empty string.
            DocumentNotFoundError: If the path does not exist.
            InvalidDocumentError: If the path is not a valid .docx document:
                wrong suffix, a directory, an empty/truncated file or not a
                zip archive, malformed XML in a part, or the required
                word/document.xml part is missing. Carries ``path`` (the
                input that failed validation).
            WorkspaceError: If the workspace cannot be created (unwritable base,
                undeterminable home directory) or an existing workspace was
                unpacked from a different document.
            WorkspaceSyncError: If the source document was modified since the
                workspace was created, or if a leftover workspace holds unsaved
                changes from a previous session (e.g. it saved to a different
                path, or a save failed, and the session never closed cleanly).
                The message includes the workspace path. Pass
                force_recreate=True to discard the workspace and re-unpack from
                the current source.
            WorkspaceLockedError: If a live session — another process, or an
                unclosed Document in this one — already holds the document's
                workspace. Close the other session, or pass force_recreate=True
                to take the workspace over, discarding its unsaved edits.

        Example:
            doc = Document.open("contract.docx")
            doc = Document.open("contract.docx", author="Legal Team")
        """
        # Deliberately not resolved: Workspace resolves internally for source_path,
        # but it also needs the name the caller actually opened. If that is a symlink,
        # it is the name Word was told to open, and therefore the name its ~$ owner
        # file sits beside — the save-time guard has no other way to find that stub.
        path = Path(path)

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

    def find_text(self, text: str, occurrence: int = 0, paragraph: str | None = None) -> SearchResult | None:
        """Find text in the document, including across element boundaries.

        Args:
            text: Text to search for (must be non-empty)
            occurrence: Which occurrence (0 = first, 1 = second, etc.).
                Counts document-wide when ``paragraph`` is None, and within
                the paragraph when scoped — the same convention as the edit
                methods and ``add_comment``.
            paragraph: Optional paragraph reference (e.g., "P2#f3c1") to scope
                the search to one paragraph. None searches the whole document.

        Returns:
            A SearchResult, or None if the text (or that occurrence) is not
            found. Fields:

            - ``start`` / ``end``: character offsets of the match in the
              *containing paragraph's* visible text (not document-wide).
            - ``text``: the matched text.
            - ``paragraph_ref``: hash-anchored ref like "P3#a7b2", directly
              usable as the ``paragraph=`` argument of follow-up edits. Valid
              until that paragraph is edited.
            - ``paragraph_occurrence``: occurrence index of this match within
              its paragraph — pass as the ``occurrence=`` of a follow-up edit
              (edit methods count occurrences within the paragraph, not
              document-wide).
            - ``spans_revision``: True if the match crosses a tracked-revision
              boundary (e.g. part of it is inside a tracked insertion).
            - ``paragraph_index``: 1-based index of the containing paragraph —
              the same integer embedded in ``paragraph_ref``, provided so you
              never need to string-parse the ref.

        Example:
            match = doc.find_text("30 days")
            if match:
                doc.replace(
                    "30 days",
                    "60 days",
                    paragraph=match.paragraph_ref,
                    occurrence=match.paragraph_occurrence,
                )

        To enumerate every hit in one call, use :meth:`find_all`.

        Raises:
            ValueError: If ``text`` is not a non-empty string, ``occurrence``
                is not a non-negative integer (None included — the default is
                0, not None), or ``paragraph`` is malformed.
            ParagraphIndexError: If ``paragraph``'s index is out of range.
            HashMismatchError: If ``paragraph``'s hash is stale.
        """
        self._ensure_open()
        return self._revision_manager.find_text(text, occurrence, paragraph=paragraph)

    def find_all(self, text: str, paragraph: str | None = None) -> list[SearchResult]:
        """Find every match of ``text``, in document order.

        One call replaces the N+1 ``find_text`` probes needed to enumerate N
        hits, and each result carries exactly what a follow-up edit needs:
        pass ``paragraph_ref`` as ``paragraph=`` and ``paragraph_occurrence``
        as ``occurrence=`` to target that specific match.

        Args:
            text: Text to search for (must be non-empty).
            paragraph: Optional paragraph reference (e.g., "P2#f3c1") to scope
                the search. None searches the whole document.

        Returns:
            A list of SearchResult (see :meth:`find_text` for the fields),
            empty when nothing matches — no-match is not an error here.

        Raises:
            ValueError: If ``text`` is not a non-empty string, or
                ``paragraph`` is malformed.
            ParagraphIndexError: If ``paragraph``'s index is out of range.
            HashMismatchError: If ``paragraph``'s hash is stale.

        Example:
            # Edit every match in one atomic batch. reversed() puts
            # same-paragraph ops in the required descending occurrence order,
            # so this is safe however the matches are distributed:
            ops = [
                EditOperation.replace(
                    r.text,
                    "60 days",
                    paragraph=r.paragraph_ref,
                    occurrence=r.paragraph_occurrence,
                )
                for r in reversed(doc.find_all("30 days"))
            ]
            doc.batch_edit(ops)

        Editing one match at a time also works when every paragraph holds at
        most one match; with several matches in one paragraph, an edit
        invalidates the paragraph's remaining refs and shifts the occurrence
        numbers of the matches after it, so either re-run find_all after each
        edit or batch the same-paragraph ops in *descending* occurrence order
        as above — an edit never shifts the matches before it. (Ascending
        order mis-targets; descending is not valid for search strings that
        overlap themselves, e.g. "aa" in "aaaa".)
        """
        self._ensure_open()
        return self._revision_manager.find_all(text, paragraph=paragraph)

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

    def _compute_new_ref(self, old_ref: str, paragraphs: list | None = None) -> str:
        """Compute a fresh paragraph reference after mutation.

        ``paragraphs`` is an optional pre-fetched <w:p> list so batch callers
        pay for one full-DOM walk per batch; None fetches fresh.
        """
        ref = ParagraphRef.parse(old_ref)
        if paragraphs is None:
            paragraphs = self._document_editor.dom.getElementsByTagName("w:p")
        p = paragraphs[ref.index - 1]
        new_hash = compute_paragraph_hash(p)
        return f"P{ref.index}#{new_hash}"

    def _edit_result(self, old_ref: str, group_id: int | None, paragraphs: list | None = None) -> EditResult:
        """Build an EditResult from a mutated paragraph's old ref and its group."""
        revision_ids = self._revision_manager.group_revisions(group_id) if group_id is not None else ()
        return EditResult(self._compute_new_ref(old_ref, paragraphs), group_id=group_id, revision_ids=revision_ids)

    def paragraph_count(self) -> int:
        """Return the total number of paragraphs in the document.

        Cheap bounds check for pagination — avoids building the full
        :meth:`list_paragraphs` result just to learn the count.

        Returns:
            Total number of paragraphs (the highest valid 1-based ref index).
        """
        self._ensure_open()
        return len(self._document_editor.dom.getElementsByTagName("w:p"))

    def list_paragraphs(
        self, max_chars: int = 80, *, start: int = 1, limit: int | None = _DEFAULT_LIST_LIMIT
    ) -> list[str]:
        """List paragraphs with hash-anchored references.

        Returns a list of strings like "P1#a7b2| Introduction to the..."
        for use as stable paragraph references in editing operations. Refs
        are **1-based global** indexes (P1, P2, …) and stay correct across
        pages — a slice starting at paragraph 51 emits "P51#…", not "P1#…".

        Note:
            Changed in 0.6.1: a bare call now returns at most 200 paragraphs
            (it used to return all of them). Whenever paragraphs remain beyond
            the returned window, the last list entry is a truncation notice
            instead of a paragraph, e.g. ``"... 50 more paragraphs; use
            start=201 or limit=None"``. Notice lines always start with
            ``"..."`` and never match the ``P{i}#{hash}`` ref shape, so
            ref-consuming code can filter them out. Pass ``limit=None`` to
            restore the full listing.

        Args:
            max_chars: Maximum characters for the preview text (default 80).
                Must be >= 0. Use 0 to get only the hash refs (e.g. "P1#a7b2"),
                with no preview or "| " separator.
            start: 1-based index of the first paragraph to return (default 1).
                Must be >= 1. A ``start`` beyond the last paragraph yields an
                empty list.
            limit: Maximum number of paragraphs to return (default 200), or
                ``None`` for all paragraphs from ``start`` onward. Must be
                >= 0 when given.

        Returns:
            List of hash-tagged paragraph preview strings, plus one trailing
            ``"... N more paragraphs; use start=… or limit=None"`` notice
            when the window did not reach the end of the document.

        Raises:
            ValueError: If ``max_chars``, ``start``, or ``limit`` is not an
                integer (bool included), ``max_chars`` < 0, ``start`` < 1, or
                ``limit`` < 0.

        Example:
            # Walk a large document page by page; the trailing notice on each
            # page tells you the next start.
            page = doc.list_paragraphs()                # P1..P200 + notice
            page = doc.list_paragraphs(start=201)       # P201.. and so on
            everything = doc.list_paragraphs(limit=None)  # no cap, no notice
        """
        self._ensure_open()
        if isinstance(max_chars, bool) or not isinstance(max_chars, int):
            raise ValueError(f"'max_chars' must be an integer, got {max_chars!r}")
        if max_chars < 0:
            raise ValueError(f"max_chars must be >= 0, got {max_chars}")
        result = []
        # _iter_paragraph_slice validates start/limit eagerly — call it before
        # the arithmetic below can hit a non-int start.
        slice_pairs = self._iter_paragraph_slice(start, limit)
        last_index = start - 1  # highest index emitted; start-1 when the slice is empty
        for i, p in slice_pairs:
            last_index = i
            h = compute_paragraph_hash(p)
            if max_chars == 0:
                result.append(f"P{i}#{h}")
                continue
            tm = build_text_map(p)
            preview = tm.text[:max_chars]
            if len(tm.text) > max_chars:
                preview += "..."
            result.append(f"P{i}#{h}| {preview}")
        remaining = self.paragraph_count() - last_index
        if remaining > 0:
            noun = "paragraph" if remaining == 1 else "paragraphs"
            result.append(f"... {remaining} more {noun}; use start={last_index + 1} or limit=None")
        return result

    def _iter_paragraph_slice(self, start: int, limit: int | None) -> Iterator[tuple[int, Element]]:
        """Return ``(index, paragraph_element)`` pairs for a 1-based slice.

        Shared pagination logic for :meth:`list_paragraphs` and
        :meth:`list_paragraphs_structured`. ``index`` is the 1-based global
        paragraph index, preserved across slices. Callers handle
        ``_ensure_open()`` themselves. Argument validation is eager so callers
        get a ``ValueError`` immediately, not on first iteration.

        Raises:
            ValueError: If ``start`` or ``limit`` is not an integer (bool
                included), ``start`` < 1, or ``limit`` < 0 when given.
        """
        if isinstance(start, bool) or not isinstance(start, int):
            raise ValueError(f"'start' must be an integer, got {start!r}")
        if start < 1:
            raise ValueError(f"start must be >= 1, got {start}")
        if limit is not None:
            if isinstance(limit, bool) or not isinstance(limit, int):
                raise ValueError(f"'limit' must be an integer or None, got {limit!r}")
            if limit < 0:
                raise ValueError(f"limit must be >= 0, got {limit}")
        paragraphs = self._document_editor.dom.getElementsByTagName("w:p")
        begin = start - 1
        end = begin + limit if limit is not None else None
        return enumerate(paragraphs[begin:end], start=begin + 1)

    def list_paragraphs_structured(
        self, *, start: int = 1, limit: int | None = _DEFAULT_LIST_LIMIT
    ) -> list[ParagraphInfo]:
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

        Note:
            Changed in 0.6.1: a bare call now returns at most 200 records (it
            used to return all of them). Unlike :meth:`list_paragraphs`, **no
            truncation notice is appended** — every entry is a
            :class:`ParagraphInfo`, never a string — so a capped result is
            silent. To detect truncation, check whether the last record's
            ``index`` is still below :meth:`paragraph_count` (robust for any
            ``start``; with ``start=1``, comparing ``len(result)`` works too).
            Pass ``limit=None`` for the full listing.

        Args:
            start: 1-based index of the first paragraph to return (default 1).
                Must be >= 1. A ``start`` beyond the last paragraph yields an
                empty list.
            limit: Maximum number of paragraphs to return (default 200), or
                ``None`` for all paragraphs from ``start`` onward. Must be
                >= 0 when given; ``0`` yields an empty list.

        Returns:
            List of :class:`ParagraphInfo` records (no notice entries).

        Raises:
            ValueError: If ``start`` or ``limit`` is not an integer (bool
                included), ``start`` < 1, or ``limit`` < 0.

        Example:
            infos = doc.list_paragraphs_structured()  # bounded: at most 200
            if infos and infos[-1].index < doc.paragraph_count():
                ...  # truncated — continue from start=infos[-1].index + 1
            everything = doc.list_paragraphs_structured(limit=None)
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

    def context(self, ref: str, window: int = 2) -> list[ParagraphInfo]:
        """Return the paragraphs surrounding ``ref``, in document order.

        Fetches the referenced paragraph plus up to ``window`` paragraphs on
        each side (fewer at the document edges) — the "show me what's around
        this match" helper for search results: pass a
        :class:`~docx_editor.track_changes.SearchResult`'s ``paragraph_ref``
        straight in. The records are identical to what
        :meth:`list_paragraphs_structured` would emit for the same indexes.

        Args:
            ref: Paragraph reference (e.g., "P3#a7b2") from
                :meth:`list_paragraphs`, :meth:`find_text`/:meth:`find_all`,
                or an edit result.
            window: Number of paragraphs to include on *each side* of the
                referenced one (default 2, so up to 5 records). Must be >= 0;
                ``0`` returns just the referenced paragraph. Clamped at the
                document edges — no padding, no wrap-around.

        Returns:
            List of :class:`ParagraphInfo` records covering
            ``max(1, i - window) .. min(paragraph_count(), i + window)``,
            where ``i`` is the referenced paragraph's index.

        Raises:
            ValueError: If ``ref`` has an invalid format, or ``window`` is
                not an integer (bool included) or is < 0.
            ParagraphIndexError: If the paragraph index is out of range.
            HashMismatchError: If the hash no longer matches current
                paragraph content (paragraph was modified after the ref
                was captured).

        Example:
            match = doc.find_text("Termination")
            for info in doc.context(match.paragraph_ref, window=2):
                print(info)
        """
        self._ensure_open()
        if isinstance(window, bool) or not isinstance(window, int):
            raise ValueError(f"'window' must be an integer, got {window!r}")
        if window < 0:
            raise ValueError(f"window must be >= 0, got {window}")
        index, _ = self._resolve_validated_ref(ref)
        first = max(1, index - window)
        last = min(self.paragraph_count(), index + window)
        return self.list_paragraphs_structured(start=first, limit=last - first + 1)

    def _resolve_validated_ref(self, ref: str) -> tuple[int, list[Element]]:
        """Parse ``ref``, bounds-check its index, and verify its hash.

        Shared validation for the ref-taking read methods
        (:meth:`get_paragraph_location`, :meth:`context`).

        Returns:
            ``(index, paragraphs)`` — the ref's 1-based index and the full
            document paragraph list it was validated against, so callers that
            need the elements don't re-query the DOM.

        Raises:
            ValueError: If ``ref`` has an invalid format.
            ParagraphIndexError: If the paragraph index is out of range.
            HashMismatchError: If the hash no longer matches current
                paragraph content.
        """
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
        return parsed.index, paragraphs

    def _style_maps(self) -> tuple[dict[str, int], dict[str, ListItem]]:
        """Outline-level and numbering maps defined by paragraph styles.

        One ``word/styles.xml`` parse serves both maps. Parsed per call
        (no caching, matching the per-call table rescan philosophy); a
        document without a styles part degrades to ``({}, {})``.
        """
        styles_path = self._workspace.word_path / "styles.xml"
        if not styles_path.exists():
            return {}, {}
        styles_dom = XMLEditor(styles_path).dom
        return _build_style_outline_map(styles_dom), _build_style_numbering_map(styles_dom)

    def get_paragraph_location(self, ref: str) -> ParagraphLocation:
        """Return the structural location of the paragraph identified by ``ref``.

        Tells the caller whether a paragraph lives in the document body or
        inside a table cell, and — when in a table — gives its 1-based
        coordinates (table index, row, logical column, depth). Also reports
        list membership: ``location.list`` is a ``ListItem(num_id, ilvl)``
        for list paragraphs, else ``None``.

        ``location.table.col`` is the *logical-grid* column, accounting for
        ``w:gridSpan`` of preceding cells in the same row. A cell that
        visually sits in column 4 reports ``col=4`` even when an earlier
        cell in the row spans 2 grid columns.

        For ``location.list``, a direct ``w:pPr/w:numPr`` wins when present
        — including Word's ``numId=0`` "numbering disabled" marker, which
        reports ``None`` with no style fallback; otherwise the numbering
        defined by the paragraph's style in ``word/styles.xml`` applies,
        with ``w:basedOn`` chains resolved. Rendered display numbers
        (e.g. "7.2(a)") are not computed.

        ``location.style`` is the raw ``w:pStyle`` style id (e.g.
        ``"Heading1"``), ``None`` when absent. ``location.outline_level``
        is the 0-based outline level (0 == Heading 1): a direct
        ``w:outlineLvl`` on the paragraph wins (the spec's ``w:val="9"``
        means body text → ``None``); otherwise the level defined by the
        paragraph's style in ``word/styles.xml`` applies, with ``w:basedOn``
        chains resolved. ``location.heading_path`` is the chain of nearest
        preceding headings containing the paragraph, outermost first,
        using each heading's current visible text; a heading's own path
        excludes itself. Headings inside table cells participate in
        document order.

        ``location.section`` is the paragraph's 1-based section index. A
        paragraph carrying a direct ``w:pPr/w:sectPr`` closes a section
        and belongs to the section it closes; the next paragraph starts
        the following one. The body-level ``w:sectPr`` defines the final
        section. Single-section documents report ``1`` everywhere.

        Heading and section context is derived from whole-document scans
        on every call; to locate many paragraphs, prefer
        :meth:`list_paragraph_locations`, which precomputes it once.

        Args:
            ref: Paragraph reference from :meth:`list_paragraphs` (e.g.,
                ``"P3#a7b2"``).

        Returns:
            :class:`ParagraphLocation`. ``location.in_table`` is ``False``
            for body paragraphs; ``True`` when the paragraph is inside a
            ``<w:tc>`` cell (in which case ``location.table`` is populated).
            ``location.list`` is a :class:`ListItem` for list paragraphs,
            ``None`` otherwise. ``location.style``,
            ``location.outline_level`` and ``location.heading_path`` carry
            the paragraph's heading context, and ``location.section`` its
            1-based section index, as described above.

        Raises:
            ValueError: If ``ref`` has an invalid format.
            ParagraphIndexError: If the paragraph index is out of range.
            HashMismatchError: If the hash no longer matches current
                paragraph content (paragraph was modified after the ref
                was captured).

        Example:
            loc = doc.get_paragraph_location("P3#a7b2")
            if loc.in_table:
                cell = loc.table
                print(f"table {cell.index} r{cell.row} c{cell.col}")
            if loc.list:
                print(f"list numId={loc.list.num_id} level={loc.list.ilvl}")
            if loc.outline_level is not None:
                print(f"heading level {loc.outline_level + 1}")
            print(f"under {' > '.join(loc.heading_path) or '(no heading)'}")
            print(f"section {loc.section}")
        """
        self._ensure_open()
        index, paragraphs = self._resolve_validated_ref(ref)
        p = paragraphs[index - 1]
        style_outlines, style_numbering = self._style_maps()
        heading_path = _compute_heading_paths(paragraphs[:index], style_outlines)[-1]
        section = _compute_section_indexes(paragraphs[:index])[-1]
        return _compute_paragraph_location(
            p,
            style_outlines=style_outlines,
            style_numbering=style_numbering,
            heading_path=heading_path,
            section=section,
        )

    def list_paragraph_locations(self) -> list[tuple[str, ParagraphLocation]]:
        """List every paragraph paired with its structural location.

        Batch counterpart to :meth:`get_paragraph_location`: precomputes
        table indexes, style outline levels, style numbering, heading
        paths, and section indexes once instead of re-scanning the
        document per ref. Each entry is ``(ref, location)`` where ``ref``
        is the same ``"P{i}#{hash}"`` token emitted by
        :meth:`list_paragraphs` (the part before ``|``) and accepted by
        :meth:`get_paragraph_location` and the editing methods.

        Returns:
            List of ``(ref, ParagraphLocation)`` tuples in document order.
            ``location.in_table`` is ``False`` for body paragraphs; ``True``
            when the paragraph is inside a ``<w:tc>`` cell.
            ``location.list`` is a :class:`ListItem` for list paragraphs,
            ``None`` otherwise (a direct ``w:numPr`` wins, else the
            paragraph style's numbering applies with ``w:basedOn`` chains
            resolved; rendered display numbers are not computed).
            ``location.style``, ``location.outline_level``,
            ``location.heading_path`` and ``location.section`` carry the
            paragraph's heading and section context with the same
            semantics as :meth:`get_paragraph_location`.

        Example:
            for ref, loc in doc.list_paragraph_locations():
                if loc.in_table:
                    cell = loc.table
                    print(f"{ref}: table {cell.index} r{cell.row} c{cell.col}")
                if loc.list:
                    print(f"{ref}: list numId={loc.list.num_id} level={loc.list.ilvl}")
                print(f"{ref}: under {' > '.join(loc.heading_path) or '(no heading)'}")
                print(f"{ref}: section {loc.section}")
        """
        self._ensure_open()
        dom = self._document_editor.dom
        table_index = _build_table_index(dom)
        style_outlines, style_numbering = self._style_maps()
        paragraphs = dom.getElementsByTagName("w:p")
        heading_paths = _compute_heading_paths(paragraphs, style_outlines)
        section_indexes = _compute_section_indexes(paragraphs)
        result = []
        for i, (p, path, section) in enumerate(zip(paragraphs, heading_paths, section_indexes, strict=True), start=1):
            ref = f"P{i}#{compute_paragraph_hash(p)}"
            result.append((
                ref,
                _compute_paragraph_location(
                    p,
                    table_index,
                    style_outlines=style_outlines,
                    style_numbering=style_numbering,
                    heading_path=path,
                    section=section,
                ),
            ))
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

    def get_original_text(self) -> str:
        """Get the original (pre-revision) text of the document.

        Returns flattened text with paragraphs separated by newlines.
        Deleted text is included, inserted text is excluded — the inverse
        of get_visible_text().

        For intra-paragraph revisions this equals what get_visible_text()
        would return after reject_all(), without modifying the document.
        Read-only: paragraph references, hashes, and all editing operations
        keep working on the accepted (visible) view.

        Returns:
            The original text content
        """
        self._ensure_open()
        paragraphs = self._document_editor.dom.getElementsByTagName("w:p")
        parts = []
        for p in paragraphs:
            tm = build_text_map(p, view="original")
            parts.append(tm.text)
        return "\n".join(parts)

    def get_markup_text(self) -> str:
        """Get document text with tracked changes rendered inline.

        Paragraphs are separated by newlines; insertions render as
        ``[ins#{id}:{author}]...[/ins]`` and deletions as
        ``[del#{id}:{author}]...[/del]``, nesting included — a foreign
        deletion inside a pending insertion renders as
        ``[ins#1:A]kept [del#9:B]gone[/del][/ins]``.

        A verification view for humans and agents (e.g. checking redlines
        without accepting them), not a parseable format: author names are
        not escaped, tabs/breaks are not rendered, and text inside a
        drawing's text box appears both inline in the host paragraph's line
        and again as its own line (same as get_text()).

        Returns:
            The marked-up text content

        Example:
            doc.replace("30 days", "60 days", paragraph="P2#f3c1")
            print(doc.get_markup_text())
            # ... [del#3:Reviewer]30 days[/del][ins#4:Reviewer]60 days[/ins] ...
        """
        self._ensure_open()
        return self._revision_manager.get_markup_text()

    def replace(self, find: str, replace_with: str, *, paragraph: str, occurrence: int | None = None) -> EditResult:
        """Replace text with tracked changes.

        Creates a tracked deletion of the old text and insertion of the new
        text. Words shared by ``find`` and ``replace_with`` at either end are
        trimmed first, so only the changed words become revisions — a replace
        that only adds or only removes words is written as a pure insertion
        or deletion. The insertion carries the formatting (rPr) that covers
        the most characters of the replaced span — runs sharing identical
        formatting tally together, with ties breaking to the earliest-seen
        formatting.

        When ``replace_with`` equals the found text, the call is a no-op: no
        revisions are created and the returned EditResult equals the input
        ``paragraph`` ref with ``group_id=None`` and ``revision_ids=()`` —
        that triple is how callers detect the no-op.

        Args:
            find: Text to find and replace
            replace_with: Replacement text
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence within the paragraph (0 = first,
                1 = second, etc.). Omitted → ``find`` must be unique in the
                paragraph, else AmbiguousTextError (use find_all() to
                enumerate the matches, or pass an explicit occurrence).

        Returns:
            EditResult — the new paragraph reference with updated hash (e.g.,
            "P2#c3d4"; usable anywhere a ref string is expected), carrying
            ``group_id``/``revision_ids`` of the revisions this edit created
            for accept_group()/reject_group().

        Raises:
            ValueError: If ``find`` is not a non-empty string,
                ``replace_with`` is not a string, ``paragraph`` is not a ref
                string, or ``occurrence`` is negative or not an integer.
            TextNotFoundError: If ``find`` is absent or ``occurrence`` is out
                of range for the paragraph.
            AmbiguousTextError: If ``occurrence`` is omitted and ``find``
                matches more than once in the paragraph.
            HashMismatchError: If the paragraph hash is stale.

        Example:
            new_ref = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
            doc.replace("other text", "new text", paragraph=new_ref)
        """
        self._ensure_open()
        _require_ref_string(paragraph)
        change_id = self._revision_manager.replace_text(find, replace_with, occurrence=occurrence, paragraph=paragraph)
        return self._edit_result(paragraph, self._revision_manager.group_id_of(change_id))

    def delete(self, text: str, *, paragraph: str, occurrence: int | None = None) -> EditResult:
        """Mark text as deleted with tracked changes.

        Args:
            text: Text to mark as deleted
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence within the paragraph (0 = first,
                1 = second, etc.). Omitted → ``text`` must be unique in the
                paragraph, else AmbiguousTextError.

        Returns:
            EditResult — the new paragraph reference with updated hash (e.g.,
            "P2#c3d4"), carrying ``group_id``/``revision_ids`` of the
            revisions this edit created.

        Raises:
            ValueError: If ``text`` is not a non-empty string, ``paragraph``
                is not a ref string, or ``occurrence`` is negative or not an integer.
            TextNotFoundError: If ``text`` is absent or ``occurrence`` is out
                of range for the paragraph.
            AmbiguousTextError: If ``occurrence`` is omitted and ``text``
                matches more than once in the paragraph.
            HashMismatchError: If the paragraph hash is stale.

        Example:
            new_ref = doc.delete("obsolete clause", paragraph="P2#f3c1")
        """
        self._ensure_open()
        _require_ref_string(paragraph)
        change_id = self._revision_manager.suggest_deletion(text, occurrence=occurrence, paragraph=paragraph)
        return self._edit_result(paragraph, self._revision_manager.group_id_of(change_id))

    def insert_after(self, anchor: str, text: str, *, paragraph: str, occurrence: int | None = None) -> EditResult:
        """Insert text after anchor with tracked changes.

        Args:
            anchor: Text to find as insertion point
            text: Text to insert after the anchor
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence of anchor within the paragraph
                (0 = first). Omitted → ``anchor`` must be unique in the
                paragraph, else AmbiguousTextError.

        Returns:
            EditResult — the new paragraph reference with updated hash (e.g.,
            "P2#c3d4"), carrying ``group_id``/``revision_ids`` of the
            revisions this edit created.

        Raises:
            ValueError: If ``anchor`` is not a non-empty string, ``text`` is
                not a string, ``paragraph`` is not a ref string, or
                ``occurrence`` is negative or not an integer.
            TextNotFoundError: If ``anchor`` is absent or ``occurrence`` is
                out of range for the paragraph.
            AmbiguousTextError: If ``occurrence`` is omitted and ``anchor``
                matches more than once in the paragraph.
            HashMismatchError: If the paragraph hash is stale.

        Example:
            new_ref = doc.insert_after("Section 5", " (as amended)", paragraph="P2#f3c1")
        """
        self._ensure_open()
        _require_ref_string(paragraph)
        change_id = self._revision_manager.insert_text_after(anchor, text, occurrence=occurrence, paragraph=paragraph)
        return self._edit_result(paragraph, self._revision_manager.group_id_of(change_id))

    def insert_before(self, anchor: str, text: str, *, paragraph: str, occurrence: int | None = None) -> EditResult:
        """Insert text before anchor with tracked changes.

        Args:
            anchor: Text to find as insertion point
            text: Text to insert before the anchor
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence of anchor within the paragraph
                (0 = first). Omitted → ``anchor`` must be unique in the
                paragraph, else AmbiguousTextError.

        Returns:
            EditResult — the new paragraph reference with updated hash (e.g.,
            "P2#c3d4"), carrying ``group_id``/``revision_ids`` of the
            revisions this edit created.

        Raises:
            ValueError: If ``anchor`` is not a non-empty string, ``text`` is
                not a string, ``paragraph`` is not a ref string, or
                ``occurrence`` is negative or not an integer.
            TextNotFoundError: If ``anchor`` is absent or ``occurrence`` is
                out of range for the paragraph.
            AmbiguousTextError: If ``occurrence`` is omitted and ``anchor``
                matches more than once in the paragraph.
            HashMismatchError: If the paragraph hash is stale.

        Example:
            new_ref = doc.insert_before("Section 6", "New clause: ", paragraph="P2#f3c1")
        """
        self._ensure_open()
        _require_ref_string(paragraph)
        change_id = self._revision_manager.insert_text_before(anchor, text, occurrence=occurrence, paragraph=paragraph)
        return self._edit_result(paragraph, self._revision_manager.group_id_of(change_id))

    @overload
    def batch_edit(self, operations: list[EditOperation], *, dry_run: Literal[False] = ...) -> list[EditResult]: ...

    @overload
    def batch_edit(self, operations: list[EditOperation], *, dry_run: Literal[True]) -> list[EditValidationResult]: ...

    def batch_edit(
        self, operations: list[EditOperation], *, dry_run: bool = False
    ) -> list[EditResult] | list[EditValidationResult]:
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
            When dry_run is False: list of EditResult (new paragraph references
            with updated hashes), in input order. Each operation gets its own
            revision group — accept one op and reject another via
            accept_group()/reject_group().
            When dry_run is True: list of EditValidationResult, one per operation.

        Raises:
            ValueError: If ``operations`` is not a list at all (e.g. None or
                a bare EditOperation) — raised before any validation, in both
                dry-run and apply modes.
            BatchOperationError: The only exception a non-dry-run batch raises
                for a failing operation — validation (element is not an
                EditOperation, malformed ref, stale hash, bad index) and apply
                (missing text, ambiguous target)
                failures alike. ``operation_index`` names the failing op and
                ``original`` (also ``__cause__``) holds the underlying typed
                exception (e.g. a HashMismatchError with ``actual_hash``).
                The document is left unchanged.

        Example:
            new_refs = doc.batch_edit([
                EditOperation.replace("old", "new", paragraph="P20#a7b2"),
                EditOperation.delete("remove", paragraph="P15#f3c1"),
            ])

            # Pre-flight the batch (note: same-paragraph sequential effects are
            # not simulated — see the dry_run note above):
            results = doc.batch_edit(ops, dry_run=True)
            if all(r.valid for r in results):
                doc.batch_edit(ops)
        """
        self._ensure_open()
        if not isinstance(operations, list):
            raise ValueError(f"batch_edit(): 'operations' must be a list of EditOperation, got {operations!r}")
        if dry_run:
            return self._revision_manager.validate_batch(operations)
        change_ids = self._revision_manager.batch_edit(operations)
        # One <w:p> walk for all result refs — batch ops never change the
        # paragraph set, so the snapshot is valid for every op.
        paragraphs = self._document_editor.dom.getElementsByTagName("w:p")
        return [
            self._edit_result(op.paragraph, self._revision_manager.group_id_of(change_id), paragraphs)
            for op, change_id in zip(operations, change_ids, strict=True)
        ]

    def rewrite_paragraph(self, ref: str, new_text: str) -> EditResult:
        """Rewrite a paragraph's text with automatic fine-grained tracked changes.

        Diffs the current paragraph text against new_text at word level and
        generates minimal tracked insertions, deletions, and replacements.
        All revisions from one rewrite share a revision group, so the rewrite
        can be accepted or rejected as a unit — accepting only some of a
        rewrite's revisions by id garbles the paragraph (each one is a diff
        hunk, not a self-contained edit).

        Args:
            ref: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            new_text: Desired new text for the paragraph

        Returns:
            EditResult — the new paragraph reference with updated hash (e.g.,
            "P2#c3d4"), carrying ``group_id``/``revision_ids`` of all the
            revisions the rewrite created (``group_id`` is None when
            new_text equals the current text, or when every change landed
            inside your own pending insertions and was merged in place).

        Raises:
            ValueError: If ``new_text`` is not a string (empty string is
                allowed — it deletes all text), or ``ref`` is malformed.
            ParagraphIndexError: If ``ref``'s index is out of range.
            HashMismatchError: If ``ref``'s hash is stale.

        Example:
            result = doc.rewrite_paragraph("P2#f3c1", "The board shall approve the proposal.")
            doc.reject_group(result.group_id)  # undo the whole rewrite
        """
        self._ensure_open()
        group_id = self._revision_manager.rewrite_paragraph(ref, new_text)
        return self._edit_result(ref, group_id)

    def batch_rewrite(self, rewrites: list[tuple[str, str]]) -> list[EditResult]:
        """Rewrite multiple paragraphs with upfront hash validation.

        All paragraph hashes are validated before any rewrites are applied.
        If any hash is stale, the entire batch is rejected before any changes
        are made. Once validation passes, rewrites are applied sequentially.
        Each rewrite gets its own revision group, or ``group_id=None`` when
        it created no revisions (see rewrite_paragraph).

        Args:
            rewrites: List of (ref, new_text) tuples

        Returns:
            List of EditResult (new paragraph references with updated hashes),
            in input order, each carrying its rewrite's
            ``group_id``/``revision_ids`` (``group_id`` is None for a rewrite
            that made no change or whose changes fully merged into your own
            pending insertions).

        Raises:
            ValueError: If ``rewrites`` is not a list at all (e.g. None) —
                raised before any validation.
            BatchOperationError: The only exception raised for a failing
                rewrite; carries ``operation_index`` and ``original``.

        Example:
            refs = doc.list_paragraphs()
            new_refs = doc.batch_rewrite([
                ("P1#a7b2", "Updated first paragraph."),
                ("P3#c3d4", "Updated third paragraph."),
            ])
        """
        self._ensure_open()
        if not isinstance(rewrites, list):
            raise ValueError(f"batch_rewrite(): 'rewrites' must be a list of (ref, new_text) tuples, got {rewrites!r}")
        group_ids = self._revision_manager.batch_rewrite(rewrites)
        return [self._edit_result(ref, group_id) for (ref, _), group_id in zip(rewrites, group_ids, strict=True)]

    # ==================== Comments API ====================

    def add_comment(
        self,
        anchor_text: str,
        comment: str,
        *,
        paragraph: str | None = None,
        occurrence: int | None = None,
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
            occurrence: Which occurrence to anchor to (0 = first). Omitted →
                ``anchor_text`` must be unique in the search scope, else
                AmbiguousTextError.

        Returns:
            The comment ID.

        Raises:
            TextNotFoundError: If ``anchor_text`` is absent or ``occurrence``
                is out of range for the scope.
            AmbiguousTextError: If ``occurrence`` is omitted and
                ``anchor_text`` matches more than once in the search scope.
            HashMismatchError: If ``paragraph``'s hash is stale.
            CommentError: If ``anchor_text`` is not a non-empty string, or
                ``comment`` is not a string.
            ValueError: If ``occurrence`` is negative or not an integer.

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

        Raises:
            ValueError: If ``comment_id`` is not an integer (bool included),
                or ``reply`` is not a non-empty string.
            CommentError: If no comment with ``comment_id`` exists; carries
                ``comment_id``.

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

        Raises:
            ValueError: If ``comment_id`` is not an integer (bool included).

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

        Raises:
            ValueError: If ``comment_id`` is not an integer (bool included).

        Example:
            doc.delete_comment(0)
        """
        self._ensure_open()
        return self._comment_manager.delete_comment(comment_id)

    # ==================== Revision Management API ====================

    def list_revisions(self, author: str | None = None, paragraph: str | None = None) -> list[Revision]:
        """List all tracked changes in the document.

        Args:
            author: If provided, filter by author name
            paragraph: If provided, a paragraph reference from
                list_paragraphs() (e.g. "P2#f3c1"); only revisions inside
                that paragraph are returned.

        Returns:
            List of Revision objects sorted by id. Each carries location
            fields: ``paragraph_ref`` (hash-anchored ref of its containing
            paragraph), ``occurrence`` (0-based index of the revision's text
            within that paragraph — for insertions it plugs into the
            ``occurrence=`` parameter of replace()/delete()/add_comment();
            for deletions it counts in the original, pre-revision text and
            must not be passed to those APIs; None when the text is not
            locatable, e.g. nested revisions), plus ``nested_under`` and
            ``contains_ids`` describing revision nesting (e.g. a foreign
            deletion inside another author's pending insertion), and
            ``group_id``/``group_source`` linking revisions from the same
            logical edit — recorded for this session's edits, inferred by
            parse-time reconstruction for revisions already in the file
            (``group_id`` is None only for ungroupable revisions, e.g.
            missing author/date).

        Raises:
            ValueError: If ``paragraph`` is malformed
            ParagraphIndexError: If the paragraph index is out of range
            HashMismatchError: If the paragraph hash doesn't match current content

        Example:
            # Reviewer workflow: inspect one paragraph's revisions, then act.
            # limit=None: every entry must be a real ref, never a truncation
            # notice, because each one is passed as paragraph= below.
            for ref in doc.list_paragraphs(max_chars=0, limit=None):
                for r in doc.list_revisions(paragraph=ref):
                    print(f"{r.id}: {r.type} '{r.text}' by {r.author}")
            doc.accept_revision(3)
        """
        self._ensure_open()
        return self._revision_manager.list_revisions(author=author, paragraph=paragraph)

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

    def accept_group(self, group_id: int) -> int:
        """Accept every revision created by one logical edit operation.

        Each edit method (replace, delete, insert_after/before,
        rewrite_paragraph, and each operation of a batch) registers the
        revisions it creates as one revision group; its EditResult carries
        the ``group_id``. Accepting the group applies the whole edit —
        resolving a multi-revision edit (especially a rewrite) revision by
        revision can leave the text garbled if only some are applied.

        Group ids are in-memory and per-open-Document, renumbered on each
        open. Revisions already in the file (previous sessions, foreign
        reviewers) get inferred groups reconstructed at parse time —
        contiguous same-paragraph revisions sharing identical author and
        date — so whole logical edits resolve as a unit after reopen too
        (see ``Revision.group_source``). Always use a group id from this
        session's EditResult or list_revisions(); a stale id from a
        previous session may resolve to a different group. save() does not
        invalidate groups.

        Args:
            group_id: Group id from an EditResult (or a Revision's
                ``group_id``)

        Returns:
            Number of revisions accepted. Members already resolved
            individually are skipped (and not counted).

        Raises:
            RevisionError: If the group id is unknown to this open Document.

        Example:
            result = doc.rewrite_paragraph(ref, "New text.")
            doc.accept_group(result.group_id)  # apply the whole rewrite
        """
        self._ensure_open()
        return self._revision_manager.accept_group(group_id)

    def reject_group(self, group_id: int) -> int:
        """Reject every revision created by one logical edit operation.

        The counterpart of :meth:`accept_group` — rejecting the group undoes
        the whole edit, restoring the exact pre-edit text (deletions are
        restored, insertions removed). Same group semantics and lifetime as
        accept_group(), including inferred groups after reopen.

        Args:
            group_id: Group id from an EditResult (or a Revision's
                ``group_id``)

        Returns:
            Number of revisions rejected. Members already resolved
            individually are skipped (and not counted).

        Raises:
            RevisionError: If the group id is unknown to this open Document.

        Example:
            result = doc.rewrite_paragraph(ref, "New text.")
            doc.reject_group(result.group_id)  # undo the whole rewrite
        """
        self._ensure_open()
        return self._revision_manager.reject_group(group_id)

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

    def save(self, path: str | Path | None = None, validate: bool = False, force: bool = False) -> Path:
        """Save the document.

        The workspace is flagged as holding unsaved changes before anything is
        written, and the flag is cleared only by a successful save back to the
        source. So after a save to a different path, or a save that raised
        partway, a later open() of the source refuses to adopt the workspace
        (WorkspaceSyncError) instead of silently carrying this session's edits
        over. Recover with force_recreate=True.

        Args:
            path: Output path (defaults to original source path)
            validate: If True, validate with LibreOffice before saving
            force: If True, skip save-time safety checks. By default save()
                refuses to overwrite the source if it changed on disk since it
                was opened (raising WorkspaceSyncError) or if the destination
                appears open in Word — a ``~$`` owner file exists next to it
                (raising DocumentOpenError). Pass force=True only for a
                confirmed-stale lock left by a crashed session.

        Returns:
            Path to the saved document

        Raises:
            WorkspaceSyncError: If the source document changed on disk since
                it was opened (protects long-lived sessions from overwriting
                edits made in Word). Pass force=True to overwrite anyway.
            DocumentOpenError: If the destination appears open in Word (a ``~$``
                owner file exists) and force is False, or if the OS denies the
                final replace because another program holds the destination open.
                force=True skips the ``~$`` check but cannot suppress the latter —
                the OS still refuses the write.

        Example:
            doc.save()  # Save to original path
            doc.save("contract_v2.docx")  # Save to new path
        """
        self._ensure_open()

        # Write-ahead: flag the workspace before any editor flush touches it,
        # so a save that fails (or a process that dies) after the flushes still
        # leaves the flag on disk and a later open() refuses to adopt the
        # diverged workspace. A successful save back to the source clears it.
        self._workspace.mark_dirty()

        # Ensure comment relationships and content types
        self._ensure_comment_relationships()
        self._ensure_comment_content_types()

        # Save all editors
        self._document_editor.save()
        self._comment_manager.save_all()

        # Pack and save
        return self._workspace.save(destination=path, validate=validate, force=force)

    def close(self, cleanup: bool = True) -> None:
        """Close the document and clean up workspace.

        Releases the advisory workspace lock in both cleanup modes — closing
        is what frees the document for another session to open.

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
        """Raise DocumentClosedError if the document is closed."""
        if self._closed:
            raise DocumentClosedError(
                f"Document is closed. Reopen it with Document.open({str(self.source_path)!r}) to continue.",
                path=self.source_path,
            )

    def _setup_tracking(self) -> None:
        """Set up tracked changes infrastructure in the document.

        Runs at every open. Its writes (people.xml, [Content_Types].xml,
        document.xml.rels, settings.xml rsids) deliberately do NOT mark the
        workspace dirty: they are deterministic bookkeeping that an adopting
        session re-produces identically (each helper checks before adding; the
        rsid comes from meta), not unsaved user content. Marking dirty here
        would flag every workspace the moment it is opened, so any session
        that crashed without editing would force force_recreate on the next
        open with no data-loss risk behind it. All post-open writes DO mark
        dirty first — see the on_save/on_write hooks in __init__.
        """
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
            on_save=self._workspace.mark_dirty,
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
            on_save=self._workspace.mark_dirty,
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
