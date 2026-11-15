"""Comment management for docx_editor.

Provides CommentManager for creating and managing document comments.
"""

import html
import posixpath
import shutil
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import NamedTuple
from xml.dom.minidom import Element
from xml.parsers.expat import ExpatError

import defusedxml.minidom
from defusedxml.common import DefusedXmlException

from .exceptions import (
    AmbiguousTextError,
    CommentError,
    HashMismatchError,
    NodeNotFoundError,
    ParagraphIndexError,
    TextNotFoundError,
    XMLError,
    _truncate_preview,
)
from .track_changes.dom import _addressable_paragraph, _adjudicable_id
from .track_changes.listing import _RevisionLocationContext
from .xml_editor import (
    W14_NAMESPACE,
    W15_NAMESPACE,
    W16CEX_NAMESPACE,
    W16CID_NAMESPACE,
    DocxXMLEditor,
    ParagraphRef,
    TextMapMatch,
    _escape_xml,
    _generate_hex_id,
    _reject_control_chars,
    _require_valid_occurrence,
    body_paragraphs,
    build_text_map,
    compute_paragraph_hash,
    count_in_text_map,
    find_in_text_map,
    get_rPr_xml,
    get_text_node_data,
    rebuild_run_fragments,
    render_plain_wt,
)

# Path to template files
TEMPLATE_DIR = Path(__file__).parent / "ooxml" / "templates"


@dataclass
class Comment:
    """Represents a document comment.

    ``paragraph_ref`` is the hash-anchored ref ("P{i}#{hash}") of the body
    paragraph holding the comment's ``w:commentRangeStart``; for a range that
    spans paragraphs, the one it starts in. A reply carries the ref of the
    comment it answers, so a whole thread reports one paragraph. It is None
    when the comment has no range start inside a body paragraph (the marker is
    missing, sits between block elements, or is in a text box).

    A reply to such a comment is None too, wherever its own marker sits.

    Like every ref it is a snapshot taken at list time: usable as
    ``paragraph=`` until that paragraph is edited, stale afterwards.
    """

    id: int
    text: str
    author: str
    date: datetime | None
    resolved: bool = False
    replies: list["Comment"] = field(default_factory=list)
    paragraph_ref: str | None = None

    def __repr__(self) -> str:
        status = "[RESOLVED] " if self.resolved else ""
        reply_count = f" ({len(self.replies)} replies)" if self.replies else ""
        location = f" @{self.paragraph_ref}" if self.paragraph_ref else ""
        return f"Comment({self.id}: {status}'{self.text[:30]}...' by {self.author}{reply_count}{location})"


class CommentPart(NamedTuple):
    """One of the four comment parts: how it is named, wired and rooted."""

    name: str  # default file name in word/
    rel_type: str  # relationship type in document.xml.rels
    content_type: str  # override in [Content_Types].xml
    root_tag: str  # the part's document element
    prefix: str  # prefix of the markup written into the part
    namespace: str  # what that prefix must be bound to


COMMENT_PARTS = (
    CommentPart(
        "comments.xml",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        "w:comments",
        "w14",
        W14_NAMESPACE,
    ),
    CommentPart(
        "commentsExtended.xml",
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
        "w15:commentsEx",
        "w15",
        W15_NAMESPACE,
    ),
    CommentPart(
        "commentsIds.xml",
        "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
        "w16cid:commentsIds",
        "w16cid",
        W16CID_NAMESPACE,
    ),
    CommentPart(
        "commentsExtensible.xml",
        "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml",
        "w16cex:commentsExtensible",
        "w16cex",
        W16CEX_NAMESPACE,
    ),
)


def has_root(path: Path, root_tag: str) -> bool:
    """Whether ``path`` is a well-formed XML file whose document element is ``root_tag``.

    Parsed on its own, never through an editor: a file that fails this test
    must not end up in an editor cache.
    """
    try:
        return path.is_file() and defusedxml.minidom.parse(str(path)).documentElement.tagName == root_tag
    except (OSError, ExpatError, DefusedXmlException):
        # OSError: a target no file system can hold, e.g. an overlong name.
        return False


def _is_inside(path: Path, directory: Path) -> bool:
    """Whether ``path``, with symlinks followed, lies inside ``directory``."""
    try:
        return path.resolve().is_relative_to(directory.resolve())
    except (OSError, RuntimeError):
        # A link that cannot be followed (a loop raises RuntimeError before
        # Python 3.13) leads nowhere inside the directory.
        return False


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
        *,
        on_write: Callable[[], None] | None = None,
    ):
        """Initialize with workspace path and document editor.

        Args:
            workspace_path: Path to the unpacked workspace folder
            document_editor: DocxXMLEditor for word/document.xml
            author: Author name for new comments
            initials: Author initials for new comments
            on_write: Optional write-ahead callback (e.g.
                Workspace.mark_dirty) invoked before any comment-file write
                touches the workspace on disk. Broader than XMLEditor's
                ``on_save`` (which it is wired into for editor flushes) —
                hence the different name — because it also covers non-editor
                writes like template copies
        """
        self.workspace_path = workspace_path
        self.word_path = workspace_path / "word"
        self.document_editor = document_editor
        self.author = author
        self.initials = initials
        self._on_write = on_write

        # Comment file paths, in COMMENT_PARTS order, and why a new comment
        # cannot be wired (empty when it can).
        self.part_paths, self._unusable_relationships = self._resolve_comment_parts()
        (
            self.comments_path,
            self.comments_extended_path,
            self.comments_ids_path,
            self.comments_extensible_path,
        ) = self.part_paths
        # A part that is a link out of the workspace is never read or written:
        # it counts as absent, and a new comment is refused.
        self._outside_parts = frozenset(path for path in self.part_paths if not _is_inside(path, self.workspace_path))

        # Cache for lazy-loaded editors
        self._editors: dict[str, DocxXMLEditor] = {}

        # Load existing comments for reply support
        self.existing_comments = self._load_existing_comments()
        self.next_comment_id = self._get_next_comment_id()

    def part_exists(self, path: Path) -> bool:
        """Whether the comment part at ``path`` is there to be read or written.

        A part that leaves the package (a symlink out of the workspace) counts
        as absent, so nothing is read from or written through it.
        """
        return path not in self._outside_parts and path.exists()

    def _resolve_comment_parts(self) -> tuple[tuple[Path, ...], list[str]]:
        """Return each comment part's path (``COMMENT_PARTS`` order) and the relationships that cannot be honoured.

        A comment part is the one ``document.xml.rels`` names for its
        relationship type, whatever its file name (``word/comments1.xml``).
        A type with no relationship, or one that names the default file, gets
        the default name, where the part is created on the first comment and
        wired at save.

        A target under another name is adopted only when it is a file whose
        document element is the part's expected root. Anything else — a
        missing file, another kind of part such as ``document.xml``, an
        external target, a target with a backslash or one that resolves
        outside the workspace, two different parts for one type, one file for
        two types — cannot be honoured: adopting it would put a second editor on a
        part the document already edits, and a part created under the default
        name instead would stay unwired, because its type already has a
        relationship. Each such relationship is described in the second list,
        and :meth:`_prepare_comment_parts` refuses to add a comment while it
        is non-empty.
        """
        resolved: dict[str, set[Path]] = {part.rel_type: set() for part in COMMENT_PARTS}
        unusable: list[str] = []
        rels_path = self.word_path / "_rels" / "document.xml.rels"
        if rels_path.exists():
            parts_by_type = {part.rel_type: part for part in COMMENT_PARTS}
            rels = defusedxml.minidom.parse(str(rels_path))
            for rel_elem in rels.getElementsByTagName("Relationship"):
                part = parts_by_type.get(rel_elem.getAttribute("Type"))
                if part is None:
                    continue
                target = rel_elem.getAttribute("Target")
                if rel_elem.getAttribute("TargetMode") == "External":
                    unusable.append(f"{part.name} relationship has the external target {target!r}")
                    continue
                # A backslash is not a separator in an OPC part name, but a
                # file system may treat it as one and climb out of the package.
                if "\\" in target:
                    unusable.append(f"{part.name} relationship target {target!r} contains a backslash")
                    continue
                # An OPC target is a reference relative to the source part,
                # /word/document.xml. normpath cannot climb above the root of
                # an absolute path; a symlink in the workspace still can.
                part_name = posixpath.normpath(posixpath.join("/word", target))
                part_path = self.workspace_path / part_name.lstrip("/")
                if not _is_inside(part_path, self.workspace_path):
                    unusable.append(f"{part.name} relationship target {target!r} leaves the package")
                elif part_path == self.word_path / part.name or has_root(part_path, part.root_tag):
                    resolved[part.rel_type].add(part_path)
                else:
                    unusable.append(f"{part.name} relationship target {target!r} is not a comment part of this package")
        paths = []
        for part in COMMENT_PARTS:
            candidates = resolved[part.rel_type]
            if len(candidates) > 1:
                unusable.append(f"{part.name} relationships name more than one part")
            paths.append(candidates.pop() if len(candidates) == 1 else self.word_path / part.name)
        for part, path in zip(COMMENT_PARTS, paths, strict=True):
            if paths.count(path) > 1:
                unusable.append(f"{part.name} resolves to {path.name}, which another comment part also resolves to")
        return tuple(paths), unusable

    def _get_editor(self, xml_path: Path) -> DocxXMLEditor:
        """Get or create an editor for the specified XML file."""
        path_str = str(xml_path)
        if path_str not in self._editors:
            self._editors[path_str] = DocxXMLEditor(
                xml_path,
                rsid=self.document_editor.rsid,
                author=self.author,
                initials=self.initials,
                on_save=self._on_write,
            )
        return self._editors[path_str]

    def add_comment(
        self,
        anchor_text: str | None,
        comment_text: str | None,
        *,
        paragraph: str | None = None,
        occurrence: int | None = None,
    ) -> int:
        """Add a comment anchored to specific text, or to a whole paragraph.

        The anchor lookup uses the same text-map search as ``count_matches`` and
        the tracked-change edit methods, so anchors that span ``w:t`` run
        boundaries (formatting changes, smart-quote splits, ``w:ins`` wrappers)
        are found.

        With ``anchor_text=None`` the comment covers the whole of ``paragraph``:
        its full accepted-view text (pending insertions included, pending
        deletions excluded), with the markers placed exactly where an anchor
        of that full text would place them.

        Args:
            anchor_text: Text to attach the comment to. ``None`` comments on
                the whole paragraph named by ``paragraph``.
            comment_text: The comment content. ``None`` is refused like any
                other non-string.
            paragraph: Optional paragraph reference (e.g., ``"P3#a7b2"``) to
                scope the search. ``None`` searches the whole document.
                Required when ``anchor_text`` is ``None``.
            occurrence: Which occurrence to anchor (0 = first). Omitted →
                ``anchor_text`` must be unique in the search scope, else
                AmbiguousTextError. Not allowed when ``anchor_text`` is
                ``None``.

        Returns:
            The comment ID.

        Raises:
            TextNotFoundError: If the anchor text is not found.
            AmbiguousTextError: If ``occurrence`` is omitted and
                ``anchor_text`` matches more than once in the search scope.
            HashMismatchError: If a paragraph reference's hash is stale.
            ParagraphIndexError: If a paragraph reference's index is out of range.
            CommentError: If ``anchor_text`` is neither ``None`` nor a
                non-empty string, ``comment_text`` is not a string, or either
                holds a control character (``anchor_text`` may contain ``\\t``
                — a tab mark is one text-map character — but ``comment_text``
                may not). Also if ``anchor_text`` is ``None`` and ``paragraph``
                is omitted, or the paragraph has no text to anchor on.
            ValueError: If ``occurrence`` is negative or not an integer, or is
                given with ``anchor_text=None``.
            XMLError: If an existing comment part cannot take the comment (its
                root is not the expected element, or it binds ``w14`` to
                another namespace). Raised before document.xml is touched.
        """
        # Everything is validated here, before _place_markers mutates
        # document.xml — a crash later (html.escape) would leave orphaned range
        # markers. Comments carry no split semantics, so '\n' is rejected in
        # both: in the anchor it could never match, and in the body it would be
        # an invisible, unreviewable literal.
        if anchor_text is None:
            if paragraph is None:
                raise CommentError(
                    "add_comment(): anchor_text=None comments on a whole paragraph; pass paragraph= to say which one"
                )
            if occurrence is not None:
                raise ValueError(
                    "add_comment(): occurrence= selects among matches of anchor_text; "
                    "it cannot be combined with a whole-paragraph comment (anchor_text=None)"
                )
            comment_text = self._validate_comment_text(comment_text, field="comment_text", ctx="add_comment(): ")
            text_map = build_text_map(self._resolve_paragraph(paragraph))
            # The whole paragraph is the anchored form with the paragraph's full
            # text as the anchor, so placement is _place_markers' and nothing new.
            match = find_in_text_map(text_map, text_map.text, 0) if text_map.text else None
            if match is None:
                raise CommentError(f"add_comment(): paragraph {paragraph} has no text to anchor a comment on")
        else:
            if not isinstance(anchor_text, str) or not anchor_text:
                raise CommentError(f"anchor_text must be a non-empty string, got {anchor_text!r}")
            comment_text = self._validate_comment_text(comment_text, field="comment_text", ctx="add_comment(): ")
            try:
                _reject_control_chars(anchor_text, field="anchor_text", ctx="add_comment(): ", allow_tab=True)
            except ValueError as e:
                raise CommentError(str(e)) from e
            _, match = self._locate_anchor(anchor_text, paragraph, occurrence)

        comment_id = self.next_comment_id
        self._prepare_comment_parts()

        # Place range markers at character-precise positions in document.xml
        self._place_markers(match, comment_id)

        return self._commit_comment(comment_id, comment_text)

    @staticmethod
    def _validate_comment_text(text: str | None, *, field: str, ctx: str) -> str:
        """Reject comment/note text that cannot be written into a comment part.

        One rule for the two element-or-text anchored bodies — ``add_comment``'s
        ``comment_text`` and ``add_comment_on_elements``' — so they cannot
        drift: it must be a string, and must carry no control characters
        (``\\n`` included: a comment body is a single ``<w:t>``, where a literal
        newline is an invisible, unreviewable artifact). ``reply_to_comment``
        keeps its own check because a reply may not be empty and reports a
        wrong type as ``ValueError``, not ``CommentError``.

        Returns:
            ``text``, narrowed to ``str``.

        Raises:
            CommentError: If ``text`` is not a string or holds a control character.
        """
        if not isinstance(text, str):
            raise CommentError(f"{field} must be a string, got {text!r}")
        try:
            _reject_control_chars(text, field=field, ctx=ctx, allow_newline=False)
        except ValueError as e:
            raise CommentError(str(e)) from e
        return text

    def add_comment_on_elements(self, first: Element, last: Element, comment_text: str) -> int:
        """Anchor a comment on the span from ``first`` to ``last`` (document order).

        The element-anchored counterpart of ``add_comment``: the range markers
        become *siblings* of whole elements instead of being spliced into a run
        at character offsets. That is what lets a comment bracket a tracked
        revision — a ``w:del``'s text is not in the accepted text map at all,
        and an inserted string may repeat in its paragraph, so text search
        cannot address either one unambiguously.

        ``first`` and ``last`` may be the same element, and may live in
        different paragraphs (a tracked split); ``w:commentRangeStart`` /
        ``w:commentRangeEnd`` are range marks, valid anywhere run-level content
        is.

        Markers sit *outside* the bracketed elements, so rejecting the revision
        does not carry them away — the caller owns cleanup (see
        ``Document._reap_note_comments``).

        Args:
            first: First element of the span; the start marker goes before it.
            last: Last element of the span; the end marker and the comment
                reference run go after it.
            comment_text: The comment content.

        Returns:
            The comment ID.

        Raises:
            CommentError: If ``comment_text`` is not a string or holds a
                control character.
            XMLError: If an existing comment part cannot take the comment.
                Raised before document.xml is touched.
        """
        self._validate_comment_text(comment_text, field="comment_text", ctx="add_comment_on_elements(): ")

        comment_id = self.next_comment_id
        self._prepare_comment_parts()
        self.document_editor.insert_before(first, self._comment_range_start_xml(comment_id))
        # _comment_range_end_xml emits the end marker *and* the reference run.
        self.document_editor.insert_after(last, self._comment_range_end_xml(comment_id))
        return self._commit_comment(comment_id, comment_text)

    def _commit_comment(self, comment_id: int, comment_text: str, *, parent_para_id: str | None = None) -> int:
        """Write the four comment parts for an already-placed marker pair.

        The shared tail of every comment creation path — ``add_comment``,
        ``add_comment_on_elements`` and ``reply_to_comment`` (which passes its
        parent's ``para_id``): the markers are already in document.xml, this
        records the body. ``next_comment_id`` moves here and nowhere else, so a
        note comment and a user comment can never be allocated the same id.
        """
        para_id = _generate_hex_id()
        durable_id = _generate_hex_id()
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        # Add to all comment XML files
        self._add_to_comments_xml(comment_id, para_id, comment_text, timestamp)
        self._add_to_comments_extended_xml(para_id, parent_para_id=parent_para_id)
        self._add_to_comments_ids_xml(para_id, durable_id)
        self._add_to_comments_extensible_xml(durable_id)

        # Track for reply support
        self.existing_comments[comment_id] = {"para_id": para_id}
        self.next_comment_id += 1

        return comment_id

    def _resolve_paragraph(self, paragraph: str) -> Element:
        """Return the body paragraph a ``P<index>#<hash>`` reference names.

        Raises:
            ParagraphIndexError: If the index is out of range.
            HashMismatchError: If the paragraph's content no longer matches the hash.
        """
        ref = ParagraphRef.parse(paragraph)
        paragraphs = body_paragraphs(self.document_editor.dom)
        if ref.index < 1 or ref.index > len(paragraphs):
            raise ParagraphIndexError(ref.index, len(paragraphs))
        p = paragraphs[ref.index - 1]
        actual = compute_paragraph_hash(p)
        if actual != ref.hash:
            raise HashMismatchError(ref.index, ref.hash, actual, _truncate_preview(build_text_map(p).text))
        return p

    def _locate_anchor(
        self, anchor_text: str, paragraph: str | None, occurrence: int | None
    ) -> tuple[Element, TextMapMatch]:
        """Locate ``anchor_text`` and return its paragraph + text-map match.

        Mirrors ``RevisionManager._locate_in_paragraph`` /
        ``_locate_document_wide`` so comment anchors and edit anchors find the
        same text and fail the same way: ``occurrence=None`` requires a unique
        anchor (else AmbiguousTextError), an explicit out-of-range occurrence
        reports the actual total, and a negative or non-int occurrence is
        rejected with ValueError.
        """
        _require_valid_occurrence(occurrence)
        occ = occurrence if occurrence is not None else 0

        if paragraph is not None:
            p = self._resolve_paragraph(paragraph)
            text_map = build_text_map(p)
            total = count_in_text_map(text_map, anchor_text)
            match = find_in_text_map(text_map, anchor_text, occ)
            if match is not None:
                if occurrence is None and total > 1:
                    raise AmbiguousTextError(
                        anchor_text,
                        paragraph_ref=paragraph,
                        paragraph_preview=text_map.text,
                        total_occurrences=total,
                    )
                return p, match
            if total > 0:
                raise TextNotFoundError(
                    anchor_text,
                    paragraph_ref=paragraph,
                    paragraph_preview=text_map.text,
                    occurrence=occ,
                    total_occurrences=total,
                )
            raise TextNotFoundError(
                anchor_text,
                paragraph_ref=paragraph,
                paragraph_preview=text_map.text,
            )

        # Document-wide search. With an explicit occurrence the scan returns at
        # the requested match; with occurrence=None it runs to the end so
        # `current` ends up as the document-wide total, which the ambiguity and
        # out-of-range errors need.
        current = 0
        found: tuple[Element, TextMapMatch] | None = None
        for p in body_paragraphs(self.document_editor.dom):
            text_map = build_text_map(p)
            local = 0
            while True:
                m = find_in_text_map(text_map, anchor_text, local)
                if m is None:
                    break
                if current == occ:
                    if occurrence is not None:
                        return p, m
                    found = (p, m)
                current += 1
                local += 1
        if found is not None:
            if occurrence is None and current > 1:
                raise AmbiguousTextError(anchor_text, total_occurrences=current)
            return found
        if current > 0:
            raise TextNotFoundError(
                anchor_text,
                occurrence=occ,
                total_occurrences=current,
            )
        raise TextNotFoundError(anchor_text)

    def _place_markers(self, match: TextMapMatch, comment_id: int) -> None:
        """Insert range start/end markers at the precise match boundaries."""
        start_pos = match.positions[0]
        end_pos = match.positions[-1]
        start_node = start_pos.node
        start_offset = start_pos.offset_in_node
        end_node = end_pos.node
        end_offset = end_pos.offset_in_node

        start_run = _find_ancestor_run(start_node)
        end_run = _find_ancestor_run(end_node)
        if start_run is None or end_run is None:
            raise CommentError("Could not find parent run for comment anchor")

        start_marker = self._comment_range_start_xml(comment_id)
        end_marker = self._comment_range_end_xml(comment_id)

        if start_run is end_run:
            self._rebuild_run_with_markers(
                start_run,
                start_node,
                start_offset,
                end_node,
                end_offset,
                start_marker,
                end_marker,
            )
            return

        # Independent runs — place END first (after start_run in document order),
        # then START. Either order works since the two runs are siblings, but
        # this matches the plan and reads naturally.
        self._rebuild_run_with_markers(
            end_run,
            None,
            0,
            end_node,
            end_offset,
            "",
            end_marker,
        )
        self._rebuild_run_with_markers(
            start_run,
            start_node,
            start_offset,
            None,
            0,
            start_marker,
            "",
        )

    def _rebuild_run_with_markers(
        self,
        run: Element,
        start_node,
        start_offset: int,
        end_node,
        end_offset: int,
        start_marker: str,
        end_marker: str,
    ) -> None:
        """Rebuild ``run``'s XML, splitting ``w:t`` children to insert markers.

        ``start_marker`` / ``end_marker`` are inserted right before the
        character at ``start_offset`` of ``start_node`` and right after the
        character at ``end_offset`` of ``end_node`` respectively. Pass an
        empty string (and any int offset, e.g. ``0``) when the corresponding
        marker lives in a different run.

        Iterates *direct* children of ``run`` (not descendants) so non-``w:t``
        content like ``<w:tab/>``, ``<w:br/>``, ``<w:drawing/>``, and field
        markers is preserved, and so ``w:t`` nodes nested inside a drawing's
        text box are left untouched. A ``<w:tab/>`` may itself be an edge: it
        is one text-map character, so the marker simply lands before or after
        the tab's own run.
        """
        rPr_xml = get_rPr_xml(run)

        def render_wt(wt) -> list[str]:
            wt_text = get_text_node_data(wt)
            is_start_here = bool(start_marker) and wt is start_node
            is_end_here = bool(end_marker) and wt is end_node
            parts: list[str] = []

            if not is_start_here and not is_end_here:
                parts.extend(render_plain_wt(wt, rPr_xml))
            elif is_start_here and is_end_here:
                before = wt_text[:start_offset]
                matched = wt_text[start_offset : end_offset + 1]
                after = wt_text[end_offset + 1 :]
                if before:
                    parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before)}</w:t></w:r>")
                parts.append(start_marker)
                if matched:
                    parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(matched)}</w:t></w:r>")
                parts.append(end_marker)
                if after:
                    parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after)}</w:t></w:r>")
            elif is_start_here:
                before = wt_text[:start_offset]
                after = wt_text[start_offset:]
                if before:
                    parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before)}</w:t></w:r>")
                parts.append(start_marker)
                if after:
                    parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after)}</w:t></w:r>")
            else:  # is_end_here only
                before = wt_text[: end_offset + 1]
                after = wt_text[end_offset + 1 :]
                if before:
                    parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before)}</w:t></w:r>")
                parts.append(end_marker)
                if after:
                    parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after)}</w:t></w:r>")
            return parts

        def render_other(child) -> list[str]:
            # Range markers are run-level marks, so a <w:tab/> edge needs no
            # split: the marker brackets the tab's own run.
            parts = [f"<w:r>{rPr_xml}{child.toxml()}</w:r>"]
            if start_marker and child is start_node:
                parts.insert(0, start_marker)
            if end_marker and child is end_node:
                parts.append(end_marker)
            return parts

        xml_parts = rebuild_run_fragments(run, rPr_xml, render_wt, render_other)
        self.document_editor.replace_node(run, "".join(xml_parts))

    def reply_to_comment(self, parent_comment_id: int, reply_text: str) -> int:
        """Add a reply to an existing comment.

        Args:
            parent_comment_id: The ID of the comment to reply to
            reply_text: The reply content

        Returns:
            The new comment ID for the reply

        Raises:
            ValueError: If ``parent_comment_id`` is not an integer (bool
                included), or ``reply_text`` is not a non-empty string.
            CommentError: If the parent comment is not found, or has no
                ``w:commentReference`` in the document to seat the reply on;
                carries ``comment_id``.
            XMLError: If an existing comment part cannot take the reply, or
                the parent has more than one reference or range start in the
                document. Raised before document.xml is touched.
        """
        # Reject junk ids before the membership test: None/"1" would surface
        # as a misleading "not found" CommentError, and True aliases comment
        # id 1 (bool == int). The message names the public field 'comment_id',
        # not this internal parameter.
        if isinstance(parent_comment_id, bool) or not isinstance(parent_comment_id, int):
            raise ValueError(f"reply_to_comment(): 'comment_id' must be an integer, got {parent_comment_id!r}")
        # Must fail before any marker is inserted into document.xml — a None
        # reply would otherwise be serialized into comments.xml as "None".
        if not isinstance(reply_text, str) or not reply_text:
            raise ValueError(f"reply_to_comment(): 'reply' must be a non-empty string, got {reply_text!r}")
        try:
            _reject_control_chars(reply_text, field="'reply'", ctx="reply_to_comment(): ", allow_newline=False)
        except ValueError as e:
            raise CommentError(str(e)) from e
        if parent_comment_id not in self.existing_comments:
            raise CommentError(
                f"Parent comment with id={parent_comment_id} not found",
                comment_id=parent_comment_id,
            )

        comment_id = self.next_comment_id

        # Find the parent's markers in document.xml. A point comment (Word
        # 2007, LibreOffice) has a reference and no range; a comment with no
        # reference gives the reply nowhere to sit.
        parent_ref_elem = self._find_marker("w:commentReference", parent_comment_id)
        if parent_ref_elem is None:
            raise CommentError(
                f"Parent comment with id={parent_comment_id} has no w:commentReference in the document; "
                "cannot anchor a reply",
                comment_id=parent_comment_id,
            )
        parent_start_elem = self._find_marker("w:commentRangeStart", parent_comment_id)

        # Everything that can refuse the reply has now run: the parent is only
        # given a paraId once the reply is certain to be written.
        self._prepare_comment_parts()
        parent_para_id = self._ensure_para_id(parent_comment_id)
        self._ensure_extended_row(parent_para_id)
        self._place_reply_markers(comment_id, parent_start_elem, parent_ref_elem)

        return self._commit_comment(comment_id, reply_text, parent_para_id=parent_para_id)

    def _find_marker(self, tag: str, comment_id: int) -> Element | None:
        """The one ``tag`` marker of ``comment_id`` in document.xml, or None when it has none.

        Several markers for the one id raise ``MultipleNodesFoundError``.
        """
        try:
            return self.document_editor.get_node(tag=tag, attrs={"w:id": str(comment_id)})
        except NodeNotFoundError:
            return None

    def _place_reply_markers(self, comment_id: int, parent_start: Element | None, parent_ref: Element) -> None:
        """Seat a reply's markers around its parent's, where Word expects them.

        A parent with no range (``parent_start`` is None) is a point comment;
        its reply is anchored the same way, by a reference run alone.
        """
        # A reference written outside any run is followed directly: its parent
        # is then a paragraph, and a run placed after that would land in w:body.
        ref_anchor = _find_ancestor_run(parent_ref) or parent_ref
        if parent_start is None:
            self.document_editor.insert_after(ref_anchor, self._comment_ref_run_xml(comment_id))
            return
        self.document_editor.insert_after(parent_start, self._comment_range_start_xml(comment_id))
        # One fragment, so the range end precedes the reference run: two
        # insert_after calls on the same anchor would land the second ahead
        # of the first and seat the reference run inside its own range.
        self.document_editor.insert_after(ref_anchor, self._comment_range_end_xml(comment_id))

    def list_comments(self, author: str | None = None) -> list[Comment]:
        """List all comments in the document.

        Args:
            author: If provided, filter by author name

        Returns:
            List of Comment objects (with replies nested), each carrying the
            ``paragraph_ref`` of the paragraph it sits on
        """
        if not self.part_exists(self.comments_path):
            return []

        refs = self._range_start_refs()
        editor = self._get_editor(self.comments_path)
        comments_dict: dict[int, Comment] = {}
        parent_map: dict[str, str] = {}  # para_id -> parent_para_id

        # Build parent map from commentsExtended.xml
        if self.part_exists(self.comments_extended_path):
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
                if para_id and self.part_exists(self.comments_extended_path):
                    ext_editor = self._get_editor(self.comments_extended_path)
                    for ex_elem in ext_editor.dom.getElementsByTagName("w15:commentEx"):
                        if ex_elem.getAttribute("w15:paraId") == para_id:
                            comment.resolved = ex_elem.getAttribute("w15:done") == "1"
                            break

                comments_dict[comment.id] = comment

        # Build reply tree
        para_to_id: dict[str, int] = {}
        for comment_id, info in self.existing_comments.items():
            if info["para_id"] is not None:
                para_to_id[info["para_id"]] = comment_id

        def parent_of(comment_id: int) -> int | None:
            para_id = self.existing_comments.get(comment_id, {}).get("para_id")
            parent_para = parent_map.get(para_id) if para_id else None
            return para_to_id.get(parent_para) if parent_para else None

        # Nest replies
        root_comments = []
        for comment_id, comment in comments_dict.items():
            parent_id = parent_of(comment_id)
            if parent_id in comments_dict:
                comments_dict[parent_id].replies.append(comment)
            else:
                root_comments.append(comment)

            # The ref is the thread root's, found through the unfiltered
            # hierarchy so ``author=`` never changes what a comment reports.
            # ``seen`` ends the climb on a cyclic paraIdParent chain.
            root_id, seen = comment_id, {comment_id}
            while (up := parent_of(root_id)) is not None and up not in seen:
                seen.add(up)
                root_id = up
            comment.paragraph_ref = refs.get(root_id)

        return sorted(root_comments, key=lambda c: c.id)

    def _range_start_refs(self) -> dict[int, str | None]:
        """Paragraph ref of each comment's range start, keyed by comment id.

        One paragraph-index walk and one marker walk per call, however many
        comments there are. A marker outside any body paragraph maps to None;
        of duplicate markers for one id, the first in document order wins.
        """
        context = _RevisionLocationContext(self.document_editor.dom)
        refs: dict[int, str | None] = {}
        for marker in self.document_editor.dom.getElementsByTagName("w:commentRangeStart"):
            comment_id = _adjudicable_id(marker)
            if comment_id is None:
                continue
            paragraph = _addressable_paragraph(marker)
            refs.setdefault(comment_id, None if paragraph is None else context.paragraph_ref(paragraph))
        return refs

    def resolve_comment(self, comment_id: int) -> bool:
        """Mark a comment as resolved.

        Args:
            comment_id: The comment ID to resolve

        Returns:
            True if resolved, False if the document has no such comment

        Raises:
            ValueError: If ``comment_id`` is not an integer (bool included).
            XMLError: If the comment has no ``w15:commentEx`` row yet and an
                existing comment part cannot take one. Raised before anything
                is changed.

        Note:
            The done flag lives only in commentsExtended.xml, keyed by
            ``w14:paraId``. A comment from a producer that writes neither
            (Word 2007/2010, LibreOffice, Google Docs) is given a paraId and a
            ``w15:commentEx`` row here, on first need.
        """
        # A junk id must not silently return False ("not found"), and True
        # would alias comment id 1 (bool == int).
        if isinstance(comment_id, bool) or not isinstance(comment_id, int):
            raise ValueError(f"resolve_comment(): 'comment_id' must be an integer, got {comment_id!r}")
        if comment_id not in self.existing_comments:
            return False

        para_id = self.existing_comments[comment_id]["para_id"]
        row = self._extended_row(para_id) if para_id is not None else None
        if row is None:
            self._prepare_comment_parts()
            row = self._ensure_extended_row(self._ensure_para_id(comment_id))
        row.setAttribute("w15:done", "1")
        return True

    def delete_comment(self, comment_id: int) -> bool:
        """Delete a comment, and every reply threaded under it, from the document.

        A reply is linked to its parent only by ``w15:paraIdParent``: left
        behind, it would point at a paraId no part of the document still holds,
        and its own descendants would become unreachable — so the thread goes
        as a unit, the way Word deletes one.

        Args:
            comment_id: The comment ID to delete

        Returns:
            True if deleted, False if the document has no such comment

        Raises:
            ValueError: If ``comment_id`` is not an integer (bool included).
        """
        # Same guard as resolve_comment: no silent False for junk, no bool
        # aliasing comment id 1.
        if isinstance(comment_id, bool) or not isinstance(comment_id, int):
            raise ValueError(f"delete_comment(): 'comment_id' must be an integer, got {comment_id!r}")
        if comment_id not in self.existing_comments:
            return False

        for reply_id in self.reply_ids(comment_id):
            self.delete_comment(reply_id)

        para_id = self.existing_comments[comment_id]["para_id"]

        self._remove_range_markers(comment_id)

        # Remove from comments.xml — the part exists: the comment was read
        # from it or written into it.
        comment_elem = self._comment_element(comment_id)
        comment_elem.parentNode.removeChild(comment_elem)

        # A comment with no paraId has no row in the parts keyed by one.
        if para_id is None:
            del self.existing_comments[comment_id]
            return True

        # Remove from commentsExtended.xml
        if self.part_exists(self.comments_extended_path):
            editor = self._get_editor(self.comments_extended_path)
            for ex_elem in editor.dom.getElementsByTagName("w15:commentEx"):
                if ex_elem.getAttribute("w15:paraId") == para_id:
                    ex_elem.parentNode.removeChild(ex_elem)
                    break

        # Remove from commentsIds.xml, keeping the durable id it holds — that
        # entry is the only place the two ids are linked, so the extensible
        # part can only be found before this element goes.
        # A comment authored elsewhere whose commentsIds.xml row is missing
        # leaves no way to reach its extensible entry; that orphan stays.
        durable_id = None
        if self.part_exists(self.comments_ids_path):
            editor = self._get_editor(self.comments_ids_path)
            for id_elem in editor.dom.getElementsByTagName("w16cid:commentId"):
                if id_elem.getAttribute("w16cid:paraId") == para_id:
                    durable_id = id_elem.getAttribute("w16cid:durableId")
                    id_elem.parentNode.removeChild(id_elem)
                    break

        # Remove from commentsExtensible.xml
        if durable_id and self.part_exists(self.comments_extensible_path):
            editor = self._get_editor(self.comments_extensible_path)
            for ex_elem in editor.dom.getElementsByTagName("w16cex:commentExtensible"):
                if ex_elem.getAttribute("w16cex:durableId") == durable_id:
                    ex_elem.parentNode.removeChild(ex_elem)
                    break

        del self.existing_comments[comment_id]
        return True

    def _remove_range_markers(self, comment_id: int) -> None:
        """Drop a comment's range markers and reference run from document.xml.

        Each lookup is guarded: a marker can already be gone (carried away
        inside a rejected revision that hosted it), and a comment missing one
        marker must still lose the others rather than keep a half-range.
        """
        for tag in ("w:commentRangeStart", "w:commentRangeEnd"):
            try:
                marker = self.document_editor.get_node(tag=tag, attrs={"w:id": str(comment_id)})
                marker.parentNode.removeChild(marker)
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

    def move_comment_markers(self, comment_id: int, first: Element, last: Element) -> None:
        """Re-anchor an existing comment's range on the span ``first``..``last``.

        The comment bodies and ids are untouched; the whole thread's range
        markers move together, because a reply's markers are seated relative to
        its parent's and would otherwise be stranded at the old anchor. Used
        when the revisions a note comment was anchored on are resolved while
        other revisions it explains are still pending: leaving the range behind
        would point the rationale at text nobody changed.
        """
        thread = self._thread(comment_id)
        for reply_id, _ in thread:
            self._remove_range_markers(reply_id)
        self._remove_range_markers(comment_id)
        self.document_editor.insert_before(first, self._comment_range_start_xml(comment_id))
        self.document_editor.insert_after(last, self._comment_range_end_xml(comment_id))
        # Parent-first order, so each reply lands on its own parent's new
        # markers — the same relative placement reply_to_comment gave it.
        for reply_id, parent_id in thread:
            self._place_reply_markers(
                reply_id,
                self.document_editor.get_node(tag="w:commentRangeStart", attrs={"w:id": str(parent_id)}),
                self.document_editor.get_node(tag="w:commentReference", attrs={"w:id": str(parent_id)}),
            )

    def reply_ids(self, comment_id: int) -> list[int]:
        """Ids of every comment threaded under ``comment_id``, descendants included.

        A reply is linked to its parent only by ``w15:paraIdParent`` in
        commentsExtended.xml; deleting the parent alone would leave that
        attribute pointing at a paraId no part of the document still holds.
        """
        return [reply_id for reply_id, _ in self._thread(comment_id)]

    def _thread(self, comment_id: int) -> list[tuple[int, int]]:
        """``(reply, its parent)`` for the whole thread under ``comment_id``.

        Ordered so a reply always follows the comment it answers, which is what
        lets ``move_comment_markers`` re-seat each reply on its own parent's
        freshly placed markers rather than flattening the thread onto the root.
        """
        if not self.part_exists(self.comments_extended_path):
            return []
        para_of_id = {cid: info["para_id"] for cid, info in self.existing_comments.items()}
        id_of_para = {para_id: cid for cid, para_id in para_of_id.items() if para_id is not None}
        editor = self._get_editor(self.comments_extended_path)
        children: dict[str, list[int]] = {}
        for ex_elem in editor.dom.getElementsByTagName("w15:commentEx"):
            parent_para = ex_elem.getAttribute("w15:paraIdParent")
            child_id = id_of_para.get(ex_elem.getAttribute("w15:paraId"))
            if parent_para and child_id is not None:
                children.setdefault(parent_para, []).append(child_id)

        found: list[tuple[int, int]] = []
        seen = {comment_id}
        queue = [comment_id]
        while queue:
            current = queue.pop(0)
            for child_id in children.get(para_of_id.get(current) or "", ()):
                # A malformed commentsExtended could name the same child twice,
                # or point a thread back at itself; either would loop forever.
                if child_id in seen:
                    continue
                seen.add(child_id)
                found.append((child_id, current))
                queue.append(child_id)
        return found

    def save_all(self) -> None:
        """Save all modified XML files."""
        for editor in self._editors.values():
            editor.save()

    # ==================== Private: Loading ====================

    def _get_next_comment_id(self) -> int:
        """Get the next available comment ID."""
        if not self.part_exists(self.comments_path):
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
        """Register every comment that has an integer ``w:id``, with its paraId.

        ``para_id`` is None for a comment whose producer wrote no
        ``w14:paraId``; :meth:`_ensure_para_id` gives it one when a reply or a
        resolve first needs it.
        """
        if not self.part_exists(self.comments_path):
            return {}

        editor = self._get_editor(self.comments_path)
        existing = {}

        for comment_elem in editor.dom.getElementsByTagName("w:comment"):
            comment_id = comment_elem.getAttribute("w:id")
            if not comment_id:
                continue

            try:
                existing[int(comment_id)] = {"para_id": self._get_comment_para_id(comment_elem)}
            except ValueError:
                continue

        return existing

    def _get_comment_para_id(self, comment_elem) -> str | None:
        """The ``w14:paraId`` of the comment's last paragraph, or None.

        The last paragraph's is the one commentsExtended.xml and
        commentsIds.xml refer to a comment by.
        """
        paragraphs = comment_elem.getElementsByTagName("w:p")
        return (paragraphs[-1].getAttribute("w14:paraId") or None) if paragraphs else None

    def _comment_element(self, comment_id: int) -> Element:
        """The ``w:comment`` element of a registered comment.

        Of several elements sharing the ``w:id``, the last: the one
        :meth:`_load_existing_comments` registers and ``list_comments`` returns.
        """
        comments = self._get_editor(self.comments_path).dom.getElementsByTagName("w:comment")
        return [elem for elem in comments if _adjudicable_id(elem) == comment_id][-1]

    def _ensure_para_id(self, comment_id: int) -> str:
        """Return the paraId of a registered comment, giving it one if it has none.

        The new id goes on the comment's last paragraph (an empty one is
        appended to a comment that has no paragraph). Call
        :meth:`_prepare_comment_parts` first: it declares ``xmlns:w14`` on a
        comments.xml that never used it.
        """
        info = self.existing_comments[comment_id]
        if info["para_id"] is None:
            comment_elem = self._comment_element(comment_id)
            paragraphs = comment_elem.getElementsByTagName("w:p")
            paragraph = (
                paragraphs[-1]
                if paragraphs
                else self._get_editor(self.comments_path).append_to(comment_elem, "<w:p/>")[0]
            )
            info["para_id"] = _generate_hex_id()
            paragraph.setAttribute("w14:paraId", info["para_id"])
        return info["para_id"]

    def _extended_row(self, para_id: str) -> Element | None:
        """The ``w15:commentEx`` row for ``para_id``, or None."""
        if not self.part_exists(self.comments_extended_path):
            return None
        editor = self._get_editor(self.comments_extended_path)
        for ex_elem in editor.dom.getElementsByTagName("w15:commentEx"):
            if ex_elem.getAttribute("w15:paraId") == para_id:
                return ex_elem
        return None

    def _ensure_extended_row(self, para_id: str) -> Element:
        """Return the ``w15:commentEx`` row for ``para_id``, adding a top-level one if it has none.

        commentsExtended.xml is created when absent. No commentsIds.xml or
        commentsExtensible.xml row is added: both are optional, and Word 2013
        documents carry neither.
        """
        return self._extended_row(para_id) or self._add_to_comments_extended_xml(para_id, parent_para_id=None)

    def _parse_comment(self, comment_elem) -> Comment | None:
        """Parse a w:comment element into a Comment object.

        Returns None for a comment with no integer ``w:id``: it cannot be
        addressed, so it is skipped here as in :meth:`_load_existing_comments`.
        """
        try:
            comment_id = int(comment_elem.getAttribute("w:id"))
        except ValueError:
            return None

        author = comment_elem.getAttribute("w:author") or "Unknown"
        date_str = comment_elem.getAttribute("w:date")

        try:
            date = datetime.fromisoformat(date_str.replace("Z", "+00:00")) if date_str else None
        except ValueError:
            date = None

        text_parts = [get_text_node_data(t_elem) for t_elem in comment_elem.getElementsByTagName("w:t")]

        return Comment(
            id=comment_id,
            text="".join(text_parts),
            author=author,
            date=date,
        )

    # ==================== Private: XML File Creation ====================

    def _ensure_comment_file(self, path: Path, template_name: str) -> None:
        """Ensure a comment XML file exists, creating from template if needed.

        Runs during add_comment(), i.e. after open — a post-open workspace
        write, so it is write-ahead flagged like any editor flush, even though
        the copied template holds no user content yet.
        """
        if not path.exists():
            if self._on_write is not None:
                self._on_write()
            shutil.copy(TEMPLATE_DIR / template_name, path)

    def _prepare_comment_parts(self) -> None:
        """Make the comment parts that already exist ready to receive a comment.

        Called before the first marker goes into document.xml, so everything
        that can fail for a pre-existing part fails while the document is still
        untouched:

        - a comment part that is a link out of the workspace;
        - a comment relationship that cannot be honoured (see
          :meth:`_resolve_comment_parts`);
        - a part whose document element is not the expected prefixed element;
        - a part that binds the prefix its new markup uses to another
          namespace.

        A ``comments.xml`` written by a producer that never emits
        ``w14:paraId`` declares no ``xmlns:w14`` at all; it is added here.

        A missing part is not created here: the ``_add_to_*`` helpers copy it
        from its template once the markers are in. A call that fails therefore
        leaves no new part behind.

        Raises:
            XMLError: If a new comment cannot be written and wired
                (the subclasses ``NodeNotFoundError`` and
                ``MultipleNodesFoundError`` when a part holds the expected
                element not at all or more than once).
        """
        if self._outside_parts:
            names = ", ".join(sorted(path.name for path in self._outside_parts))
            raise XMLError(f"{names}: comment part leaves the package; cannot add a comment")
        if self._unusable_relationships:
            raise XMLError(f"document.xml.rels: {'; '.join(self._unusable_relationships)}; cannot add a comment")
        existing = [
            (self._get_editor(path), part)
            for path, part in zip(self.part_paths, COMMENT_PARTS, strict=True)
            if self.part_exists(path)
        ]
        # Check every part before changing any: a rejected part leaves the
        # others exactly as they were.
        for editor, part in existing:
            root = editor.dom.documentElement
            if editor.get_node(tag=part.root_tag) is not root:
                raise XMLError(f"{editor.xml_path.name}: root element is <{root.tagName}>, expected <{part.root_tag}>")
            if root.hasAttribute(f"xmlns:{part.prefix}"):
                editor.ensure_namespace(part.prefix, part.namespace)
        for editor, part in existing:
            editor.ensure_namespace(part.prefix, part.namespace)

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

    def _add_to_comments_extended_xml(self, para_id: str, parent_para_id: str | None) -> Element:
        """Add a single comment to commentsExtended.xml and return its new row."""
        self._ensure_comment_file(self.comments_extended_path, "commentsExtended.xml")

        editor = self._get_editor(self.comments_extended_path)
        root = editor.get_node(tag="w15:commentsEx")

        if parent_para_id:
            xml = f'<w15:commentEx w15:paraId="{para_id}" w15:paraIdParent="{parent_para_id}" w15:done="0"/>'
        else:
            xml = f'<w15:commentEx w15:paraId="{para_id}" w15:done="0"/>'
        return editor.append_to(root, xml)[0]

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
        return f'<w:commentRangeEnd w:id="{comment_id}"/>\n' + self._comment_ref_run_xml(comment_id)

    def _comment_ref_run_xml(self, comment_id: int) -> str:
        """Generate XML for comment reference run."""
        return f"""<w:r>
  <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
  <w:commentReference w:id="{comment_id}"/>
</w:r>"""


def _find_ancestor_run(node) -> Element | None:
    """Walk parents until reaching a ``w:r`` element, or ``None``."""
    parent = node.parentNode
    while parent:
        if parent.nodeType == parent.ELEMENT_NODE and parent.tagName == "w:r":
            return parent
        parent = parent.parentNode
    return None
