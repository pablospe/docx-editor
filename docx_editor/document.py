"""Main Document class for docx_editor.

Provides the primary user-facing API for editing Word documents with track changes
and comments.
"""

import html
import shutil
import warnings
from collections.abc import Iterable, Iterator
from pathlib import Path
from typing import Literal, overload
from xml.dom.minidom import Attr, Element

import defusedxml.minidom

from .comments import Comment, CommentManager
from .exceptions import (
    DocumentClosedError,
    DocumentProtectedError,
    HashMismatchError,
    ParagraphIndexError,
    UnanchoredNoteWarning,
    _truncate_preview,
)
from .track_changes import (
    EditOperation,
    EditResult,
    EditValidationResult,
    ResolveResult,
    Revision,
    RevisionManager,
    SearchResult,
    UnhandledRevision,
    _ancestor_paragraph,
    _paragraph_hint,
    _resolve_search_target,
    _validate_note,
)
from .workspace import Workspace
from .xml_editor import (
    _TEXTBOX_CONTENT,
    DocxXMLEditor,
    ListItem,
    ParagraphInfo,
    ParagraphLocation,
    ParagraphRef,
    _build_paragraph_info,
    _build_style_numbering_map,
    _build_style_outline_map,
    _build_table_index,
    _compute_heading_paths,
    _compute_paragraph_location,
    _compute_section_indexes,
    _is_inside_element,
    body_paragraphs,
    build_text_map,
    compute_paragraph_hash,
)

# Default page size for list_paragraphs / list_paragraphs_structured. Bounds
# the output of a bare call on large documents; pass limit=None for everything.
_DEFAULT_LIST_LIMIT = 200

# CT_Settings is a sequence (ECMA-376 Part 1 §17.15.1.78), so w:trackRevisions has
# exactly one legal slot. These are the elements the schema puts *before* it —
# the flag goes after the last of them that the document actually has. Only the
# prefix is listed, on purpose: an element we do not recognize (a w14:/w15:
# extension, a producer oddity) is never used as an anchor, so an unknown
# trailing element can never push the flag out of its slot.
_SETTINGS_BEFORE_TRACK_REVISIONS: tuple[str, ...] = (
    "writeProtection",
    "view",
    "zoom",
    "removePersonalInformation",
    "removeDateAndTime",
    "doNotDisplayPageBoundaries",
    "displayBackgroundShape",
    "printPostScriptOverText",
    "printFractionalCharacterWidth",
    "printFormsData",
    "embedTrueTypeFonts",
    "embedSystemFonts",
    "saveSubsetFonts",
    "saveFormsData",
    "mirrorMargins",
    "alignBordersAndEdges",
    "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter",
    "gutterAtTop",
    "hideSpellingErrors",
    "hideGrammaticalErrors",
    "activeWritingStyle",
    "proofState",
    "formsDesign",
    "attachedTemplate",
    "linkStyles",
    "stylePaneFormatFilter",
    "stylePaneSortMethod",
    "documentType",
    "mailMerge",
    "revisionView",
)


def _on_off(value: str) -> bool | None:
    """Read an ST_OnOff attribute value (ECMA-376 Part 1 §17.17.4).

    Returns None for anything outside the six spellings the schema allows.
    "The attribute is not there" and "the attribute holds something we cannot
    read" are different facts, and the two callers want opposite things from
    the second one — the protection guard has to fail closed, the track-changes
    read has to stop reporting an unreadable flag as already on — so neither
    can be folded into a shared default here.

    Args:
        value: The raw attribute value, from an attribute that is present.

    Returns:
        True or False for a legal value, None for one this schema does not
        define.
    """
    normalized = value.strip().lower()
    if normalized in {"1", "true", "on"}:
        return True
    if normalized in {"0", "false", "off"}:
        return False
    return None


def _local_name(tag: str) -> str:
    """Strip a namespace prefix from a tag or attribute name, so reads work
    whatever prefix the producing application chose ("w:zoom" and "wx:zoom" are
    both "zoom")."""
    return tag.split(":")[-1]


def _attr_node(elem: Element, local_name: str) -> Attr | None:
    """The attribute node named ``local_name``, whatever prefix it carries.

    Returns the node rather than its value because a caller that removes the
    attribute needs its actual name; None when the element has no such
    attribute, which the ST_OnOff readers must tell apart from an empty value.
    """
    for attr in elem.attributes.values():
        if _local_name(attr.name) == local_name:
            return attr
    return None


# Why a note= rationale had nothing to anchor to. Each is inferred where the
# information actually exists — the edit method knows whether it asked for a
# no-op, only the span lookup knows a group is paragraph-mark-only — and lands
# verbatim in the UnanchoredNoteWarning message.
_NOTE_NOOP_REPLACE = "the replace was a no-op ('find' equals 'replace_with'), so no revision was created"
_NOTE_AMENDED_INSERTION = "the edit amended your own pending insertion, which creates no new revision"
_NOTE_REWRITE_NO_REVISION = (
    "the rewrite created no revisions (text unchanged, or every change amended your own pending insertions)"
)
_NOTE_PARAGRAPH_MARK_ONLY = "the only revision is a tracked paragraph mark, which cannot carry a comment anchor"


def _batch_note_reason(op: EditOperation) -> str:
    """Why a batch operation's note had no group, inferred from the operation."""
    if op.action == "replace" and op.find == op.replace_with:
        return _NOTE_NOOP_REPLACE
    return _NOTE_AMENDED_INSERTION


def _unanchored_note_message(note: str, reason: str) -> str:
    """The one UnanchoredNoteWarning message, with the cause filled in."""
    shown = note if len(note) <= 60 else note[:57] + "..."
    return (
        f"note={shown!r} was not recorded: {reason}. The edit itself applied — only the note "
        f"was dropped, and this EditResult's comment_id is None. Use add_comment() to attach "
        f"it as an ordinary, document-scoped comment if it is still wanted."
    )


def _require_ref_string(paragraph: str | None, field: str | None = None) -> str:
    """Reject a missing or non-string paragraph ref before it can silently select
    the RevisionManager's document-wide search branch (its ``paragraph=None``
    mode is intentional at that layer, not at this one).

    ``paragraph`` is Optional in the edit methods' signatures only so a
    SearchResult target can supply it; those callers pass ``field`` (the name of
    their target parameter) to get the "or pass a SearchResult as …" hint.
    Callers whose ref is genuinely required and positional (split_paragraph)
    omit it. Returns the ref, narrowed to ``str`` for the call that follows.
    """
    if not isinstance(paragraph, str):
        message = f"'paragraph' must be a paragraph ref string like 'P3#a7b2', got {type(paragraph).__name__}"
        raise ValueError(message if field is None else f"{message} — {_paragraph_hint(field)}")
    return paragraph


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

    def __init__(self, workspace: Workspace, *, allow_protected: bool = False):
        """Initialize Document with a workspace.

        Use Document.open() instead of calling this directly.

        Args:
            workspace: Workspace instance for the document
            allow_protected: If True, open a document whose editing protection
                is enforced instead of raising DocumentProtectedError.

        Raises:
            DocumentProtectedError: If the document enforces a body-locking
                protection mode and allow_protected is False.
        """
        # First, before any editor exists: a document we are going to refuse
        # must not be written to (_setup_tracking writes four parts).
        self._check_protection(workspace, allow_protected)

        self._workspace = workspace
        self._closed = False
        # Lazily parsed word/styles.xml maps (see _style_maps).
        self._style_maps_cache: tuple[dict[str, int], dict[str, ListItem]] | None = None
        # note= rationale channel. A note comment is revision-scoped: it is
        # deleted as soon as every group it explains is gone (accepted,
        # rejected, or carried away with a foreign host), so rationale never
        # outlives the proposal. All three maps are per-open-Document, like the
        # group registry they key on, and stay empty for callers that never
        # pass a note — which is what keeps _reap_note_comments free on that
        # path.
        self._note_comments: dict[int, int] = {}  # group id -> comment id
        self._note_groups: dict[int, set[int]] = {}  # comment id -> every group it explains
        self._note_anchors: dict[int, int] = {}  # comment id -> the group its markers bracket

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
        *,
        allow_protected: bool = False,
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
            allow_protected: If True, open a document whose editing protection
                is enforced (Word's *Restrict Editing*) instead of raising
                DocumentProtectedError. The protection is left in place, so the
                saved document is still restricted when it reaches Word.

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
            DocumentProtectedError: If the document enforces an editing
                protection that locks its body text (``readOnly``, ``forms`` or
                ``comments``). Carries ``path`` and ``mode``. Unprotect it in
                Word, or pass allow_protected=True. A document enforcing
                ``trackedChanges`` opens normally — that mode asks for exactly
                what this library does.

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
        try:
            return cls(workspace, allow_protected=allow_protected)
        except BaseException:
            # The workspace is live from here on: it holds the advisory lock and,
            # when this call unpacked it, a directory. A raise inside __init__
            # hands the caller no object to close(), so without this the document
            # would be locked against its own retry — the same guard
            # Workspace.__init__ keeps over its own failure path. The lock goes
            # either way; the directory only if we made it, since an adopted one
            # is the caller's (see Workspace.created and unpack_document's own
            # created_output_dir guard).
            try:
                workspace.close(cleanup=workspace.created)
            except Exception:
                # close() releases the lock in a finally, so a failed rmtree
                # (a scanner still holding a handle on Windows, say) leaves a
                # clean directory and an unlocked document behind. Not worth
                # replacing the error the caller actually has to act on.
                pass
            raise

    @classmethod
    def discard_workspace(cls, path: str | Path, *, workspace_dir: str | Path | None = None) -> bool:
        """Delete a document's workspace so the next :meth:`open` starts clean.

        The recovery tool for a crashed or killed script: that session's
        workspace stays behind flagged as holding unsaved changes, and every
        later ``open()`` of the same document raises
        :class:`~docx_editor.exceptions.WorkspaceSyncError` (or
        :class:`~docx_editor.exceptions.WorkspaceLockedError` if its advisory
        lock is still on disk) until the workspace is dealt with. One call here
        resets that state, instead of passing ``force_recreate=True`` on every
        open from then on or deleting the cache directory by hand.

        **This discards whatever unsaved edits the workspace holds** — same
        destruction as ``open(force_recreate=True)``, just without opening. To
        rescue those edits first, save the orphaned workspace elsewhere and
        inspect it before discarding (see the Example below).

        The workspace's advisory lock sidecar goes with it, including a lock
        held by a live session — so a wedged process cannot keep the document
        unopenable. Make sure that process is really gone first: two sessions
        sharing one workspace silently overwrite each other's saves.

        Args:
            path: Path to the .docx file whose workspace should be deleted. Not
                required to exist — a workspace outlives its source, so this
                still cleans up after a moved or deleted document.
            workspace_dir: Base directory override, matching :meth:`open`'s
                argument. Must be the same value the workspace was created
                with, or a different workspace is targeted (and nothing is
                found).

        Returns:
            True if a workspace was deleted, False if there was none — so this
            is safe to call unconditionally at the start of a script.

        Example:
            # Idempotent reset before a fresh run:
            Document.discard_workspace("contract.docx")
            doc = Document.open("contract.docx")

            # Rescue the orphaned workspace's edits first:
            from docx_editor.workspace import Workspace
            Workspace("contract.docx", create=False).save("rescued.docx")
            Document.discard_workspace("contract.docx")
        """
        return Workspace.delete(path, workspace_dir=workspace_dir)

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

    @property
    def has_textbox_content(self) -> bool:
        """Whether any paragraph is hidden inside a drawing's text box.

        Text boxes are not an editing surface: their paragraphs are absent
        from every listing, their text from every view, search and hash (see
        :func:`~docx_editor.xml_editor.body_paragraphs`). That makes an
        all-text-box document — a poster, a flyer, a certificate — read as
        blank, which is indistinguishable from a genuinely empty document
        without this flag.

        Exactly the complement of what ``body_paragraphs`` drops: True means
        at least one ``w:p`` was excluded, so any text those paragraphs carry
        is not reachable from here. To read it, go through HTML —
        ``soffice --headless --convert-to html file.docx`` then
        ``pandoc file.html -t plain`` — rather than reporting the document
        as empty; pandoc may render a ``[ShapeN]`` label beside a box's text,
        from the placeholder LibreOffice exports for a named shape — ignore
        those. Its ``txt:Text`` filter and
        pandoc reading the ``.docx`` directly both drop text boxes silently.
        """
        self._ensure_open()
        dom = self._document_editor.dom
        return any(_is_inside_element(p, _TEXTBOX_CONTENT) for p in dom.getElementsByTagName("w:p"))

    # ==================== Track Changes API ====================

    def find_text(self, text: str, occurrence: int = 0, paragraph: str | None = None) -> SearchResult | None:
        """Find text in the document, including across element boundaries.

        Args:
            text: Text to search for (must be non-empty; a tab mark is
                ``\\t``, as in ``get_visible_text()``)
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

    def _compute_new_ref(self, old_ref: str, paragraphs: list[Element]) -> str:
        """Compute a fresh paragraph reference after mutation.

        ``paragraphs`` is the caller's body-paragraph snapshot (see
        :func:`~docx_editor.xml_editor.body_paragraphs`), so a batch pays for
        one full-DOM walk rather than one per result.
        """
        ref = ParagraphRef.parse(old_ref)
        p = paragraphs[ref.index - 1]
        new_hash = compute_paragraph_hash(p)
        return f"P{ref.index}#{new_hash}"

    def _edit_result(
        self,
        old_ref: str,
        group_id: int | None,
        paragraphs: list[Element] | None = None,
        element_index: dict[str, list[Element]] | None = None,
        comment_id: int | None = None,
    ) -> EditResult:
        """Build an EditResult from a mutated paragraph's old ref, group, and changeset."""
        revision_ids = self._revision_manager.group_revisions(group_id) if group_id is not None else ()
        changeset_id = self._revision_manager.changeset_id_of(group_id) if group_id is not None else None
        if paragraphs is None:
            paragraphs = body_paragraphs(self._document_editor.dom)
        refs = self._resulting_refs(old_ref, group_id, paragraphs, element_index)
        return EditResult(
            refs[0],
            group_id=group_id,
            revision_ids=revision_ids,
            changeset_id=changeset_id,
            refs=refs,
            comment_id=comment_id,
        )

    def _attach_notes(self, pairs: list[tuple[int | None, str | None, str]]) -> list[int | None]:
        """Anchor each operation's ``note=`` as a comment on its revisions.

        The single implementation behind ``note=`` on every edit method and on
        ``batch_edit``. Each pair is ``(group_id, note, reason)`` for one
        operation, in input order; ``reason`` is that operation's explanation
        for a ``None`` group, supplied by the caller because only the caller
        knows whether it asked for a no-op.

        Operations of one call that carry the *same* note text share one
        comment, anchored on the first of them with an anchorable revision —
        20 redlines explaining themselves the same way leave the reviewer one
        comment, not 20, while genuinely different rationales stay distinct.

        Returns one ``comment_id | None`` per input pair. Every ``None`` where a
        note was actually given comes with exactly one ``UnanchoredNoteWarning``
        naming why: a dropped rationale is never silent.

        Comment markers change no paragraph text, so this can run before or
        after the caller builds its refs; it costs one DOM walk for the whole
        call, and none at all when no pair carries a note.
        """
        # Called after the caller's own mutation, and by every edit path — the
        # one place that can notice an edit which amended an *earlier* note's
        # insertion out of existence. That empties a group with no accept or
        # reject call to key cleanup on, so without this the rationale would
        # outlive its revision and ship in the saved file. Free until a note
        # has actually been attached.
        self._reap_note_comments()

        results: list[int | None] = [None] * len(pairs)
        grouped = [(gid, note) for gid, note, _ in pairs if note is not None and gid is not None]
        spans = self._revision_manager.group_spans({gid for gid, _ in grouped}) if grouped else {}

        # Dedupe by note text: the first group *with a span* owns the comment,
        # so a note shared with a paragraph-mark-only operation still lands.
        owners: dict[str, int] = {}
        for gid, note in grouped:
            if note not in owners and gid in spans:
                owners[note] = gid
        comment_of_note: dict[str, int] = {}
        for note, gid in owners.items():
            first, last = spans[gid]
            comment_of_note[note] = self._comment_manager.add_comment_on_elements(first, last, note)
            self._note_anchors[comment_of_note[note]] = gid

        for i, (gid, note, reason) in enumerate(pairs):
            if note is None:
                continue
            comment_id = comment_of_note.get(note)
            if comment_id is None:
                # This note reached no anchorable revision anywhere in the call.
                self._warn_unanchored_note(note, reason if gid is None else _NOTE_PARAGRAPH_MARK_ONLY)
                continue
            # Recorded — by this operation or by a sibling sharing its text.
            # Either way the rationale is in the document, so no warning.
            results[i] = comment_id
            # Only a group the markers could sit on is registered: one with no
            # span (a pure paragraph-mark split) can neither host the anchor if
            # the current one is resolved nor honestly keep the comment alive.
            if gid in spans:
                self._note_comments[gid] = comment_id
                self._note_groups.setdefault(comment_id, set()).add(gid)
        return results

    def _warn_unanchored_note(self, note: str, reason: str) -> None:
        """Report a note that had no revision to anchor to, at the caller's line."""
        warnings.warn(
            _unanchored_note_message(note, reason),
            UnanchoredNoteWarning,
            # 4 frames out of warnings.warn: _warn_unanchored_note ->
            # _attach_notes -> Document.replace (or any other edit method) ->
            # caller, so the warning points at the caller's edit line rather
            # than at library internals.
            stacklevel=4,
        )

    def _reap_note_comments(self) -> None:
        """Keep every note comment pointing at a redline that is still pending.

        Called from every resolution entry point — one invariant instead of a
        special case per verb: a note explains a *proposal*, and accepting ends
        the proposal just as rejecting does, so the comment goes either way.
        Also called from the edit paths, where an amendment can empty a group
        with no resolution call at all (see ``_attach_notes``).

        Every registered group is examined, never a narrowed candidate set:
        which revisions a call removed is not the same question as which groups
        it emptied. Rejecting *another author's* insertion carries away any of
        our revisions nested inside it, so even a sweep filtered to somebody
        else can leave one of our groups with nothing left to explain. And a
        group can survive with nothing a marker could bracket — its content
        revisions amended away, only a tracked paragraph mark left — so a group
        counts here only while ``group_spans`` still offers a span for it.

        A comment whose anchor is gone moves onto another group it explains;
        one with no such group left is deleted, replies included. Registered
        groups only ever lose revisions, so a group dropped here can never
        become anchorable again.

        Returns before any DOM work when no note was ever attached, so callers
        that never pass ``note=`` pay nothing.
        """
        if not self._note_comments:
            return
        dead = self._revision_manager.groups_are_dead(set(self._note_comments))
        spans = self._revision_manager.group_spans(set(self._note_comments) - dead)
        for comment_id in list(self._note_groups):
            # Every registered comment still exists: delete_comment() forgets
            # the ones a caller removes, so nothing here can re-anchor a comment
            # with no body — which would make the file unreadable.
            groups = self._note_groups[comment_id]
            live = {gid for gid in groups if gid in spans}
            if live == groups and self._note_anchors[comment_id] in live:
                continue
            for group_id in groups - live:
                del self._note_comments[group_id]
            if not live:
                # Nothing left it can honestly bracket: the note goes, and
                # delete_comment takes the replies threaded under it with it.
                self._comment_manager.delete_comment(comment_id)
                self._forget_note_comment(comment_id, groups)
                continue
            self._note_groups[comment_id] = live
            if self._note_anchors[comment_id] in live:
                continue
            # The markers sat on an edit that is now resolved, where they would
            # bracket text nobody changed. Move the thread onto one that is not.
            anchor = min(live)
            self._comment_manager.move_comment_markers(comment_id, *spans[anchor])
            self._note_anchors[comment_id] = anchor

    def _forget_note_comment(self, comment_id: int, groups: Iterable[int]) -> None:
        """Drop every trace of one note comment from the three note maps."""
        self._note_groups.pop(comment_id, None)
        self._note_anchors.pop(comment_id, None)
        for group_id in groups:
            if self._note_comments.get(group_id) == comment_id:
                del self._note_comments[group_id]

    def _resulting_refs(
        self,
        old_ref: str,
        group_id: int | None,
        paragraphs: list[Element],
        element_index: dict[str, list[Element]] | None = None,
    ) -> tuple[str, ...]:
        """Refs of every paragraph a group's revisions landed in, document order.

        A ``\\n`` edit splits one paragraph into several, so the group spans
        more than one; this reads the *live* positions of the group's own
        paragraphs, staying correct even when a split elsewhere in the same
        batch shifted every later index. ``element_index`` (a pre-built ``w:id``
        -> element map) is passed by a batch in which some op split, forcing the
        identity path for *every* op so a shifted non-split op still resolves to
        its live paragraph. Falls back to the single recomputed ref at
        ``old_ref``'s index when the edit created no group (a no-op).
        """
        mgr = self._revision_manager
        # Cheap path: no group, or (outside a splitting batch) a group that made
        # no split — one paragraph at ``old_ref``'s index. split_count() is
        # DOM-walk-free, so a no-split edit never pays for an element index.
        if group_id is None or (element_index is None and mgr.split_count(group_id) == 0):
            return (self._compute_new_ref(old_ref, paragraphs),)
        # Locate the group's paragraphs by identity so refs stay correct even
        # after a split shifted every later index.
        if element_index is None:
            element_index = mgr._revision_element_index()
        index_map = {id(p): i for i, p in enumerate(paragraphs, start=1)}
        # First resulting paragraph = the lowest live index among the group's
        # own paragraphs. A pure split (split_paragraph) leaves the tail
        # paragraph with no revision of its own, so the resulting paragraphs are
        # `split_count + 1` consecutive siblings starting there — walk them.
        member_paras = [
            para
            for rev_id in mgr.group_revisions(group_id)
            for elem in element_index.get(str(rev_id), ())
            if (para := _ancestor_paragraph(elem)) is not None and id(para) in index_map
        ]
        # Both fallbacks below are defensive: a split group's revisions always
        # resolve to live, consecutive paragraphs, so member_paras is non-empty
        # and the sibling walk never runs short.
        if not member_paras:  # pragma: no cover
            return (self._compute_new_ref(old_ref, paragraphs),)
        current = min(member_paras, key=lambda p: index_map[id(p)])
        refs: list[str] = []
        for _ in range(mgr.split_count(group_id) + 1):
            if current is None or id(current) not in index_map:  # pragma: no cover
                break
            refs.append(f"P{index_map[id(current)]}#{compute_paragraph_hash(current)}")
            sibling = current.nextSibling
            while sibling is not None and (sibling.nodeType != sibling.ELEMENT_NODE or sibling.tagName != "w:p"):
                sibling = sibling.nextSibling
            current = sibling
        return tuple(refs) if refs else (self._compute_new_ref(old_ref, paragraphs),)

    def paragraph_count(self) -> int:
        """Return the total number of paragraphs in the document.

        Cheap bounds check for pagination — avoids building the full
        :meth:`list_paragraphs` result just to learn the count.

        Paragraphs inside a drawing's text box are not counted — text boxes
        are excluded from the ref index space entirely, so no ref addresses
        one (see :func:`~docx_editor.xml_editor.body_paragraphs`).

        Returns:
            Total number of paragraphs (the highest valid 1-based ref index).
        """
        self._ensure_open()
        return len(body_paragraphs(self._document_editor.dom))

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
        paragraphs = body_paragraphs(self._document_editor.dom)
        begin = start - 1
        end = begin + limit if limit is not None else None
        return enumerate(paragraphs[begin:end], start=begin + 1)

    def list_paragraphs_structured(
        self, *, start: int = 1, limit: int | None = _DEFAULT_LIST_LIMIT
    ) -> list[ParagraphInfo]:
        """List paragraphs as structured :class:`ParagraphInfo` records.

        Like :meth:`list_paragraphs`, but returns named records instead of
        pipe-delimited preview strings. The ``text`` field is always the full,
        untruncated paragraph text — there is no ``max_chars`` parameter.
        ``str(info)`` uses the same ``"P{i}#{hash}| {text}"`` delimiter format
        as :meth:`list_paragraphs`, but always with the full text (it matches
        :meth:`list_paragraphs` output only when that call's ``max_chars`` is
        large enough to avoid truncation).

        Each record also carries the paragraph's cheap structural facts —
        ``in_table``, ``style`` and ``outline_level`` (see
        :class:`ParagraphInfo`) — so building a table of contents or skipping
        table-cell paragraphs needs no second pass:
        :meth:`list_paragraph_locations` is only necessary for table
        coordinates, list numbering, heading paths and section indexes.

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

            # Table of contents, one pass, no locations call:
            toc = [(i.outline_level, i.text) for i in everything if i.outline_level is not None]
        """
        self._ensure_open()
        # One styles.xml parse for the whole slice resolves style-defined outline
        # levels; the per-paragraph work stays a single text map (ISSUES.md #52).
        style_outlines, _ = self._style_maps()
        return [_build_paragraph_info(p, i, style_outlines) for i, p in self._iter_paragraph_slice(start, limit)]

    def get_paragraph(self, index: int) -> ParagraphInfo:
        """Return one paragraph as a structured :class:`ParagraphInfo` record.

        Single-item counterpart to :meth:`list_paragraphs_structured`. The
        returned record — index, hash-anchored ref, full untruncated text, plus
        ``in_table``/``style``/``outline_level`` — is identical to the one that
        method would emit for the same paragraph.

        Note:
            Fixed-size output, but O(document) work: every call walks all
            ``<w:p>`` elements to reach one. (``word/styles.xml`` is parsed once
            per open Document, not per call.) Fine for a one-off lookup; for
            many paragraphs call ``list_paragraphs_structured(limit=None)`` once
            and index the result instead of looping over this.

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
        paragraphs = body_paragraphs(self._document_editor.dom)
        if index < 1 or index > len(paragraphs):
            raise ParagraphIndexError(index, len(paragraphs))
        style_outlines, _ = self._style_maps()
        return _build_paragraph_info(paragraphs[index - 1], index, style_outlines)

    def context(self, ref: str, window: int = 2) -> list[ParagraphInfo]:
        """Return the paragraphs surrounding ``ref``, in document order.

        Fetches the referenced paragraph plus up to ``window`` paragraphs on
        each side (fewer at the document edges) — the "show me what's around
        this match" helper for search results: pass a
        :class:`~docx_editor.track_changes.SearchResult`'s ``paragraph_ref``
        straight in. The records are identical to what
        :meth:`list_paragraphs_structured` would emit for the same indexes.

        Note:
            Fixed-size output, but O(document) work: resolving the ref walks all
            ``<w:p>`` elements, twice over (hash check, then the slice). Fine per
            match; to annotate many matches, call
            ``list_paragraphs_structured(limit=None)`` once and slice around
            each ``paragraph_index`` instead.

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
        paragraphs = body_paragraphs(self._document_editor.dom)
        if parsed.index < 1 or parsed.index > len(paragraphs):
            raise ParagraphIndexError(parsed.index, len(paragraphs))
        p = paragraphs[parsed.index - 1]
        actual_hash = compute_paragraph_hash(p)
        if actual_hash != parsed.hash:
            tm = build_text_map(p)
            preview = _truncate_preview(tm.text)
            raise HashMismatchError(parsed.index, parsed.hash, actual_hash, preview)
        return parsed.index, paragraphs

    def _style_maps(self) -> tuple[dict[str, int], dict[str, ListItem]]:
        """Outline-level and numbering maps defined by paragraph styles.

        One ``word/styles.xml`` parse serves both maps, memoized for the life of
        this Document: nothing in this library ever writes ``styles.xml`` (it is
        read-only workspace content), so the maps cannot go stale within a
        session. The parse is expensive enough to matter — ~340 ms on a
        3300-paragraph document — which is what makes memoizing it necessary now
        that :meth:`get_paragraph` needs it too (ISSUES.md #52); the location
        APIs get the same win for free. A document without a styles part
        degrades to ``({}, {})``.

        Table indexes, heading paths and section indexes are deliberately *not*
        cached: those derive from ``document.xml``, which edits do change.

        The returned maps are the cached objects themselves, not copies —
        callers must treat them as read-only, since a mutation would outlive
        the call and be seen by every later consumer.

        The parse deliberately skips :class:`XMLEditor`'s line-tracking parser
        (halving the cost on a large styles part): the map builders read style
        ids and attributes only, and nothing ever edits ``styles.xml``, so
        ``parse_position`` metadata would go unused.
        """
        if self._style_maps_cache is None:
            styles_path = self._workspace.word_path / "styles.xml"
            if not styles_path.exists():
                self._style_maps_cache = ({}, {})
            else:
                styles_dom = defusedxml.minidom.parse(str(styles_path))
                self._style_maps_cache = (
                    _build_style_outline_map(styles_dom),
                    _build_style_numbering_map(styles_dom),
                )
        return self._style_maps_cache

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
        paragraphs = body_paragraphs(dom)
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
        Inserted text is included, deleted text is excluded, and a tab mark
        (``<w:tab/>``) is one ``\\t`` — the coordinate space ``find_text()``
        searches and ``SearchResult`` offsets index. Text a tracked move took
        away (``w:moveFrom``) is excluded like a deletion; its destination
        (``w:moveTo``) is included. Text inside a
        drawing's text box is excluded too — it belongs to the box, not to
        any addressable paragraph. A document whose content lives entirely in
        text boxes therefore returns nothing but the separators between its
        host paragraphs, so test it with ``.strip()``;
        :attr:`has_textbox_content` tells that case from a genuinely empty
        document.

        Returns:
            The visible text content
        """
        self._ensure_open()
        paragraphs = body_paragraphs(self._document_editor.dom)
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
        Text inside a drawing's text box is excluded, exactly as in
        get_visible_text().
        Read-only: paragraph references, hashes, and all editing operations
        keep working on the accepted (visible) view.

        Returns:
            The original text content
        """
        self._ensure_open()
        paragraphs = body_paragraphs(self._document_editor.dom)
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
        not escaped and tabs/breaks are not rendered (unlike
        ``get_visible_text()``, where a tab mark is a ``\\t``). Text inside a
        drawing's text box does not appear at all — box content is excluded
        from every text view and from paragraph enumeration.

        Returns:
            The marked-up text content

        Example:
            doc.replace("30 days", "60 days", paragraph="P2#f3c1")
            print(doc.get_markup_text())
            # ... [del#3:Reviewer]30 days[/del][ins#4:Reviewer]60 days[/ins] ...
        """
        self._ensure_open()
        return self._revision_manager.get_markup_text()

    def replace(
        self,
        find: str | SearchResult,
        replace_with: str,
        *,
        paragraph: str | None = None,
        occurrence: int | None = None,
        note: str | None = None,
    ) -> EditResult:
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

        A replace landing wholly inside your own pending insertion *amends*
        that insertion: your unsaved text is rewritten in place rather than
        counter-proposed, whether the match covers part of the insertion or
        all of it. No revision is created, so the result carries
        ``group_id=None`` and ``revision_ids=()`` (with an updated ref) — to
        undo the amendment, reject the group of the insertion it amended:
        the one holding the end of the match, which keeps its id and its
        group. A match spanning two of your own adjacent insertions
        consolidates into that one, dropping any insertion consumed whole.

        Args:
            find: Text to find and replace, or a
                :class:`~docx_editor.track_changes.SearchResult` from
                :meth:`find_text`/:meth:`find_all`. A SearchResult already
                carries the matched text, its paragraph and its occurrence, so
                pass neither ``paragraph`` nor ``occurrence`` with one.
            replace_with: Replacement text
            paragraph: Paragraph reference from list_paragraphs() (e.g.,
                "P2#f3c1"). Required unless ``find`` is a SearchResult.
            occurrence: Which occurrence within the paragraph (0 = first,
                1 = second, etc.). Omitted → ``find`` must be unique in the
                paragraph, else AmbiguousTextError (use find_all() to
                enumerate the matches, or pass an explicit occurrence).
            note: Rationale for this edit, anchored as a comment bracketing
                the revisions it creates, with its id on the result's
                ``comment_id``. The comment is revision-scoped: it is deleted
                as soon as the edit is resolved, accepted or rejected alike.
                Operations of one call that share the same note text share one
                comment. For rationale that must survive resolution, use
                ``add_comment()`` instead — those comments are
                document-scoped.

        Returns:
            EditResult — the new paragraph reference with updated hash (e.g.,
            "P2#c3d4"; usable anywhere a ref string is expected), carrying
            ``group_id``/``revision_ids`` of the revisions this edit created
            for accept_group()/reject_group(), and ``comment_id`` when a
            ``note`` was anchored.

        Warns:
            UnanchoredNoteWarning: If ``note`` was given but the call created
                no revision to anchor it on (a no-op, or an amendment to your
                own pending insertion). The edit still applies; the note is
                dropped and ``comment_id`` is None.

        Raises:
            ValueError: If ``find`` is not a non-empty string or contains a
                tab (``\\t`` — a tab mark can be matched but not replaced yet,
                ISSUES.md #6), ``replace_with`` is not a string, ``paragraph`` is missing or
                not a ref string, ``occurrence`` is negative or not an integer,
                ``note`` is neither None nor a non-empty control-character-free
                string, or ``find`` is a SearchResult and
                ``paragraph``/``occurrence`` was given too.
            TextNotFoundError: If ``find`` is absent or ``occurrence`` is out
                of range for the paragraph.
            AmbiguousTextError: If ``occurrence`` is omitted and ``find``
                matches more than once in the paragraph.
            HashMismatchError: If the paragraph hash is stale — including a
                SearchResult whose paragraph was edited after the search.

        Example:
            new_ref = doc.replace("30 days", "60 days", paragraph="P2#f3c1")
            doc.replace("other text", "new text", paragraph=new_ref)

            # Straight from a search — no ref or occurrence bookkeeping:
            match = doc.find_text("30 days")
            doc.replace(match, "60 days")
        """
        self._ensure_open()
        find, paragraph, occurrence = _resolve_search_target(
            find, paragraph, occurrence, ctx="replace(): ", field="'find'"
        )
        paragraph = _require_ref_string(paragraph, "'find'")
        # Before the edit: a bad note must not leave an applied edit behind it.
        _validate_note(note, ctx="replace(): ")
        change_id = self._revision_manager.replace_text(find, replace_with, occurrence=occurrence, paragraph=paragraph)
        group_id = self._revision_manager.group_id_of(change_id)
        reason = _NOTE_NOOP_REPLACE if find == replace_with else _NOTE_AMENDED_INSERTION
        comment_id = self._attach_notes([(group_id, note, reason)])[0]
        return self._edit_result(paragraph, group_id, comment_id=comment_id)

    def delete(
        self,
        text: str | SearchResult,
        *,
        paragraph: str | None = None,
        occurrence: int | None = None,
        note: str | None = None,
    ) -> EditResult:
        """Mark text as deleted with tracked changes.

        Args:
            text: Text to mark as deleted, or a
                :class:`~docx_editor.track_changes.SearchResult` from
                :meth:`find_text`/:meth:`find_all` — which supplies the text,
                the paragraph and the occurrence, so pass neither
                ``paragraph`` nor ``occurrence`` with one.
            paragraph: Paragraph reference from list_paragraphs() (e.g.,
                "P2#f3c1"). Required unless ``text`` is a SearchResult.
            occurrence: Which occurrence within the paragraph (0 = first,
                1 = second, etc.). Omitted → ``text`` must be unique in the
                paragraph, else AmbiguousTextError.
            note: Rationale for this edit, anchored as a comment bracketing
                the deletion it creates, and deleted with it when the deletion
                is resolved (see :meth:`replace`).

        Returns:
            EditResult — the new paragraph reference with updated hash (e.g.,
            "P2#c3d4"), carrying ``group_id``/``revision_ids`` of the
            revisions this edit created, and ``comment_id`` when a ``note``
            was anchored.

        Warns:
            UnanchoredNoteWarning: If ``note`` was given but the call created
                no revision to anchor it on. The edit still applies; the note
                is dropped and ``comment_id`` is None.

        Raises:
            ValueError: If ``text`` is not a non-empty string or contains a
                tab (``\\t`` — a tab mark can be matched but not deleted yet,
                ISSUES.md #6), ``paragraph`` is missing or not a ref string,
                ``occurrence`` is negative or not an integer, ``note`` is
                neither None nor a non-empty control-character-free string,
                or ``text`` is a SearchResult
                and ``paragraph``/``occurrence`` was given too.
            TextNotFoundError: If ``text`` is absent or ``occurrence`` is out
                of range for the paragraph.
            AmbiguousTextError: If ``occurrence`` is omitted and ``text``
                matches more than once in the paragraph.
            HashMismatchError: If the paragraph hash is stale.

        Example:
            new_ref = doc.delete("obsolete clause", paragraph="P2#f3c1")
            doc.delete(doc.find_text("obsolete clause"))  # same edit, from a search
        """
        self._ensure_open()
        text, paragraph, occurrence = _resolve_search_target(
            text, paragraph, occurrence, ctx="delete(): ", field="'text'"
        )
        paragraph = _require_ref_string(paragraph, "'text'")
        _validate_note(note, ctx="delete(): ")
        change_id = self._revision_manager.suggest_deletion(text, occurrence=occurrence, paragraph=paragraph)
        group_id = self._revision_manager.group_id_of(change_id)
        comment_id = self._attach_notes([(group_id, note, _NOTE_AMENDED_INSERTION)])[0]
        return self._edit_result(paragraph, group_id, comment_id=comment_id)

    def insert_after(
        self,
        anchor: str | SearchResult,
        text: str,
        *,
        paragraph: str | None = None,
        occurrence: int | None = None,
        note: str | None = None,
    ) -> EditResult:
        """Insert text after anchor with tracked changes.

        Args:
            anchor: Text to find as insertion point, or a
                :class:`~docx_editor.track_changes.SearchResult` from
                :meth:`find_text`/:meth:`find_all` — which supplies the anchor
                text, its paragraph and its occurrence, so pass neither
                ``paragraph`` nor ``occurrence`` with one.
            text: Text to insert after the anchor
            paragraph: Paragraph reference from list_paragraphs() (e.g.,
                "P2#f3c1"). Required unless ``anchor`` is a SearchResult.
            occurrence: Which occurrence of anchor within the paragraph
                (0 = first). Omitted → ``anchor`` must be unique in the
                paragraph, else AmbiguousTextError.
            note: Rationale for this edit, anchored as a comment bracketing
                the insertion it creates, and deleted with it when the
                insertion is resolved (see :meth:`replace`).

        Returns:
            EditResult — the new paragraph reference with updated hash (e.g.,
            "P2#c3d4"), carrying ``group_id``/``revision_ids`` of the
            revisions this edit created, and ``comment_id`` when a ``note``
            was anchored.

        Warns:
            UnanchoredNoteWarning: If ``note`` was given but the call created
                no revision a comment can bracket — an amendment to your own
                pending insertion, or a bare ``"\\n"`` whose only revision is
                the paragraph mark. The edit still applies; the note is
                dropped and ``comment_id`` is None.

        Raises:
            ValueError: If ``anchor`` is not a non-empty string, ``text`` is
                not a string, ``paragraph`` is missing or not a ref string,
                ``occurrence`` is negative or not an integer, ``note`` is
                neither None nor a non-empty control-character-free string, or
                ``anchor`` is a SearchResult and ``paragraph``/``occurrence``
                was given too.
            TextNotFoundError: If ``anchor`` is absent or ``occurrence`` is
                out of range for the paragraph.
            AmbiguousTextError: If ``occurrence`` is omitted and ``anchor``
                matches more than once in the paragraph.
            HashMismatchError: If the paragraph hash is stale.

        Example:
            new_ref = doc.insert_after("Section 5", " (as amended)", paragraph="P2#f3c1")
            doc.insert_after(doc.find_text("Section 5"), " (as amended)")
        """
        self._ensure_open()
        anchor, paragraph, occurrence = _resolve_search_target(
            anchor, paragraph, occurrence, ctx="insert_after(): ", field="'anchor'"
        )
        paragraph = _require_ref_string(paragraph, "'anchor'")
        _validate_note(note, ctx="insert_after(): ")
        change_id = self._revision_manager.insert_text_after(anchor, text, occurrence=occurrence, paragraph=paragraph)
        group_id = self._revision_manager.group_id_of(change_id)
        comment_id = self._attach_notes([(group_id, note, _NOTE_AMENDED_INSERTION)])[0]
        return self._edit_result(paragraph, group_id, comment_id=comment_id)

    def insert_before(
        self,
        anchor: str | SearchResult,
        text: str,
        *,
        paragraph: str | None = None,
        occurrence: int | None = None,
        note: str | None = None,
    ) -> EditResult:
        """Insert text before anchor with tracked changes.

        Args:
            anchor: Text to find as insertion point, or a
                :class:`~docx_editor.track_changes.SearchResult` from
                :meth:`find_text`/:meth:`find_all` — which supplies the anchor
                text, its paragraph and its occurrence, so pass neither
                ``paragraph`` nor ``occurrence`` with one.
            text: Text to insert before the anchor
            paragraph: Paragraph reference from list_paragraphs() (e.g.,
                "P2#f3c1"). Required unless ``anchor`` is a SearchResult.
            occurrence: Which occurrence of anchor within the paragraph
                (0 = first). Omitted → ``anchor`` must be unique in the
                paragraph, else AmbiguousTextError.
            note: Rationale for this edit, anchored as a comment bracketing
                the insertion it creates, and deleted with it when the
                insertion is resolved (see :meth:`replace`).

        Returns:
            EditResult — the new paragraph reference with updated hash (e.g.,
            "P2#c3d4"), carrying ``group_id``/``revision_ids`` of the
            revisions this edit created, and ``comment_id`` when a ``note``
            was anchored.

        Warns:
            UnanchoredNoteWarning: If ``note`` was given but the call created
                no revision a comment can bracket — an amendment to your own
                pending insertion, or a bare ``"\\n"`` whose only revision is
                the paragraph mark. The edit still applies; the note is
                dropped and ``comment_id`` is None.

        Raises:
            ValueError: If ``anchor`` is not a non-empty string, ``text`` is
                not a string, ``paragraph`` is missing or not a ref string,
                ``occurrence`` is negative or not an integer, ``note`` is
                neither None nor a non-empty control-character-free string, or
                ``anchor`` is a SearchResult and ``paragraph``/``occurrence``
                was given too.
            TextNotFoundError: If ``anchor`` is absent or ``occurrence`` is
                out of range for the paragraph.
            AmbiguousTextError: If ``occurrence`` is omitted and ``anchor``
                matches more than once in the paragraph.
            HashMismatchError: If the paragraph hash is stale.

        Example:
            new_ref = doc.insert_before("Section 6", "New clause: ", paragraph="P2#f3c1")
            doc.insert_before(doc.find_text("Section 6"), "New clause: ")
        """
        self._ensure_open()
        anchor, paragraph, occurrence = _resolve_search_target(
            anchor, paragraph, occurrence, ctx="insert_before(): ", field="'anchor'"
        )
        paragraph = _require_ref_string(paragraph, "'anchor'")
        _validate_note(note, ctx="insert_before(): ")
        change_id = self._revision_manager.insert_text_before(anchor, text, occurrence=occurrence, paragraph=paragraph)
        group_id = self._revision_manager.group_id_of(change_id)
        comment_id = self._attach_notes([(group_id, note, _NOTE_AMENDED_INSERTION)])[0]
        return self._edit_result(paragraph, group_id, comment_id=comment_id)

    def split_paragraph(self, ref: str, *, before: str, occurrence: int | None = None) -> EditResult:
        """Split a paragraph into two with a tracked paragraph break.

        Explicit sugar for the ``\\n``-means-split behavior: the paragraph is
        cut immediately before ``before``, its paragraph mark flagged as an
        inserted revision and the tail (from ``before`` on) moved into a new
        following paragraph. Accepting keeps the split; rejecting the group
        rejoins the two paragraphs. Equivalent to
        ``insert_before(before, "\\n", ...)``.

        Args:
            ref: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1").
            before: Text to split before (keyword-only); the break lands at its
                start. Must be a non-empty string present in the paragraph.
            occurrence: Which occurrence of ``before`` within the paragraph
                (0 = first). Omitted → ``before`` must be unique in the
                paragraph, else AmbiguousTextError.

        Returns:
            EditResult — the first paragraph's ref (with updated hash); its
            ``refs`` tuple carries the refs of both resulting paragraphs, and
            ``group_id``/``revision_ids`` cover the whole split.

        Raises:
            ValueError: If ``before`` is not a non-empty string, ``ref`` is not
                a ref string, or ``occurrence`` is negative or not an integer.
            TextNotFoundError: If ``before`` is absent or ``occurrence`` is out
                of range for the paragraph.
            AmbiguousTextError: If ``occurrence`` is omitted and ``before``
                matches more than once in the paragraph.
            HashMismatchError: If the paragraph hash is stale.

        Example:
            result = doc.split_paragraph("P2#f3c1", before="However,")
            r1, r2 = result.refs
        """
        self._ensure_open()
        _require_ref_string(ref)
        change_id = self._revision_manager.insert_text_before(before, "\n", occurrence=occurrence, paragraph=ref)
        # Takes no note= of its own, but every edit path reaps (see _attach_notes).
        self._reap_note_comments()
        return self._edit_result(ref, self._revision_manager.group_id_of(change_id))

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
            accept_group()/reject_group(). Operations carrying ``note=`` also
            carry the resulting ``comment_id``; ops of this call that share the
            same note text share one comment, deleted when the last operation
            it explains is resolved (see :meth:`replace`).
            When dry_run is True: list of EditValidationResult, one per
            operation. A row that failed on a stale hash carries
            ``current_ref``, the ref for that paragraph's current content —
            rebuild the operation with it instead of parsing ``error``.

        Warns:
            UnanchoredNoteWarning: Once per operation whose ``note`` had no
                revision to anchor to. The batch still applies in full.

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

            # Repair the stale-hash rows and retry, no message parsing:
            for row in results:
                if row.current_ref:
                    ops[row.index] = EditOperation.replace(
                        "old", "new", paragraph=row.current_ref
                    )
        """
        self._ensure_open()
        if not isinstance(operations, list):
            raise ValueError(f"batch_edit(): 'operations' must be a list of EditOperation, got {operations!r}")
        if dry_run:
            return self._revision_manager.validate_batch(operations)
        change_ids = self._revision_manager.batch_edit(operations)
        if not change_ids:
            return []
        # One shared <w:p> walk for all result refs. A \n split shifts every
        # later paragraph's index, so if ANY op split, resolve EVERY op's ref by
        # element identity (one shared revision-element index) — otherwise a
        # shifted non-split op would report a stale index. A no-split batch skips
        # the index entirely, so its DOM-walk count stays constant (ISSUES.md #51).
        mgr = self._revision_manager
        group_ids = [mgr.group_id_of(change_id) for change_id in change_ids]
        # One _attach_notes call for the whole batch: that is what lets ops
        # sharing a note share one comment, and costs one DOM walk rather than
        # one per op. Comment markers add no <w:p>, so the walks below stay
        # valid whichever side of them this runs on.
        comment_ids = self._attach_notes([
            (gid, op.note, _batch_note_reason(op)) for op, gid in zip(operations, group_ids, strict=True)
        ])
        any_split = any(gid is not None and mgr.split_count(gid) for gid in group_ids)
        paragraphs = body_paragraphs(self._document_editor.dom)
        element_index = mgr._revision_element_index() if any_split else None
        return [
            self._edit_result(op.paragraph, gid, paragraphs, element_index, comment_id=cid)
            for op, gid, cid in zip(operations, group_ids, comment_ids, strict=True)
        ]

    def rewrite_paragraph(self, ref: str, new_text: str, *, note: str | None = None) -> EditResult:
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
            note: Rationale for the rewrite, anchored as one comment
                spanning its first through last revision — one comment for the
                whole rewrite, not one per diff hunk — and deleted with it when
                the rewrite is resolved (see :meth:`replace`).

        Returns:
            EditResult — the new paragraph reference with updated hash (e.g.,
            "P2#c3d4"), carrying ``group_id``/``revision_ids`` of all the
            revisions the rewrite created (``group_id`` is None when
            new_text equals the current text, or when every change landed
            inside your own pending insertions and amended them in place —
            undo those by rejecting the group of the amended insertion), and
            ``comment_id`` when a ``note`` was anchored.

        Warns:
            UnanchoredNoteWarning: If ``note`` was given but the rewrite
                created no revisions. The rewrite still applies; the note is
                dropped and ``comment_id`` is None.

        Raises:
            ValueError: If ``new_text`` is not a string (empty string is
                allowed — it deletes all text of a tab-free paragraph) or does
                not hold the same number of tab marks (``\\t``) as the
                paragraph (ISSUES.md #6),
                ``note`` is neither None nor a
                non-empty control-character-free string, or ``ref`` is
                malformed.
            ParagraphIndexError: If ``ref``'s index is out of range.
            HashMismatchError: If ``ref``'s hash is stale.

        Example:
            result = doc.rewrite_paragraph("P2#f3c1", "The board shall approve the proposal.")
            doc.reject_group(result.group_id)  # undo the whole rewrite
        """
        self._ensure_open()
        _validate_note(note, ctx="rewrite_paragraph(): ")
        group_id = self._revision_manager.rewrite_paragraph(ref, new_text)
        comment_id = self._attach_notes([(group_id, note, _NOTE_REWRITE_NO_REVISION)])[0]
        return self._edit_result(ref, group_id, comment_id=comment_id)

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
        # No note= of its own, but a rewrite here can amend an earlier note's
        # insertion out of existence just as batch_edit can (see _attach_notes).
        self._reap_note_comments()
        # Shared <w:p> walk; if any rewrite split, resolve every ref by element
        # identity so a shifted later rewrite never reports a stale index.
        mgr = self._revision_manager
        any_split = any(gid is not None and mgr.split_count(gid) for gid in group_ids)
        paragraphs = body_paragraphs(self._document_editor.dom)
        element_index = mgr._revision_element_index() if any_split else None
        return [
            self._edit_result(ref, group_id, paragraphs, element_index)
            for (ref, _), group_id in zip(rewrites, group_ids, strict=True)
        ]

    # ==================== Comments API ====================

    def add_comment(
        self,
        anchor_text: str | SearchResult,
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
            anchor_text: Text to attach the comment to, or a
                :class:`~docx_editor.track_changes.SearchResult` from
                :meth:`find_text`/:meth:`find_all` — which supplies the anchor
                text, its paragraph and its occurrence, so pass neither
                ``paragraph`` nor ``occurrence`` with one.
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
            ValueError: If ``occurrence`` is negative or not an integer, or
                ``anchor_text`` is a SearchResult and
                ``paragraph``/``occurrence`` was given too.

        Example:
            doc.add_comment("Section 5", "Please review this section")
            doc.add_comment("foo", "Note", paragraph="P3#a7b2", occurrence=1)

            # Comment exactly the match a search found (2nd "foo", say):
            doc.add_comment(doc.find_all("foo")[1], "Note")
        """
        self._ensure_open()
        anchor_text, paragraph, occurrence = _resolve_search_target(
            anchor_text, paragraph, occurrence, ctx="add_comment(): ", field="'anchor_text'"
        )
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
        """Delete a comment, and every reply threaded under it, from the document.

        Args:
            comment_id: ID of the comment to delete

        Returns:
            True if deleted, False if not found

        Raises:
            ValueError: If ``comment_id`` is not an integer (bool included).

        Note:
            Deleting a ``note=`` rationale by its ``EditResult.comment_id`` is
            allowed and final: the note stops tracking the revisions it
            explained, so resolving them later neither resurrects nor re-anchors
            it.

        Example:
            doc.delete_comment(0)
        """
        self._ensure_open()
        deleted = self._comment_manager.delete_comment(comment_id)
        if deleted:
            self._forget_note_comment(comment_id, list(self._note_groups.get(comment_id, ())))
        return deleted

    # ==================== Revision Management API ====================

    def list_revisions(self, author: str | None = None, paragraph: str | None = None) -> list[Revision]:
        """List the document's tracked insertions and deletions.

        Insertions and deletions only. Every other revision type in the OOXML
        schema — format changes, moves, table-structure revisions — is listed
        by ``list_unhandled_revisions()`` instead, because none of it can be
        passed to ``accept_revision()``.

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
            locatable, e.g. nested revisions, or when ``paragraph_ref`` is
            itself None — a revision inside a drawing's text box lists and
            accepts by id, but has no addressable location), plus
            ``nested_under`` and ``contains_ids`` describing revision
            nesting (e.g. a foreign deletion inside another author's
            pending insertion), and
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

        Note:
            Any ``note=`` rationale left with no live revision to explain is
            deleted with it, replies included (see :meth:`replace`).

        Args:
            revision_id: ID of the revision to accept

        Returns:
            True if accepted, False if not found

        Example:
            doc.accept_revision(1)
        """
        self._ensure_open()
        resolved = self._revision_manager.accept_revision(revision_id)
        self._reap_note_comments()
        return resolved

    def reject_revision(self, revision_id: int) -> bool:
        """Reject a revision by ID.

        For insertions: removes the inserted content.
        For deletions: restores the deleted content.

        Note:
            Any ``note=`` rationale left with no live revision to explain is
            deleted with it, replies included (see :meth:`replace`).

        Args:
            revision_id: ID of the revision to reject

        Returns:
            True if rejected, False if not found

        Example:
            doc.reject_revision(1)
        """
        self._ensure_open()
        resolved = self._revision_manager.reject_revision(revision_id)
        self._reap_note_comments()
        return resolved

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

        Note:
            Any ``note=`` rationale left with no live revision to explain is
            deleted with it, replies included (see :meth:`replace`).

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
        count = self._revision_manager.accept_group(group_id)
        self._reap_note_comments()
        return count

    def reject_group(self, group_id: int) -> int:
        """Reject every revision created by one logical edit operation.

        The counterpart of :meth:`accept_group` — rejecting the group undoes
        the whole edit, restoring the exact pre-edit text (deletions are
        restored, insertions removed). Same group semantics and lifetime as
        accept_group(), including inferred groups after reopen.

        Note:
            Any ``note=`` rationale left with no live revision to explain is
            deleted with it, replies included (see :meth:`replace`).

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
        count = self._revision_manager.reject_group(group_id)
        self._reap_note_comments()
        return count

    def accept_changeset(self, changeset_id: int) -> int:
        """Accept every revision created by one whole call (a changeset).

        A changeset is the intent tier — one level above a group:

            one call (a single edit, or an entire batch_edit/batch_rewrite)
            = one changeset ⊇ one-or-more groups ⊇ revisions

        Accepting the changeset applies every group the call created. Its
        ``changeset_id`` is carried by each EditResult the call returned (and
        by every Revision's ``changeset_id``). There is no tier above this —
        the model stops at three (revision < group < changeset).

        A changeset is the ``(author, date)`` equivalence class over groups:
        for edits made here it bundles the whole call; for revisions already
        in the file it partitions the reconstructed groups by identical
        author + date (``changeset_source == "inferred"``). A ``batch_edit``
        whose ops land in different paragraphs is therefore one changeset even
        though its groups are non-contiguous.

        Changeset ids are in-memory and per-open-Document, renumbered on each
        open, exactly like group ids — always use one from this session's
        EditResult or list_revisions(). Rump-tolerant: after Word has already
        resolved part of the changeset, accepting resolves whatever survives.

        Note:
            Any ``note=`` rationale left with no live revision to explain is
            deleted with it, replies included (see :meth:`replace`).

        Args:
            changeset_id: Changeset id from an EditResult (or a Revision's
                ``changeset_id``)

        Returns:
            Number of revisions accepted across the changeset's groups.
            Members already resolved individually are skipped (and not
            counted).

        Raises:
            RevisionError: If the changeset id is unknown to this open Document.

        Example:
            results = doc.batch_edit([...])
            doc.accept_changeset(results[0].changeset_id)  # accept the whole batch
        """
        self._ensure_open()
        count = self._revision_manager.accept_changeset(changeset_id)
        self._reap_note_comments()
        return count

    def reject_changeset(self, changeset_id: int) -> int:
        """Reject every revision created by one whole call (a changeset).

        The counterpart of :meth:`accept_changeset` — rejecting the changeset
        undoes the entire call (every group), restoring the exact pre-call
        text. Same changeset semantics and lifetime as accept_changeset(),
        including inferred changesets after reopen. No tier exists above this.

        Note:
            Any ``note=`` rationale left with no live revision to explain is
            deleted with it, replies included (see :meth:`replace`).

        Args:
            changeset_id: Changeset id from an EditResult (or a Revision's
                ``changeset_id``)

        Returns:
            Number of revisions rejected across the changeset's groups.
            Members already resolved individually are skipped (and not
            counted).

        Raises:
            RevisionError: If the changeset id is unknown to this open Document.

        Example:
            results = doc.batch_edit([...])
            doc.reject_changeset(results[0].changeset_id)  # undo the whole batch
        """
        self._ensure_open()
        count = self._revision_manager.reject_changeset(changeset_id)
        self._reap_note_comments()
        return count

    def list_unhandled_revisions(self, author: str | None = None) -> list[UnhandledRevision]:
        """List the revision types this library does not accept or reject.

        The complement of ``list_revisions()``: everything in the OOXML
        revision schema except insertions and deletions — property changes
        (``w:pPrChange``, ``w:rPrChange``, ``w:sectPrChange``, the table
        ``*PrChange`` family), content moves (``w:moveFrom``/``w:moveTo`` and
        their range marks), table-structure revisions (``w:cellIns``,
        ``w:cellDel``, ``w:cellMerge``), ``w:numberingChange`` and the
        custom-XML range marks. These survive open/edit/save unchanged and are
        left pending by ``accept_all()``/``reject_all()``.

        Call this before telling a human "all changes accepted": on a
        format-only redline ``accept_all()`` returns 0 because there was
        nothing it *could* accept, not because there was nothing to do.

        These rows are deliberately not ``Revision`` objects — they carry
        nothing ``accept_revision()`` could act on, so they are not
        interchangeable with ``list_revisions()`` output.

        Only ``word/document.xml`` is inspected; headers, footers and
        footnotes are the container-parts epic (ISSUES.md #30).

        Args:
            author: If provided, filter by author name. Marks with no
                ``w:author`` attribute read as ``"Unknown"``, so they match
                only ``author="Unknown"``, are excluded from every other
                filtered call, and are included in an unfiltered one.

        Returns:
            List of UnhandledRevision in document order, each with ``tag``,
            ``id`` (None when the mark carries no numeric ``w:id``),
            ``author``, ``date`` and ``paragraph_ref`` (None outside any
            paragraph).

        Example:
            result = doc.accept_all()
            if result.unhandled:
                for row in doc.list_unhandled_revisions():
                    print(f"still pending: {row.tag} by {row.author}")
        """
        self._ensure_open()
        return self._revision_manager.list_unhandled_revisions(author=author)

    def accept_all(self, author: str | None = None) -> ResolveResult:
        """Accept all insertions and deletions.

        Resolves ``w:ins``/``w:del`` only. Every other revision type is left
        pending and reported on the result rather than silently ignored — see
        ``list_unhandled_revisions()`` for the full list and
        :class:`ResolveResult` for the counting rule.

        Only ``word/document.xml`` is inspected; headers, footers and
        footnotes are the container-parts epic (ISSUES.md #30).

        Note:
            Any ``note=`` rationale left with no live revision to explain is
            deleted with it, replies included, so a pipeline that calls
            ``accept_all()`` to produce a clean deliverable does not ship agent
            rationale as live comments. A sweep filtered to another author
            leaves the notes whose revisions survive it alone — but one of ours
            nested inside a rejected foreign insertion goes with its host, and
            so does its note.

        Args:
            author: If provided, only accept revisions by this author

        Returns:
            :class:`ResolveResult` — an ``int`` whose value is the number of
            revisions accepted (so existing ``count = doc.accept_all()`` code
            is unaffected), carrying ``.unhandled`` (how many revision
            elements this library never resolves are still in the document,
            counted after resolution) and ``.unhandled_types`` (tag -> count).

        Warns:
            UnhandledRevisionWarning: If ``.unhandled`` is nonzero — the point
                at which "everything is accepted" would be a false claim.

        Example:
            result = doc.accept_all()
            print(f"Accepted {result} revisions")
            if result.unhandled:
                print(f"Still pending: {result.unhandled_types}")
        """
        self._ensure_open()
        result = self._revision_manager.accept_all(author=author)
        self._reap_note_comments()
        return result

    def reject_all(self, author: str | None = None) -> ResolveResult:
        """Reject all insertions and deletions.

        Resolves ``w:ins``/``w:del`` only; every other revision type is left
        pending and reported exactly as in ``accept_all()``.

        Note:
            Any ``note=`` rationale left with no live revision to explain is
            deleted with it, replies included (see :meth:`accept_all`).

        Args:
            author: If provided, only reject revisions by this author

        Returns:
            :class:`ResolveResult` — an ``int`` whose value is the number of
            revisions rejected, carrying ``.unhandled`` and
            ``.unhandled_types``.

        Warns:
            UnhandledRevisionWarning: If ``.unhandled`` is nonzero.

        Example:
            result = doc.reject_all(author="OtherUser")
            if result.unhandled:
                print(f"Still pending: {result.unhandled_types}")
        """
        self._ensure_open()
        result = self._revision_manager.reject_all(author=author)
        self._reap_note_comments()
        return result

    # ==================== Save/Close API ====================

    def save(
        self,
        path: str | Path | None = None,
        validate: bool = False,
        force: bool = False,
        *,
        track_changes: bool | None = None,
    ) -> Path:
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
            track_changes: Whether to turn Word's track-changes switch
                (``<w:trackRevisions/>`` in settings.xml) on in the saved file.
                None (the default) writes it exactly when this document carries
                a revision authored by us — so the human who keeps typing in
                Word after our redline stays tracked, and the return leg is
                still adjudicatable. The test is the document's state, not
                whether this session edited: a pending redline of ours reopened
                from an earlier session still counts, because it is still
                waiting for a reply. A document holding no revision of ours —
                one we did not redline, or whose revisions we accepted — is
                saved untouched. True writes
                the flag whether or not we redlined anything; False leaves
                settings.xml alone, and never removes a flag the document
                already had.

                A document that turns tracking *off* explicitly
                (``<w:trackRevisions w:val="false"/>``) is respected, not
                overridden: under the default the element is left as it is and
                the save warns that the recipient's edits will not be tracked.
                ``track_changes=True`` overrides it. Ownership is by ``w:author``,
                so a foreign redline by an author with our name reads as ours.

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
        self._ensure_track_changes_flag(track_changes)

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

    @staticmethod
    def _check_protection(workspace: Workspace, allow_protected: bool) -> None:
        """Refuse a document whose editing protection locks its body text.

        Reads ``w:documentProtection`` from settings.xml (read-only — no editor,
        so nothing in the workspace is touched by a document we are about to
        refuse). Only an *enforced* protection counts: Word writes the element
        with ``w:enforcement="0"`` when the mode is configured but switched off,
        and that document is editable in Word, so it is editable here. An
        enforcement value outside ST_OnOff fails *closed* — a guard over locked
        content cannot read "we could not parse the switch" as "the switch is
        off" — and ``allow_protected=True`` is still the way past it.

        Only ``w:documentProtection`` is read. ``w:writeProtection`` (Word's
        "Password to modify" and "Always Open Read-Only") is a different
        element and deliberately out of scope: it restricts saving over the
        original rather than editing the body, which ``save()`` already
        surfaces when it refuses to replace a read-only file.

        Takes the workspace explicitly because it runs before ``__init__``
        assigns ``self._workspace``.
        """
        settings_path = workspace.word_path / "settings.xml"
        if not settings_path.exists():
            return

        dom = defusedxml.minidom.parse(str(settings_path))
        root = dom.documentElement
        protection = None
        for child in root.childNodes:
            if child.nodeType == child.ELEMENT_NODE and _local_name(child.tagName) == "documentProtection":
                protection = child
                break
        if protection is None:
            return

        enforcement = _attr_node(protection, "enforcement")
        # No attribute at all is the schema default, which is off: a protection
        # Word never switched on is not protection. A value we cannot read is a
        # different matter — see the docstring, it fails closed.
        if enforcement is None or _on_off(enforcement.value) is False:
            return

        edit = _attr_node(protection, "edit")
        mode = edit.value if edit is not None else None
        # "trackedChanges" enforcement asks every editor to leave a redline —
        # which is all this library ever does — so it must never raise. Neither
        # does "none", nor a mode from a schema we do not know: the guard exists
        # to protect content that was locked, not to police unknown values.
        if mode not in {"readOnly", "forms", "comments"}:
            return

        if allow_protected:
            return

        what = {
            "readOnly": "is protected against all editing (Restrict Editing: 'No changes (Read only)')",
            "forms": "only allows typing in form fields (Restrict Editing: 'Filling in forms')",
            "comments": "only allows commenting (Restrict Editing: 'Comments')",
        }[mode]
        extra = (
            " Comments are permitted by this mode, so add_comment() works once it is open."
            if mode == "comments"
            else ""
        )
        raise DocumentProtectedError(
            f"{workspace.source_path} {what}, and the protection is enforced. "
            f"Turn protection off in Word (Review > Restrict Editing > Stop Protection), "
            f"or open it anyway with "
            f"Document.open({str(workspace.source_path)!r}, allow_protected=True)."
            f"{extra}",
            path=workspace.source_path,
            mode=mode,
        )

    def _ensure_track_changes_flag(self, track_changes: bool | None) -> None:
        """Turn Word's track-changes switch on in settings.xml at save time.

        A ``<w:trackRevisions/>`` in settings.xml is what makes Word keep tracking
        after our redline lands: without it the recipient's own edits are
        untracked and the two rounds can no longer be told apart.

        Args:
            track_changes: True to write the flag unconditionally, False to
                leave settings.xml alone, None (the default from ``save()``) to
                write it only when this document carries a revision we authored.
        """
        if track_changes is False:
            return

        wanted = track_changes if track_changes is not None else self._revision_manager.has_own_revisions()
        if not wanted:
            return

        settings_path = self._workspace.word_path / "settings.xml"
        if not settings_path.exists():
            # Same tolerance as _update_settings: a document without the part
            # keeps saving, it just gets no flag. Under the default that is the
            # whole story, but a caller who asked for the flag in as many words
            # is owed the news that it did not happen.
            if track_changes is True:
                warnings.warn(
                    f"{self._workspace.source_path} has no word/settings.xml, so track changes "
                    f"could not be turned on. The revisions saved here stay visible, but edits "
                    f"the recipient makes in Word will not be tracked.",
                    UserWarning,
                    stacklevel=3,
                )
            return

        editor = DocxXMLEditor(
            settings_path,
            rsid=self._workspace.rsid,
            author=self._workspace.author,
            on_save=self._workspace.mark_dirty,
        )
        root = editor.dom.documentElement
        prefix = root.tagName.split(":")[0] if ":" in root.tagName else "w"

        children = [c for c in root.childNodes if c.nodeType == c.ELEMENT_NODE]
        existing = next((c for c in children if _local_name(c.tagName) == "trackRevisions"), None)

        if existing is not None:
            # No w:val means on: a bare <w:trackRevisions/> is how Word writes
            # "tracking is on". A w:val we cannot read is not on — reporting it
            # as on would make an explicit track_changes=True a silent no-op.
            val = _attr_node(existing, "val")
            if val is None or _on_off(val.value) is True:
                # Already on — nothing to write, and nothing to reserialize.
                return
            if track_changes is None:
                warnings.warn(
                    f"{self._workspace.source_path} does not have track changes on "
                    f'(<w:trackRevisions w:val="{val.value}"/> in settings.xml), so it is left as it '
                    f"is: the revisions saved here stay visible, but edits the recipient makes in "
                    f"Word will not be tracked. Save with track_changes=True to turn tracking on "
                    f"instead.",
                    UserWarning,
                    stacklevel=3,
                )
                return
            # An explicit track_changes=True is the caller saying it in as many
            # words, which outranks whatever the document's w:val held. Dropping
            # the attribute leaves Word's own canonical <w:trackRevisions/>.
            existing.removeAttribute(val.name)
            editor.save()
            return

        flag_xml = f"<{prefix}:trackRevisions/>"
        # CT_Settings is a sequence: land after the last element the schema puts
        # before trackRevisions, never after an element we do not recognize.
        anchor = None
        for child in children:
            if _local_name(child.tagName) in _SETTINGS_BEFORE_TRACK_REVISIONS:
                anchor = child
        if anchor is not None:
            editor.insert_after(anchor, flag_xml)
        elif children:
            # Nothing the schema puts before trackRevisions is here, so every
            # sibling belongs after it: go first, never last.
            editor.insert_before(children[0], flag_xml)
        else:  # pragma: no cover - _update_settings guarantees a w:rsids sibling by save time
            editor.append_to(root, flag_xml)
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
