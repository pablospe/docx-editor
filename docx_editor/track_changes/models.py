"""Data model for track changes: result and revision dataclasses, tag
constants, validators, and the revision-element walk."""

from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from datetime import datetime
from typing import Literal
from xml.dom.minidom import Element

from ..xml_editor import ParagraphRef, TextMapMatch, _reject_control_chars, _require_valid_occurrence

# Provenance of a revision group: created by an edit through this open
# Document ("recorded") vs reconstructed at parse time ("inferred").
GroupSource = Literal["recorded", "inferred"]


def _validate_edit_target(value: str, *, field: str, ctx: str) -> None:
    """Validate a replace/delete target: the control-character rule for search
    text, then refuse a target that would consume a ``<w:tab/>``.

    The generic scan lets ``\\t`` through (``allow_tab=True``, as for any
    search text) so the tab gets its own message here: a tab is searchable
    (it is a ``"\\t"`` in the text map) and an insertion may land on either
    side of it, but deleting or replacing one would have to remove the
    ``<w:tab/>`` element under tracking, which no edit path supports yet
    (ISSUES.md #6). Exact by construction: ``replace_with`` can never hold a
    tab, so affix trimming cannot narrow a tab-free target onto a tab.
    """
    _reject_control_chars(value, field=field, ctx=ctx, allow_tab=True)
    if isinstance(value, str) and "\t" in value:
        raise ValueError(
            f"{ctx}{field} contains a tab ('\\t') — tabs can be searched and matched but not deleted "
            f"or replaced yet; target the text on either side of the tab instead (ISSUES.md #6)."
        )


def _validate_note(note: str | None, *, ctx: str) -> None:
    """Reject a ``note=`` rationale that could not become a comment body.

    ``None`` means "no note". Anything else must be a non-empty string free of
    control characters: the note is written into a single ``<w:t>`` of a comment
    part, exactly like ``add_comment``'s text, so a newline or a tab would land
    there as an invisible, unreviewable artifact.

    Called *before* the edit runs (and at ``EditOperation`` construction), so a
    bad note never leaves an applied edit with a dropped rationale behind it.

    Raises:
        ValueError: If ``note`` is neither None nor a valid note string.
    """
    if note is None:
        return
    if not isinstance(note, str) or not note:
        raise ValueError(
            f"{ctx}'note' must be a non-empty string or None — the rationale to attach as a comment, got {note!r}"
        )
    # The shared newline message explains a *search* failure ("it can never
    # match"), which says nothing to someone writing a comment body — so say
    # why a newline is wrong here, and leave the rest to the shared rule.
    if "\n" in note:
        raise ValueError(
            f"{ctx}'note' must not contain a newline ('\\n') — a comment body is a single "
            f"<w:t>, where a literal newline is an invisible, unreviewable artifact."
        )
    _reject_control_chars(note, field="'note'", ctx=ctx, allow_newline=True)


@dataclass
class _RegistrySnapshot:
    """Copy of the group + changeset registry, for rollback with a DOM snapshot."""

    counter: int
    groups: dict[int, tuple[int, ...]]
    revision_groups: dict[int, int | None]
    group_sources: dict[int, GroupSource]
    changeset_counter: int
    changesets: dict[int, tuple[int, ...]]
    group_changesets: dict[int, int]
    changeset_sources: dict[int, GroupSource]


@dataclass
class EditOperation:
    """A single edit operation for batch processing.

    Prefer the typed constructors (:meth:`replace`, :meth:`delete`,
    :meth:`insert_after`, :meth:`insert_before`) — they validate arguments at
    construction time with the same rules ``batch_edit`` applies, so mistakes
    surface immediately instead of at apply time. The raw
    ``EditOperation(action=..., ...)`` form remains supported.

    Each typed constructor also accepts a
    :class:`~docx_editor.track_changes.SearchResult` in place of its target
    text, which fills in ``paragraph`` and ``occurrence`` from the match —
    ``[EditOperation.replace(m, "60 days") for m in doc.find_all("30 days")]``
    needs no ref or occurrence bookkeeping. The dataclass fields stay plain
    strings either way.
    """

    action: Literal["replace", "delete", "insert_after", "insert_before"]
    paragraph: str  # Required: hash-anchored ref like "P3#a7b2"
    find: str | None = None  # For replace
    replace_with: str | None = None  # For replace
    text: str | None = None  # For delete (text to delete) or insert (text to insert)
    anchor: str | None = None  # For insert_after/insert_before
    occurrence: int | None = None  # None = target must be unique in the paragraph
    note: str | None = None  # Rationale to anchor as a comment on this op's revisions

    @staticmethod
    def _validate_common(constructor: str, paragraph: str | None, occurrence: int | None, field: str) -> str:
        """Construction-time checks shared by all typed constructors.

        Returns the validated ref, so callers can hand a narrowed ``str`` to the
        dataclass field (``paragraph`` is Optional in the signatures only to
        allow the SearchResult form).
        """
        if paragraph is None:
            # Mirrors ParagraphRef.parse's non-string wording (which handles every
            # other bad ref below) so the None case reads identically, plus the
            # hint that a SearchResult would have supplied the ref.
            raise ValueError(
                f"Invalid paragraph reference None: expected a string like 'P3#a7b2', got NoneType — "
                f"{_paragraph_hint(field)}"
            )
        ParagraphRef.parse(paragraph)
        _require_valid_occurrence(occurrence, f"EditOperation.{constructor}(): ")
        return paragraph

    @classmethod
    def replace(
        cls,
        find: "str | SearchResult",
        replace_with: str,
        *,
        paragraph: str | None = None,
        occurrence: int | None = None,
        note: str | None = None,
    ) -> "EditOperation":
        """Build a validated replace operation (mirrors ``Document.replace``).

        Args:
            find: Text to find and replace (must be non-empty), or a
                :class:`SearchResult` from find_text()/find_all() — which also
                supplies ``paragraph`` and ``occurrence``, so pass neither.
            replace_with: Replacement text (empty string allowed — replacing
                with nothing is a valid tracked deletion)
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1").
                Required unless ``find`` is a SearchResult.
            occurrence: Which occurrence within the paragraph (0 = first).
                Omitted → ``find`` must be unique in the paragraph, else the
                batch fails with a wrapped AmbiguousTextError at apply time.
            note: Rationale for this edit, anchored as a comment on the
                revisions it creates (see ``Document.replace``).

        Raises:
            ValueError: If ``paragraph`` is missing or malformed, ``occurrence``
                is not a non-negative integer, ``find`` is not a non-empty
                string or contains a tab (``\\t`` — a tab mark can be matched
                but not replaced yet, ISSUES.md #6), ``replace_with`` is not a string,
                ``note`` is neither None nor a non-empty control-character-free
                string, or ``find`` is a SearchResult and
                ``paragraph``/``occurrence`` was given too.
        """
        find, paragraph, occurrence = _resolve_search_target(
            find, paragraph, occurrence, ctx="EditOperation.replace(): ", field="'find'"
        )
        paragraph = cls._validate_common("replace", paragraph, occurrence, "'find'")
        if not isinstance(find, str) or not find:
            raise ValueError(
                f"EditOperation.replace(): 'find' must be a non-empty string — the text to search for, got {find!r}"
            )
        if not isinstance(replace_with, str):
            raise ValueError(
                f"EditOperation.replace(): 'replace_with' must be a string (empty string is allowed), "
                f"got {replace_with!r}"
            )
        _validate_edit_target(find, field="'find'", ctx="EditOperation.replace(): ")
        _reject_control_chars(replace_with, field="'replace_with'", ctx="EditOperation.replace(): ", allow_newline=True)
        _validate_note(note, ctx="EditOperation.replace(): ")
        return cls(
            action="replace",
            paragraph=paragraph,
            find=find,
            replace_with=replace_with,
            occurrence=occurrence,
            note=note,
        )

    @classmethod
    def delete(
        cls,
        text: "str | SearchResult",
        *,
        paragraph: str | None = None,
        occurrence: int | None = None,
        note: str | None = None,
    ) -> "EditOperation":
        """Build a validated delete operation (mirrors ``Document.delete``).

        Args:
            text: Text to mark as deleted (must be non-empty), or a
                :class:`SearchResult` from find_text()/find_all() — which also
                supplies ``paragraph`` and ``occurrence``, so pass neither.
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1").
                Required unless ``text`` is a SearchResult.
            occurrence: Which occurrence within the paragraph (0 = first).
                Omitted → ``text`` must be unique in the paragraph, else the
                batch fails with a wrapped AmbiguousTextError at apply time.
            note: Rationale for this edit, anchored as a comment on the
                revisions it creates (see ``Document.delete``).

        Raises:
            ValueError: If ``paragraph`` is missing or malformed, ``occurrence``
                is not a non-negative integer, ``text`` is not a non-empty
                string or contains a tab (``\\t`` — a tab mark can be matched
                but not deleted yet, ISSUES.md #6), ``note`` is neither None nor a non-empty
                control-character-free string, or ``text`` is a SearchResult and
                ``paragraph``/``occurrence`` was given too.
        """
        text, paragraph, occurrence = _resolve_search_target(
            text, paragraph, occurrence, ctx="EditOperation.delete(): ", field="'text'"
        )
        paragraph = cls._validate_common("delete", paragraph, occurrence, "'text'")
        if not isinstance(text, str) or not text:
            raise ValueError(
                f"EditOperation.delete(): 'text' must be a non-empty string — the text to mark as deleted, got {text!r}"
            )
        _validate_edit_target(text, field="'text'", ctx="EditOperation.delete(): ")
        _validate_note(note, ctx="EditOperation.delete(): ")
        return cls(action="delete", paragraph=paragraph, text=text, occurrence=occurrence, note=note)

    @classmethod
    def _insert(
        cls,
        action: Literal["insert_after", "insert_before"],
        anchor: "str | SearchResult",
        text: str,
        paragraph: str | None,
        occurrence: int | None,
        note: str | None,
    ) -> "EditOperation":
        anchor, paragraph, occurrence = _resolve_search_target(
            anchor, paragraph, occurrence, ctx=f"EditOperation.{action}(): ", field="'anchor'"
        )
        paragraph = cls._validate_common(action, paragraph, occurrence, "'anchor'")
        if not isinstance(anchor, str) or not anchor:
            raise ValueError(
                f"EditOperation.{action}(): 'anchor' must be a non-empty string — the text to insert near, "
                f"got {anchor!r}"
            )
        if not isinstance(text, str):
            raise ValueError(
                f"EditOperation.{action}(): 'text' must be a string (empty string is allowed), got {text!r}"
            )
        _reject_control_chars(anchor, field="'anchor'", ctx=f"EditOperation.{action}(): ", allow_tab=True)
        _reject_control_chars(text, field="'text'", ctx=f"EditOperation.{action}(): ", allow_newline=True)
        _validate_note(note, ctx=f"EditOperation.{action}(): ")
        return cls(action=action, paragraph=paragraph, anchor=anchor, text=text, occurrence=occurrence, note=note)

    @classmethod
    def insert_after(
        cls,
        anchor: "str | SearchResult",
        text: str,
        *,
        paragraph: str | None = None,
        occurrence: int | None = None,
        note: str | None = None,
    ) -> "EditOperation":
        """Build a validated insert_after operation (mirrors ``Document.insert_after``).

        Args:
            anchor: Text to find as insertion point (must be non-empty), or a
                :class:`SearchResult` from find_text()/find_all() — which also
                supplies ``paragraph`` and ``occurrence``, so pass neither.
            text: Text to insert after the anchor
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1").
                Required unless ``anchor`` is a SearchResult.
            occurrence: Which occurrence of anchor within the paragraph
                (0 = first). Omitted → ``anchor`` must be unique in the
                paragraph, else the batch fails with a wrapped
                AmbiguousTextError at apply time.
            note: Rationale for this edit, anchored as a comment on the
                revisions it creates (see ``Document.insert_after``).

        Raises:
            ValueError: If ``paragraph`` is missing or malformed, ``occurrence``
                is not a non-negative integer, ``anchor`` is not a non-empty
                string, ``text`` is not a string, ``note`` is neither None nor a
                non-empty control-character-free string, or ``anchor`` is a
                SearchResult and ``paragraph``/``occurrence`` was given too.
        """
        return cls._insert("insert_after", anchor, text, paragraph, occurrence, note)

    @classmethod
    def insert_before(
        cls,
        anchor: "str | SearchResult",
        text: str,
        *,
        paragraph: str | None = None,
        occurrence: int | None = None,
        note: str | None = None,
    ) -> "EditOperation":
        """Build a validated insert_before operation (mirrors ``Document.insert_before``).

        Args:
            anchor: Text to find as insertion point (must be non-empty), or a
                :class:`SearchResult` from find_text()/find_all() — which also
                supplies ``paragraph`` and ``occurrence``, so pass neither.
            text: Text to insert before the anchor
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1").
                Required unless ``anchor`` is a SearchResult.
            occurrence: Which occurrence of anchor within the paragraph
                (0 = first). Omitted → ``anchor`` must be unique in the
                paragraph, else the batch fails with a wrapped
                AmbiguousTextError at apply time.
            note: Rationale for this edit, anchored as a comment on the
                revisions it creates (see ``Document.insert_before``).

        Raises:
            ValueError: If ``paragraph`` is missing or malformed, ``occurrence``
                is not a non-negative integer, ``anchor`` is not a non-empty
                string, ``text`` is not a string, ``note`` is neither None nor a
                non-empty control-character-free string, or ``anchor`` is a
                SearchResult and ``paragraph``/``occurrence`` was given too.
        """
        return cls._insert("insert_before", anchor, text, paragraph, occurrence, note)


@dataclass(frozen=True)
class _ValidationOutcome:
    """Internal result of validating one operation: why it failed, and the ref
    that re-points it at the paragraph's current content when the failure was a
    stale hash.

    ``current_ref`` clears the *hash* check only — it says nothing about whether
    the operation's target text still exists, or is still unique, in that
    paragraph. A rebuilt operation is re-validated like any other (and fails
    loudly, atomically, at apply time if the target moved on).

    Named rather than a bare 2-tuple because both fields are ``str | None`` —
    positional unpacking would be one transposition away from reporting the
    recovery ref as the error message.
    """

    error: str | None  # None when the operation would apply cleanly
    current_ref: str | None = None  # set only for stale-hash failures


def _not_an_edit_operation_message(op: object) -> str:
    """Shared batch_edit/validate_batch message for a non-EditOperation element,
    so the raising and never-raises paths cannot drift apart."""
    return (
        f"expected EditOperation, got {type(op).__name__} — build operations with "
        "EditOperation.replace()/.delete()/.insert_after()/.insert_before()"
    )


@dataclass
class EditValidationResult:
    """Outcome of validating one EditOperation in a dry-run batch.

    ``current_ref`` is the recovery field for the one failure a caller can fix
    mechanically: a stale hash. It holds the ref that targets the same paragraph
    at its *current* content (e.g. ``"P7#c4d8"`` when the operation carried
    ``"P7#a7b2"``), so a retry is ``EditOperation.replace(..., paragraph=
    row.current_ref)`` — no parsing the hash out of ``error``'s prose. It is
    ``None`` for every other outcome: valid rows, malformed refs, out-of-range
    indexes, missing or ambiguous target text, and non-EditOperation elements.
    """

    index: int  # 0-based position in the input operations list
    paragraph: str | None  # the operation's paragraph ref (None if it was missing)
    valid: bool  # True if the op would apply cleanly
    error: str | None = None  # human-readable reason when not valid
    current_ref: str | None = None  # ref for the paragraph's current content, stale-hash rows only


class EditResult(str):
    """Result of a tracked edit: the new paragraph ref plus revision-group info.

    Subclasses ``str`` — the string value *is* the new hash-anchored
    paragraph reference (e.g. ``"P2#c3d4"``), so an EditResult works
    unchanged anywhere a ref string is expected (``paragraph=`` of
    follow-up edits, equality with plain strings, dict keys).

    Extra attributes:

    - ``group_id``: id of the revision group holding every revision this
      operation created, usable with ``accept_group``/``reject_group``.
      None when the operation created no new revisions — e.g. text amended
      into one of your own pending insertions (physically merged, so it is
      inseparable from the earlier operation at the XML level: undo it by
      rejecting the group of the insertion it amended), a rewrite that
      found no differences, or a rewrite whose changes all landed inside
      your own pending insertions. Group ids are per-open-Document
      and renumbered on each open — after close()/reopen the same revisions
      belong to a freshly inferred group with a new id, so never carry a
      group_id across sessions (see ``Document.accept_group``).
    - ``changeset_id``: id of the changeset (one whole call: this single
      edit, or the entire ``batch_edit``/``batch_rewrite``) that this
      operation's group belongs to, usable with
      ``accept_changeset``/``reject_changeset``. One changeset contains ≥1
      group; a single edit is a one-group changeset. None whenever
      ``group_id`` is None. Per-open-Document and renumbered on each open,
      exactly like ``group_id``.
    - ``revision_ids``: the w:ids of the group's member revisions, in
      creation order, as of this edit's return; ``()`` when ``group_id`` is
      None. A later edit that splits one of these insertions adds the
      split-off half to the group — ``Document.list_revisions`` reflects
      live membership.
    - ``refs``: every resulting paragraph ref, in document order. A normal
      edit stays inside one paragraph, so ``refs`` is ``(str(self),)``. A
      ``\\n`` edit is a tracked paragraph split, so ``refs`` carries the first
      paragraph (== the string value) plus one ref per new paragraph the split
      created (``("P2#…", "P3#…", …)``). Like all refs these are valid until
      the next structural edit; after a split, later paragraphs' indexes have
      shifted, so re-resolve (``list_paragraphs``/``find_text``) before reusing
      stale refs.
    - ``comment_id``: id of the comment holding this edit's ``note=``
      rationale, anchored on the revisions above — a live comment id, usable
      with ``reply_to_comment()``/``delete_comment()``/``resolve_comment()``.
      None when no ``note=`` was given, and None with an
      :class:`~docx_editor.exceptions.UnanchoredNoteWarning` when a note was
      given but the edit created no revision to anchor it on. Operations of one
      call that share the same note text share one comment, so several
      EditResults can carry the same id. The comment is deleted when the last
      revision it explains is resolved, accepted or rejected alike.
    """

    group_id: int | None
    changeset_id: int | None
    revision_ids: tuple[int, ...]
    refs: tuple[str, ...]
    comment_id: int | None

    def __new__(
        cls,
        ref: str,
        group_id: int | None = None,
        revision_ids: tuple[int, ...] = (),
        changeset_id: int | None = None,
        refs: tuple[str, ...] | None = None,
        comment_id: int | None = None,
    ) -> "EditResult":
        result = super().__new__(cls, ref)
        result.group_id = group_id
        result.changeset_id = changeset_id
        result.revision_ids = revision_ids
        result.refs = refs if refs is not None else (str(ref),)
        result.comment_id = comment_id
        return result


@dataclass(frozen=True)
class SearchResult:
    """Public result of ``Document.find_text`` / ``find_all`` — no DOM internals.

    ``start``/``end`` are character offsets in the *containing paragraph's*
    visible text (text maps are per-paragraph), not document-wide offsets.

    ``paragraph_ref`` is computed at search time and is directly usable as the
    ``paragraph=`` argument of follow-up edits; like refs from
    ``list_paragraphs()``, it is valid until that paragraph is edited.

    ``find_text``'s ``occurrence`` counts matches document-wide, while edit
    methods count within one paragraph — ``paragraph_occurrence`` bridges the
    two: pass it as the ``occurrence=`` of a follow-up edit to target exactly
    the match ``find_text`` located.

    ``paragraph_index`` is the 1-based document-order index of the containing
    paragraph — the same integer embedded in ``paragraph_ref`` — so consumers
    never need to string-parse the ref to compare or sort by position.

    ``repr()``/``str()`` are compact one-liners
    (``SearchResult(P3#a7b2 occ=0 '30 days')``) so printing a list of results
    stays cheap; matched text longer than 60 characters is elided with
    ``"..."`` in the display only. Every field — including the full ``text``
    — remains accessible as an attribute.
    """

    start: int  # Start offset in the paragraph's visible text
    end: int  # Exclusive end offset, same coordinate space
    text: str  # The matched text
    paragraph_ref: str  # Hash-anchored ref like "P3#a7b2"
    paragraph_occurrence: int  # Occurrence index of this match within its paragraph
    spans_revision: bool  # True if the match crosses a tracked-revision boundary
    paragraph_index: int  # 1-based index of the containing paragraph (same as in paragraph_ref)

    def __repr__(self) -> str:
        text = self.text if len(self.text) <= 60 else self.text[:57] + "..."
        spans = " spans_rev" if self.spans_revision else ""
        return f"SearchResult({self.paragraph_ref} occ={self.paragraph_occurrence} {text!r}{spans})"


def _resolve_search_target(
    target: "str | SearchResult",
    paragraph: str | None,
    occurrence: int | None,
    *,
    ctx: str,
    field: str,
) -> tuple[str, str | None, int | None]:
    """Normalize a text-or-SearchResult edit target into ``(text, paragraph, occurrence)``.

    A plain string passes through untouched, so every existing call site keeps
    its exact semantics. A :class:`SearchResult` supplies all three values —
    its ``text``, its ``paragraph_ref`` and its ``paragraph_occurrence`` —
    which is what makes ``doc.replace(match, "60 days")`` mean exactly what
    spelling those three fields out by hand means.

    Passing a SearchResult *and* ``paragraph=``/``occurrence=`` is a
    contradiction rather than a merge (which of the two paragraphs wins?), so
    it raises instead of silently preferring one.

    Args:
        target: Text to locate, or a SearchResult that already located it.
        paragraph: The call's ``paragraph=`` argument (None when omitted).
        occurrence: The call's ``occurrence=`` argument (None when omitted).
        ctx: Message prefix naming the caller, e.g. ``"replace(): "``.
        field: Display name of the target parameter, e.g. ``"'find'"``.

    Returns:
        ``(text, paragraph, occurrence)``, for the caller to validate exactly
        as it validates hand-written arguments. A plain tuple rather than a
        named object (cf. :class:`_ValidationOutcome`) because it is never
        carried: every call site spreads it straight back into the three
        distinctly-named parameters it came from, on the next line.

    Raises:
        ValueError: If ``target`` is a SearchResult and ``paragraph`` or
            ``occurrence`` was given too.
    """
    if not isinstance(target, SearchResult):
        return target, paragraph, occurrence
    redundant = [
        name
        for name, given in (("paragraph=", paragraph is not None), ("occurrence=", occurrence is not None))
        if given
    ]
    if redundant:
        raise ValueError(
            f"{ctx}{field} is a SearchResult, which already pins the paragraph ({target.paragraph_ref!r}) "
            f"and the occurrence ({target.paragraph_occurrence}) — drop {' and '.join(redundant)}"
        )
    return target.text, target.paragraph_ref, target.paragraph_occurrence


def _paragraph_hint(field: str) -> str:
    """Trailing hint shared by the "paragraph ref is missing" messages.

    ``paragraph`` is optional in the *signatures* only because a SearchResult
    target supplies it (see :func:`_resolve_search_target`); an edit still has to
    know which paragraph it targets. Document's methods and EditOperation's
    constructors keep their own long-standing message prefixes (both are pinned
    by tests) and share this recovery hint, so the advice cannot drift.
    """
    return (
        f"pass a ref from list_paragraphs()/find_text(), or pass a SearchResult as {field} "
        f"(it carries both the paragraph and the occurrence)"
    )


@dataclass(frozen=True)
class _LocatedMatch:
    """Internal: a text-map match plus the paragraph identity needed for refs."""

    match: TextMapMatch
    paragraph_index: int  # 1-based document-order index of the containing w:p
    paragraph: Element  # The containing w:p element
    paragraph_occurrence: int  # Occurrence index of the match within that paragraph


RevisionType = Literal["insertion", "deletion", "move_from", "move_to", "property_change"]

# ``Revision.type`` -> the short kind shown by ``Revision.__repr__``.
_REPR_KIND_BY_TYPE: dict[str, str] = {
    "insertion": "ins",
    "deletion": "del",
    "move_from": "moveFrom",
    "move_to": "moveTo",
    "property_change": "pPrChange",
}


@dataclass
class Revision:
    """Represents a tracked change: an insertion, deletion, move half or
    paragraph-property change.

    ``type`` is one of five values (see ``RevisionType``):

    - ``"insertion"`` / ``"deletion"``: a ``w:ins``/``w:del``; ``text`` is the
      inserted/deleted text.
    - ``"move_from"`` / ``"move_to"``: the two halves of a content move
      (``w:moveFrom``/``w:moveTo``), listed as two entries exactly as Word's
      revision pane shows "Moved from"/"Moved to"; ``text`` is the moved text
      as it appears in that half. Word never pairs the halves by id, so they
      are independent rows sharing an inferred ``changeset_id`` when they
      carry the same author and date — ``accept_changeset``/
      ``reject_changeset``, ``accept_all``/``reject_all`` resolve a move as a
      unit. Resolving one half by id is allowed but is what Word's per-entry
      resolution also permits: accepting the ``move_to`` and rejecting the
      ``move_from`` duplicates the text, the inverse loses it; a lone half
      in a damaged file behaves as the deletion/insertion it structurally is.
    - ``"property_change"``: a ``w:pPrChange`` — the paragraph's previous
      properties, recorded inside its ``w:pPr``. ``text`` is ``""``;
      accepting drops the record, rejecting restores the recorded
      properties.

    Location and nesting fields are populated by ``list_revisions``:

    - ``paragraph_ref``: hash-anchored reference (``"P{i}#{hash}"``) of the
      containing paragraph, or None when the revision sits in no *addressable*
      paragraph — outside any ``<w:p>`` (e.g. ``<w:trPr>`` row markers), or
      inside a drawing's text box, whose paragraphs are excluded from the ref
      index space (see ``body_paragraphs``). Such a revision is still
      listed, and ``accept_all``/``reject_all`` always resolve it. Anything
      narrower depends on how the box is stored. Word normally writes a box
      twice — an ``mc:Choice`` copy and an ``mc:Fallback`` copy — and then
      the revision is listed once per copy, so one ``accept_revision``/
      ``accept_group`` call resolves only the copy it lands on and leaves the
      twins out of step. Copies carrying distinct ids and identical
      ``w:author``/``w:date`` join one inferred changeset, so
      ``accept_changeset``/``reject_changeset`` reach both — along with every
      other group carrying that author and the identical raw ``w:date``
      string. Copies that share a ``w:id`` (a producer that duplicated the
      ``mc:Choice`` content verbatim) are ungroupable: ``group_id`` and
      ``changeset_id`` are both None, so no group- or changeset-keyed call
      can reach them and ``accept_all``/``reject_all`` is the single call
      that takes both. A box stored in one form only (VML ``w:pict`` with no
      ``mc:Fallback`` twin) is listed once and behaves like any other
      revision.
    - ``occurrence``: 0-based occurrence index of ``text`` within the
      containing paragraph, counted in the view where the revision's text
      lives — the accepted (visible) view for insertions and ``move_to``
      halves, the original (pre-revision) view for deletions and
      ``move_from`` halves. For insertions and ``move_to`` it plugs directly
      into the ``occurrence=`` parameter of replace()/delete()/
      add_comment(). None whenever targeting-by-text does not apply: empty
      text (property changes, paragraph-mark markers), a host insertion
      whose original text no longer matches its visible span (a nested
      deletion consumed part of it), a nested deletion (its text never
      existed in the original document), or a None ``paragraph_ref`` (an
      occurrence with no ref cannot be acted on).
    - ``nested_under``: id of the nearest enclosing revision (e.g. a foreign
      deletion inside another author's pending insertion), else None.
    - ``contains_ids``: ids of revisions nested inside this one, in document
      order. Both nesting fields report *structural* containment and so,
      unlike ``text``, still cross into a text box: a host insertion
      wrapping a run that carries a box lists the box's own revisions here,
      and they name it in ``nested_under``. Accepting the host does not
      resolve them — accepting an insertion unwraps only that element.
    - ``group_id``: id of the revision group this revision belongs to (all
      revisions from one logical edit share it — see
      ``Document.accept_group``). Edits made through the open Document
      record their group directly; revisions already in the file get an
      *inferred* group reconstructed at parse time (see ``group_source``).
      None only when the revision is ungroupable: it sits outside any
      paragraph, lacks an author or date, shares a duplicated id with
      another revision (id-keyed lookup cannot tell the occurrences
      apart), or is a mid-session split half of a foreign insertion.
      Revisions with non-numeric ids are omitted from ``list_revisions()``
      entirely — no id-keyed operation could target them.
    - ``group_source``: provenance of ``group_id`` — ``"recorded"`` for
      groups created by edits through this open Document, ``"inferred"``
      for groups reconstructed at parse time by the heuristic (contiguous
      same-paragraph revisions sharing identical ``w:author`` + ``w:date``
      are one group). Our own writes stamp collision-bumped dates (within
      one open session, two changesets by one author never share a
      second), so inferred groups match the original changesets;
      same-paragraph ops of one ``batch_edit``/``batch_rewrite`` call
      share their date and merge by design. The counter is per-session,
      so own writes from a previous session merge like foreign revisions:
      identical author + date can still over-merge (``w:date`` has second
      precision). None iff ``group_id`` is None.
    - ``changeset_id``: id of the changeset (one whole call — see
      ``EditResult.changeset_id``) this revision's group belongs to,
      resolvable with ``Document.accept_changeset``/``reject_changeset``.
      A changeset is the ``(author, date)`` equivalence class over groups:
      every group sharing this revision's ``w:author`` + ``w:date`` is in
      the same changeset (recorded edits bundle one call; inferred
      changesets partition reconstructed groups by that key). None iff
      ``group_id`` is None.
    - ``changeset_source``: provenance of ``changeset_id`` — ``"recorded"``
      for changesets bundled by a call through this open Document,
      ``"inferred"`` for changesets partitioned at parse time. None iff
      ``changeset_id`` is None.
    """

    id: int
    type: RevisionType
    author: str
    date: datetime | None
    text: str
    paragraph_ref: str | None = None
    occurrence: int | None = None
    nested_under: int | None = None
    contains_ids: tuple[int, ...] = ()
    group_id: int | None = None
    group_source: GroupSource | None = None
    changeset_id: int | None = None
    changeset_source: GroupSource | None = None

    def __repr__(self) -> str:
        kind = _REPR_KIND_BY_TYPE.get(self.type, self.type)
        location = f" @{self.paragraph_ref}" if self.paragraph_ref else ""
        preview = self.text[:30] + ("..." if len(self.text) > 30 else "")
        nested = f", nested_under={self.nested_under}" if self.nested_under is not None else ""
        contains = f", contains={list(self.contains_ids)}" if self.contains_ids else ""
        inferred = "(inferred)" if self.group_source == "inferred" else ""
        group = f", group={self.group_id}{inferred}" if self.group_id is not None else ""
        changeset = f", cs={self.changeset_id}" if self.changeset_id is not None else ""
        return f"Revision({kind} {self.id}{location}: '{preview}' by {self.author}{nested}{contains}{group}{changeset})"


# Revision-bearing element tags, per ECMA-376 Part 1 §17.13 (Annotations).
#
# HANDLED are the tags this library adjudicates by id: accept_revision /
# reject_revision / accept_all / reject_all / list_revisions all walk exactly
# these: insertions and deletions, the two halves of a content move, and a
# paragraph-property change (ISSUES.md #68 — the two foreign families the
# corpus census actually found in the wild, see benchmarks/corpus/README.md).
#
# MOVE_RANGE are scaffolding, not revisions: a move's *RangeStart/*RangeEnd
# pair brackets the moved content but carries no text and nothing to
# adjudicate, so the marks are never listed and never counted as unhandled —
# ``_sweep_move_range_marks`` removes a pair once no pending move content of
# its family remains between them.
#
# UNHANDLED are the rest of the schema's revision marks — parsed, carried
# through a save unchanged, and never resolved. Splitting them into named
# constants is what lets accept_all report what it could not touch (the
# "honesty floor", ISSUES.md #64) instead of silently claiming success.
HANDLED_REVISION_TAGS: tuple[str, ...] = ("w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:pPrChange")

MOVE_RANGE_TAGS: tuple[str, ...] = (
    "w:moveFromRangeStart",
    "w:moveFromRangeEnd",
    "w:moveToRangeStart",
    "w:moveToRangeEnd",
)

UNHANDLED_REVISION_TAGS: tuple[str, ...] = (
    # Property changes other than the paragraph's: the previous formatting,
    # recorded alongside the element whose properties changed.
    "w:rPrChange",
    "w:sectPrChange",
    "w:tblPrChange",
    "w:tblPrExChange",
    "w:trPrChange",
    "w:tcPrChange",
    "w:tblGridChange",
    # Table-structure revisions and numbering.
    "w:cellIns",
    "w:cellDel",
    "w:cellMerge",
    "w:numberingChange",
    # Custom-XML range marks.
    "w:customXmlInsRangeStart",
    "w:customXmlInsRangeEnd",
    "w:customXmlDelRangeStart",
    "w:customXmlDelRangeEnd",
    "w:customXmlMoveFromRangeStart",
    "w:customXmlMoveFromRangeEnd",
    "w:customXmlMoveToRangeStart",
    "w:customXmlMoveToRangeEnd",
)

ALL_REVISION_TAGS: tuple[str, ...] = HANDLED_REVISION_TAGS + MOVE_RANGE_TAGS + UNHANDLED_REVISION_TAGS

# Handled tag -> the ``Revision.type`` it is listed as.
_REVISION_TYPE_BY_TAG: dict[str, RevisionType] = {
    "w:ins": "insertion",
    "w:del": "deletion",
    "w:moveFrom": "move_from",
    "w:moveTo": "move_to",
    "w:pPrChange": "property_change",
}

# The handled tags that wrap run content (CT_RunTrackChange). The same four
# appear empty under ``w:pPr/w:rPr`` as paragraph-mark markers.
_RUN_TRACK_CHANGE_TAGS: tuple[str, ...] = ("w:ins", "w:del", "w:moveFrom", "w:moveTo")

# Tag -> the bracket kind ``get_markup_text`` renders it with.
_MARKUP_KIND_BY_TAG: dict[str, str] = {
    "w:ins": "ins",
    "w:del": "del",
    "w:moveFrom": "moveFrom",
    "w:moveTo": "moveTo",
}

# The property-change family: each of these records the element's *previous*
# properties as its child subtree, so everything inside one describes state
# that is already gone, not a pending revision.
#
# Three of the recorded types can legally hold revision marks:
#   - w:tcPrChange's w:tcPr is CT_TcPrInner -> w:cellIns/w:cellDel/w:cellMerge
#   - a paragraph-mark w:rPrChange's w:rPr is CT_ParaRPrOriginal, which opens
#     with EG_ParaRPrTrackChanges -> w:ins/w:del/w:moveFrom/w:moveTo
#   - w:pPrChange's w:pPr is CT_PPrBase -> w:numPr -> w:ins/w:numberingChange
# The rest (CT_SectPrBase, CT_TblPrBase, CT_TblPrExBase, CT_TrPrBase,
# CT_TblGridBase, and CT_RPrOriginal for a run-level w:rPrChange) cannot. The
# rule is stated for the whole family anyway, because "a change record
# describes the past" is what makes any of them historical.
#
# w:numberingChange is deliberately absent: it carries its previous value in a
# w:original attribute and has no recorded subtree to skip.
#
# NOTE: only the unhandled/pending path uses this. list_revisions() and
# accept_all()/reject_all() still walk every handled tag, so a historical
# w:del recorded inside a change record is listed and resolved like a live
# one — pre-existing behavior, pinned as a known gap by
# tests/test_unhandled_revisions.py::test_handled_path_still_adjudicates_
# marks_inside_change_records.
CHANGE_RECORD_TAGS: tuple[str, ...] = (
    "w:pPrChange",
    "w:rPrChange",
    "w:sectPrChange",
    "w:tblPrChange",
    "w:tblPrExChange",
    "w:trPrChange",
    "w:tcPrChange",
    "w:tblGridChange",
)


def iter_revision_elements(root, tags: Iterable[str], *, skip_change_records: bool = False) -> Iterator[Element]:
    """Yield every element under ``root`` whose tag is in ``tags``, document order.

    One recursive pre-order traversal for the whole tag set, rather than one
    ``getElementsByTagName`` per tag: with 30 revision tags the per-tag form
    would cost 30 full-document walks, which the ISSUES.md #56/#62 walk-count
    pin (``tests/test_revision_groups.py``) exists to prevent.

    By default no subtree is skipped, so a mark nested inside another revision
    — a ``w:del`` inside a foreign ``w:ins`` — is yielded after its host,
    exactly as ``_revision_elements`` does for the handled tags.

    Args:
        root: any DOM node; its descendants are searched, not itself.
        tags: the tag names to yield.
        skip_change_records: when True, a ``CHANGE_RECORD_TAGS`` element is
            still yielded but its subtree is not descended into. That subtree
            is the *recorded previous state*, so a ``w:cellIns`` sitting in a
            ``w:tcPrChange``'s recorded ``w:tcPr`` is a historical marker, not
            a second pending revision. Callers reporting what is still pending
            want True; a raw inventory of what the XML contains wants False.
    """
    tag_set = frozenset(tags)
    change_records = frozenset(CHANGE_RECORD_TAGS) if skip_change_records else frozenset()

    def walk(node) -> Iterator[Element]:
        for child in node.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            if child.tagName in tag_set:
                yield child
            if child.tagName in change_records:
                continue
            yield from walk(child)

    yield from walk(root)


@dataclass(frozen=True)
class RevisionCensus:
    """Counts of revision-bearing elements in one XML part, by tag.

    Produced by :func:`count_revision_elements`. Purely descriptive — it says
    what a document contains, not what any operation did with it. A raw
    inventory, so unlike the pending-revision count on
    :class:`ResolveResult` it *does* include marks recorded inside a change
    record (a ``w:cellIns`` in a ``w:tcPrChange``'s historical ``w:tcPr``).

    Attributes:
        by_tag: tag name -> number of elements, for every tag in
            ``ALL_REVISION_TAGS`` that occurs at least once. Tags with no
            occurrences are absent rather than zero, so the mapping stays a
            readable inventory.
        ins_del_contexts: parent tag name -> number of ``w:ins``/``w:del``
            elements directly under it. ``w:ins`` and ``w:del`` are also
            *structural* markers: under ``w:rPr`` they usually mark an
            inserted/deleted paragraph mark (``w:pPr/w:rPr``), though the same
            parent tag also covers a change record's recorded
            ``w:rPrChange/w:rPr``; under ``w:trPr`` they mark an
            inserted/deleted table row. Those cases resolve approximately today
            (see ``RevisionManager.accept_all``), so the context breakdown is
            the evidence for whether they occur in real documents (ISSUES.md
            #68) — an ordinary content revision shows up here as ``"w:p"`` or
            ``"w:tc"``.
    """

    by_tag: dict[str, int]
    ins_del_contexts: dict[str, int]

    @property
    def total(self) -> int:
        """Total revision elements counted, across all tags."""
        return sum(self.by_tag.values())


def count_revision_elements(root) -> RevisionCensus:
    """Census of every revision-bearing element under ``root``.

    One traversal covering all of ``ALL_REVISION_TAGS`` (see
    :func:`iter_revision_elements`). ``root`` is any DOM node — a whole parsed
    part, a ``w:body``, a single paragraph.
    """
    by_tag: dict[str, int] = {}
    contexts: dict[str, int] = {}
    for elem in iter_revision_elements(root, ALL_REVISION_TAGS):
        tag = elem.tagName
        by_tag[tag] = by_tag.get(tag, 0) + 1
        if tag in ("w:ins", "w:del"):
            parent = elem.parentNode
            parent_tag = getattr(parent, "tagName", "(root)")
            contexts[parent_tag] = contexts.get(parent_tag, 0) + 1
    return RevisionCensus(by_tag=by_tag, ins_del_contexts=contexts)


@dataclass(frozen=True)
class UnhandledRevision:
    """One revision element this library does not accept or reject.

    Returned by ``Document.list_unhandled_revisions()``. Deliberately not a
    :class:`Revision`: nothing here is adjudicable, so there is no id to hand
    ``accept_revision()`` — see ``list_unhandled_revisions``. Where ``id`` is
    not None it is reported for identification only: an unhandled mark's
    ``w:id`` may coincide with a *handled* revision's, and passing it to
    ``accept_revision``/``reject_revision`` then resolves that other element
    instead (ROADMAP.md #87).

    Attributes:
        tag: the element's tag — one of ``UNHANDLED_REVISION_TAGS``, or a
            ``HANDLED_REVISION_TAGS`` tag whose element carries no numeric
            ``w:id`` (nothing id-keyed can resolve it, so it is reported here
            rather than omitted from both listings)
            (e.g. ``"w:rPrChange"``).
        id: the element's ``w:id``, or None when it carries none or a
            non-numeric one. Unlike :class:`Revision`, an id-less mark is still
            listed — nothing here is targeted by id, so there is nothing to
            omit it from.
        author: ``w:author``, or ``"Unknown"`` when the attribute is absent —
            matching :class:`Revision`. ``w:tblGridChange`` and the range
            ``*End`` marks carry only ``w:id`` in the schema, so they always
            read as ``"Unknown"``.
        date: parsed ``w:date``, or None when absent or unparseable.
        paragraph_ref: hash-anchored ref of the containing ``<w:p>``, or None
            when the mark sits in no *addressable* paragraph — outside any
            paragraph (e.g. a ``w:tblPrChange`` in a table's properties, or a
            ``w:sectPrChange`` in a section break), or inside a drawing's text
            box, whose paragraphs are excluded from the ref index space (see
            ``body_paragraphs``). A mark inside a box is listed once per
            stored copy, exactly as :class:`Revision` is.
    """

    tag: str
    id: int | None
    author: str
    date: datetime | None
    paragraph_ref: str | None = None

    def __repr__(self) -> str:
        kind = self.tag.removeprefix("w:")
        ident = f" {self.id}" if self.id is not None else ""
        location = f" @{self.paragraph_ref}" if self.paragraph_ref else ""
        return f"UnhandledRevision({kind}{ident}{location} by {self.author})"


class ResolveResult(int):
    """Result of ``accept_all``/``reject_all``: the count plus what was skipped.

    Subclasses ``int``, and the int value *is* the number of revisions
    resolved — so ``count = doc.accept_all()`` keeps working unchanged in
    comparisons, arithmetic, f-strings and ``json.dumps``.

    Extra attributes:

    - ``unhandled``: how many revision elements the document still holds that
      this library never resolves — run/section/table property changes,
      table-structure revisions, ``w:numberingChange``, custom-XML range
      marks (see ``UNHANDLED_REVISION_TAGS``) — plus any handled-type mark
      whose ``w:id`` is missing or non-numeric, which no id-keyed call can
      reach. ``0`` on a redline made of insertions, deletions, moves and
      paragraph-property changes.
    - ``unhandled_types``: tag -> count for those elements, e.g.
      ``{"w:rPrChange": 3, "w:cellIns": 1}``. Empty when ``unhandled`` is 0.

    Both are counted *after* resolution, which is the honest measure of the
    claim being made ("everything is resolved"): a foreign mark inside a
    rejected insertion's subtree is removed with it, so it correctly does not
    appear. It is not a census of what the document held on entry — for that,
    read the counts before resolving.
    """

    unhandled: int
    unhandled_types: dict[str, int]

    def __new__(cls, count: int, unhandled_types: dict[str, int] | None = None) -> "ResolveResult":
        result = super().__new__(cls, count)
        types = dict(unhandled_types or {})
        result.unhandled_types = types
        result.unhandled = sum(types.values())
        return result

    def __repr__(self) -> str:
        if not self.unhandled:
            return f"ResolveResult({int(self)})"
        return f"ResolveResult({int(self)}, unhandled={self.unhandled} {self.unhandled_types})"

    # int has no __str__ of its own, so without this str()/format()/f-strings
    # would fall through to __repr__ and print "ResolveResult(2)" where every
    # existing caller expects "2".
    __str__ = int.__repr__


@dataclass
class _GroupCapture:
    """Filled in by ``RevisionManager._grouped`` when its with-block exits."""

    group_id: int | None = None
