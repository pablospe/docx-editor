"""Track changes management for docx_editor.

Provides RevisionManager for creating and managing tracked changes (insertions/deletions).
"""

import difflib
import re
from collections import OrderedDict
from dataclasses import dataclass
from datetime import datetime
from typing import Literal
from xml.dom.minidom import Element

from .exceptions import (
    BatchOperationError,
    HashMismatchError,
    ParagraphIndexError,
    RevisionError,
    TextNotFoundError,
)
from .xml_editor import (
    DocxXMLEditor,
    ParagraphRef,
    TextMap,
    TextMapMatch,
    TextPosition,
    _escape_xml,
    build_text_map,
    compute_paragraph_hash,
    compute_text_hash,
    find_in_text_map,
    get_rPr_xml,
    get_text_node_data,
    rebuild_run_fragments,
    render_plain_wt,
)


@dataclass
class EditOperation:
    """A single edit operation for batch processing.

    Prefer the typed constructors (:meth:`replace`, :meth:`delete`,
    :meth:`insert_after`, :meth:`insert_before`) — they validate arguments at
    construction time with the same rules ``batch_edit`` applies, so mistakes
    surface immediately instead of at apply time. The raw
    ``EditOperation(action=..., ...)`` form remains supported.
    """

    action: Literal["replace", "delete", "insert_after", "insert_before"]
    paragraph: str  # Required: hash-anchored ref like "P3#a7b2"
    find: str | None = None  # For replace
    replace_with: str | None = None  # For replace
    text: str | None = None  # For delete (text to delete) or insert (text to insert)
    anchor: str | None = None  # For insert_after/insert_before
    occurrence: int = 0

    @staticmethod
    def _validate_common(constructor: str, paragraph: str, occurrence: int) -> None:
        """Construction-time checks shared by all typed constructors."""
        ParagraphRef.parse(paragraph)
        if occurrence < 0:
            raise ValueError(f"EditOperation.{constructor}(): occurrence must be >= 0, got {occurrence}")

    @classmethod
    def replace(cls, find: str, replace_with: str, *, paragraph: str, occurrence: int = 0) -> "EditOperation":
        """Build a validated replace operation (mirrors ``Document.replace``).

        Args:
            find: Text to find and replace (must be non-empty)
            replace_with: Replacement text (empty string allowed — replacing
                with nothing is a valid tracked deletion)
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence within the paragraph (0 = first)

        Raises:
            ValueError: If the paragraph ref is malformed, ``occurrence`` is
                negative, ``find`` is empty, or ``replace_with`` is None.
        """
        cls._validate_common("replace", paragraph, occurrence)
        if not find:
            raise ValueError("EditOperation.replace(): 'find' must be a non-empty string — the text to search for")
        if replace_with is None:
            raise ValueError("EditOperation.replace(): 'replace_with' must be a string (empty string is allowed)")
        return cls(
            action="replace",
            paragraph=paragraph,
            find=find,
            replace_with=replace_with,
            occurrence=occurrence,
        )

    @classmethod
    def delete(cls, text: str, *, paragraph: str, occurrence: int = 0) -> "EditOperation":
        """Build a validated delete operation (mirrors ``Document.delete``).

        Args:
            text: Text to mark as deleted (must be non-empty)
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence within the paragraph (0 = first)

        Raises:
            ValueError: If the paragraph ref is malformed, ``occurrence`` is
                negative, or ``text`` is empty.
        """
        cls._validate_common("delete", paragraph, occurrence)
        if not text:
            raise ValueError("EditOperation.delete(): 'text' must be a non-empty string — the text to mark as deleted")
        return cls(action="delete", paragraph=paragraph, text=text, occurrence=occurrence)

    @classmethod
    def _insert(
        cls,
        action: Literal["insert_after", "insert_before"],
        anchor: str,
        text: str,
        paragraph: str,
        occurrence: int,
    ) -> "EditOperation":
        cls._validate_common(action, paragraph, occurrence)
        if not anchor:
            raise ValueError(f"EditOperation.{action}(): 'anchor' must be a non-empty string — the text to insert near")
        if text is None:
            raise ValueError(f"EditOperation.{action}(): 'text' must be a string (empty string is allowed)")
        return cls(action=action, paragraph=paragraph, anchor=anchor, text=text, occurrence=occurrence)

    @classmethod
    def insert_after(cls, anchor: str, text: str, *, paragraph: str, occurrence: int = 0) -> "EditOperation":
        """Build a validated insert_after operation (mirrors ``Document.insert_after``).

        Args:
            anchor: Text to find as insertion point (must be non-empty)
            text: Text to insert after the anchor
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence of anchor within the paragraph (0 = first)

        Raises:
            ValueError: If the paragraph ref is malformed, ``occurrence`` is
                negative, ``anchor`` is empty, or ``text`` is None.
        """
        return cls._insert("insert_after", anchor, text, paragraph, occurrence)

    @classmethod
    def insert_before(cls, anchor: str, text: str, *, paragraph: str, occurrence: int = 0) -> "EditOperation":
        """Build a validated insert_before operation (mirrors ``Document.insert_before``).

        Args:
            anchor: Text to find as insertion point (must be non-empty)
            text: Text to insert before the anchor
            paragraph: Paragraph reference from list_paragraphs() (e.g., "P2#f3c1")
            occurrence: Which occurrence of anchor within the paragraph (0 = first)

        Raises:
            ValueError: If the paragraph ref is malformed, ``occurrence`` is
                negative, ``anchor`` is empty, or ``text`` is None.
        """
        return cls._insert("insert_before", anchor, text, paragraph, occurrence)


@dataclass
class EditValidationResult:
    """Outcome of validating one EditOperation in a dry-run batch."""

    index: int  # 0-based position in the input operations list
    paragraph: str | None  # the operation's paragraph ref (None if it was missing)
    valid: bool  # True if the op would apply cleanly
    error: str | None = None  # human-readable reason when not valid


@dataclass(frozen=True)
class SearchResult:
    """Public result of ``Document.find_text`` — no DOM internals.

    ``start``/``end`` are character offsets in the *containing paragraph's*
    visible text (text maps are per-paragraph), not document-wide offsets.

    ``paragraph_ref`` is computed at search time and is directly usable as the
    ``paragraph=`` argument of follow-up edits; like refs from
    ``list_paragraphs()``, it is valid until that paragraph is edited.

    ``find_text``'s ``occurrence`` counts matches document-wide, while edit
    methods count within one paragraph — ``paragraph_occurrence`` bridges the
    two: pass it as the ``occurrence=`` of a follow-up edit to target exactly
    the match ``find_text`` located.
    """

    start: int  # Start offset in the paragraph's visible text
    end: int  # Exclusive end offset, same coordinate space
    text: str  # The matched text
    paragraph_ref: str  # Hash-anchored ref like "P3#a7b2"
    paragraph_occurrence: int  # Occurrence index of this match within its paragraph
    spans_revision: bool  # True if the match crosses a tracked-revision boundary


@dataclass(frozen=True)
class _LocatedMatch:
    """Internal: a text-map match plus the paragraph identity needed for refs."""

    match: TextMapMatch
    paragraph_index: int  # 1-based document-order index of the containing w:p
    paragraph: Element  # The containing w:p element
    paragraph_occurrence: int  # Occurrence index of the match within that paragraph


@dataclass
class Revision:
    """Represents a tracked change (insertion or deletion).

    Location and nesting fields are populated by ``list_revisions``:

    - ``paragraph_ref``: hash-anchored reference (``"P{i}#{hash}"``) of the
      containing paragraph, or None when the revision sits outside any
      ``<w:p>`` (e.g. ``<w:trPr>`` row markers).
    - ``occurrence``: 0-based occurrence index of ``text`` within the
      containing paragraph, counted in the view where the revision's text
      lives — the accepted (visible) view for insertions, the original
      (pre-revision) view for deletions. For insertions it plugs directly
      into the ``occurrence=`` parameter of replace()/delete()/
      add_comment(). None whenever targeting-by-text does not apply: empty
      text, a host insertion whose original text no longer matches its
      visible span (a nested deletion consumed part of it), or a nested
      deletion (its text never existed in the original document).
    - ``nested_under``: id of the nearest enclosing revision (e.g. a foreign
      deletion inside another author's pending insertion), else None.
    - ``contains_ids``: ids of revisions nested inside this one, in document
      order.
    """

    id: int
    type: Literal["insertion", "deletion"]
    author: str
    date: datetime | None
    text: str
    paragraph_ref: str | None = None
    occurrence: int | None = None
    nested_under: int | None = None
    contains_ids: tuple[int, ...] = ()

    def __repr__(self) -> str:
        kind = "ins" if self.type == "insertion" else "del"
        location = f" @{self.paragraph_ref}" if self.paragraph_ref else ""
        preview = self.text[:30] + ("..." if len(self.text) > 30 else "")
        nested = f", nested_under={self.nested_under}" if self.nested_under is not None else ""
        contains = f", contains={list(self.contains_ids)}" if self.contains_ids else ""
        return f"Revision({kind} {self.id}{location}: '{preview}' by {self.author}{nested}{contains})"


def _ancestor_paragraph(elem) -> Element | None:
    """Nearest <w:p> ancestor of ``elem``, or None if outside any paragraph.

    Not replaceable by ``xml_editor._innermost_ancestor``: this loop also
    stops at the first non-element ancestor, so it terminates even on node
    chains whose ``parentNode`` never yields None (mock DOMs in tests).
    """
    node = elem.parentNode
    while node is not None and node.nodeType == node.ELEMENT_NODE:
        if node.tagName == "w:p":
            return node
        node = node.parentNode
    return None


def _nearest_revision_ancestor_id(elem) -> int | None:
    """id of the closest <w:ins>/<w:del> ancestor carrying a w:id, else None."""
    node = elem.parentNode
    while node is not None and node.nodeType == node.ELEMENT_NODE:
        if node.tagName in ("w:ins", "w:del"):
            rev_id = node.getAttribute("w:id")
            if rev_id:
                return int(rev_id)
        node = node.parentNode
    return None


def _descendant_revision_ids(elem) -> tuple[int, ...]:
    """ids of all <w:ins>/<w:del> descendants of ``elem``, in document order."""
    ids: list[int] = []

    def walk(node) -> None:
        for child in node.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            if child.tagName in ("w:ins", "w:del"):
                rev_id = child.getAttribute("w:id")
                if rev_id:
                    ids.append(int(rev_id))
            walk(child)

    walk(elem)
    return tuple(ids)


def _insertion_text_nodes(elem) -> list:
    """All <w:t>/<w:delText> descendants of a <w:ins>, in document order.

    Including <w:delText> means a host insertion whose content was later
    deleted by a nested <w:del> still reports the full text it originally
    inserted (plain <w:delText> never appears under <w:ins> otherwise).
    """
    nodes: list = []

    def walk(node) -> None:
        for child in node.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            if child.tagName in ("w:t", "w:delText"):
                nodes.append(child)
            else:
                walk(child)

    walk(elem)
    return nodes


def _has_ancestor(node, ancestor) -> bool:
    """True if ``ancestor`` is ``node`` itself or one of its ancestors."""
    current = node
    while current is not None:
        if current is ancestor:
            return True
        current = current.parentNode
    return False


def _occurrence_in_text_map(tm: TextMap, elem, text: str) -> int | None:
    """0-based occurrence index of ``text`` at ``elem``'s own span in ``tm``.

    Mirrors ``find_in_text_map``'s stepping (``idx + 1`` between matches) so
    the result plugs directly into the ``occurrence=`` parameter of the
    anchor APIs. Returns None when the revision's span cannot be equated
    with ``text``: no position in the map belongs to ``elem``, the map
    text at the span's start doesn't spell ``text``, or the spelled-out
    span extends beyond ``elem`` (e.g. a partially consumed host insertion
    whose missing suffix happens to be spelled by the following text —
    anchoring there would silently cross the revision boundary).
    """
    if not text:
        return None
    start = next((i for i, pos in enumerate(tm.positions) if _has_ancestor(pos.node, elem)), None)
    if start is None:
        return None
    if not tm.text.startswith(text, start):
        return None
    if not all(_has_ancestor(pos.node, elem) for pos in tm.positions[start : start + len(text)]):
        return None
    count = 0
    idx = tm.text.find(text)
    while idx != -1 and idx < start:
        count += 1
        idx = tm.text.find(text, idx + 1)
    return count


class _RevisionLocationContext:
    """Per-``list_revisions``-call cache of paragraph indexes, refs, and text maps."""

    def __init__(self, dom):
        self._p_index = {id(p): i for i, p in enumerate(dom.getElementsByTagName("w:p"), start=1)}
        self._refs: dict[int, str] = {}
        self._maps: dict[tuple[int, str], TextMap] = {}

    def paragraph_ref(self, p) -> str | None:
        """Hash-anchored ref ("P{i}#{hash}") of ``p``; None if not indexed."""
        key = id(p)
        index = self._p_index.get(key)
        if index is None:
            return None
        if key not in self._refs:
            # Hash from the cached accepted map, shared with the occurrence
            # path, instead of compute_paragraph_hash (which builds its own).
            self._refs[key] = f"P{index}#{compute_text_hash(self.text_map(p, 'accepted').text)}"
        return self._refs[key]

    def text_map(self, p, view: Literal["accepted", "original"]) -> TextMap:
        """Cached text map of ``p`` for ``view``."""
        key = (id(p), view)
        if key not in self._maps:
            self._maps[key] = build_text_map(p, view=view)
        return self._maps[key]


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

    def _resolve_paragraph(self, ref: ParagraphRef):
        """Resolve a ParagraphRef to its <w:p> element, validating the hash.

        Args:
            ref: Parsed paragraph reference

        Returns:
            The <w:p> DOM element

        Raises:
            ParagraphIndexError: If paragraph index is out of range
            HashMismatchError: If the hash doesn't match current content
        """
        paragraphs = self.editor.dom.getElementsByTagName("w:p")
        if ref.index < 1 or ref.index > len(paragraphs):
            raise ParagraphIndexError(ref.index, len(paragraphs))
        p = paragraphs[ref.index - 1]
        actual_hash = compute_paragraph_hash(p)
        if actual_hash != ref.hash:
            tm = build_text_map(p)
            preview = tm.text[:80]
            if len(tm.text) > 80:
                preview += "..."
            raise HashMismatchError(ref.index, ref.hash, actual_hash, preview)
        return p

    def _find_in_paragraph(self, paragraph, text: str, occurrence: int = 0) -> TextMapMatch | None:
        """Find the nth occurrence of text within a single paragraph."""
        text_map = build_text_map(paragraph)
        return find_in_text_map(text_map, text, occurrence)

    def _paragraph_preview(self, paragraph) -> str:
        """Visible text of a paragraph, used for scoped TextNotFoundError previews.

        Truncation to 80 chars is handled by ``TextNotFoundError._truncate_preview``;
        callers must not truncate here, to keep a single source of truth.
        """
        return build_text_map(paragraph).text

    def batch_edit(self, operations: list[EditOperation]) -> list[int]:
        """Apply multiple edits atomically with upfront hash validation.

        Validates all paragraph hashes before applying any edits.
        Applies edits in reverse paragraph order so earlier paragraphs'
        hashes remain valid throughout.

        Args:
            operations: List of EditOperation objects (each must have paragraph set)

        Returns:
            List of change IDs, one per operation (in original input order)

        Raises:
            HashMismatchError: If any paragraph hash is stale (no edits applied)
            BatchOperationError: If any operation fails validation (carries
                ``operation_index`` so the caller knows which op failed).
        """
        if not operations:
            return []

        # Parse and validate all refs upfront
        parsed: list[tuple[int, ParagraphRef, EditOperation]] = []
        for i, op in enumerate(operations):
            if not op.paragraph:
                raise BatchOperationError(i, "paragraph reference is required for batch mode")
            ref = ParagraphRef.parse(op.paragraph)
            self._resolve_paragraph(ref)  # Raises HashMismatchError if stale
            parsed.append((i, ref, op))

        # Sort by paragraph index descending (reverse order) for application
        # Stable sort preserves original order for same-paragraph edits
        parsed.sort(key=lambda x: x[1].index, reverse=True)

        # Snapshot DOM before any mutation so we can roll back on partial failure.
        snapshot = self.editor.dom.toxml(encoding=self.editor.encoding)

        try:
            results = [0] * len(operations)
            for original_idx, _ref, op in parsed:
                try:
                    change_id = self._apply_single_edit(op)
                except ValueError as e:
                    raise BatchOperationError(original_idx, str(e)) from e
                results[original_idx] = change_id
            return results
        except Exception:
            # Restore via the line-tracking parser so parse_position is preserved.
            # If rollback itself fails, surface the original edit error — it is
            # the actionable one; a rollback failure is a secondary symptom.
            try:
                self.editor._reload_dom_from_bytes(snapshot)
            except Exception:
                pass
            raise

    def _resolve_action_target(self, op: EditOperation) -> str:
        """Validate op's required args and return the text this op must locate.

        Shared by ``_apply_single_edit`` and ``_validate_single`` so the two
        paths cannot drift out of sync. Rejects a negative ``occurrence`` up
        front (the one non-well-formed input ``_find_in_paragraph`` chokes on)
        so both paths fail cleanly before the search.

        Raises:
            ValueError: If ``occurrence`` is negative, required arguments for
                op.action are missing, or the action is unrecognized.
        """
        if op.occurrence < 0:
            raise ValueError(f"occurrence must be >= 0, got {op.occurrence}")

        if op.action == "replace":
            if not op.find or op.replace_with is None:
                raise ValueError("replace requires 'find' and 'replace_with'")
            return op.find
        elif op.action == "delete":
            if not op.text:
                raise ValueError("delete requires 'text'")
            return op.text
        elif op.action in ("insert_after", "insert_before"):
            if not op.anchor or op.text is None:
                raise ValueError(f"{op.action} requires 'anchor' and 'text'")
            return op.anchor
        else:
            raise ValueError(f"Unknown action: {op.action}")

    def _apply_single_edit(self, op: EditOperation) -> int:
        """Apply a single edit operation. Paragraph hash was already validated."""
        ref = ParagraphRef.parse(op.paragraph)
        p = self.editor.dom.getElementsByTagName("w:p")[ref.index - 1]

        target = self._resolve_action_target(op)
        match = self._find_in_paragraph(p, target, op.occurrence)
        if match is None:
            raise TextNotFoundError(
                target,
                paragraph_ref=op.paragraph,
                paragraph_preview=self._paragraph_preview(p),
            )

        if op.action == "replace":
            assert op.replace_with is not None  # guaranteed by _resolve_action_target
            return self._replace_across_nodes(match, op.replace_with)
        elif op.action == "delete":
            return self._delete_across_nodes(match)
        else:  # insert_after / insert_before
            assert op.text is not None  # guaranteed by _resolve_action_target
            position = "after" if op.action == "insert_after" else "before"
            return self._insert_near_match(match, op.text, position)

    def validate_batch(self, operations: list[EditOperation]) -> list[EditValidationResult]:
        """Validate a batch of edits without applying any of them.

        Mirrors the checks in ``batch_edit`` / ``_apply_single_edit`` (paragraph
        ref format, hash freshness, per-action argument requirements, and target
        text existence) but never raises and never mutates the document. Each
        operation gets its own result so the caller sees the full picture even
        when some ops are valid and others are not.

        Limitation: each operation is validated independently against the
        *current* document state; sequential effects are not simulated. A batch
        with multiple operations on the same paragraph (where one op's edit
        changes what a later op would see) may validate differently than it
        applies. Cross-paragraph batches are unaffected, since edits never
        change the paragraph count.

        Args:
            operations: List of EditOperation objects (each should have paragraph set)

        Returns:
            One EditValidationResult per operation, in input order.
        """
        results = []
        for i, op in enumerate(operations):
            error = self._validate_single(op)
            results.append(
                EditValidationResult(
                    index=i,
                    paragraph=op.paragraph,
                    valid=error is None,
                    error=error,
                )
            )
        return results

    def _validate_single(self, op: EditOperation) -> str | None:
        """Return an error message if ``op`` would fail, or None if it is valid.

        Reuses ``_resolve_paragraph``, ``_resolve_action_target``, and
        ``_find_in_paragraph`` — the same helpers ``_apply_single_edit`` uses —
        so dry-run validation cannot drift from real application semantics.
        Reads only.
        """
        if not op.paragraph:
            return "paragraph reference is required for batch mode"

        try:
            ref = ParagraphRef.parse(op.paragraph)
        except ValueError as e:
            return str(e)

        try:
            p = self._resolve_paragraph(ref)
        except (ParagraphIndexError, HashMismatchError) as e:
            return str(e)

        # Resolve required args + the text this op must locate via the same
        # helper _apply_single_edit uses (which also rejects a negative
        # occurrence), so validation cannot drift from application semantics and
        # the find below never raises.
        try:
            target = self._resolve_action_target(op)
        except ValueError as e:
            return str(e)

        if self._find_in_paragraph(p, target, op.occurrence) is None:
            return f"text {target!r} not found in paragraph {op.paragraph} ({self._paragraph_preview(p)!r})"

        return None

    def batch_rewrite(self, rewrites: list[tuple[str, str]]) -> None:
        """Rewrite multiple paragraphs with upfront hash validation."""
        if not rewrites:
            return

        # Parse and validate all refs upfront
        parsed: list[tuple[ParagraphRef, str]] = []
        seen_indices: set[int] = set()
        for i, (ref_str, new_text) in enumerate(rewrites):
            ref = ParagraphRef.parse(ref_str)
            if ref.index in seen_indices:
                raise BatchOperationError(
                    i,
                    f"duplicate paragraph P{ref.index}. Each paragraph can appear at most once in a batch rewrite.",
                )
            seen_indices.add(ref.index)
            self._resolve_paragraph(ref)  # Raises HashMismatchError if stale
            parsed.append((ref, new_text))

        # Sort by paragraph index descending
        parsed.sort(key=lambda x: x[0].index, reverse=True)

        # Apply rewrites in reverse paragraph order
        for ref, new_text in parsed:
            self.rewrite_paragraph(f"P{ref.index}#{ref.hash}", new_text)

    def rewrite_paragraph(self, ref_str: str, new_text: str) -> None:
        """Rewrite a paragraph's text, generating fine-grained tracked changes.

        Diffs old vs new text at word level and applies minimal tracked changes
        (insertions, deletions, replacements) to transform the paragraph.

        Args:
            ref_str: Paragraph reference string (e.g., "P3#a7b2")
            new_text: Desired new text for the paragraph

        Raises:
            HashMismatchError: If the paragraph hash doesn't match
            IndexError: If paragraph index is out of range
        """
        ref = ParagraphRef.parse(ref_str)
        p = self._resolve_paragraph(ref)
        text_map = build_text_map(p)
        old_text = text_map.text

        if old_text == new_text:
            return

        old_tokens = _tokenize_words(old_text)
        new_tokens = _tokenize_words(new_text)

        sm = difflib.SequenceMatcher(None, old_tokens, new_tokens)
        opcodes = sm.get_opcodes()

        # Convert token indices to character offsets in old_text
        old_token_offsets = []
        pos = 0
        for tok in old_tokens:
            old_token_offsets.append(pos)
            pos += len(tok)

        # Build hunks: list of (tag, old_char_start, old_char_end, new_fragment)
        hunks = []
        for tag, i1, i2, j1, j2 in opcodes:
            if tag == "equal":
                continue

            # Character range in old_text
            if i1 < len(old_tokens):
                old_char_start = old_token_offsets[i1]
            else:
                old_char_start = len(old_text)

            if i2 > 0:
                last_tok = old_tokens[i2 - 1]
                old_char_end = old_token_offsets[i2 - 1] + len(last_tok)
            else:
                old_char_end = old_char_start

            new_fragment = "".join(new_tokens[j1:j2])

            hunks.append((tag, old_char_start, old_char_end, new_fragment))

        # Process hunks in reverse order for position stability
        for tag, old_char_start, old_char_end, new_fragment in reversed(hunks):
            # Rebuild text_map each iteration since DOM changes
            text_map = build_text_map(p)

            if tag == "replace":
                match_text = old_text[old_char_start:old_char_end]
                match = self._find_match_at_position(text_map, match_text, old_char_start)
                self._replace_across_nodes(match, new_fragment)

            elif tag == "delete":
                match_text = old_text[old_char_start:old_char_end]
                match = self._find_match_at_position(text_map, match_text, old_char_start)
                self._delete_across_nodes(match)

            elif tag == "insert":
                self._rewrite_insert_at(p, text_map, old_char_start, new_fragment)

    def _find_match_at_position(self, text_map: TextMap, search: str, expected_pos: int) -> TextMapMatch:
        """Find text at an expected character position in the text map.

        Unlike find_in_text_map which finds the first occurrence, this
        verifies the match is at the expected position. Used by
        rewrite_paragraph() to avoid matching the wrong occurrence when
        the same text appears multiple times in a paragraph.

        Raises RevisionError if the text is not found at the expected position.
        """
        idx = text_map.find(search, expected_pos)
        if idx == -1 or idx != expected_pos:
            raise RevisionError(f"Rewrite failed: could not locate '{search}' at position {expected_pos}")
        end = idx + len(search)
        positions = text_map.get_nodes_for_range(idx, end)
        if positions:
            first_ins = positions[0].is_inside_ins
            spans = any(p.is_inside_ins != first_ins for p in positions)
        else:
            spans = False
        return TextMapMatch(
            start=idx,
            end=end,
            text=search,
            positions=positions,
            spans_boundary=spans,
        )

    def _rewrite_insert_at(self, paragraph, text_map: TextMap, char_pos: int, text: str) -> None:
        """Insert text at a character position within a paragraph.

        Used by rewrite_paragraph() for 'insert' opcodes.

        Args:
            paragraph: The <w:p> DOM element
            text_map: Current text map for the paragraph
            char_pos: Character position in visible text to insert at
            text: Text to insert
        """
        if not text_map.positions:
            # Empty paragraph — append insertion
            # Get rPr from any existing run, or use empty
            runs = paragraph.getElementsByTagName("w:r")
            rPr_xml = get_rPr_xml(runs[0]) if runs else ""

            ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"
            # Insert before w:sectPr if present, else append
            sect_prs = paragraph.getElementsByTagName("w:sectPr")
            if sect_prs:
                self.editor.insert_before(sect_prs[0], ins_xml)
            else:
                self.editor.append_to(paragraph, ins_xml)
            return

        if char_pos >= len(text_map.positions):
            # Insert at end — after last character's run
            last_pos = text_map.positions[-1]
            run, rPr_xml = self._get_run_info(last_pos.node)
            if not run:
                return

            # Inside our own <w:ins>: splice text directly; a foreign
            # author's insertion gets our own sibling <w:ins> instead
            ins_ancestor = self._find_ancestor(run, "w:ins")
            if ins_ancestor:
                if self._owns_ins(ins_ancestor):
                    node_text = self._get_node_text(last_pos.node)
                    self._set_node_text(last_pos.node, node_text + text)
                    _set_xml_space_preserve(last_pos.node)
                else:
                    node_text = self._get_node_text(last_pos.node)
                    self._insert_own_ins_within_foreign_ins(ins_ancestor, last_pos.node, len(node_text), text, rPr_xml)
                return

            ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"
            self.editor.insert_after(run, ins_xml)
            return

        # Insert at a position within the text
        pos = text_map.positions[char_pos]
        run, rPr_xml = self._get_run_info(pos.node)
        if not run:
            return

        # Inside our own <w:ins>: splice text directly; a foreign author's
        # insertion gets our own sibling <w:ins> (splitting theirs mid-content)
        ins_ancestor = self._find_ancestor(run, "w:ins")
        if ins_ancestor:
            if self._owns_ins(ins_ancestor):
                node_text = self._get_node_text(pos.node)
                offset = pos.offset_in_node
                self._set_node_text(pos.node, node_text[:offset] + text + node_text[offset:])
                _set_xml_space_preserve(pos.node)
            else:
                self._insert_own_ins_within_foreign_ins(ins_ancestor, pos.node, pos.offset_in_node, text, rPr_xml)
            return

        # Split the run at the offset and insert <w:ins> between
        node_text = self._get_node_text(pos.node)
        offset = pos.offset_in_node
        before_text = node_text[:offset]
        after_text = node_text[offset:]

        xml_parts = []
        if before_text:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before_text)}</w:t></w:r>")
        xml_parts.append(f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>")
        if after_text:
            xml_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after_text)}</w:t></w:r>")

        new_xml = "".join(xml_parts)
        self.editor.replace_node(run, new_xml)

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

    def _find_document_wide(self, text: str, occurrence: int = 0) -> TextMapMatch:
        """Document-wide nth-occurrence lookup via text maps.

        Raises:
            TextNotFoundError: If the text is not found or occurrence doesn't
                exist; ``total_occurrences`` matches :meth:`count_matches`.
        """
        match = self._find_across_boundaries(text, occurrence)
        if match is not None:
            return match
        total = self.count_matches(text)
        if total:
            raise TextNotFoundError(text, occurrence=occurrence, total_occurrences=total)
        raise TextNotFoundError(text)

    def _find_across_boundaries_located(self, text: str, occurrence: int = 0) -> _LocatedMatch | None:
        """Find the nth occurrence of text across element boundaries.

        Searches across all paragraphs using text maps, keeping paragraph
        identity so callers can build hash-anchored refs.

        Returns:
            A _LocatedMatch, or None if not found.
        """
        current_occurrence = 0
        for idx, paragraph in enumerate(self.editor.dom.getElementsByTagName("w:p"), start=1):
            text_map = build_text_map(paragraph)
            local_occ = 0
            while True:
                match = find_in_text_map(text_map, text, local_occ)
                if match is None:
                    break
                if current_occurrence == occurrence:
                    return _LocatedMatch(
                        match=match,
                        paragraph_index=idx,
                        paragraph=paragraph,
                        paragraph_occurrence=local_occ,
                    )
                current_occurrence += 1
                local_occ += 1
        return None

    def _find_across_boundaries(self, text: str, occurrence: int = 0) -> TextMapMatch | None:
        """Find the nth occurrence of text across element boundaries.

        Searches across all paragraphs using text maps.
        Returns TextMapMatch or None.
        """
        located = self._find_across_boundaries_located(text, occurrence)
        return located.match if located is not None else None

    def find_text(self, text: str, occurrence: int = 0) -> SearchResult | None:
        """Find the nth occurrence of text, as a public SearchResult.

        Searches across element boundaries; ``occurrence`` counts matches
        document-wide (0 = first). Returns None if not found.
        """
        located = self._find_across_boundaries_located(text, occurrence)
        if located is None:
            return None
        return SearchResult(
            start=located.match.start,
            end=located.match.end,
            text=located.match.text,
            paragraph_ref=f"P{located.paragraph_index}#{compute_paragraph_hash(located.paragraph)}",
            paragraph_occurrence=located.paragraph_occurrence,
            spans_revision=located.match.spans_boundary,
        )

    def replace_text(self, find: str, replace_with: str, occurrence: int = 0, paragraph: str | None = None) -> int:
        """Replace text with tracked changes (deletion + insertion).

        Finds the specified occurrence of `find` text and replaces it with `replace_with`,
        creating a tracked deletion for the old text and insertion for the new text.

        Args:
            find: Text to find and replace
            replace_with: Replacement text
            occurrence: Which occurrence to replace (0 = first, 1 = second, etc.)
            paragraph: Optional paragraph reference (e.g., "P2#f3c1") to scope the search

        Returns:
            The change ID of the insertion

        Raises:
            TextNotFoundError: If the text is not found or occurrence doesn't exist
            HashMismatchError: If the paragraph hash doesn't match
        """
        if paragraph is not None:
            ref = ParagraphRef.parse(paragraph)
            p = self._resolve_paragraph(ref)
            match = self._find_in_paragraph(p, find, occurrence)
            if match is None:
                raise TextNotFoundError(
                    find,
                    paragraph_ref=paragraph,
                    paragraph_preview=self._paragraph_preview(p),
                )
            return self._replace_across_nodes(match, replace_with)

        match = self._find_document_wide(find, occurrence)
        return self._replace_across_nodes(match, replace_with)

    def suggest_deletion(self, text: str, occurrence: int = 0, paragraph: str | None = None) -> int:
        """Mark text as deleted with tracked changes.

        Args:
            text: Text to mark as deleted
            occurrence: Which occurrence to delete (0 = first, 1 = second, etc.)
            paragraph: Optional paragraph reference (e.g., "P2#f3c1") to scope the search

        Returns:
            The change ID of the deletion

        Raises:
            TextNotFoundError: If the text is not found or occurrence doesn't exist
            HashMismatchError: If the paragraph hash doesn't match
        """
        if paragraph is not None:
            ref = ParagraphRef.parse(paragraph)
            p = self._resolve_paragraph(ref)
            match = self._find_in_paragraph(p, text, occurrence)
            if match is None:
                raise TextNotFoundError(
                    text,
                    paragraph_ref=paragraph,
                    paragraph_preview=self._paragraph_preview(p),
                )
            return self._delete_across_nodes(match)

        match = self._find_document_wide(text, occurrence)
        return self._delete_across_nodes(match)

    def _get_run_info(self, node) -> tuple[Element | None, str]:
        """Get the parent w:r element and its rPr XML for a w:t node."""
        run = node.parentNode
        while run and run.nodeName != "w:r":
            run = run.parentNode
        if not run:
            return None, ""
        return run, get_rPr_xml(run)

    def _get_node_text(self, node) -> str:
        """Get text content of a w:t node by concatenating ALL child text nodes.

        Thin wrapper around :func:`get_text_node_data` — kept as a method so
        existing call sites (and external subclasses) don't have to change.
        """
        return get_text_node_data(node)

    def _set_node_text(self, node, text: str) -> None:
        """Replace all text content of a w:t/w:delText element with ``text``.

        Removes every existing TEXT_NODE child and appends a single new one
        carrying the full content. Necessary because assigning to
        ``firstChild.data`` would leave any sibling text nodes behind,
        corrupting the document when the element holds split text (issue #9).
        """
        for child in list(node.childNodes):
            if child.nodeType == child.TEXT_NODE:
                node.removeChild(child)
        node.appendChild(node.ownerDocument.createTextNode(text))

    def _build_cross_boundary_parts(self, match: TextMapMatch) -> list[tuple[Element, str, str, str, str, int]]:
        """Build per-node data for a cross-boundary match.

        Returns list of (run, rPr_xml, before_text, matched_part, after_text, node_id) tuples,
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
        for nid, info in node_data.items():
            node_text = self._get_node_text(info["node"])
            first = info["first_offset"]
            last = info["last_offset"]
            before = node_text[:first]
            matched = node_text[first : last + 1]
            after = node_text[last + 1 :]
            result.append((info["run"], info["rPr_xml"], before, matched, after, nid))
        return result

    def _classify_segments(self, match: TextMapMatch) -> list[tuple[bool | None, list[TextPosition]]]:
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

        # Site D: all positions inside <w:ins> — dispatch on insertion ownership
        if all(p.is_inside_ins for p in match.positions):
            first_node = match.positions[0].node
            ins_elem = self._find_ancestor(first_node, "w:ins")
            ins_groups = self._group_positions_by_ins(match.positions)

            if all(g_ins is None or self._owns_ins(g_ins) for g_ins, _ in ins_groups):
                # All our own — edit in place (historical behavior)
                # Save parent/next sibling before removal may detach ins_elem
                ins_parent = ins_elem.parentNode if ins_elem else None
                ins_next = ins_elem.nextSibling if ins_elem else None

                self._remove_from_insertion(match.positions)

                first_rPr = parts[0][1]
                new_run_xml = f"<w:r>{first_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r>"

                if ins_elem and ins_elem.parentNode:
                    # ins_elem still in DOM — insert replacement inside it
                    first_child = ins_elem.firstChild
                    if first_child:
                        self.editor.insert_before(first_child, new_run_xml)
                    else:  # pragma: no cover - removing all content deletes the ins outright
                        self.editor.append_to(ins_elem, new_run_xml)
                elif ins_parent:  # pragma: no branch - an ins removed from the DOM had a parent
                    # ins_elem was fully removed — wrap replacement in a new <w:ins>
                    ins_wrapper_xml = f"<w:ins>{new_run_xml}</w:ins>"
                    if ins_next:
                        self.editor.insert_before(ins_next, ins_wrapper_xml)
                    else:
                        self.editor.append_to(ins_parent, ins_wrapper_xml)
                return -1

            # Foreign insertion(s) involved — preserve them: nest our deletion
            # inside, then place our replacement <w:ins> right after it,
            # splitting the foreign ins when trailing content follows.
            first_id, last_del = self._delete_from_ins_positions(match.positions)

            first_rPr = parts[0][1]
            replacement_xml = f"<w:ins><w:r>{first_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r></w:ins>"
            if last_del is None:  # pragma: no cover - a foreign group always creates a del
                return first_id
            del_ins = self._find_ancestor(last_del, "w:ins")
            if del_ins is not None:
                self._split_ins_after_child(del_ins, last_del)
                new_nodes = self.editor.insert_after(del_ins, replacement_xml)
            else:  # pragma: no cover - Site D positions are inside ins, so the del is nested
                new_nodes = self.editor.insert_after(last_del, replacement_xml)
            for node in new_nodes:
                if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":  # pragma: no branch
                    return int(node.getAttribute("w:id"))
            return first_id  # pragma: no cover - the fragment always yields a w:ins

        # Use first run's rPr for the insertion
        first_rPr = parts[0][1]

        # Group parts by run for multi-w:t preservation
        run_order: list[int] = []
        run_map: dict[int, dict] = {}
        for run, rPr_xml, before, matched, after, nid in parts:
            rid = id(run)
            if rid not in run_map:
                run_order.append(rid)
                run_map[rid] = {"run": run, "rPr_xml": rPr_xml, "parts": []}
            run_map[rid]["parts"].append((before, matched, after, nid))

        xml_parts = []
        part_idx = 0
        total_parts = len(parts)
        for rid in run_order:
            info = run_map[rid]
            run = info["run"]
            rPr_xml = info["rPr_xml"]

            # Build deterministic node-to-part mapping using node ids from parts
            node_to_part = {nid: (before, matched, after) for before, matched, after, nid in info["parts"]}
            parts_emitted = 0

            # Keyword-only defaults bind this iteration's state (B023)
            def render_wt(wt, *, node_to_part=node_to_part, run_rPr=rPr_xml, base_idx=part_idx) -> list[str]:
                nonlocal parts_emitted
                fragments: list[str] = []
                if id(wt) in node_to_part:
                    before, matched, after = node_to_part[id(wt)]
                    parts_emitted += 1

                    if before:
                        fragments.append(f"<w:r>{run_rPr}<w:t>{_escape_xml(before)}</w:t></w:r>")
                    fragments.append(
                        f"<w:del><w:r>{run_rPr}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>"
                    )

                    # Insert replacement after the last deletion
                    if base_idx + parts_emitted == total_parts:
                        fragments.append(f"<w:ins><w:r>{first_rPr}<w:t>{_escape_xml(replace_with)}</w:t></w:r></w:ins>")

                    if after:
                        fragments.append(f"<w:r>{run_rPr}<w:t>{_escape_xml(after)}</w:t></w:r>")
                else:
                    # Unmatched sibling — preserve
                    fragments.extend(render_plain_wt(wt, run_rPr))
                return fragments

            # Emit the run's children in document order (w:t split around the
            # match; w:tab/w:br/w:drawing/… preserved in place)
            xml_parts.extend(rebuild_run_fragments(run, rPr_xml, render_wt))
            part_idx += len(node_to_part)

        # Replace all affected runs: insert new XML before first run, remove all runs
        first_run = parts[0][0]
        new_xml = "".join(xml_parts)
        nodes = self.editor.insert_before(first_run, new_xml)

        seen = set()
        for run, _, _, _, _, _ in parts:
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
        marker = self.editor.dom.createComment("replace-marker")
        ref_node.parentNode.insertBefore(marker, ref_node)  # type: ignore[union-attr]

        # Process each segment to delete/remove the matched text
        # (author-aware: foreign insertions get a nested <w:del>, our own are
        # edited in place)
        for is_inside_ins, positions in segments:
            if is_inside_ins:
                self._delete_from_ins_positions(positions)
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

    def _find_ancestor(self, node, tag_name: str) -> Element | None:
        """Find the nearest ancestor with the given tag name."""
        parent = node.parentNode
        while parent:
            if parent.nodeType == parent.ELEMENT_NODE and parent.tagName == tag_name:
                return parent
            parent = parent.parentNode
        return None

    def _owns_ins(self, ins_elem) -> bool:
        """Whether ``ins_elem`` is the current author's own pending insertion.

        A missing or empty ``w:author`` reads as foreign: we must never
        destructively edit an insertion we cannot attribute to ourselves.
        Comparison is exact string equality — differing Unicode normalization
        or case reads as foreign, which fails safe (we nest instead of
        destroy).
        """
        author = ins_elem.getAttribute("w:author")
        return bool(author) and author == self.editor.author

    def _ins_identity_attrs(self, ins_elem) -> str:
        """Serialize the identity attributes of an insertion for re-creation.

        Returns ``w:author``/``w:date``/``w16du:dateUtc`` (those present) as an
        XML attribute string, so a re-created half of a split ``w:ins`` keeps
        the original author's identity; attribute injection only adds a
        fresh ``w:id``.
        """
        parts = []
        for attr in ("w:author", "w:date", "w16du:dateUtc"):
            value = ins_elem.getAttribute(attr)
            if value:
                parts.append(f' {attr}="{_escape_xml(value)}"')
        return "".join(parts)

    def _group_positions_by_ins(self, positions: list) -> list[tuple[Element | None, list[TextPosition]]]:
        """Group contiguous positions by their ancestor <w:ins> element.

        Positions are in document order, so positions sharing an ins element
        are contiguous; adjacent distinct ins elements form separate groups.
        A group's element is None for positions outside any insertion.
        """
        groups: list[tuple[Element | None, list[TextPosition]]] = []
        current_ins = None
        for pos in positions:
            ins_elem = self._find_ancestor(pos.node, "w:ins")
            if not groups or ins_elem is not current_ins:
                groups.append((ins_elem, [pos]))
                current_ins = ins_elem
            else:
                groups[-1][1].append(pos)
        return groups

    def _delete_from_ins_positions(self, positions: list) -> tuple[int, Element | None]:
        """Author-aware deletion of match positions that sit inside <w:ins>.

        Our own insertions are edited in place (text physically removed, as
        before). A foreign author's insertion is preserved: the matched text
        is wrapped in a nested <w:del> carrying our authorship — Word's own
        representation for deleting another reviewer's pending insertion.

        Returns (first created del id or -1, last created del element or None).
        """
        first_del_id = -1
        last_del: Element | None = None
        for ins_elem, group in self._group_positions_by_ins(positions):
            if ins_elem is None or self._owns_ins(ins_elem):
                self._remove_from_insertion(group)
            else:
                del_id, group_last_del = self._delete_regular_segment(group)
                if first_del_id == -1:
                    first_del_id = del_id
                if group_last_del is not None:  # pragma: no branch - a foreign group always creates a del
                    last_del = group_last_del
        return first_del_id, last_del

    def _split_ins_after_child(self, ins_elem, child) -> None:
        """Split ``ins_elem`` after ``child``, keeping the author's identity.

        Everything following ``child`` moves into a fresh sibling <w:ins>
        that copies this insertion's w:author/w:date (fresh w:id via
        attribute injection). ``ins_elem`` is typically another author's
        insertion — the copied identity keeps both halves attributed to
        them. No-op when nothing follows ``child``. ``child`` may be a
        descendant; the split happens after the direct child containing it.
        """
        while child.parentNode is not ins_elem:
            child = child.parentNode
        trailing = []
        node = child.nextSibling
        while node is not None:
            trailing.append(node)
            node = node.nextSibling
        if not any(n.nodeType == n.ELEMENT_NODE for n in trailing):
            return
        children_xml = "".join(n.toxml() for n in trailing)
        for n in trailing:
            ins_elem.removeChild(n)
        identity_xml = self._ins_identity_attrs(ins_elem)
        self.editor.insert_after(ins_elem, f"<w:ins{identity_xml}>{children_xml}</w:ins>")

    def _split_foreign_ins_at(self, edge_node, offset: int) -> Element | None:
        """Make (edge_node, offset) fall on a child boundary of its <w:ins>.

        Splits the run containing ``edge_node`` at ``offset`` when the split
        point falls mid-run. Returns the last element that belongs to the
        left side of the split point (None when the split point is at the
        very start of the insertion's content).
        """
        run, rPr_xml = self._get_run_info(edge_node)
        if not run:  # pragma: no cover - a w:t node always sits inside a run
            return None
        node_text = self._get_node_text(edge_node)

        # This site splits a run into left/right halves rather than rendering
        # per-w:t, and must know which side each child lands on, so it keeps
        # its own direct-children walk instead of rebuild_run_fragments.
        # Non-text children (w:tab, w:br, w:drawing, …) stay in document
        # order on whichever side of the split point they fall.
        left_parts: list[str] = []
        right_parts: list[str] = []
        side = left_parts
        for child in run.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            tag = getattr(child, "tagName", "")
            if tag == "w:rPr":
                continue
            if child is edge_node:
                if node_text[:offset]:
                    left_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(node_text[:offset])}</w:t></w:r>")
                if node_text[offset:]:
                    right_parts.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(node_text[offset:])}</w:t></w:r>")
                side = right_parts
                continue
            if tag == "w:t":
                wt_text = self._get_node_text(child)
                if wt_text:
                    side.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(wt_text)}</w:t></w:r>")
                continue
            side.append(f"<w:r>{rPr_xml}{child.toxml()}</w:r>")

        if not right_parts:
            # Split point is at the end of this run — run boundary already
            return run
        if not left_parts:
            # Split point is immediately before this run
            # (isinstance so ty narrows minidom's sibling union — the usual
            # nodeType comparison doesn't)
            prev = run.previousSibling
            while prev is not None:
                if isinstance(prev, Element):
                    return prev
                prev = prev.previousSibling
            return None

        new_nodes = self.editor.replace_node(run, "".join(left_parts + right_parts))
        elements = [n for n in new_nodes if n.nodeType == n.ELEMENT_NODE]
        return elements[len(left_parts) - 1]

    def _insert_own_ins_within_foreign_ins(self, ins_elem, edge_node, offset: int, text: str, rPr_xml: str) -> int:
        """Insert our own <w:ins> at (edge_node, offset) inside a foreign ins.

        Never splices into the foreign insertion (that would credit the other
        author) and never nests <w:ins> in <w:ins> (invalid OOXML). Boundary
        offsets produce a plain sibling; mid-insertion offsets split the
        foreign ins into two identity-preserving halves with our ins between.

        Returns the new insertion's change id.
        """
        own_ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"
        wt_nodes = self._get_wt_nodes_in_ancestor(ins_elem)
        node_text = self._get_node_text(edge_node)

        if edge_node is wt_nodes[0] and offset == 0:
            new_nodes = self.editor.insert_before(ins_elem, own_ins_xml)
        elif edge_node is wt_nodes[-1] and offset == len(node_text):
            new_nodes = self.editor.insert_after(ins_elem, own_ins_xml)
        else:
            boundary = self._split_foreign_ins_at(edge_node, offset)
            if boundary is None:  # pragma: no cover - non-boundary offsets never split at the start
                new_nodes = self.editor.insert_before(ins_elem, own_ins_xml)
            else:
                self._split_ins_after_child(ins_elem, boundary)
                new_nodes = self.editor.insert_after(ins_elem, own_ins_xml)

        for node in new_nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":  # pragma: no branch
                return int(node.getAttribute("w:id"))
        return -1  # pragma: no cover - the fragment always yields a w:ins

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
                        # Other w:t nodes exist — remove just this w:t (and run if empty)
                        self._remove_wt_and_maybe_run(first_node)
            elif not before_text:
                self._set_node_text(first_node, after_text)
                _set_xml_space_preserve(first_node)
            elif not after_text:
                self._set_node_text(first_node, before_text)
                _set_xml_space_preserve(first_node)
            else:
                # Middle split
                self._set_node_text(first_node, before_text)
                _set_xml_space_preserve(first_node)
                run = self._find_ancestor(first_node, "w:r")
                if ins_elem and run:
                    rPr_xml = get_rPr_xml(run)
                    after_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(after_text)}</w:t></w:r></w:ins>"
                    self.editor.insert_after(ins_elem, after_xml)
        else:
            # Multi-node: truncate first node to before, last node to after,
            # remove intermediate nodes entirely.
            # Only remove the w:t node; remove the run only if no w:t children remain.
            if before:
                self._set_node_text(first_node, before)
                _set_xml_space_preserve(first_node)
            else:
                self._remove_wt_and_maybe_run(first_node)

            if after:
                self._set_node_text(last_node, after)
                _set_xml_space_preserve(last_node)
            else:
                self._remove_wt_and_maybe_run(last_node)

            # Remove intermediate nodes unconditionally (entire text is matched)
            for group in groups[1:-1]:
                self._remove_wt_and_maybe_run(group["node"])

    def _remove_wt_and_maybe_run(self, wt_node) -> None:
        """Remove a w:t node, and its parent w:r if no meaningful children remain.

        Preserves the run if it still contains non-text content children
        like w:tab, w:br, w:drawing, etc.
        """
        run = self._find_ancestor(wt_node, "w:r")
        if wt_node.parentNode:
            wt_node.parentNode.removeChild(wt_node)
        if run and not run.getElementsByTagName("w:t") and run.parentNode:
            has_content_children = any(
                n.nodeType == n.ELEMENT_NODE and getattr(n, "tagName", None) not in ("w:t", "w:rPr")
                for n in run.childNodes
            )
            if not has_content_children:
                run.parentNode.removeChild(run)

    def _get_wt_nodes_in_ancestor(self, ancestor) -> list:
        """Get all w:t nodes inside an ancestor element."""
        if ancestor is None:
            return []
        return ancestor.getElementsByTagName("w:t")

    def _delete_regular_segment(self, positions: list) -> tuple[int, Element | None]:
        """Wrap matched text in <w:del> in place, run by run.

        Groups positions by run first, then by w:t node within each run,
        so that each run is removed exactly once even when it contains
        multiple w:t nodes involved in the match. The rebuilt runs stay at
        the original location, so this serves both regular top-level text
        and nesting a deletion inside a foreign author's <w:ins> (the new
        <w:del> is stamped with the current author either way).

        Returns (first created del id or -1, last created del element or None).
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
        last_del: Element | None = None
        processed_runs: set[int] = set()

        for _global_idx, (run_info, _) in enumerate(all_node_groups):
            run = run_info["run"]
            rPr_xml = run_info["rPr_xml"]
            rid = id(run)

            if rid in processed_runs:
                continue

            node_items = list(run_info["nodes"].values())

            # Render ALL w:t nodes in this run, preserving unmatched ones.
            # Keyword-only defaults bind this iteration's state (B023).
            def render_wt(wt, *, run_info=run_info, run_rPr=rPr_xml, node_items=node_items, rid=rid) -> list[str]:
                fragments: list[str] = []
                if id(wt) not in run_info["nodes"]:
                    # Unmatched sibling — preserve as-is
                    return render_plain_wt(wt, run_rPr)

                ng = run_info["nodes"][id(wt)]
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
                    fragments.append(f"<w:r>{run_rPr}<w:t>{_escape_xml(before)}</w:t></w:r>")
                fragments.append(f"<w:del><w:r>{run_rPr}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>")
                if after:
                    fragments.append(f"<w:r>{run_rPr}<w:t>{_escape_xml(after)}</w:t></w:r>")
                return fragments

            # Emit the run's children in document order (w:tab/w:br/w:drawing/…
            # preserved in place)
            new_xml = "".join(rebuild_run_fragments(run, rPr_xml, render_wt))
            nodes = self.editor.insert_before(run, new_xml)
            if run.parentNode:
                run.parentNode.removeChild(run)
            processed_runs.add(rid)

            for n in nodes:
                if n.nodeType == n.ELEMENT_NODE and n.tagName == "w:del":
                    if first_del_id == -1:
                        first_del_id = int(n.getAttribute("w:id"))
                    last_del = n

        return first_del_id, last_del

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

        # Site F: all positions inside <w:ins> — author-aware dispatch (our own
        # insertions edited in place, foreign ones get a nested <w:del>)
        if all(p.is_inside_ins for p in match.positions):
            first_id, _ = self._delete_from_ins_positions(match.positions)
            return first_id

        # Group parts by run, using node ids from _build_cross_boundary_parts
        run_parts: OrderedDict[int, list] = OrderedDict()
        for part in parts:
            rid = id(part[0])
            if rid not in run_parts:
                run_parts[rid] = []
            run_parts[rid].append(part)

        xml_parts = []
        for _rid, rparts in run_parts.items():
            run = rparts[0][0]
            rPr_xml = rparts[0][1]

            # Build deterministic node-to-part mapping using node ids from parts
            node_to_part = {nid: (rp_xml, before, matched, after) for _, rp_xml, before, matched, after, nid in rparts}

            # Keyword-only defaults bind this iteration's state (B023)
            def render_wt(wt, *, node_to_part=node_to_part, run_rPr=rPr_xml) -> list[str]:
                fragments: list[str] = []
                if id(wt) in node_to_part:
                    rp_xml, before, matched, after = node_to_part[id(wt)]
                    if before:
                        fragments.append(f"<w:r>{rp_xml}<w:t>{_escape_xml(before)}</w:t></w:r>")
                    fragments.append(f"<w:del><w:r>{rp_xml}<w:delText>{_escape_xml(matched)}</w:delText></w:r></w:del>")
                    if after:
                        fragments.append(f"<w:r>{rp_xml}<w:t>{_escape_xml(after)}</w:t></w:r>")
                else:
                    # Unmatched sibling — preserve
                    fragments.extend(render_plain_wt(wt, run_rPr))
                return fragments

            # Emit the run's children in document order (matched w:t as
            # <w:del>; w:tab/w:br/w:drawing/… preserved in place)
            xml_parts.extend(rebuild_run_fragments(run, rPr_xml, render_wt))

        first_run = parts[0][0]
        new_xml = "".join(xml_parts)
        nodes = self.editor.insert_before(first_run, new_xml)

        seen = set()
        for run, _, _, _, _, _ in parts:
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
                del_id, _ = self._delete_from_ins_positions(positions)
            else:
                del_id, _ = self._delete_regular_segment(positions)
            if first_del_id == -1:
                first_del_id = del_id

        return first_del_id

    def insert_text_after(self, anchor: str, text: str, occurrence: int = 0, paragraph: str | None = None) -> int:
        """Insert text after anchor with tracked changes.

        Args:
            anchor: Text to find as the anchor point
            text: Text to insert after the anchor
            occurrence: Which occurrence of anchor to use (0 = first, 1 = second, etc.)
            paragraph: Optional paragraph reference (e.g., "P2#f3c1") to scope the search

        Returns:
            The change ID of the insertion

        Raises:
            TextNotFoundError: If the anchor text is not found or occurrence doesn't exist
            HashMismatchError: If the paragraph hash doesn't match
        """
        return self._insert_text(anchor, text, position="after", occurrence=occurrence, paragraph=paragraph)

    def insert_text_before(self, anchor: str, text: str, occurrence: int = 0, paragraph: str | None = None) -> int:
        """Insert text before anchor with tracked changes.

        Args:
            anchor: Text to find as the anchor point
            text: Text to insert before the anchor
            occurrence: Which occurrence of anchor to use (0 = first, 1 = second, etc.)
            paragraph: Optional paragraph reference (e.g., "P2#f3c1") to scope the search

        Returns:
            The change ID of the insertion

        Raises:
            TextNotFoundError: If the anchor text is not found or occurrence doesn't exist
            HashMismatchError: If the paragraph hash doesn't match
        """
        return self._insert_text(anchor, text, position="before", occurrence=occurrence, paragraph=paragraph)

    def _insert_text(
        self,
        anchor: str,
        text: str,
        position: Literal["before", "after"],
        occurrence: int = 0,
        paragraph: str | None = None,
    ) -> int:
        """Insert text before or after anchor with tracked changes."""
        if paragraph is not None:
            ref = ParagraphRef.parse(paragraph)
            p = self._resolve_paragraph(ref)
            match = self._find_in_paragraph(p, anchor, occurrence)
            if match is None:
                raise TextNotFoundError(
                    anchor,
                    paragraph_ref=paragraph,
                    paragraph_preview=self._paragraph_preview(p),
                )
            return self._insert_near_match(match, text, position)

        match = self._find_document_wide(anchor, occurrence)
        return self._insert_near_match(match, text, position)

    def _insert_near_match(self, match: TextMapMatch, text: str, position: Literal["before", "after"]) -> int:
        """Insert text before/after a match, splitting the edge w:t at the match boundary."""
        positions = match.positions
        if not positions:
            return -1

        if position == "after":
            edge = positions[-1]
            offset = edge.offset_in_node + 1
        else:
            edge = positions[0]
            offset = edge.offset_in_node

        run, rPr_xml = self._get_run_info(edge.node)
        if not run:
            return -1

        # Edge run inside <w:ins>: splice into our own insertion (no wrapper,
        # so no nested <w:ins>); a foreign author's insertion gets our own
        # sibling <w:ins>, splitting theirs when the anchor falls mid-content.
        ins_ancestor = self._find_ancestor(run, "w:ins")
        if ins_ancestor:
            if self._owns_ins(ins_ancestor):
                node_text = self._get_node_text(edge.node)
                self._set_node_text(edge.node, node_text[:offset] + text + node_text[offset:])
                _set_xml_space_preserve(edge.node)
                return -1
            return self._insert_own_ins_within_foreign_ins(ins_ancestor, edge.node, offset, text, rPr_xml)

        # Rebuild the edge run: split its w:t at the offset and wrap text in
        # <w:ins>; non-text children (w:tab, w:br, w:drawing, …) stay in place
        ins_xml = f"<w:ins><w:r>{rPr_xml}<w:t>{_escape_xml(text)}</w:t></w:r></w:ins>"

        def render_wt(wt) -> list[str]:
            fragments: list[str] = []
            if wt is edge.node:
                node_text = self._get_node_text(wt)
                before_text = node_text[:offset]
                after_text = node_text[offset:]
                if before_text:
                    fragments.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(before_text)}</w:t></w:r>")
                fragments.append(ins_xml)
                if after_text:
                    fragments.append(f"<w:r>{rPr_xml}<w:t>{_escape_xml(after_text)}</w:t></w:r>")
            else:
                # Preserve unmatched sibling w:t
                fragments.extend(render_plain_wt(wt, rPr_xml))
            return fragments

        nodes = self.editor.replace_node(run, "".join(rebuild_run_fragments(run, rPr_xml, render_wt)))

        for node in nodes:
            if node.nodeType == node.ELEMENT_NODE and node.tagName == "w:ins":
                return int(node.getAttribute("w:id"))
        return -1

    def list_revisions(
        self,
        author: str | None = None,
        paragraph: str | None = None,
        *,
        with_location: bool = True,
    ) -> list[Revision]:
        """List all tracked changes in the document.

        Args:
            author: If provided, filter by author name
            paragraph: If provided, a paragraph reference (e.g. "P3#a7b2")
                from list_paragraphs(); only revisions inside that paragraph
                are returned.
            with_location: If False, skip computing ``paragraph_ref`` and
                ``occurrence`` (they stay None) — the location work builds a
                text map and hash per revision-bearing paragraph, wasted on
                callers that only need ids (accept_all/reject_all re-list on
                every pass). Forced True when ``paragraph`` is given, since
                the filter matches on ``paragraph_ref``.

        Returns:
            List of Revision objects sorted by id, with location and nesting
            fields populated — see :class:`Revision`.

        Raises:
            ValueError: If ``paragraph`` is malformed
            ParagraphIndexError: If the paragraph index is out of range
            HashMismatchError: If the paragraph hash doesn't match current content
        """
        paragraph_filter = None
        if paragraph is not None:
            ref = ParagraphRef.parse(paragraph)
            self._resolve_paragraph(ref)  # validates index and hash
            paragraph_filter = f"P{ref.index}#{ref.hash}"

        ctx = None
        if with_location or paragraph_filter is not None:
            ctx = _RevisionLocationContext(self.editor.dom)

        def matches(rev: Revision | None) -> bool:
            if rev is None:
                return False
            if author is not None and rev.author != author:
                return False
            return paragraph_filter is None or rev.paragraph_ref == paragraph_filter

        revisions = []

        # Find all insertions
        for ins_elem in self.editor.dom.getElementsByTagName("w:ins"):
            rev = self._parse_revision(ins_elem, "insertion", ctx)
            if matches(rev):
                revisions.append(rev)

        # Find all deletions
        for del_elem in self.editor.dom.getElementsByTagName("w:del"):
            rev = self._parse_revision(del_elem, "deletion", ctx)
            if matches(rev):
                revisions.append(rev)

        # Sort by ID
        revisions.sort(key=lambda r: r.id)
        return revisions

    def get_markup_text(self) -> str:
        """Render document text with inline revision markup.

        Each paragraph is one line; tracked changes wrap their content as
        ``[ins#{id}:{author}]...[/ins]`` / ``[del#{id}:{author}]...[/del]``,
        nesting included (e.g. ``[ins#1:A]kept [del#9:B]gone[/del][/ins]``).

        A human/agent verification view, not a parseable format: author
        names are not escaped, tabs/breaks are not rendered, and text inside
        a drawing's text box appears both inline in the host paragraph's
        line and again as its own line (same as get_text()).
        """

        def render(node) -> str:
            parts: list[str] = []
            for child in node.childNodes:
                if child.nodeType != child.ELEMENT_NODE:
                    continue
                if child.tagName in ("w:ins", "w:del"):
                    kind = "ins" if child.tagName == "w:ins" else "del"
                    rev_id = child.getAttribute("w:id")
                    rev_author = child.getAttribute("w:author") or "Unknown"
                    parts.append(f"[{kind}#{rev_id}:{rev_author}]{render(child)}[/{kind}]")
                elif child.tagName in ("w:t", "w:delText"):
                    parts.append(get_text_node_data(child))
                else:
                    parts.append(render(child))
            return "".join(parts)

        return "\n".join(render(p) for p in self.editor.dom.getElementsByTagName("w:p"))

    def _parse_revision(
        self,
        elem,
        rev_type: Literal["insertion", "deletion"],
        ctx: _RevisionLocationContext | None = None,
    ) -> Revision | None:
        """Parse a w:ins or w:del element into a Revision object.

        Args:
            elem: The <w:ins>/<w:del> element
            rev_type: Which kind of revision ``elem`` is
            ctx: Per-call location cache from list_revisions. None (detached
                elements, unit tests) leaves paragraph_ref/occurrence unset.
        """
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
            text_elems = _insertion_text_nodes(elem)
        else:
            text_elems = elem.getElementsByTagName("w:delText")
            if not text_elems:
                # Deliberate interop fallback: nonconforming producers may
                # leave plain w:t inside w:del. Fires only when the w:del has
                # no w:delText at all; mixed content reads only w:delText.
                text_elems = elem.getElementsByTagName("w:t")

        text = "".join(self._get_node_text(t_elem) for t_elem in text_elems)

        paragraph_ref = None
        occurrence = None
        if ctx is not None:
            paragraph = _ancestor_paragraph(elem)
            if paragraph is not None:
                paragraph_ref = ctx.paragraph_ref(paragraph)
                if text:
                    # Insertions live in the visible text; deletions in the
                    # original (pre-revision) text.
                    view: Literal["accepted", "original"] = "accepted" if rev_type == "insertion" else "original"
                    occurrence = _occurrence_in_text_map(ctx.text_map(paragraph, view), elem, text)

        return Revision(
            id=int(rev_id),
            type=rev_type,
            author=author,
            date=date,
            text=text,
            paragraph_ref=paragraph_ref,
            occurrence=occurrence,
            nested_under=_nearest_revision_ancestor_id(elem),
            contains_ids=_descendant_revision_ids(elem),
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

        Repeats passes until no listed revision can be processed, fully
        resolving nested revisions in Word-authored files (e.g. a w:del inside
        a w:ins) and terminating even when an author filter leaves other
        authors' revisions in the document. Revisions are matched by w:id, so
        if Word emits duplicate ids across authors, a filtered call may also
        process a same-id revision by another author.

        Args:
            author: If provided, only accept revisions by this author

        Returns:
            Number of revisions accepted
        """
        count = 0
        while True:
            progressed = False
            # Process in reverse order by ID to avoid index issues
            revisions = self.list_revisions(author=author, with_location=False)
            for rev in sorted(revisions, key=lambda r: r.id, reverse=True):
                if self.accept_revision(rev.id):
                    count += 1
                    progressed = True
            if not progressed:
                return count

    def reject_all(self, author: str | None = None) -> int:
        """Reject all revisions, optionally filtered by author.

        Repeats passes until no listed revision can be processed, fully
        resolving nested revisions in Word-authored files (e.g. a w:del inside
        a w:ins) and terminating even when an author filter leaves other
        authors' revisions in the document. Revisions are matched by w:id, so
        if Word emits duplicate ids across authors, a filtered call may also
        process a same-id revision by another author.

        Args:
            author: If provided, only reject revisions by this author

        Returns:
            Number of revisions rejected
        """
        count = 0
        while True:
            progressed = False
            # Process in reverse order by ID to avoid index issues
            revisions = self.list_revisions(author=author, with_location=False)
            for rev in sorted(revisions, key=lambda r: r.id, reverse=True):
                if self.reject_revision(rev.id):
                    count += 1
                    progressed = True
            if not progressed:
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


def _tokenize_words(text: str) -> list[str]:
    """Split text into alternating word and whitespace tokens."""
    return re.findall(r"\S+|\s+", text)


def _set_xml_space_preserve(wt_elem) -> None:
    """Set xml:space='preserve' on a w:t element to preserve whitespace."""
    wt_elem.setAttribute("xml:space", "preserve")
