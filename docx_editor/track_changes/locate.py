"""Text location: the ``RevisionManager`` mixin that resolves paragraph refs and finds text."""

from xml.dom.minidom import Element

from ..exceptions import (
    AmbiguousTextError,
    HashMismatchError,
    ParagraphIndexError,
    TextNotFoundError,
    _truncate_preview,
)
from ..xml_editor import (
    ParagraphRef,
    TextMapMatch,
    _require_valid_occurrence,
    body_paragraphs,
    build_text_map,
    compute_paragraph_hash,
    count_in_text_map,
    find_in_text_map,
)
from .base import _RevisionManagerBase
from .models import SearchResult, _LocatedMatch


class _LocateMixin(_RevisionManagerBase):
    def _resolve_paragraph(self, ref: ParagraphRef, paragraphs: list[Element] | None = None):
        """Resolve a ParagraphRef to its <w:p> element, validating the hash.

        Args:
            ref: Parsed paragraph reference
            paragraphs: Optional pre-fetched list of every <w:p> element, so
                batch callers pay for one full-DOM walk per batch instead of
                one per operation. Default None fetches fresh.

        Returns:
            The <w:p> DOM element

        Raises:
            ParagraphIndexError: If paragraph index is out of range
            HashMismatchError: If the hash doesn't match current content
        """
        if paragraphs is None:
            paragraphs = body_paragraphs(self.editor.dom)
        if ref.index < 1 or ref.index > len(paragraphs):
            raise ParagraphIndexError(ref.index, len(paragraphs))
        p = paragraphs[ref.index - 1]
        actual_hash = compute_paragraph_hash(p)
        if actual_hash != ref.hash:
            tm = build_text_map(p)
            preview = _truncate_preview(tm.text)
            raise HashMismatchError(ref.index, ref.hash, actual_hash, preview)
        return p

    def _locate_in_paragraph(self, paragraph, paragraph_ref: str, text: str, occurrence: int | None) -> TextMapMatch:
        """The single scoped find-or-raise path shared by every paragraph-scoped edit.

        Args:
            paragraph: The resolved <w:p> element.
            paragraph_ref: The caller's ref string (for error messages).
            text: Text to locate.
            occurrence: Which occurrence (0 = first). None means the text must
                be unique within the paragraph.

        Raises:
            ValueError: If ``occurrence`` is negative or not an integer, or ``text`` is not a
                non-empty string.
            TextNotFoundError: If the text is absent, or ``occurrence`` is out
                of range (then with ``occurrence``/``total_occurrences`` set).
            AmbiguousTextError: If ``occurrence`` is None and the text matches
                more than once in the paragraph.
        """
        _require_valid_occurrence(occurrence)
        if not isinstance(text, str) or not text:
            raise ValueError(f"search text must be a non-empty string, got {text!r}")
        text_map = build_text_map(paragraph)
        total = count_in_text_map(text_map, text)

        occ = occurrence if occurrence is not None else 0
        match = find_in_text_map(text_map, text, occ)
        if match is None:
            if total > 0:
                raise TextNotFoundError(
                    text,
                    paragraph_ref=paragraph_ref,
                    paragraph_preview=text_map.text,
                    occurrence=occ,
                    total_occurrences=total,
                )
            raise TextNotFoundError(
                text,
                paragraph_ref=paragraph_ref,
                paragraph_preview=text_map.text,
            )
        if occurrence is None and total > 1:
            raise AmbiguousTextError(
                text,
                paragraph_ref=paragraph_ref,
                paragraph_preview=text_map.text,
                total_occurrences=total,
            )
        return match

    def count_matches(self, text: str) -> int:
        """Count how many times a text string appears in the document.

        Uses text maps for accurate counting across element boundaries.

        Args:
            text: Text to search for

        Returns:
            Number of occurrences found
        """
        count = 0
        for paragraph in body_paragraphs(self.editor.dom):
            count += count_in_text_map(build_text_map(paragraph), text)
        return count

    def _locate_document_wide(self, text: str, occurrence: int | None = None) -> TextMapMatch:
        """Document-wide nth-occurrence lookup via text maps.

        Totals come from :meth:`count_matches` rather than a single
        ``count_in_text_map`` call (as ``_locate_in_paragraph`` uses) because
        text maps are per-paragraph — there is no one document-wide map — and
        the exact total is part of the error contract.

        Raises:
            ValueError: If ``occurrence`` is negative or not an integer, or ``text`` is not a
                non-empty string.
            TextNotFoundError: If the text is not found or occurrence doesn't
                exist; ``total_occurrences`` matches :meth:`count_matches`.
            AmbiguousTextError: If ``occurrence`` is None and the text matches
                more than once in the document.
        """
        _require_valid_occurrence(occurrence)
        if not isinstance(text, str) or not text:
            raise ValueError(f"search text must be a non-empty string, got {text!r}")
        occ = occurrence if occurrence is not None else 0
        match = self._find_across_boundaries(text, occ)
        if match is None:
            total = self.count_matches(text)
            if total:
                raise TextNotFoundError(text, occurrence=occ, total_occurrences=total)
            raise TextNotFoundError(text)
        if occurrence is None:
            total = self.count_matches(text)
            if total > 1:
                raise AmbiguousTextError(text, total_occurrences=total)
        return match

    def _find_across_boundaries_located(self, text: str, occurrence: int = 0) -> _LocatedMatch | None:
        """Find the nth occurrence of text across element boundaries.

        Searches across all paragraphs using text maps, keeping paragraph
        identity so callers can build hash-anchored refs.

        Returns:
            A _LocatedMatch, or None if not found.
        """
        current_occurrence = 0
        for idx, paragraph in enumerate(body_paragraphs(self.editor.dom), start=1):
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

    def find_text(self, text: str, occurrence: int = 0, paragraph: str | None = None) -> SearchResult | None:
        """Find the nth occurrence of text, as a public SearchResult.

        Searches across element boundaries. With ``paragraph=None``,
        ``occurrence`` counts matches document-wide (0 = first); with a
        paragraph reference, the search is scoped to that paragraph and
        ``occurrence`` counts within it. Returns None if not found.

        Raises:
            ValueError: If ``text`` is not a non-empty string, ``occurrence``
                is not a non-negative integer (None included — the default is
                0, not None), or ``paragraph`` is malformed.
            ParagraphIndexError: If ``paragraph``'s index is out of range.
            HashMismatchError: If ``paragraph``'s hash is stale.
        """
        if not isinstance(text, str) or not text:
            raise ValueError(f"find_text(): search text must be a non-empty string, got {text!r}")
        _require_valid_occurrence(occurrence, "find_text(): ", allow_none=False)

        if paragraph is not None:
            results = self.find_all(text, paragraph=paragraph)
            # 0 <= guard: a bare results[occurrence] would let a negative
            # index silently return a match from the end.
            return results[occurrence] if 0 <= occurrence < len(results) else None

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
            paragraph_index=located.paragraph_index,
        )

    def find_all(self, text: str, paragraph: str | None = None) -> list[SearchResult]:
        """Enumerate every match of ``text`` as a list of SearchResults.

        One call replaces the N+1 ``find_text`` probes needed to enumerate N
        hits. Each result's ``paragraph_ref``/``paragraph_occurrence`` plug
        directly into a follow-up edit's ``paragraph=``/``occurrence=``.

        Args:
            text: Text to search for (must be non-empty).
            paragraph: Optional paragraph reference (e.g., "P2#f3c1") to scope
                the search. None searches the whole document.

        Returns:
            SearchResults in document order; ``[]`` when nothing matches (it
            is an enumeration API, not a lookup — no-match is not an error).

        Raises:
            ValueError: If ``text`` is not a non-empty string, or
                ``paragraph`` is malformed.
            ParagraphIndexError: If ``paragraph``'s index is out of range.
            HashMismatchError: If ``paragraph``'s hash is stale.
        """
        if not isinstance(text, str) or not text:
            raise ValueError(f"find_all(): search text must be a non-empty string, got {text!r}")

        if paragraph is not None:
            ref = ParagraphRef.parse(paragraph)
            paragraphs = [(ref.index, self._resolve_paragraph(ref))]
        else:
            paragraphs = list(enumerate(body_paragraphs(self.editor.dom), start=1))

        results: list[SearchResult] = []
        for idx, p in paragraphs:
            text_map = build_text_map(p)
            paragraph_ref: str | None = None
            local_occ = 0
            while (match := find_in_text_map(text_map, text, local_occ)) is not None:
                if paragraph_ref is None:
                    paragraph_ref = f"P{idx}#{compute_paragraph_hash(p)}"
                results.append(
                    SearchResult(
                        start=match.start,
                        end=match.end,
                        text=match.text,
                        paragraph_ref=paragraph_ref,
                        paragraph_occurrence=local_occ,
                        spans_revision=match.spans_boundary,
                        paragraph_index=idx,
                    )
                )
                local_occ += 1
        return results
