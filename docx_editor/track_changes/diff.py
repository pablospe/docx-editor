"""Pure-text tokenising and diffing used by paragraph rewrites."""

import difflib
import re


def _tokenize_words(text: str) -> list[str]:
    """Split text into alternating word and whitespace tokens."""
    return re.findall(r"\S+|\s+", text)


def _diff_hunks(old_text: str, new_text: str) -> list[tuple[str, int, int, str]]:
    """Word-level diff of ``old_text`` → ``new_text`` as edit hunks.

    Each hunk is ``(tag, old_char_start, old_char_end, new_fragment)`` with the
    character range in ``old_text``; ``equal`` opcodes are dropped. The texts
    are diffed segment by segment between their tab marks (the caller has
    checked both hold the same number of ``"\\t"``), so no hunk ever spans a
    tab: an insert at a segment's end lands right before the following tab,
    one at a segment's start right after the preceding tab.
    """
    hunks: list[tuple[str, int, int, str]] = []
    base = 0
    for old_seg, new_seg in zip(old_text.split("\t"), new_text.split("\t"), strict=True):
        old_tokens = _tokenize_words(old_seg)
        new_tokens = _tokenize_words(new_seg)
        offsets = []
        pos = 0
        for tok in old_tokens:
            offsets.append(pos)
            pos += len(tok)
        for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, old_tokens, new_tokens).get_opcodes():
            if tag == "equal":
                continue
            start = offsets[i1] if i1 < len(old_tokens) else len(old_seg)
            end = offsets[i2 - 1] + len(old_tokens[i2 - 1]) if i2 > 0 else start
            hunks.append((tag, base + start, base + end, "".join(new_tokens[j1:j2])))
        base += len(old_seg) + 1
    return hunks


def _trim_replace_affixes(document: str, search: str, replace_with: str) -> tuple[int, int]:
    """Compute the character lengths of the word-level common prefix and
    suffix that ``replace_with`` leaves unchanged in the matched text.

    ``document`` is the matched text as the document spells it and ``search``
    the text the caller searched for. They are equal up to quote style
    (:func:`~docx_editor.xml_editor.fold_quotes`) and so have the same length
    and the same word boundaries. A token of ``replace_with`` leaves its word
    unchanged when it equals that word *exactly* as searched (the caller typed
    the same thing on both sides, so the document's characters are kept) or
    exactly as the document has it (already what the caller wants). A token
    that differs from both is a change, even when it differs only in quote
    style — so a style-only replace such as ``'"Term"'`` → ``'“Term”'`` is a
    real tracked change, not a no-op (ROADMAP.md #96).

    Trimming is word-granular (tokens from :func:`_tokenize_words`) so a
    replace only revises whole changed words, matching the diff granularity
    of ``rewrite_paragraph``. The suffix scan is bounded by each side's
    remainder after the prefix, so a token shared at both ends is never
    consumed twice.

    One exception is character-granular: when what remains on *both* sides after
    the word-level trim is nothing but whitespace, the shared characters of that
    whitespace are trimmed too. A single-to-double space edit then becomes a pure
    one-space insertion instead of a ``del " "`` + ``ins "  "`` pair that renders
    as an invisible, unreviewable redline (ISSUES.md #60). The gate keeps this
    away from word redlines, where character trimming would be actively harmful:
    ``"30 days"`` → ``"60 days"`` must stay one whole-word replacement, not
    ``del "3"`` + ``ins "6"``.

    Returns:
        ``(prefix_len, suffix_len)`` in characters.
    """
    f_toks = _tokenize_words(document)
    s_toks = _tokenize_words(search)
    r_toks = _tokenize_words(replace_with)

    def unchanged(k: int, tok: str) -> bool:
        return tok == s_toks[k] or tok == f_toks[k]

    i = 0
    while i < len(f_toks) and i < len(r_toks) and unchanged(i, r_toks[i]):
        i += 1
    j = 0
    while j < len(f_toks) - i and j < len(r_toks) - i and unchanged(-(j + 1), r_toks[-(j + 1)]):
        j += 1

    prefix_len = sum(len(tok) for tok in f_toks[:i])
    suffix_len = sum(len(tok) for tok in f_toks[len(f_toks) - j :])

    f_rest = document[prefix_len : len(document) - suffix_len]
    r_rest = replace_with[prefix_len : len(replace_with) - suffix_len]
    if f_rest and r_rest and not f_rest.strip() and not r_rest.strip():
        prefix_len += _common_prefix_len(f_rest, r_rest)
        # Re-slice: the prefix just grew, and the shared characters it consumed
        # must not be counted again from the other end.
        f_rest = document[prefix_len : len(document) - suffix_len]
        r_rest = replace_with[prefix_len : len(replace_with) - suffix_len]
        suffix_len += _common_prefix_len(f_rest[::-1], r_rest[::-1])
    return prefix_len, suffix_len


def _common_prefix_len(a: str, b: str) -> int:
    """Number of leading characters ``a`` and ``b`` share."""
    n = 0
    while n < len(a) and n < len(b) and a[n] == b[n]:
        n += 1
    return n
