"""Straight and curly quotes match each other in every search (ROADMAP.md #96).

Word autoformats ``'`` and ``"`` into ``‘ ’`` and ``“ ”`` as you type, so an
agent that types the straight form must still find the text it can see.
"""

import dataclasses
import re
from pathlib import Path

import pytest
from conftest import NS, find_ref, match_for, replace_document_xml

from docx_editor import AmbiguousTextError, Document, EditOperation
from docx_editor.track_changes.diff import _trim_replace_affixes
from docx_editor.xml_editor import fold_quotes

AUTHOR = "Test Author"
OTHER = "Reviewer A"
DATE = "2024-01-01T00:00:00Z"

CURLY = "The “Term” means we don’t pay."


def _open(simple_docx: Path, tmp_path: Path, body_xml: str) -> Document:
    path = tmp_path / "quotes.docx"
    replace_document_xml(
        simple_docx, path, f'<?xml version="1.0"?><w:document {NS}><w:body>{body_xml}</w:body></w:document>'
    )
    return Document.open(path, author=AUTHOR)


def _p(*runs: str) -> str:
    return "<w:p>" + "".join(f'<w:r><w:t xml:space="preserve">{r}</w:t></w:r>' for r in runs) + "</w:p>"


def _revs(doc: Document) -> list[tuple[str, str]]:
    return sorted((r.type, r.text) for r in doc.list_revisions())


def _xml(doc: Document) -> str:
    return doc._document_editor.dom.toxml()


@pytest.fixture
def curly_doc(simple_docx, tmp_path):
    with _open(simple_docx, tmp_path, _p(CURLY) + _p('Plain "straight" text here.')) as doc:
        yield doc


class TestFoldQuotes:
    def test_only_the_four_quotes_change_and_length_is_kept(self):
        src = "‘a’ “b” ' \" ‚„′«"
        folded = fold_quotes(src)
        assert len(folded) == len(src)
        assert folded == "'a' \"b\" ' \" ‚„′«"


class TestFind:
    def test_straight_search_finds_curly_text(self, curly_doc):
        m = match_for(curly_doc, '"Term"')
        assert m.text == "“Term”"
        assert (m.start, m.end) == (4, 10)
        assert match_for(curly_doc, "don't").text == "don’t"

    def test_curly_search_finds_straight_text(self, curly_doc):
        m = match_for(curly_doc, "“straight”")
        assert m.text == '"straight"'

    def test_find_all_and_count_mix_styles(self, curly_doc):
        results = curly_doc.find_all('"')
        assert [r.text for r in results] == ["“", "”", '"', '"']
        assert curly_doc.count_matches("”") == 4

    def test_paragraph_scoped_occurrence_counts_both_styles(self, simple_docx, tmp_path):
        with _open(simple_docx, tmp_path, _p("say ‘hi’ then 'hi' again")) as doc:
            ref = find_ref(doc, "again")
            assert [r.paragraph_occurrence for r in doc.find_all("'hi'", paragraph=ref)] == [0, 1]
            assert match_for(doc, "'hi'", occurrence=1).text == "'hi'"


class TestEdits:
    def test_replace_deletes_the_documents_characters_and_inserts_as_passed(self, curly_doc):
        curly_doc.replace(match_for(curly_doc, '"Term"'), '"Word"')
        # The quotes belong to the word token, so the whole word is revised.
        assert _revs(curly_doc) == [("deletion", "“Term”"), ("insertion", '"Word"')]

    def test_replace_whole_quoted_span_records_curly_deletion(self, curly_doc):
        curly_doc.replace(match_for(curly_doc, "we don't pay."), "they pay.")
        dels = [r.text for r in curly_doc.list_revisions() if r.type == "deletion"]
        assert dels == ["we don’t"]

    def test_affix_trim_leaves_quote_mismatched_word_untouched(self, curly_doc):
        # (a) "don't" typed the same on both sides: keep the document's "don’t".
        curly_doc.replace("don't pay", "don't charge", paragraph=find_ref(curly_doc, "Term"))
        assert _revs(curly_doc) == [("deletion", "pay"), ("insertion", "charge")]
        assert "don’t charge" in curly_doc.get_visible_text()

    def test_interior_word_is_rewritten_as_passed(self, simple_docx, tmp_path):
        # Only words at the ends of the span are trimmed: "don't" sits between two
        # changed words, so it is deleted and reinserted with the caller's quote.
        with _open(simple_docx, tmp_path, _p("say old don’t pay now")) as doc:
            doc.replace("old don't pay", "new don't charge", paragraph=find_ref(doc, "say"))
            assert _revs(doc) == [("deletion", "old don’t pay"), ("insertion", "new don't charge")]
            assert doc.get_visible_text() == "say new don't charge now"

    def test_straight_to_curly_style_only_replace_is_tracked(self, curly_doc):
        # (b) The new token differs from both the search and the document token.
        curly_doc.replace('"straight"', "“straight”", paragraph=find_ref(curly_doc, "Plain"))
        assert _revs(curly_doc) == [("deletion", '"straight"'), ("insertion", "“straight”")]

    def test_curly_to_straight_style_only_replace_is_tracked(self, curly_doc):
        # (b) Deleted text is the document's characters, inserted text as passed.
        curly_doc.replace('“Term"', '"Term"', paragraph=find_ref(curly_doc, "Term"))
        assert _revs(curly_doc) == [("deletion", "“Term”"), ("insertion", '"Term"')]

    @pytest.mark.parametrize(
        ("find", "new"),
        [
            ("“Term”", "“Term”"),  # exact, as on main
            ('"Term"', '"Term"'),  # same on both sides: nothing asked to change
            ('"Term"', "“Term”"),  # already what the document has
        ],
    )
    def test_replace_with_what_is_there_is_a_noop(self, curly_doc, find, new):
        # (c)
        assert curly_doc.replace(find, new, paragraph=find_ref(curly_doc, "Term")).group_id is None
        assert curly_doc.list_revisions() == []

    def test_mixed_word_change_and_style_change(self, curly_doc):
        # (d) "Plain" unchanged, '"straight"' changes style, "text" changes word.
        curly_doc.replace('Plain "straight" text', "Plain “straight” words", paragraph=find_ref(curly_doc, "Plain"))
        assert _revs(curly_doc) == [("deletion", '"straight" text'), ("insertion", "“straight” words")]
        assert "Plain “straight” words here." in curly_doc.get_visible_text()

    def test_inserted_text_is_written_as_passed(self, curly_doc):
        curly_doc.replace(match_for(curly_doc, "pay"), "say 'no'")
        assert "say 'no'" in curly_doc.get_visible_text()

    def test_delete_and_inserts_with_mismatched_anchor(self, curly_doc):
        curly_doc.delete(match_for(curly_doc, ' "Term"'))
        curly_doc.insert_after(match_for(curly_doc, "don't"), " ever")
        curly_doc.insert_before(match_for(curly_doc, "“straight”"), "very ")
        text = curly_doc.get_visible_text()
        assert "The means we don’t ever pay." in text
        assert 'Plain very "straight" text' in text

    def test_match_across_smart_quote_run_split(self, simple_docx, tmp_path):
        body = (
            '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">see </w:t></w:r>'
            "<w:r><w:t>“</w:t></w:r><w:r><w:t>Annex</w:t></w:r><w:r><w:t>”</w:t></w:r>"
            '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> now</w:t></w:r></w:p>'
        )
        with _open(simple_docx, tmp_path, body) as doc:
            doc.replace(match_for(doc, '“Annex"'), "“Schedule”")
            assert "see “Schedule” now" in doc.get_visible_text()
            xml = _xml(doc)
            assert "<w:b/>" in xml and "<w:i/>" in xml
            doc.reject_all()
            assert "see “Annex” now" in doc.get_visible_text()

    def test_comment_anchor_with_straight_quotes(self, curly_doc):
        curly_doc.add_comment(match_for(curly_doc, '"Term"'), "define this")
        p = curly_doc._document_editor.dom.getElementsByTagName("w:p")[0]
        kids = [n for n in p.childNodes if n.nodeType == n.ELEMENT_NODE]
        tags = [n.tagName for n in kids]
        start, end = tags.index("w:commentRangeStart"), tags.index("w:commentRangeEnd")
        inside = "".join(t.firstChild.data for k in kids[start:end] for t in k.getElementsByTagName("w:t"))
        assert inside == "“Term”"

    def test_ambiguous_error_lists_both_styles(self, simple_docx, tmp_path):
        with _open(simple_docx, tmp_path, _p("say ‘hi’ then 'hi' again")) as doc:
            with pytest.raises(AmbiguousTextError) as exc:
                doc.replace("'hi'", "'yo'", paragraph=find_ref(doc, "again"))
            err = exc.value
            assert err.total_occurrences == 2
            assert [(m.paragraph_occurrence, m.start) for m in err.matches] == [(0, 4), (1, 14)]

    @pytest.mark.parametrize("dry_run", [True, False])
    def test_batch_edit_with_mismatched_targets(self, curly_doc, dry_run):
        ref = find_ref(curly_doc, "Term")
        ops = [EditOperation.replace('"Term"', '"Word"', paragraph=ref), EditOperation.delete(" don't", paragraph=ref)]
        results = curly_doc.batch_edit(ops, dry_run=dry_run)
        assert len(results) == 2
        if dry_run:
            assert all(r.valid for r in results)
        else:
            assert 'The "Word" means we pay.' in curly_doc.get_visible_text()


class TestSearchResultQuery:
    """A SearchResult carries the typed query, so edits given it behave like the string form."""

    def _edited_xml(self, simple_docx, tmp_path, name, edit) -> str:
        sub = tmp_path / name
        sub.mkdir()
        with _open(simple_docx, sub, _p(CURLY)) as doc:
            edit(doc)
            # Edit timestamps and the per-session rsid differ by construction;
            # everything else must match.
            return re.sub(r'(w:date|w16du:dateUtc|w:rsid\w*)="[^"]*"', "", _xml(doc))

    def test_find_text_and_find_all_record_the_query(self, curly_doc):
        assert match_for(curly_doc, "don't pay").query == "don't pay"
        assert match_for(curly_doc, "don't pay", paragraph=find_ref(curly_doc, "Term")).query == "don't pay"
        assert [r.query for r in curly_doc.find_all('"Term"')] == ['"Term"']

    def test_replace_with_search_result_matches_string_form(self, simple_docx, tmp_path):
        def by_string(doc):
            doc.replace("don't pay", "don't charge", paragraph=find_ref(doc, "Term"))

        def by_result(doc):
            doc.replace(match_for(doc, "don't pay"), "don't charge")
            assert _revs(doc) == [("deletion", "pay"), ("insertion", "charge")]

        assert self._edited_xml(simple_docx, tmp_path, "s", by_string) == self._edited_xml(
            simple_docx, tmp_path, "r", by_result
        )

    def test_batch_edit_with_search_result_matches_string_form(self, simple_docx, tmp_path):
        def by_string(doc):
            doc.batch_edit([EditOperation.replace("don't pay", "don't charge", paragraph=find_ref(doc, "Term"))])

        def by_result(doc):
            doc.batch_edit([EditOperation.replace(match_for(doc, "don't pay"), "don't charge")])

        assert self._edited_xml(simple_docx, tmp_path, "s", by_string) == self._edited_xml(
            simple_docx, tmp_path, "r", by_result
        )

    def test_result_without_query_falls_back_to_document_text(self, curly_doc):
        m = dataclasses.replace(match_for(curly_doc, "don't pay"), query=None)
        assert m.text == "don’t pay"
        curly_doc.replace(m, "don't charge")
        # Searched as "don’t pay": the straight "don't" differs from both, so it is revised.
        assert _revs(curly_doc) == [("deletion", "don’t pay"), ("insertion", "don't charge")]


class TestRevisionOccurrence:
    def test_revision_occurrence_counts_folded_matches(self, simple_docx, tmp_path):
        body = (
            '<w:p><w:r><w:t xml:space="preserve">say ‘hi’ then </w:t></w:r>'
            f'<w:ins w:id="1" w:author="{OTHER}" w:date="{DATE}"><w:r><w:t>\'hi\'</w:t></w:r></w:ins>'
            '<w:r><w:t xml:space="preserve"> again</w:t></w:r></w:p>'
        )
        with _open(simple_docx, tmp_path, body) as doc:
            (rev,) = doc.list_revisions()
            assert rev.occurrence == 1
            assert rev.paragraph_ref is not None
            m = doc.find_all(rev.text, paragraph=rev.paragraph_ref)[rev.occurrence]
            assert m.text == "'hi'" and m.spans_revision is False and m.start == 14


@pytest.mark.parametrize(
    ("document", "search", "new", "expected"),
    [
        ("don’t pay", "don't pay", "don't charge", (6, 0)),  # new == search token
        ("don’t pay", "don't pay", "don’t charge", (6, 0)),  # new == document token
        ("don't pay", "don't pay", "don’t pay", (0, 4)),  # differs from both: a change
    ],
)
def test_trim_affixes_token_rule(document, search, new, expected):
    assert _trim_replace_affixes(document, search, new) == expected
