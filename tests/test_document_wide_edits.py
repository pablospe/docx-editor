"""Edits without ``paragraph=``, ``find_all(expected_count=)`` and ``replace_all`` (ROADMAP.md #95).

A plain-string target with no ``paragraph`` is located document-wide: one
match edits, several raise AmbiguousTextError with a document-wide pick-list,
none raises TextNotFoundError — and nothing is written unless the edit applies.
"""

from pathlib import Path

import pytest
from conftest import NS, replace_document_xml

from docx_editor import (
    AmbiguousTextError,
    Document,
    MatchCountError,
    TextNotFoundError,
)

AUTHOR = "Test Author"


def _p(*runs: str) -> str:
    return "<w:p>" + "".join(f'<w:r><w:t xml:space="preserve">{r}</w:t></w:r>' for r in runs) + "</w:p>"


def _open(simple_docx: Path, tmp_path: Path, *paragraphs: str) -> Document:
    path = tmp_path / "wide.docx"
    body = "".join(paragraphs)
    replace_document_xml(
        simple_docx, path, f'<?xml version="1.0"?><w:document {NS}><w:body>{body}</w:body></w:document>'
    )
    return Document.open(path, author=AUTHOR)


def _xml(doc: Document) -> str:
    return doc._document_editor.dom.toxml()


@pytest.fixture
def doc(simple_docx, tmp_path):
    with _open(
        simple_docx,
        tmp_path,
        _p("Payment is due in 30 days."),
        _p("The unique clause stays."),
        _p("Notice within 30 days, or 30 days after receipt."),
    ) as d:
        yield d


class TestUniqueMatch:
    def test_replace_edits_and_returns_the_touched_paragraphs_post_edit_ref(self, doc):
        result = doc.replace("unique clause", "only clause")
        assert str(result) == doc.list_paragraphs()[1].split("|")[0]
        assert str(result).startswith("P2#")
        assert result.group_id is not None and result.revision_ids
        # The returned ref is live: a follow-up scoped edit uses it directly.
        doc.replace("stays", "remains", paragraph=str(result))
        assert "The only clause remains." in doc.get_visible_text()

    def test_delete_insert_after_and_insert_before(self, doc):
        doc.delete("unique ")
        doc.insert_after("clause", " here")
        doc.insert_before("Payment", "Final ")
        text = doc.get_visible_text()
        assert "The clause here stays." in text
        assert "Final Payment is due" in text

    def test_cross_run_match(self, simple_docx, tmp_path):
        with _open(simple_docx, tmp_path, _p("Split ac", "ross runs."), _p("Other.")) as d:
            result = d.replace("across", "over")
            assert str(result).startswith("P1#")
            assert d.get_visible_text().startswith("Split over runs.")

    def test_quote_insensitive_match(self, simple_docx, tmp_path):
        with _open(simple_docx, tmp_path, _p("First."), _p("The “Term” applies.")) as d:
            result = d.replace('"Term"', '"Word"')
            assert str(result).startswith("P2#")
            assert '"Word"' in d.get_visible_text()

    def test_note_is_anchored(self, doc):
        result = doc.replace("unique clause", "only clause", note="clarity")
        assert result.comment_id is not None
        assert [c.text for c in doc.list_comments()] == ["clarity"]

    def test_newline_split_reports_both_halves(self, doc):
        result = doc.insert_after("unique clause", "\nNew paragraph")
        assert len(result.refs) == 2
        assert [r.split("#")[0] for r in result.refs] == ["P2", "P3"]


class TestRefusals:
    def test_ambiguous_lists_document_wide_occurrences_and_writes_nothing(self, doc):
        before = _xml(doc)
        with pytest.raises(AmbiguousTextError) as exc:
            doc.replace("30 days", "60 days")
        err = exc.value
        assert err.paragraph_ref is None
        assert err.total_occurrences == 3
        assert [m.occurrence for m in err.matches] == [0, 1, 2]
        assert [m.paragraph_ref.split("#")[0] for m in err.matches] == ["P1", "P3", "P3"]
        assert [m.paragraph_occurrence for m in err.matches] == [0, 0, 1]
        assert _xml(doc) == before

    def test_pick_list_occurrence_round_trips(self, doc):
        with pytest.raises(AmbiguousTextError) as exc:
            doc.delete("30 days")
        pick = exc.value.matches[2]
        result = doc.replace("30 days", "60 days", occurrence=pick.occurrence)
        assert str(result).startswith("P3#")
        assert doc.list_paragraphs()[2].endswith("Notice within 30 days, or 60 days after receipt.")

    def test_occurrence_counts_across_paragraphs(self, doc):
        result = doc.insert_before("30 days", "about ", occurrence=1)
        assert str(result).startswith("P3#")
        assert "Notice within about 30 days, or 30 days" in doc.get_visible_text()

    def test_not_found(self, doc):
        before = _xml(doc)
        with pytest.raises(TextNotFoundError) as exc:
            doc.delete("absent text")
        assert exc.value.next == "list_paragraphs()"
        assert _xml(doc) == before

    def test_occurrence_out_of_range(self, doc):
        with pytest.raises(TextNotFoundError) as exc:
            doc.replace("30 days", "x", occurrence=3)
        assert exc.value.total_occurrences == 3
        assert doc.list_revisions() == []

    def test_bad_note_is_refused_before_locating(self, doc):
        with pytest.raises(ValueError, match="note"):
            doc.replace("absent text", "x", note="")
        assert doc.list_revisions() == []

    def test_control_character_in_target_is_a_value_error(self, doc):
        with pytest.raises(ValueError, match="control character"):
            doc.replace("30\x01days", "x")

    @pytest.mark.parametrize(
        ("call", "message"),
        [
            # The target is absent: each must still raise the bad-argument
            # ValueError, not TextNotFoundError from the document-wide search.
            (lambda d: d.replace("absent text", 5), "replace_with"),
            (lambda d: d.insert_after("absent text", 5), "'text' must be a string"),
            (lambda d: d.insert_before("absent text", "a\x01b"), "control character"),
            (lambda d: d.delete("30 days\t"), "tab"),
        ],
    )
    def test_bad_input_is_refused_before_locating(self, doc, call, message):
        with pytest.raises(ValueError, match=message):
            call(doc)
        assert doc.list_revisions() == []


class TestFindAllExpectedCount:
    def test_matching_count_returns_results(self, doc):
        assert len(doc.find_all("30 days", expected_count=3)) == 3
        assert doc.find_all("absent", expected_count=0) == []

    @pytest.mark.parametrize("expected", [0, 2, 4])
    def test_other_count_raises_with_the_matches(self, doc, expected):
        with pytest.raises(MatchCountError) as exc:
            doc.find_all("30 days", expected_count=expected)
        err = exc.value
        assert (err.expected_count, err.total_occurrences) == (expected, 3)
        assert [m.paragraph_ref.split("#")[0] for m in err.matches] == ["P1", "P3", "P3"]
        assert err.next == "find_all('30 days')"
        assert "matches 3 times in the document, expected" in str(err)

    def test_too_few_with_no_match(self, doc):
        with pytest.raises(MatchCountError) as exc:
            doc.find_all("absent", expected_count=1)
        assert exc.value.total_occurrences == 0
        assert exc.value.matches == []

    def test_scoped_count(self, doc):
        ref = doc.list_paragraphs()[2].split("|")[0]
        assert len(doc.find_all("30 days", ref, expected_count=2)) == 2
        with pytest.raises(MatchCountError) as exc:
            doc.find_all("30 days", ref, expected_count=1)
        assert exc.value.paragraph_ref == ref
        assert [m.occurrence for m in exc.value.matches] == [0, 1]

    @pytest.mark.parametrize("bad", [-1, True, 1.0, "2"])
    def test_bad_expected_count(self, doc, bad):
        with pytest.raises(ValueError, match="expected_count"):
            doc.find_all("30 days", expected_count=bad)


class TestReplaceAll:
    def test_replaces_every_match_in_one_changeset(self, doc):
        results = doc.replace_all("30 days", "60 days", expected_count=3)
        assert [r.split("#")[0] for r in results] == ["P1", "P3", "P3"]
        assert len({r.changeset_id for r in results}) == 1
        assert len({r.group_id for r in results}) == 3
        assert results[1] == results[2] == doc.list_paragraphs()[2].split("|")[0]
        text = doc.get_visible_text()
        assert "30 days" not in text
        assert text.count("60 days") == 3

    def test_count_mismatch_writes_nothing(self, doc):
        before = _xml(doc)
        with pytest.raises(MatchCountError):
            doc.replace_all("30 days", "60 days", expected_count=2)
        assert _xml(doc) == before

    def test_zero_expected_and_absent(self, doc):
        assert doc.replace_all("absent", "x", expected_count=0) == []
        assert doc.list_revisions() == []

    def test_overlapping_needle_is_refused(self, simple_docx, tmp_path):
        with _open(simple_docx, tmp_path, _p("aaaa")) as d:
            with pytest.raises(ValueError, match="overlap"):
                d.replace_all("aa", "b", expected_count=3)
            assert d.list_revisions() == []

    def test_note_is_shared(self, doc):
        results = doc.replace_all("30 days", "60 days", expected_count=3, note="new term")
        assert len({r.comment_id for r in results}) == 1
        assert [c.text for c in doc.list_comments()] == ["new term"]

    @pytest.mark.parametrize(
        ("find", "new", "message"),
        [("a\tb", "x", "tab"), ("30 days", 5, "replace_with"), ("", "x", "non-empty")],
    )
    def test_bad_input_is_refused_before_the_count(self, doc, find, new, message):
        with pytest.raises(ValueError, match=message):
            doc.replace_all(find, new, expected_count=99)
        assert doc.list_revisions() == []

    @pytest.mark.parametrize("bad", [None, -1, True, 1.0])
    def test_expected_count_is_required_and_validated(self, doc, bad):
        # None would otherwise reach find_all as "no guard" and replace every match.
        with pytest.raises(ValueError, match=r"replace_all\(\): expected_count"):
            doc.replace_all("30 days", "60 days", expected_count=bad)
        assert doc.list_revisions() == []
