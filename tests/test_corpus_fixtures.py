"""Moves and paragraph-property changes against the real corpus files.

The two ``benchmarks/corpus`` files that carry ``w:moveFrom``/``w:moveTo`` and
``w:pPrChange`` in the wild (see ``benchmarks/corpus/README.md``). Corpus
files are never committed (provenance policy) — build them locally with
``uv run python benchmarks/corpus/build_corpus.py``; until then these skip.
``tests/test_move_and_ppr_change_resolution.py`` carries the same assertions
over hand-authored mirrors, so the contract is tested either way; the weekly
corpus workflow additionally asserts that ``accept_all()`` leaves no
resolvable revision element in any saved file.
"""

import warnings
import zipfile
from pathlib import Path

import defusedxml.minidom
import pytest

from docx_editor import Document, UnhandledRevisionWarning
from docx_editor.track_changes import count_revision_elements

CORPUS_FILES = Path(__file__).resolve().parents[1] / "benchmarks" / "corpus" / "files"
TABLE_MOVE = CORPUS_FILES / "locore_TC-table-DnD-move.docx"
UNKNOWN_STYLE = CORPUS_FILES / "locore_UnknownStyleInRedline.docx"

needs_corpus = pytest.mark.skipif(
    not (TABLE_MOVE.exists() and UNKNOWN_STYLE.exists()),
    reason="corpus not built: run `uv run python benchmarks/corpus/build_corpus.py`",
)

VISIBLE_MOVED = "\n\n\n\nText\nA1\nB1\nA2\nB2\n"
VISIBLE_ORIGINAL = "A1\nB1\nA2\nB2\nText\n\n\n\n\n"


def _saved_census(path: Path) -> dict[str, int]:
    with zipfile.ZipFile(path) as z:
        dom = defusedxml.minidom.parseString(z.read("word/document.xml"))
    return count_revision_elements(dom).by_tag


@pytest.fixture(autouse=True)
def _no_unhandled_warnings():
    with warnings.catch_warnings():
        warnings.simplefilter("error", UnhandledRevisionWarning)
        yield


@pytest.fixture
def corpus_copy(tmp_path):
    """Open a private copy so the corpus file itself is never touched."""

    def _open(src: Path) -> Document:
        dest = tmp_path / src.name
        dest.write_bytes(src.read_bytes())
        return Document.open(dest, author="Tester", force_recreate=True)

    return _open


@needs_corpus
class TestTableDragAndDropMove:
    def test_listing(self, corpus_copy):
        with corpus_copy(TABLE_MOVE) as doc:
            revs = doc.list_revisions()
            assert len(revs) == 16
            assert [r.type for r in revs] == ["move_from"] * 8 + ["move_to"] * 8
            assert [r.text for r in revs] == ["", "A1", "", "B1", "", "A2", "", "B2"] * 2
            assert len({r.changeset_id for r in revs}) == 1
            assert doc.list_unhandled_revisions() == []
            assert doc.get_visible_text() == VISIBLE_MOVED

    def test_accept_all(self, corpus_copy, tmp_path):
        out = tmp_path / "accepted.docx"
        with corpus_copy(TABLE_MOVE) as doc:
            result = doc.accept_all()
            assert result == 16
            assert result.unhandled == 0
            assert doc.get_visible_text() == VISIBLE_MOVED
            doc.save(out)
        assert _saved_census(out) == {}

    def test_reject_all(self, corpus_copy, tmp_path):
        out = tmp_path / "rejected.docx"
        with corpus_copy(TABLE_MOVE) as doc:
            result = doc.reject_all()
            assert result == 16
            assert result.unhandled == 0
            assert doc.get_visible_text() == VISIBLE_ORIGINAL
            doc.save(out)
        assert _saved_census(out) == {}


@needs_corpus
class TestUnknownStyleInRedline:
    def test_listing(self, corpus_copy):
        with corpus_copy(UNKNOWN_STYLE) as doc:
            revs = doc.list_revisions()
            assert [(r.id, r.type, r.text) for r in revs] == [(0, "property_change", ""), (0, "property_change", "")]
            assert doc.list_unhandled_revisions() == []

    def test_accept_all_keeps_current_styles(self, corpus_copy, tmp_path):
        out = tmp_path / "accepted.docx"
        with corpus_copy(UNKNOWN_STYLE) as doc:
            styles_before = [doc.get_paragraph(i).style for i in (1, 3)]
            result = doc.accept_all()
            assert result == 2
            assert result.unhandled == 0
            assert [doc.get_paragraph(i).style for i in (1, 3)] == styles_before == ["Cmsor1", "Cmsor3"]
            doc.save(out)
        assert _saved_census(out) == {}

    def test_reject_all_restores_the_recorded_styles(self, corpus_copy, tmp_path):
        out = tmp_path / "rejected.docx"
        with corpus_copy(UNKNOWN_STYLE) as doc:
            result = doc.reject_all()
            assert result == 2
            assert result.unhandled == 0
            # P1's record is self-closing (previously no properties); P3's
            # records a style id styles.xml does not define.
            assert doc.get_paragraph(1).style is None
            assert doc.get_paragraph(3).style == "UnknownStyle"
            doc.save(out)
        assert _saved_census(out) == {}
