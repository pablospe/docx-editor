"""Tests against real-world fixtures containing revisions authored outside this library.

Fixtures (see ``tests/test_data/THIRD_PARTY_NOTICES``):

* ``OXML_TrackChanges_Test.docx`` — 8 foreign revisions (4 ``w:ins`` + 4 ``w:del``)
  from 6 distinct authors, covering a del+ins replacement pair (colour→color),
  an em-dash insertion adjoining smart-quote runs, and identical text
  (``"DRAFT "``) inserted by one author and deleted by another in the same
  paragraph.
* ``tricky-track-changes.docx`` — no revisions, but heavily fragmented runs
  (``w:proofErr``, bookmarks, ``w:lastRenderedPageBreak``, formatting splits)
  that exercise cross-run text search the way real Word output does.

Each fixture paragraph ends with a bracketed note run describing the expected
accept/reject outcomes (approximately — some notes normalize spacing or drop
punctuation), so text assertions go through ``_body_text`` (which strips the
notes) — matching against the full visible text would let a note satisfy an
assertion about the body. The double spaces in several expected strings are
load-bearing: they come from adjacent runs the revisions never touch (e.g. the
separator run between the colour→color ``w:del``/``w:ins`` pair), so a "fix"
that collapses them would be wrong.
"""

import shutil
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import pytest
from conftest import find_ref, replace_document_xml

from docx_editor import Document, Revision

# Ground truth for OXML_TrackChanges_Test.docx. Paragraph indexes account for
# the fixture's two empty <w:p/> elements (before the title and before the
# em-dash paragraph); hashes are CRC32 of each paragraph's visible text
# including its trailing bracketed note run. None of the fixture's revisions
# nest, so nested_under/contains_ids keep their defaults. Ids 1007/1008 carry
# identical text in the same paragraph — type + occurrence (in each type's own
# view) is what tells them apart. Every revision differs from its document-
# order neighbor by paragraph, author, or date, so group reconstruction
# infers eight singleton groups, numbered 1-8 in document order. Every
# revision also has a distinct (author, date) pair, so the changeset tier
# partitions them into eight singleton inferred changesets — changeset_id
# equals group_id here (1-8), each changeset_source "inferred".
EXPECTED_REVISIONS = [
    Revision(
        id=1001,
        type="insertion",
        author="Test Author",
        date=datetime(2026, 1, 29, 16, 55, tzinfo=timezone.utc),
        text="inserted",
        paragraph_ref="P3#e6b4",
        occurrence=0,
        group_id=1,
        group_source="inferred",
        changeset_id=1,
        changeset_source="inferred",
    ),
    Revision(
        id=1002,
        type="deletion",
        author="Test Author",
        date=datetime(2026, 1, 29, 16, 56, tzinfo=timezone.utc),
        text="old ",
        paragraph_ref="P4#1750",
        occurrence=0,
        group_id=2,
        group_source="inferred",
        changeset_id=2,
        changeset_source="inferred",
    ),
    Revision(
        id=1003,
        type="deletion",
        author="Editor A",
        date=datetime(2026, 1, 29, 16, 57, tzinfo=timezone.utc),
        text="colour",
        paragraph_ref="P5#a192",
        occurrence=0,
        group_id=3,
        group_source="inferred",
        changeset_id=3,
        changeset_source="inferred",
    ),
    Revision(
        id=1004,
        type="insertion",
        author="Editor A",
        date=datetime(2026, 1, 29, 16, 57, 1, tzinfo=timezone.utc),
        text="color",
        paragraph_ref="P5#a192",
        occurrence=0,
        group_id=4,
        group_source="inferred",
        changeset_id=4,
        changeset_source="inferred",
    ),
    Revision(
        id=1005,
        type="deletion",
        author="Reviewer",
        date=datetime(2026, 1, 29, 16, 58, tzinfo=timezone.utc),
        text="reenter",
        paragraph_ref="P6#975f",
        occurrence=0,
        group_id=5,
        group_source="inferred",
        changeset_id=5,
        changeset_source="inferred",
    ),
    Revision(
        id=1006,
        type="insertion",
        author="Reviewer B",
        date=datetime(2026, 1, 29, 16, 59, tzinfo=timezone.utc),
        text="— or is it?",
        paragraph_ref="P8#1b23",
        occurrence=0,
        group_id=6,
        group_source="inferred",
        changeset_id=6,
        changeset_source="inferred",
    ),
    Revision(
        id=1007,
        type="insertion",
        author="Author A",
        date=datetime(2026, 1, 29, 17, 0, tzinfo=timezone.utc),
        text="DRAFT ",
        paragraph_ref="P9#5772",
        occurrence=0,
        group_id=7,
        group_source="inferred",
        changeset_id=7,
        changeset_source="inferred",
    ),
    Revision(
        id=1008,
        type="deletion",
        author="Author B",
        date=datetime(2026, 1, 29, 17, 0, 10, tzinfo=timezone.utc),
        text="DRAFT ",
        paragraph_ref="P9#5772",
        occurrence=0,
        group_id=8,
        group_source="inferred",
        changeset_id=8,
        changeset_source="inferred",
    ),
]

AUTHOR_REVISION_COUNTS = {
    "Test Author": 2,
    "Editor A": 2,
    "Reviewer": 1,
    "Reviewer B": 1,
    "Author A": 1,
    "Author B": 1,
}


@pytest.fixture
def foreign_docx(test_data_dir, tmp_path) -> Path:
    """Copy of the foreign-revision fixture (never open fixtures in place)."""
    dest = tmp_path / "OXML_TrackChanges_Test.docx"
    shutil.copy(test_data_dir / "OXML_TrackChanges_Test.docx", dest)
    return dest


@pytest.fixture
def tricky_docx(test_data_dir, tmp_path) -> Path:
    """Copy of the fragmented-runs fixture."""
    dest = tmp_path / "tricky-track-changes.docx"
    shutil.copy(test_data_dir / "tricky-track-changes.docx", dest)
    return dest


def _body_text(doc: Document) -> str:
    """Visible text with each paragraph's trailing bracketed note run stripped.

    The fixture's notes spell out expected outcomes, so an assertion against
    the full visible text could be satisfied by a note instead of the body it
    is meant to check.
    """
    return "\n".join(line.split("[")[0] for line in doc.get_visible_text().splitlines())


class TestForeignRevisionParsing:
    """list_revisions() on revisions this library did not create."""

    def test_all_revisions_parsed(self, foreign_docx):
        with Document.open(foreign_docx, author="Test Editor") as doc:
            assert doc.list_revisions() == EXPECTED_REVISIONS

    def test_author_filter_counts(self, foreign_docx):
        with Document.open(foreign_docx, author="Test Editor") as doc:
            for author, count in AUTHOR_REVISION_COUNTS.items():
                revs = doc.list_revisions(author=author)
                assert len(revs) == count, author
                assert all(r.author == author for r in revs)

    def test_special_characters_survive_parsing(self, foreign_docx):
        """Em dash in revision text; smart quote in surrounding paragraph text."""
        with Document.open(foreign_docx, author="Test Editor") as doc:
            rev_1006 = next(r for r in doc.list_revisions() if r.id == 1006)
            assert rev_1006.text == "— or is it?"
            assert "It’s complicated." in _body_text(doc)


class TestForeignAcceptReject:
    """accept_all()/reject_all() across all 8 foreign revisions."""

    def test_accept_all(self, foreign_docx):
        with Document.open(foreign_docx, author="Test Editor") as doc:
            assert doc.accept_all() == 8
            assert doc.list_revisions() == []
            text = _body_text(doc)
            assert "This sentence has an inserted word." in text
            assert "This sentence remains." in text
            assert "We prefer  color spelling." in text
            assert "Please  your credentials." in text
            assert "It’s complicated.— or is it?" in text
            assert "DRAFT Specification follows." in text
            assert "DRAFT DRAFT" not in text

    def test_reject_all(self, foreign_docx):
        with Document.open(foreign_docx, author="Test Editor") as doc:
            assert doc.reject_all() == 8
            assert doc.list_revisions() == []
            text = _body_text(doc)
            assert "This sentence has an  word." in text
            assert "This old sentence remains." in text
            assert "We prefer colour  spelling." in text
            assert "Please reenter your credentials." in text
            # The rejected insertion is gone from the body.
            assert "It’s complicated." in text
            assert "— or is it?" not in text
            assert "DRAFT Specification follows." in text

    def test_accept_all_author_filter(self, foreign_docx):
        with Document.open(foreign_docx, author="Test Editor") as doc:
            assert doc.accept_all(author="Editor A") == 2
            remaining = doc.list_revisions()
            assert len(remaining) == 6
            assert all(r.author != "Editor A" for r in remaining)
            assert "We prefer  color spelling." in _body_text(doc)

    def test_reject_all_author_filter(self, foreign_docx):
        with Document.open(foreign_docx, author="Test Editor") as doc:
            assert doc.reject_all(author="Test Author") == 2
            assert len(doc.list_revisions()) == 6
            text = _body_text(doc)
            # Discriminating: before the reject the body reads "This sentence
            # remains." — "old " only returns if the deletion was restored.
            assert "This old sentence remains." in text
            # Test Author's rejected insertion is gone from the body too.
            assert "This sentence has an  word." in text

    def test_accept_conflicting_insert_delete_pair(self, foreign_docx):
        """Author A inserted "DRAFT "; Author B deleted the pre-existing "DRAFT ".

        Accepting both must leave exactly one "DRAFT " in the paragraph.
        """
        with Document.open(foreign_docx, author="Test Editor") as doc:
            assert doc.accept_revision(1007) is True
            assert doc.accept_revision(1008) is True
            assert {r.id for r in doc.list_revisions()}.isdisjoint({1007, 1008})
            text = _body_text(doc)
            assert "DRAFT Specification follows." in text
            assert "DRAFT DRAFT" not in text
            draft_para = next(line for line in text.splitlines() if "Specification follows" in line)
            assert draft_para.count("DRAFT ") == 1


class TestAcceptSaveReopenRoundTrip:
    """Partial accept must survive save + reopen with the rest intact."""

    def test_accept_save_reopen_finish(self, foreign_docx):
        with Document.open(foreign_docx, author="Test Editor") as doc:
            assert doc.accept_all(author="Editor A") == 2
            doc.save()

        with Document.open(foreign_docx, author="Test Editor") as doc:
            remaining = doc.list_revisions()
            assert len(remaining) == 6
            assert "Editor A" not in {r.author for r in remaining}
            # Spot-check one untouched revision survived the round trip intact.
            rev_1006 = next(r for r in remaining if r.id == 1006)
            assert rev_1006.author == "Reviewer B"
            assert rev_1006.text == "— or is it?"
            assert rev_1006.date == datetime(2026, 1, 29, 16, 59, tzinfo=timezone.utc)
            assert "We prefer  color spelling." in _body_text(doc)
            # Finish the job on the reopened document.
            assert doc.accept_all() == 6
            assert doc.list_revisions() == []


_W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _xmlspace_document() -> str:
    """One paragraph with a Word-realistic deletion: delText carries xml:space."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_W_NS}>"
        "<w:body>"
        "<w:p>"
        '<w:r><w:t xml:space="preserve">This </w:t></w:r>'
        '<w:del w:id="99" w:author="Foreign Author" w:date="2026-01-29T16:56:00Z">'
        '<w:r><w:delText xml:space="preserve">old </w:delText></w:r>'
        "</w:del>"
        "<w:r><w:t>sentence remains.</w:t></w:r>"
        "</w:p>"
        "</w:body>"
        "</w:document>"
    )


class TestRejectDeletionXmlSpace:
    """Rejecting a deletion must not lose whitespace-significant text."""

    def test_rejected_deltext_keeps_xml_space_attribute(self, simple_docx, tmp_path):
        docx_path = tmp_path / "xmlspace.docx"
        replace_document_xml(simple_docx, docx_path, _xmlspace_document())

        with Document.open(docx_path, author="Test Editor") as doc:
            assert doc.reject_revision(99) is True
            assert doc.get_visible_text() == "This old sentence remains."
            doc.save()

        with zipfile.ZipFile(docx_path) as z:
            document_xml = z.read("word/document.xml").decode("utf-8")
        assert '<w:t xml:space="preserve">old </w:t>' in document_xml

        with Document.open(docx_path, author="Test Editor") as doc:
            assert doc.get_visible_text() == "This old sentence remains."
            assert doc.list_revisions() == []

    def test_foreign_deletion_without_attribute_round_trips(self, foreign_docx):
        """The fixture's delText lacks xml:space; text still survives save/reopen."""
        with Document.open(foreign_docx, author="Test Editor") as doc:
            assert doc.reject_revision(1002) is True
            # Discriminating: the body reads "This sentence remains." until
            # the deleted "old " (trailing space and all) is restored.
            assert "This old sentence remains." in _body_text(doc)
            doc.save()

        with Document.open(foreign_docx, author="Test Editor") as doc:
            assert "This old sentence remains." in _body_text(doc)
            assert all(r.id != 1002 for r in doc.list_revisions())


def _table_document() -> str:
    """Body paragraphs plus a table whose second cell contains a nested table."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_W_NS}>"
        "<w:body>"
        "<w:p><w:r><w:t>Body para one</w:t></w:r></w:p>"
        "<w:tbl>"
        "<w:tr>"
        "<w:tc><w:p><w:r><w:t>Cell alpha text</w:t></w:r></w:p></w:tc>"
        "<w:tc>"
        "<w:p><w:r><w:t>Cell beta text</w:t></w:r></w:p>"
        "<w:tbl>"
        "<w:tr><w:tc><w:p><w:r><w:t>Nested cell text</w:t></w:r></w:p></w:tc></w:tr>"
        "</w:tbl>"
        "</w:tc>"
        "</w:tr>"
        "</w:tbl>"
        "<w:p><w:r><w:t>Body para two</w:t></w:r></w:p>"
        "</w:body>"
        "</w:document>"
    )


def _make_table_edits(doc: Document) -> None:
    """Tracked edits in body, both cells, and the nested cell (6 revisions)."""
    doc.replace("para", "PARA", paragraph=find_ref(doc, "Body para one"))
    doc.delete("alpha ", paragraph=find_ref(doc, "Cell alpha"))
    doc.insert_after("text", " INSERTED", paragraph=find_ref(doc, "Cell beta"))
    doc.replace("Nested", "NESTED-EDIT", paragraph=find_ref(doc, "Nested cell"))


class TestTableRevisionLifecycle:
    """Create/accept/reject revisions inside table cells and nested tables."""

    @pytest.fixture
    def table_docx(self, simple_docx, tmp_path) -> Path:
        dest = tmp_path / "tables.docx"
        replace_document_xml(simple_docx, dest, _table_document())
        return dest

    def test_accept_all_in_tables(self, table_docx):
        with Document.open(table_docx, author="Test Editor") as doc:
            _make_table_edits(doc)
            assert len(doc.list_revisions()) == 6
            assert doc.accept_all() == 6
            assert doc.list_revisions() == []
            assert doc.get_visible_text() == (
                "Body PARA one\nCell text\nCell beta text INSERTED\nNESTED-EDIT cell text\nBody para two"
            )

    def test_reject_all_in_tables(self, table_docx):
        with Document.open(table_docx, author="Test Editor") as doc:
            _make_table_edits(doc)
            assert doc.reject_all() == 6
            assert doc.list_revisions() == []
            assert doc.get_visible_text() == (
                "Body para one\nCell alpha text\nCell beta text\nNested cell text\nBody para two"
            )


class TestTrickyFixtureSearch:
    """Cross-run search over real Word run fragmentation."""

    def test_document_shape(self, tricky_docx):
        with Document.open(tricky_docx, author="Test Editor") as doc:
            assert doc.paragraph_count() == 22
            assert doc.list_revisions() == []

    @pytest.mark.parametrize(
        ("needle", "count"),
        [
            ("reenter", 2),  # fragmented "re" + "enter"
            ("microservice", 2),  # fragmented "micro" + "ser" + "vice"
            ("Quick Brown", 1),  # single bold run; the note's "ick Brown" must not match
            ("H2O", 2),  # "H" + superscript "2" + "O"
        ],
    )
    def test_count_matches_across_fragmented_runs(self, tricky_docx, needle, count):
        """Fragmented terms count correctly; single-run "Quick Brown" gets no spurious note match."""
        with Document.open(tricky_docx, author="Test Editor") as doc:
            assert doc.count_matches(needle) == count

    def test_replace_fragmented_occurrence(self, tricky_docx):
        """Occurrence 0 is the fragmented match (document order beats note runs)."""
        with Document.open(tricky_docx, author="Test Editor") as doc:
            ref = find_ref(doc, "reenter")
            doc.replace("reenter", "login", paragraph=ref, occurrence=0)
            # Cross-run replace deletes both fragments and inserts once.
            assert doc.accept_all() == 3
            assert "login your password." in _body_text(doc)
