"""Tests for revision grouping (ISSUES.md #37) and reconstruction (#46).

Every logical edit operation registers the revisions it creates as one
in-memory revision group, exposed via EditResult (a str subclass carrying
``group_id``/``revision_ids``) and resolvable as a unit with
``accept_group``/``reject_group``. The headline failure this prevents:
accepting only some of a ``rewrite_paragraph``'s revisions garbles the
paragraph, because each revision is a diff hunk, not a self-contained edit.

Group ids are per-open-Document and renumbered on each open. Revisions
already in the file (pre-session or foreign) get *inferred* groups
reconstructed at parse time: contiguous same-paragraph revisions sharing
identical raw ``w:author`` + ``w:date`` are one group
(``group_source="inferred"``); session edits record theirs
(``group_source="recorded"``). Unknown group ids raise RevisionError.
"""

import shutil
from datetime import datetime, timezone
from pathlib import Path

import pytest
from conftest import count_dom_walks, find_ref

from docx_editor import Document, EditOperation, EditResult, RevisionError
from docx_editor.exceptions import BatchOperationError, DocxEditError
from docx_editor.track_changes import RevisionManager
from docx_editor.xml_editor import DocxXMLEditor

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

AUTHOR_A = "Reviewer A"
AUTHOR_B = "Reviewer B"
DATE_A = "2024-01-01T00:00:00Z"
DATE_B = "2024-01-01T00:00:01Z"


def _ins_xml(rev_id, text: str, author: str = AUTHOR_A, date: str = DATE_A) -> str:
    """A <w:ins> wrapping one run of ``text``."""
    return f'<w:ins w:id="{rev_id}" w:author="{author}" w:date="{date}"><w:r><w:t>{text}</w:t></w:r></w:ins>'


def _del_xml(rev_id, text: str, author: str = AUTHOR_A, date: str = DATE_A) -> str:
    """A <w:del> wrapping one run of deleted ``text``."""
    return (
        f'<w:del w:id="{rev_id}" w:author="{author}" w:date="{date}"><w:r><w:delText>{text}</w:delText></w:r></w:del>'
    )


class _FrozenDatetime(datetime):
    """datetime whose now() is pinned to one instant."""

    @classmethod
    def now(cls, tz=None):
        return cls(2025, 6, 1, 12, 0, 0, tzinfo=tz)


@pytest.fixture
def frozen_clock(monkeypatch):
    """Pin every w:date this session stamps to one fixed second.

    w:date has second precision, so same-second tests against the real
    clock would be flaky across second boundaries.
    """
    monkeypatch.setattr("docx_editor.xml_editor.datetime", _FrozenDatetime)


@pytest.fixture
def ticking_clock(monkeypatch):
    """Every datetime.now() call lands in a NEW second.

    Adversarial clock for the per-operation frozen timestamp: without
    freezing, each attribute-injection call inside one edit would stamp
    its own second, and reconstruction would split the edit after reopen.
    """

    class _TickingDatetime(datetime):
        _tick = 0

        @classmethod
        def now(cls, tz=None):
            _TickingDatetime._tick += 1
            return cls(2025, 6, 1, 12, _TickingDatetime._tick // 60, _TickingDatetime._tick % 60, tzinfo=tz)

    monkeypatch.setattr("docx_editor.xml_editor.datetime", _TickingDatetime)


@pytest.fixture
def settable_clock(monkeypatch):
    """A clock the test drives explicitly.

    ``now()`` returns whatever instant the test last assigned to the
    class-level ``current``; assign a later value to advance time.
    """

    class _SettableDatetime(datetime):
        current = datetime(2025, 6, 1, 12, 0, 0)

        @classmethod
        def now(cls, tz=None):
            c = cls.current
            return cls(c.year, c.month, c.day, c.hour, c.minute, c.second, tzinfo=tz)

    monkeypatch.setattr("docx_editor.xml_editor.datetime", _SettableDatetime)
    return _SettableDatetime


@pytest.fixture
def doc(temp_docx):
    """An open Document over the simple.docx copy, closed after the test."""
    document = Document.open(temp_docx)
    yield document
    document.close()


@pytest.fixture
def temp_xml(tmp_path):
    """Return a factory writing a minimal document.xml with the given body."""

    def _create_xml(body_xml: str) -> Path:
        xml = f'<?xml version="1.0" encoding="utf-8"?><w:document {NS}><w:body>{body_xml}</w:body></w:document>'
        xml_path = tmp_path / "test_doc.xml"
        xml_path.write_text(xml)
        return xml_path

    return _create_xml


def _make_manager(xml_path: Path, author: str = AUTHOR_B) -> RevisionManager:
    """Create a RevisionManager editing as ``author``."""
    editor = DocxXMLEditor(xml_path, rsid="00000000", author=author)
    return RevisionManager(editor)


def _paragraph_text(doc: Document, index: int) -> str:
    """Visible (accepted-view) text of the index-th paragraph line."""
    return doc.get_visible_text().splitlines()[index]


class TestEditResult:
    """EditResult is a drop-in ref string with group metadata attached."""

    def test_replace_returns_editresult_with_one_group(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.replace("quick", "speedy", paragraph=ref)

        assert isinstance(result, EditResult)
        assert isinstance(result, str)
        assert result.group_id is not None
        # A replace is one deletion + one insertion, grouped together.
        assert len(result.revision_ids) == 2
        by_id = {r.id: r for r in doc.list_revisions()}
        types = sorted(by_id[rid].type for rid in result.revision_ids)
        assert types == ["deletion", "insertion"]

    def test_result_string_is_usable_as_paragraph_ref(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.replace("quick", "speedy", paragraph=ref)

        # Follow-up edit takes the EditResult where a ref string is expected.
        result2 = doc.replace("lazy", "sleepy", paragraph=result)
        assert isinstance(result2, EditResult)
        assert result2.group_id != result.group_id

    def test_single_revision_ops_get_one_member_group(self, doc):
        ref = find_ref(doc, "quick brown fox")

        inserted = doc.insert_after("fox", " swiftly", paragraph=ref)
        assert inserted.group_id is not None
        assert len(inserted.revision_ids) == 1

        deleted = doc.delete("lazy ", paragraph=inserted)
        assert deleted.group_id is not None
        assert deleted.group_id != inserted.group_id
        assert len(deleted.revision_ids) == 1

    def test_insert_before_gets_group(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.insert_before("The", "Note: ", paragraph=ref)
        assert result.group_id is not None
        assert len(result.revision_ids) == 1


class TestGroupIdOnRevisions:
    """list_revisions stamps group_id on session-created revisions."""

    def test_revisions_carry_group_id(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.replace("quick", "speedy", paragraph=ref)

        for rev in doc.list_revisions():
            assert rev.group_id == result.group_id
            assert "group=" in repr(rev)

    def test_foreign_revisions_get_inferred_groups(self, test_data_dir, temp_dir):
        # Every revision in the fixture differs from its document-order
        # neighbor by paragraph, author, or date — the heuristic partitions
        # them into eight singleton inferred groups (the full expected
        # partition is pinned in test_foreign_revisions.py).
        fixture = test_data_dir / "OXML_TrackChanges_Test.docx"
        dest = temp_dir / "foreign.docx"
        shutil.copy(fixture, dest)
        with Document.open(dest) as doc:
            revisions = doc.list_revisions()
            assert revisions
            assert all(rev.group_id is not None for rev in revisions)
            assert all(rev.group_source == "inferred" for rev in revisions)
            assert all(f"group={rev.group_id}(inferred)" in repr(rev) for rev in revisions)
            gids = [rev.group_id for rev in revisions]
            assert len(set(gids)) == len(gids)  # all singletons


class TestAcceptRejectGroup:
    def test_reject_group_restores_original_text(self, doc):
        ref = find_ref(doc, "quick brown fox")
        original = _paragraph_text(doc, 1)
        result = doc.replace("quick", "speedy", paragraph=ref)

        count = doc.reject_group(result.group_id)

        assert count == 2
        assert _paragraph_text(doc, 1) == original
        assert doc.list_revisions() == []

    def test_accept_group_applies_the_edit(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.replace("quick", "speedy", paragraph=ref)

        count = doc.accept_group(result.group_id)

        assert count == 2
        assert "speedy brown fox" in _paragraph_text(doc, 1)
        assert doc.list_revisions() == []

    def test_unknown_group_raises_revision_error(self, doc):
        with pytest.raises(RevisionError, match="Unknown revision group: 999"):
            doc.accept_group(999)
        with pytest.raises(RevisionError, match="Unknown revision group: 999"):
            doc.reject_group(999)

    def test_per_id_interplay_processes_remaining_members(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.replace("quick", "speedy", paragraph=ref)

        # Resolve one member individually first.
        assert doc.accept_revision(result.revision_ids[0])

        count = doc.accept_group(result.group_id)

        assert count == len(result.revision_ids) - 1
        assert doc.list_revisions() == []
        assert "speedy brown fox" in _paragraph_text(doc, 1)
        # A fully resolved group stays known; a further call processes nothing.
        assert doc.accept_group(result.group_id) == 0


class TestRewriteGrouping:
    """The headline case: a rewrite's many revisions resolve as one unit."""

    NEW_TEXT = "A slow red cat crawls beneath the energetic dog."

    def test_rewrite_revisions_share_one_group(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.rewrite_paragraph(ref, self.NEW_TEXT)

        assert result.group_id is not None
        revisions = doc.list_revisions()
        assert len(revisions) == len(result.revision_ids) > 2
        assert {rev.group_id for rev in revisions} == {result.group_id}

    def test_reject_group_undoes_whole_rewrite(self, doc):
        ref = find_ref(doc, "quick brown fox")
        original = _paragraph_text(doc, 1)
        result = doc.rewrite_paragraph(ref, self.NEW_TEXT)

        doc.reject_group(result.group_id)

        assert _paragraph_text(doc, 1) == original
        assert doc.list_revisions() == []

    def test_accept_group_yields_new_text(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.rewrite_paragraph(ref, self.NEW_TEXT)

        doc.accept_group(result.group_id)

        assert _paragraph_text(doc, 1) == self.NEW_TEXT
        assert doc.list_revisions() == []

    def test_noop_rewrite_creates_no_group(self, doc):
        ref = find_ref(doc, "quick brown fox")
        current = _paragraph_text(doc, 1)
        result = doc.rewrite_paragraph(ref, current)

        assert result.group_id is None
        assert result.revision_ids == ()
        assert result == ref  # unchanged paragraph keeps its hash


class TestBatchGrouping:
    def test_batch_edit_one_group_per_operation(self, doc):
        ref1 = find_ref(doc, "quick brown fox")
        ref2 = find_ref(doc, "sample document for testing")

        results = doc.batch_edit([
            EditOperation.replace("quick", "speedy", paragraph=ref1),
            EditOperation.delete("sample ", paragraph=ref2),
        ])

        assert all(isinstance(r, EditResult) for r in results)
        gids = [r.group_id for r in results]
        assert None not in gids
        assert gids[0] != gids[1]
        assert len(results[0].revision_ids) == 2  # replace: del + ins
        assert len(results[1].revision_ids) == 1  # delete: one del

        # Accept one op, reject the other — the point of per-op groups.
        doc.accept_group(gids[0])
        doc.reject_group(gids[1])
        assert "speedy brown fox" in _paragraph_text(doc, 1)
        assert "sample document for testing" in doc.get_visible_text()
        assert doc.list_revisions() == []

    def test_batch_rewrite_one_group_per_rewrite(self, doc):
        ref1 = find_ref(doc, "quick brown fox")
        ref2 = find_ref(doc, "sample document for testing")

        results = doc.batch_rewrite([
            (ref1, "A slow red cat sits."),
            (ref2, "This paragraph was fully rewritten for the test."),
        ])

        assert all(isinstance(r, EditResult) for r in results)
        assert results[0].group_id != results[1].group_id
        by_group = {rev.id: rev.group_id for rev in doc.list_revisions()}
        for result in results:
            for rid in result.revision_ids:
                assert by_group[rid] == result.group_id

    def test_batch_edit_rollback_restores_registry(self, doc):
        ref1 = find_ref(doc, "quick brown fox")
        manager = doc._revision_manager
        counter_before = manager._group_counter
        groups_before = dict(manager._groups)

        with pytest.raises(BatchOperationError):
            doc.batch_edit([
                EditOperation.replace("quick", "speedy", paragraph=ref1),
                EditOperation.delete("no such text anywhere", paragraph=ref1),
            ])

        # No ghost groups pointing at rolled-back revisions.
        assert manager._group_counter == counter_before
        assert manager._groups == groups_before
        assert manager._revision_groups == {}
        assert doc.list_revisions() == []

    def test_batch_rewrite_rollback_restores_registry(self, doc, monkeypatch):
        # Both rewrites pass upfront hash validation; the failure must happen
        # at apply time, after the first rewrite has registered its group.
        # Applies run in reverse paragraph order, so the failing rewrite goes
        # on the lower-index paragraph (P2) and the succeeding one on P3.
        ref1 = find_ref(doc, "quick brown fox")
        ref2 = find_ref(doc, "sample document for testing")
        manager = doc._revision_manager
        counter_before = manager._group_counter

        real_rewrite = RevisionManager.rewrite_paragraph

        def flaky(self, ref_str, new_text):
            if new_text == "BOOM":
                raise DocxEditError("simulated apply failure")
            return real_rewrite(self, ref_str, new_text)

        monkeypatch.setattr(RevisionManager, "rewrite_paragraph", flaky)

        with pytest.raises(BatchOperationError) as exc:
            doc.batch_rewrite([
                (ref1, "BOOM"),
                (ref2, "Rewritten paragraph three."),
            ])
        assert exc.value.operation_index == 0

        assert manager._group_counter == counter_before
        assert manager._groups == {}
        assert manager._revision_groups == {}
        assert doc.list_revisions() == []


class TestGroupSessionScope:
    def test_reopen_reconstructs_inferred_group(self, temp_docx):
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            result = doc.replace("quick", "speedy", paragraph=ref)
            assert result.group_id is not None
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            assert len(revisions) == 2  # the tracked change persisted
            # The replace's del+ins pair is contiguous in one paragraph with
            # one author+date, so it reconstructs as ONE inferred group under
            # a fresh per-open id — old ids are meaningless after reopen.
            gid = revisions[0].group_id
            assert gid is not None
            assert all(rev.group_id == gid for rev in revisions)
            assert all(rev.group_source == "inferred" for rev in revisions)
            # The reconstructed group resolves as a unit.
            assert doc.accept_group(gid) == 2
            assert "speedy brown fox" in _paragraph_text(doc, 1)
            assert doc.list_revisions() == []

    def test_save_keeps_groups_alive(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.replace("quick", "speedy", paragraph=ref)

        doc.save()

        assert doc.reject_group(result.group_id) == 2


def _mark_ins_xml(rev_id, author: str = AUTHOR_A, date: str = DATE_A) -> str:
    """A paragraph-mark insertion (empty <w:ins> inside <w:pPr><w:rPr>)."""
    return f'<w:ins w:id="{rev_id}" w:author="{author}" w:date="{date}"/>'


def _split_paragraph_xml(mark_id, content_id, text: str, **kw) -> str:
    """A first-half paragraph: mark inserted, plus one inserted content run."""
    return f"<w:p><w:pPr><w:rPr>{_mark_ins_xml(mark_id, **kw)}</w:rPr></w:pPr>{_ins_xml(content_id, text, **kw)}</w:p>"


class TestSplitReconstruction:
    """A reopened tracked split spans two paragraphs; the mark makes it one group."""

    def test_reopened_split_reconstructs_as_one_group(self, temp_xml):
        body = _split_paragraph_xml(2, 1, "first ") + f"<w:p>{_ins_xml(3, 'second')}</w:p>"
        manager = _make_manager(temp_xml(body))

        gid = manager.group_id_of(1)
        assert gid is not None
        assert {manager.group_id_of(i) for i in (1, 2, 3)} == {gid}
        assert set(manager.group_revisions(gid)) == {1, 2, 3}
        assert manager._group_sources[gid] == "inferred"

    def test_reopened_multi_split_reconstructs_as_one_group(self, temp_xml):
        body = _split_paragraph_xml(10, 1, "a") + _split_paragraph_xml(11, 3, "b") + f"<w:p>{_ins_xml(5, 'c')}</w:p>"
        manager = _make_manager(temp_xml(body))

        gid = manager.group_id_of(1)
        assert gid is not None
        assert {manager.group_id_of(i) for i in (1, 3, 5, 10, 11)} == {gid}

    def test_adjacent_paragraphs_without_mark_stay_separate(self, temp_xml):
        # Same author+date, adjacent paragraphs, but no inserted mark — the two
        # edits are unrelated and must reconstruct as two groups (the rule's
        # negative control).
        body = f"<w:p>{_ins_xml(1, 'first')}</w:p><w:p>{_ins_xml(2, 'second')}</w:p>"
        manager = _make_manager(temp_xml(body))

        assert manager.group_id_of(1) != manager.group_id_of(2)

    def test_split_continuation_needs_matching_author_and_date(self, temp_xml):
        # A mark by a different author/date than the tail's revision is not a
        # continuation — the durable signal must match on both.
        body = (
            _split_paragraph_xml(2, 1, "first ", author=AUTHOR_A)
            + f"<w:p>{_ins_xml(3, 'second', author=AUTHOR_B, date=DATE_B)}</w:p>"
        )
        manager = _make_manager(temp_xml(body))

        assert manager.group_id_of(1) != manager.group_id_of(3)

    def test_split_refuses_paragraph_with_section_mark(self, temp_xml):
        # A paragraph-level w:sectPr marks a section boundary; splitting it
        # would corrupt the section structure, so it is refused (no mutation).
        body = (
            '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr>'
            "<w:r><w:t>Section end here</w:t></w:r></w:p>"
        )
        manager = _make_manager(temp_xml(body))

        with pytest.raises(RevisionError, match="section mark"):
            manager.replace_text("end", "end\nsplit")
        # No partial mutation: the original text is intact.
        assert "Section end here" in manager.editor.dom.toxml()

    def test_split_inside_existing_revision_is_atomic(self, temp_xml):
        # A split whose boundary lands inside a pre-existing (foreign) insertion
        # is not yet supported. It must refuse BEFORE mutating: a single edit has
        # no DOM rollback, so a partial delete/insert would otherwise be left
        # behind (the tail-collection raise used to fire mid-mutation).
        body = f"<w:p><w:r><w:t>Hello</w:t></w:r>{_ins_xml(1, 'WORLD', author=AUTHOR_A)}</w:p>"
        manager = _make_manager(temp_xml(body))
        before = manager.editor.dom.toxml()

        with pytest.raises(RevisionError, match="existing revision"):
            # match.end (after "Hello") lands at the start of the foreign <w:ins>.
            manager.replace_text("Hello", "a\nb")

        # Byte-for-byte unchanged — no orphaned deletion/insertion.
        assert manager.editor.dom.toxml() == before

    def test_split_inserted_segment_inherits_run_format(self, temp_xml):
        # The segment inserted at the start of the new paragraph copies the rPr
        # of the tail it sits before, so it keeps the surrounding formatting
        # instead of dropping to document default.
        body = "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello World</w:t></w:r></w:p>"
        manager = _make_manager(temp_xml(body))

        manager.replace_text("Hello", "A\nB")

        b_runs = [
            r
            for r in manager.editor.dom.getElementsByTagName("w:r")
            if (ts := r.getElementsByTagName("w:t")) and ts[0].firstChild and ts[0].firstChild.data == "B"
        ]
        assert b_runs, "inserted 'B' run not found"
        rprs = b_runs[0].getElementsByTagName("w:rPr")
        assert rprs and rprs[0].getElementsByTagName("w:b"), "split-inserted segment lost its bold rPr"


class TestForeignInsGrouping:
    """Author/attachment filters keep foreign fragments out of our groups."""

    def test_replace_inside_foreign_ins_excludes_split_halves(self, temp_xml):
        # B replaces the middle of A's pending insertion: A's w:ins splits
        # into identity-preserving halves (fresh ids, A's author) around
        # B's nested w:del + sibling w:ins. Only B's elements join the group.
        body = (
            f'<w:p><w:ins w:id="1" w:author="{AUTHOR_A}" w:date="{DATE_A}">'
            "<w:r><w:t>alpha beta gamma</w:t></w:r></w:ins></w:p>"
        )
        manager = _make_manager(temp_xml(body), author=AUTHOR_B)

        change_id = manager.replace_text("beta", "delta")

        group_id = manager.group_id_of(change_id)
        assert group_id is not None
        members = manager.group_revisions(group_id)

        revisions = {rev.id: rev for rev in manager.list_revisions()}
        member_authors = {revisions[rid].author for rid in members}
        assert member_authors == {AUTHOR_B}
        # A's original + split-off insertions stay out of the group.
        foreign_ids = {rid for rid, rev in revisions.items() if rev.author == AUTHOR_A}
        assert foreign_ids and not (foreign_ids & set(members))
        # B's nested deletion of A's text is part of B's group.
        member_types = {revisions[rid].type for rid in members}
        assert member_types == {"deletion", "insertion"}

    def test_split_foreign_ins_does_not_claim_presession_nested_del(self, temp_xml):
        # A's pending insertion contains B's PRE-SESSION deletion (made
        # before a save/reopen). B then inserts inside A's insertion before
        # the deletion: the split re-serializes A's trailing content
        # INCLUDING B's old <w:del>, running it back through attribute
        # injection. That pre-existing deletion must not join B's insert
        # group — rejecting the insert would silently resurrect "beta ".
        body = (
            f'<w:p><w:ins w:id="1" w:author="{AUTHOR_A}" w:date="{DATE_A}">'
            "<w:r><w:t>alpha </w:t></w:r>"
            f'<w:del w:id="2" w:author="{AUTHOR_B}" w:date="{DATE_A}">'
            "<w:r><w:delText>beta </w:delText></w:r></w:del>"
            "<w:r><w:t>gamma</w:t></w:r></w:ins></w:p>"
        )
        manager = _make_manager(temp_xml(body), author=AUTHOR_B)

        change_id = manager.insert_text_after("alpha", "NEW ")

        group_id = manager.group_id_of(change_id)
        assert group_id is not None
        members = manager.group_revisions(group_id)
        assert members == (change_id,)  # only the new insertion, not del#2

        # del#2 keeps its own inferred singleton group (the author change
        # inside ins#1 broke the contiguity run), distinct from B's
        # recorded insert group.
        revisions = {rev.id: rev for rev in manager.list_revisions()}
        del_gid = revisions[2].group_id
        assert del_gid is not None
        assert del_gid != group_id
        assert revisions[2].group_source == "inferred"
        assert manager.group_revisions(del_gid) == (2,)

        # Undoing the insert must leave the pre-existing deletion untouched.
        manager.reject_group(group_id)
        remaining = {rev.id for rev in manager.list_revisions()}
        assert 2 in remaining

    def test_delete_inside_foreign_ins_groups_nested_del(self, temp_xml):
        body = (
            f'<w:p><w:ins w:id="1" w:author="{AUTHOR_A}" w:date="{DATE_A}">'
            "<w:r><w:t>alpha beta gamma</w:t></w:r></w:ins></w:p>"
        )
        manager = _make_manager(temp_xml(body), author=AUTHOR_B)

        change_id = manager.suggest_deletion("beta ")

        group_id = manager.group_id_of(change_id)
        assert group_id is not None
        members = manager.group_revisions(group_id)
        revisions = {rev.id: rev for rev in manager.list_revisions()}
        assert [revisions[rid].author for rid in members] == [AUTHOR_B]
        assert [revisions[rid].type for rid in members] == ["deletion"]
        # The nested del sits inside A's insertion.
        assert revisions[members[0]].nested_under == 1


class TestOwnInsertionSplits:
    """Physical edits inside our own pending insertions keep groups truthful."""

    def test_replace_consuming_whole_own_insertion_reports_new_group(self, doc):
        ref = find_ref(doc, "quick brown fox")
        first = doc.insert_after("fox", " ZE9", paragraph=ref)

        # No token shared with " ZE9" (not even the space), so affix trimming
        # cannot narrow the match — the whole insertion is consumed.
        result = doc.replace(" ZE9", "QX7", paragraph=first)

        # The old insertion is fully consumed and the replacement re-wrapped
        # in a fresh <w:ins> — a revision this operation created, so its
        # group must be reachable from the result (not a phantom group).
        assert result.group_id is not None
        assert result.group_id != first.group_id
        revisions = doc.list_revisions()
        assert len(revisions) == 1
        assert revisions[0].group_id == result.group_id
        assert revisions[0].id in result.revision_ids

    def test_middle_delete_adopts_split_tail_into_origin_group(self, doc):
        ref = find_ref(doc, "quick brown fox")
        original = _paragraph_text(doc, 1)
        first = doc.insert_after("fox", " AAA MID BBB", paragraph=ref)

        result = doc.delete("MID ", paragraph=first)

        # The delete physically removed pending text — no new revisions of
        # its own; the split-off "BBB" tail stays with the insert's group.
        assert result.group_id is None
        revisions = doc.list_revisions()
        assert len(revisions) == 2
        assert {rev.group_id for rev in revisions} == {first.group_id}
        members = doc._revision_manager.group_revisions(first.group_id)
        assert len(members) == 2

        # Rejecting the original insertion's group removes BOTH halves.
        doc.reject_group(first.group_id)
        assert _paragraph_text(doc, 1) == original
        assert doc.list_revisions() == []

    def test_presession_middle_delete_adopts_tail_into_inferred_group(self, temp_docx):
        # The origin insertion predates this session, but reconstruction
        # gave it an inferred group on reopen — so its split-off tail joins
        # that group (and must not be claimed into a phantom group by the
        # delete operation, which itself creates no revisions).
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            original = _paragraph_text(doc, 1)
            doc.insert_after("fox", " AAA MID BBB", paragraph=ref)
            doc.save()

        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            result = doc.delete("MID ", paragraph=ref)

            assert result.group_id is None
            revisions = doc.list_revisions()
            assert len(revisions) == 2  # truncated head + split-off tail
            gid = revisions[0].group_id
            assert gid is not None
            assert all(rev.group_id == gid for rev in revisions)
            assert all(rev.group_source == "inferred" for rev in revisions)

            # Rejecting the inferred group removes BOTH halves — the rump/
            # tail correctness win of reconstruction.
            doc.reject_group(gid)
            assert _paragraph_text(doc, 1) == original
            assert doc.list_revisions() == []

    def test_presession_rewrite_does_not_claim_split_tail(self, temp_docx):
        # Same defect through the public rewrite path: without tail
        # registration, the rewrite's EditResult.group_id would hold ONLY
        # the pre-session insertion's tail, and reject_group (the documented
        # undo idiom) would rip out pre-existing text instead of undoing
        # the rewrite.
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            doc.insert_after("fox", " alpha beta gamma", paragraph=ref)
            doc.save()

        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            current = _paragraph_text(doc, 1)
            result = doc.rewrite_paragraph(ref, current.replace("beta", "BETA"))

            # The only physical change happens inside our own pre-session
            # insertion — no trackable revisions, so no group for the
            # rewrite. Every surviving revision piece (truncated origin +
            # split-off tails) stays in the origin's inferred group.
            assert result.group_id is None
            revisions = doc.list_revisions()
            gids = {rev.group_id for rev in revisions}
            assert len(gids) == 1 and None not in gids
            assert set(doc._revision_manager._groups) == gids
            assert all(rev.group_source == "inferred" for rev in revisions)
            # Placement-agnostic content check: replacement position inside
            # an own insertion is the pre-existing ISSUES.md #31 ordering
            # bug, out of scope here.
            text = _paragraph_text(doc, 1)
            assert "BETA" in text and "beta" not in text
            assert "alpha" in text and "gamma" in text


class TestGroupReconstruction:
    """Parse-time inference of groups for revisions already in the file (#46).

    Reconstruction runs at RevisionManager construction, so these unit
    tests build a manager over raw XML — no save/reopen needed.
    """

    def test_contiguous_same_author_date_one_group(self, temp_xml):
        body = f"<w:p><w:r><w:t>keep </w:t></w:r>{_ins_xml(1, 'one ')}{_del_xml(2, 'two ')}{_ins_xml(3, 'three')}</w:p>"
        manager = _make_manager(temp_xml(body))

        assert manager.group_revisions(1) == (1, 2, 3)  # document order
        by_id = {rev.id: rev for rev in manager.list_revisions()}
        assert all(rev.group_id == 1 for rev in by_id.values())
        assert all(rev.group_source == "inferred" for rev in by_id.values())
        assert "group=1(inferred)" in repr(by_id[1])

    def test_interleaved_author_breaks_contiguity(self, temp_xml):
        # A,A,B,A with identical dates: same-paragraph contiguity is
        # load-bearing — A's revisions around B's do NOT merge.
        body = (
            f"<w:p>{_ins_xml(1, 'a1 ')}{_ins_xml(2, 'a2 ')}"
            f"{_ins_xml(3, 'b ', author=AUTHOR_B)}{_ins_xml(4, 'a3')}</w:p>"
        )
        manager = _make_manager(temp_xml(body))

        assert manager._groups == {1: (1, 2), 2: (3,), 3: (4,)}

    def test_same_key_different_paragraphs_two_groups(self, temp_xml):
        # The measured paperflow partition: identical author+date in two
        # paragraphs stays two groups.
        body = f"<w:p>{_ins_xml(1, 'one')}</w:p><w:p>{_ins_xml(2, 'two')}</w:p>"
        manager = _make_manager(temp_xml(body))

        assert manager._groups == {1: (1,), 2: (2,)}

    def test_same_paragraph_different_dates_two_groups(self, temp_xml):
        body = f"<w:p>{_ins_xml(1, 'one ')}{_ins_xml(2, 'two', date=DATE_B)}</w:p>"
        manager = _make_manager(temp_xml(body))

        assert manager._groups == {1: (1,), 2: (2,)}

    def test_missing_author_or_date_stays_ungrouped(self, temp_xml):
        body = (
            f"<w:p>{_ins_xml(1, 'one ')}"
            f'<w:ins w:id="2" w:author="{AUTHOR_A}"><w:r><w:t>no date </w:t></w:r></w:ins>'
            f'<w:ins w:id="3" w:date="{DATE_A}"><w:r><w:t>no author </w:t></w:r></w:ins>'
            f"{_ins_xml(4, 'four')}</w:p>"
        )
        manager = _make_manager(temp_xml(body))

        # Ungroupable elements stay unregistered AND break the run: 1 and 4
        # do not merge across them.
        assert manager._groups == {1: (1,), 2: (4,)}
        by_id = {rev.id: rev for rev in manager.list_revisions()}
        assert by_id[2].group_id is None and by_id[2].group_source is None
        assert by_id[3].group_id is None
        assert "group=" not in repr(by_id[2])

    def test_nonconforming_ids_stay_ungrouped(self, temp_xml):
        # Non-numeric ids (nonconforming producers) break their runs, stay
        # unregistered, and are omitted from list_revisions(). A duplicated
        # id is wholly ungrouped — every occurrence, including the first —
        # because id-keyed lookup cannot tell the occurrences apart.
        body = (
            f"<w:p>{_ins_xml(1, 'one ')}"
            f'<w:ins w:id="oops" w:author="{AUTHOR_A}" w:date="{DATE_A}"><w:r><w:t>bad </w:t></w:r></w:ins>'
            f"{_ins_xml(2, 'two ')}{_ins_xml(2, 'dup ')}{_ins_xml(3, 'three')}</w:p>"
            f"<w:p>{_ins_xml(1, 'cross-para dup ')}{_ins_xml(4, 'four')}</w:p>"
        )
        manager = _make_manager(temp_xml(body))

        assert manager._groups == {1: (3,), 2: (4,)}

        revisions = manager.list_revisions()
        assert sorted(rev.id for rev in revisions) == [1, 1, 2, 2, 3, 4]  # "oops" omitted
        for rev in revisions:
            if rev.id in (1, 2):  # both occurrences of each duplicated id
                assert rev.group_id is None and rev.group_source is None
        by_id = {rev.id: rev for rev in revisions}
        assert by_id[3].group_id == 1 and by_id[4].group_id == 2

    def test_duplicate_id_with_ungroupable_occurrence_stays_ungrouped(self, temp_xml):
        # An ungroupable occurrence (here: no w:date) of a duplicated id
        # must still bar its groupable twin from winning a group — id-keyed
        # lookup would report that group for both elements.
        body = (
            f'<w:p><w:ins w:id="5" w:author="{AUTHOR_A}"><w:r><w:t>no date </w:t></w:r></w:ins>'
            f"{_ins_xml(5, 'groupable twin ')}{_ins_xml(6, 'six')}</w:p>"
        )
        manager = _make_manager(temp_xml(body))

        assert manager._groups == {1: (6,)}
        revisions = manager.list_revisions()
        assert sorted(rev.id for rev in revisions) == [5, 5, 6]
        assert all(rev.group_id is None for rev in revisions if rev.id == 5)

    def test_revision_outside_paragraph_stays_ungrouped(self, temp_xml):
        # A <w:trPr> row-mark insertion has no ancestor <w:p>.
        body = (
            "<w:tbl><w:tr>"
            f'<w:trPr><w:ins w:id="7" w:author="{AUTHOR_A}" w:date="{DATE_A}"/></w:trPr>'
            f"<w:tc><w:p>{_ins_xml(8, 'cell')}</w:p></w:tc>"
            "</w:tr></w:tbl>"
        )
        manager = _make_manager(temp_xml(body))

        assert manager._groups == {1: (8,)}
        by_id = {rev.id: rev for rev in manager.list_revisions()}
        assert by_id[7].group_id is None
        assert by_id[8].group_id == 1

    def test_live_groups_number_after_inferred(self, temp_xml):
        body = f"<w:p>{_ins_xml(1, 'old ')}<w:r><w:t>keep it</w:t></w:r></w:p>"
        manager = _make_manager(temp_xml(body))
        assert manager._groups == {1: (1,)}

        change_id = manager.replace_text("keep", "hold")

        live_gid = manager.group_id_of(change_id)
        assert live_gid is not None and live_gid > 1
        # The live capture never claims the pre-registered revision.
        assert 1 not in manager.group_revisions(live_gid)
        by_id = {rev.id: rev for rev in manager.list_revisions()}
        assert by_id[1].group_id == 1 and by_id[1].group_source == "inferred"
        assert by_id[change_id].group_source == "recorded"
        assert "(inferred)" not in repr(by_id[change_id])

    def test_saved_cross_run_replace_reconstructs_one_group(self, temp_xml, frozen_clock):
        # ISSUES.md #46 repro at unit level: one replace spanning split runs
        # creates several revisions sharing one date; after save + fresh
        # construction they reconstruct as ONE inferred group and
        # accept_group applies the whole replace.
        body = "<w:p><w:r><w:t>the quick </w:t></w:r><w:r><w:t>brown </w:t></w:r><w:r><w:t>fox jumps</w:t></w:r></w:p>"
        xml_path = temp_xml(body)
        manager = _make_manager(xml_path)
        change_id = manager.replace_text("quick brown fox", "slow red cat")
        live_gid = manager.group_id_of(change_id)
        assert live_gid is not None
        live_members = manager.group_revisions(live_gid)
        assert len(live_members) > 1  # several dels + one ins
        manager.editor.save()

        reopened = _make_manager(xml_path)
        assert len(reopened._groups) == 1
        (gid,) = reopened._groups
        assert set(reopened._groups[gid]) == set(live_members)
        assert reopened.accept_group(gid) == len(live_members)
        assert reopened.list_revisions() == []
        xml = reopened.editor.dom.toxml()
        assert "slow red cat" in xml and "quick brown" not in xml


class TestGroupReconstructionAcrossReopen:
    """Document-level save/reopen behavior of inferred groups (#46)."""

    def test_reopen_reconstructs_one_group_per_edit(self, temp_docx, frozen_clock):
        # The paperflow repro: one rewrite = several diff hunks; after
        # reopen they must resolve as one unit, not as orphans.
        new_text = "A slow red cat crawls beneath the energetic dog."
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            result = doc.rewrite_paragraph(ref, new_text)
            assert len(result.revision_ids) > 2
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            assert len(revisions) == len(result.revision_ids)
            gid = revisions[0].group_id
            assert gid is not None
            assert all(rev.group_id == gid for rev in revisions)
            assert all(rev.group_source == "inferred" for rev in revisions)
            doc.accept_group(gid)
            assert _paragraph_text(doc, 1) == new_text
            assert doc.list_revisions() == []

    def test_three_edits_same_second_partition_into_three_groups(self, temp_docx, frozen_clock):
        with Document.open(temp_docx) as doc:
            results = doc.batch_edit([
                EditOperation.replace("quick", "speedy", paragraph=find_ref(doc, "quick brown fox")),
                EditOperation.delete("sample ", paragraph=find_ref(doc, "sample document for testing")),
            ])
            third = doc.replace("well-structured", "tidy", paragraph=find_ref(doc, "well-structured document"))
            assert len({r.group_id for r in [*results, third]}) == 3
            doc.save()

        with Document.open(temp_docx) as doc:
            # Same author; the batch's ops share one date and the third
            # (single) edit gets a collision-bumped second. All three edits
            # land in distinct paragraphs, so same-paragraph contiguity
            # partitions them into three groups regardless of dates.
            by_gid: dict[int, set[str]] = {}
            for rev in doc.list_revisions():
                assert rev.group_id is not None and rev.paragraph_ref is not None
                assert rev.group_source == "inferred"
                by_gid.setdefault(rev.group_id, set()).add(rev.paragraph_ref)
            assert len(by_gid) == 3
            assert all(len(paras) == 1 for paras in by_gid.values())
            assert len(set().union(*by_gid.values())) == 3  # one paragraph each

    def test_multi_hunk_edit_straddling_seconds_stays_one_group(self, temp_docx, ticking_clock):
        # One logical edit stamps ONE w:date even when the wall clock ticks
        # between its internal injection calls (frozen_timestamp): without
        # the freeze, a multi-hunk rewrite crossing a second boundary would
        # reconstruct as several inferred groups after reopen.
        new_text = "A slow red cat crawls beneath the energetic dog."
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            result = doc.rewrite_paragraph(ref, new_text)
            assert len(result.revision_ids) > 2
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            assert len({rev.date for rev in revisions}) == 1  # one frozen stamp
            gid = revisions[0].group_id
            assert gid is not None
            assert all(rev.group_id == gid for rev in revisions)
            doc.accept_group(gid)
            assert _paragraph_text(doc, 1) == new_text

    def test_same_paragraph_same_second_two_groups(self, temp_docx, frozen_clock):
        # Headline of ISSUES.md #53: two separate edits to the SAME
        # paragraph while the wall clock sits in one second get
        # collision-bumped dates (T and T+1), so reconstruction keeps
        # them apart after reopen instead of over-merging.
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            first = doc.replace("quick", "speedy", paragraph=ref)
            second = doc.replace("lazy", "sleepy", paragraph=first)
            assert first.group_id != second.group_id  # two live groups...
            assert {rev.date for rev in doc.list_revisions()} == {
                datetime(2025, 6, 1, 12, 0, 0, tzinfo=timezone.utc),
                datetime(2025, 6, 1, 12, 0, 1, tzinfo=timezone.utc),
            }
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            assert len(revisions) == 4
            gids = {rev.group_id for rev in revisions}
            assert len(gids) == 2 and None not in gids  # ...two after reopen

    def test_rump_group_after_partial_accept(self, temp_docx):
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            doc.replace("quick", "speedy", paragraph=ref)
            doc.save()

        # A partial accept, same XML result as Word accepting one revision.
        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            assert len(revisions) == 2
            assert doc.accept_revision(revisions[0].id)
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            assert len(revisions) == 1
            gid = revisions[0].group_id
            assert gid is not None
            assert revisions[0].group_source == "inferred"
            # The rump group resolves without error.
            assert doc.accept_group(gid) == 1
            assert "speedy brown fox" in _paragraph_text(doc, 1)
            assert doc.list_revisions() == []

    def test_group_source_recorded_in_live_session(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.replace("quick", "speedy", paragraph=ref)

        for rev in doc.list_revisions():
            assert rev.group_id == result.group_id
            assert rev.group_source == "recorded"
            assert "(inferred)" not in repr(rev)


class TestChangesetDateStamping:
    """Collision-bumped w:date stamping at changeset granularity (ISSUES.md #53).

    One changeset = one single edit call or one whole batch_edit/batch_rewrite
    call. Within one open session, distinct changesets by the same author
    never share a second (the stamp is bumped past the previous one on
    collision), so parse-time reconstruction (#46) never merges them; all ops
    of one batch call share one date by design.
    """

    def test_batch_ops_share_one_date(self, temp_docx, ticking_clock):
        # Adversarial clock: every now() lands in a new second. One
        # batch_edit call is one changeset — all its revisions must still
        # carry ONE date, and same-paragraph ops of that one call merge
        # into one inferred group after reopen (changeset granularity,
        # pinned as a contract).
        with Document.open(temp_docx) as doc:
            ref_fox = find_ref(doc, "quick brown fox")
            ref_sample = find_ref(doc, "sample document for testing")
            doc.batch_edit([
                EditOperation.replace("quick", "speedy", paragraph=ref_fox),
                EditOperation.replace("lazy", "sleepy", paragraph=ref_fox),
                EditOperation.delete("sample ", paragraph=ref_sample),
            ])
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            assert len({rev.date for rev in revisions}) == 1  # one changeset date
            by_para: dict[str, set[int]] = {}
            for rev in revisions:
                assert rev.group_id is not None and rev.paragraph_ref is not None
                by_para.setdefault(rev.paragraph_ref, set()).add(rev.group_id)
            assert len(by_para) == 2
            assert all(len(gids) == 1 for gids in by_para.values())

    def test_batch_rewrite_shares_one_date(self, temp_docx, ticking_clock):
        # One batch_rewrite call is one changeset: a 50-paragraph batch must
        # not push dates 49 s into the future. Rewrites always land in
        # distinct paragraphs (duplicates are rejected), so contiguity still
        # reconstructs one group per paragraph despite the shared date.
        with Document.open(temp_docx) as doc:
            ref_fox = find_ref(doc, "quick brown fox")
            ref_sample = find_ref(doc, "sample document for testing")
            doc.batch_rewrite([
                (ref_fox, "A slow red cat crawls beneath the energetic dog."),
                (ref_sample, "Entirely different sample text."),
            ])
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            assert len({rev.date for rev in revisions}) == 1
            by_para: dict[str, set[int]] = {}
            for rev in revisions:
                assert rev.group_id is not None and rev.paragraph_ref is not None
                by_para.setdefault(rev.paragraph_ref, set()).add(rev.group_id)
            assert len(by_para) == 2
            assert all(len(gids) == 1 for gids in by_para.values())

    def test_no_bump_when_seconds_differ(self, temp_docx, settable_clock):
        # The bump only fires on collision: once real time moves past the
        # previous stamp, the true wall clock is used, not previous+1.
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            first = doc.replace("quick", "speedy", paragraph=ref)
            settable_clock.current = datetime(2025, 6, 1, 12, 0, 5)
            doc.replace("lazy", "sleepy", paragraph=first)
            assert {rev.date for rev in doc.list_revisions()} == {
                datetime(2025, 6, 1, 12, 0, 0, tzinfo=timezone.utc),
                datetime(2025, 6, 1, 12, 0, 5, tzinfo=timezone.utc),
            }

    def test_drift_bounded_by_changesets_per_second(self, temp_docx, settable_clock):
        # Drift is bounded by changesets per real second: N single edits in
        # one second stamp T, T+1, ..., T+N-1, and the drift collapses as
        # soon as real time passes the bumped value.
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            first = doc.replace("quick", "speedy", paragraph=ref)
            doc.replace("lazy", "sleepy", paragraph=first)
            doc.delete("sample ", paragraph=find_ref(doc, "sample document for testing"))
            settable_clock.current = datetime(2025, 6, 1, 12, 0, 10)
            doc.replace("well-structured", "tidy", paragraph=find_ref(doc, "well-structured document"))
            assert {rev.date for rev in doc.list_revisions()} == {
                datetime(2025, 6, 1, 12, 0, 0, tzinfo=timezone.utc),
                datetime(2025, 6, 1, 12, 0, 1, tzinfo=timezone.utc),
                datetime(2025, 6, 1, 12, 0, 2, tzinfo=timezone.utc),
                datetime(2025, 6, 1, 12, 0, 10, tzinfo=timezone.utc),
            }

    def test_clock_regression_keeps_dates_monotonic(self, temp_docx, settable_clock):
        # A clock step back (e.g. NTP) is treated as a collision: dates
        # never go backwards, pinned at previous+1 per changeset until
        # the clock catches up.
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            first = doc.replace("quick", "speedy", paragraph=ref)
            settable_clock.current = datetime(2025, 6, 1, 11, 59, 30)
            doc.replace("lazy", "sleepy", paragraph=first)
            assert {rev.date for rev in doc.list_revisions()} == {
                datetime(2025, 6, 1, 12, 0, 0, tzinfo=timezone.utc),
                datetime(2025, 6, 1, 12, 0, 1, tzinfo=timezone.utc),
            }

    def test_two_batches_same_second_get_distinct_dates(self, temp_docx, frozen_clock):
        # Two batch_edit calls are two changesets even inside one real
        # second — distinct dates, two inferred groups after reopen.
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            (first,) = doc.batch_edit([EditOperation.replace("quick", "speedy", paragraph=ref)])
            doc.batch_edit([EditOperation.replace("lazy", "sleepy", paragraph=first)])
            assert len({rev.date for rev in doc.list_revisions()}) == 2
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            assert len(revisions) == 4
            gids = {rev.group_id for rev in revisions}
            assert len(gids) == 2 and None not in gids

    def test_nested_frozen_timestamp_reuses_outer_stamp(self, temp_xml):
        # frozen_timestamp is reentrant-by-reuse: an inner scope joins the
        # enclosing changeset (batch wrapper + per-op _grouped). Only the
        # outermost entry allocates a stamp; only the outermost exit clears.
        editor = _make_manager(temp_xml("<w:p><w:r><w:t>x</w:t></w:r></w:p>")).editor
        with editor.frozen_timestamp():
            outer = editor._frozen_timestamp
            assert outer is not None
            with editor.frozen_timestamp():
                assert editor._frozen_timestamp == outer
            assert editor._frozen_timestamp == outer
        assert editor._frozen_timestamp is None

    def test_w16du_dateutc_matches_bumped_date(self, temp_xml, frozen_clock):
        # w:date and w16du:dateUtc read the same stamp variable — a bumped
        # changeset must bump both, never just w:date.
        manager = _make_manager(temp_xml("<w:p><w:r><w:t>alpha beta gamma</w:t></w:r></w:p>"))
        manager.replace_text("alpha", "ALPHA")
        manager.replace_text("gamma", "GAMMA")
        elements = [e for tag in ("w:ins", "w:del") for e in manager.editor.dom.getElementsByTagName(tag)]
        assert all(e.getAttribute("w16du:dateUtc") == e.getAttribute("w:date") for e in elements)
        assert {e.getAttribute("w:date") for e in elements} == {
            "2025-06-01T12:00:00Z",
            "2025-06-01T12:00:01Z",
        }


class TestMonotonicChangeIds:
    def test_removed_max_id_is_not_reused(self, temp_xml):
        body = "<w:p><w:r><w:t>one two three</w:t></w:r></w:p>"
        manager = _make_manager(temp_xml(body))

        first = manager.replace_text("two", "TWO")
        first_gid = manager.group_id_of(first)
        assert first_gid is not None
        first_ids = manager.group_revisions(first_gid)

        # Reject everything: the max-id elements leave the DOM entirely.
        manager.reject_group(first_gid)
        assert manager.list_revisions() == []

        second = manager.replace_text("three", "THREE")
        second_gid = manager.group_id_of(second)
        assert second_gid is not None
        second_ids = manager.group_revisions(second_gid)

        # Without the monotonic floor these ids would be reissued and the
        # registry would point two groups at the same w:id values.
        assert not (set(first_ids) & set(second_ids))
        assert min(second_ids) > max(first_ids)

    def test_nested_grouped_raises(self, temp_xml):
        body = "<w:p><w:r><w:t>one two three</w:t></w:r></w:p>"
        manager = _make_manager(temp_xml(body))

        with pytest.raises(RuntimeError, match="does not nest"):
            with manager._grouped():
                with manager._grouped():
                    pass  # pragma: no cover


class TestChangesetTier:
    """Changeset tier (ISSUES.md #54): one whole call ⊇ ≥1 group ⊇ revisions.

    A changeset is the ``(author, date)`` equivalence class over groups — a
    global (non-contiguous) class, not a contiguous run. A single edit is a
    one-group changeset; a whole ``batch_edit``/``batch_rewrite`` is one
    changeset over all its groups. Recorded live and reconstructed on reopen,
    exactly like the group tier one level down. There is no fourth tier.
    """

    def test_single_edit_is_one_group_one_changeset(self, doc):
        ref = find_ref(doc, "quick brown fox")
        result = doc.replace("quick", "speedy", paragraph=ref)

        assert result.group_id is not None
        assert result.changeset_id is not None
        for rev in doc.list_revisions():
            assert rev.group_id == result.group_id
            assert rev.changeset_id == result.changeset_id
            assert rev.changeset_source == "recorded"
            assert f", cs={result.changeset_id}" in repr(rev)
        # accept_changeset applies the whole edit as a unit.
        assert doc.accept_changeset(result.changeset_id) == 2
        assert "speedy brown fox" in _paragraph_text(doc, 1)
        assert doc.list_revisions() == []

    def test_batch_edit_is_one_changeset_over_many_groups(self, doc):
        ref1 = find_ref(doc, "quick brown fox")
        ref2 = find_ref(doc, "sample document for testing")
        results = doc.batch_edit([
            EditOperation.replace("quick", "speedy", paragraph=ref1),
            EditOperation.delete("sample ", paragraph=ref2),
        ])

        # Two groups (one per op), non-contiguous (different paragraphs), one
        # changeset — the headline: a changeset is not a contiguous run.
        assert results[0].group_id != results[1].group_id
        assert results[0].changeset_id is not None
        assert results[0].changeset_id == results[1].changeset_id
        assert {rev.changeset_id for rev in doc.list_revisions()} == {results[0].changeset_id}

        # reject_changeset undoes the entire batch (both groups' 3 revisions).
        assert doc.reject_changeset(results[0].changeset_id) == 3
        assert "quick brown fox" in _paragraph_text(doc, 1)
        assert "sample document for testing" in doc.get_visible_text()
        assert doc.list_revisions() == []

    def test_batch_rewrite_is_one_changeset_one_group_per_rewrite(self, doc):
        ref1 = find_ref(doc, "quick brown fox")
        ref2 = find_ref(doc, "sample document for testing")
        results = doc.batch_rewrite([
            (ref1, "A slow red cat sits."),
            (ref2, "This paragraph was fully rewritten for the test."),
        ])

        # The inner rewrite_paragraph _changeset() calls no-op via reentrancy:
        # one changeset for the whole call, one group per rewrite.
        assert results[0].group_id != results[1].group_id
        assert results[0].changeset_id is not None
        assert results[0].changeset_id == results[1].changeset_id
        assert {rev.changeset_id for rev in doc.list_revisions()} == {results[0].changeset_id}

    def test_standalone_rewrite_is_its_own_changeset(self, doc):
        # The reentrancy flip side: two standalone rewrite_paragraph calls are
        # two changesets (each is its own outermost boundary).
        ref1 = find_ref(doc, "quick brown fox")
        ref2 = find_ref(doc, "sample document for testing")
        r1 = doc.rewrite_paragraph(ref1, "A slow red cat sits.")
        r2 = doc.rewrite_paragraph(ref2, "Entirely different sample text.")

        assert r1.changeset_id is not None and r2.changeset_id is not None
        assert r1.changeset_id != r2.changeset_id
        assert r1.group_id != r2.group_id

    def test_unknown_changeset_raises(self, doc):
        with pytest.raises(RevisionError, match="Unknown changeset: 999"):
            doc.accept_changeset(999)
        with pytest.raises(RevisionError, match="Unknown changeset: 999"):
            doc.reject_changeset(999)

    def test_changeset_flag_resets_after_edit(self, doc):
        ref = find_ref(doc, "quick brown fox")
        doc.replace("quick", "speedy", paragraph=ref)
        assert doc._revision_manager._in_changeset is False

    def test_batch_edit_rollback_restores_changeset_registry(self, doc):
        ref1 = find_ref(doc, "quick brown fox")
        manager = doc._revision_manager
        cs_counter_before = manager._changeset_counter

        with pytest.raises(BatchOperationError):
            doc.batch_edit([
                EditOperation.replace("quick", "speedy", paragraph=ref1),
                EditOperation.delete("no such text anywhere", paragraph=ref1),
            ])

        # No ghost changeset from the failed batch; the flag is released.
        assert manager._changeset_counter == cs_counter_before
        assert manager._changesets == {}
        assert manager._group_changesets == {}
        assert manager._changeset_sources == {}
        assert manager._in_changeset is False
        assert doc.list_revisions() == []

    def test_roundtrip_two_paragraph_batch_plus_single_edit(self, temp_docx, frozen_clock):
        # Headline round-trip: a batch over two paragraphs (one changeset,
        # two groups) plus a separate single edit (a second changeset). After
        # save + reopen: exactly 2 inferred changesets / 3 groups, and each
        # changeset resolves as a unit.
        with Document.open(temp_docx) as doc:
            ref1 = find_ref(doc, "quick brown fox")
            ref2 = find_ref(doc, "sample document for testing")
            results = doc.batch_edit([
                EditOperation.replace("quick", "speedy", paragraph=ref1),
                EditOperation.delete("sample ", paragraph=ref2),
            ])
            third = doc.replace("well-structured", "tidy", paragraph=find_ref(doc, "well-structured document"))
            # Recorded: 3 groups, 2 changesets (the batch's two groups share
            # one changeset; the single edit gets a collision-bumped date and
            # its own changeset).
            assert len({r.group_id for r in [*results, third]}) == 3
            assert results[0].changeset_id == results[1].changeset_id
            assert third.changeset_id not in {results[0].changeset_id}
            manager = doc._revision_manager
            assert len(manager._changesets) == 2
            assert len(manager._groups) == 3
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            by_cs: dict[int, set[int]] = {}
            for rev in revisions:
                assert rev.changeset_id is not None
                assert rev.group_id is not None
                assert rev.changeset_source == "inferred"
                by_cs.setdefault(rev.changeset_id, set()).add(rev.group_id)
            # Exactly two changesets; three groups total.
            assert len(by_cs) == 2
            assert len({gid for gids in by_cs.values() for gid in gids}) == 3
            # One changeset spans two groups (the batch), the other spans one.
            assert sorted(len(gids) for gids in by_cs.values()) == [1, 2]

            batch_cs = next(cs for cs, gids in by_cs.items() if len(gids) == 2)
            single_cs = next(cs for cs, gids in by_cs.items() if len(gids) == 1)
            # Each changeset resolves as a unit, independently.
            assert doc.accept_changeset(batch_cs) == 3  # replace (del+ins) + delete
            assert doc.reject_changeset(single_cs) == 2  # replace (del+ins) undone
            assert doc.list_revisions() == []
            text = doc.get_visible_text()
            assert "speedy brown fox" in text
            assert "sample " not in text
            assert "well-structured" in text  # the single edit was rejected

    def test_same_paragraph_batch_over_merges_after_reopen(self, temp_docx, frozen_clock):
        # Accepted imprecision (carried up from #46): two batch ops in the
        # SAME paragraph share (author, date) and paragraph, so #46 merges
        # them into ONE group on reopen — hence one changeset. Pinned so the
        # collapse is a known contract, not a silent regression.
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            results = doc.batch_edit([
                EditOperation.replace("quick", "speedy", paragraph=ref),
                EditOperation.replace("lazy", "sleepy", paragraph=ref),
            ])
            # Recorded: two groups, one changeset.
            assert results[0].group_id != results[1].group_id
            assert results[0].changeset_id == results[1].changeset_id
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            assert len(revisions) == 4
            gids = {rev.group_id for rev in revisions}
            cs_ids = {rev.changeset_id for rev in revisions}
            assert len(gids) == 1  # #46 merges same-paragraph same-date ops
            assert len(cs_ids) == 1 and None not in cs_ids

    def test_rump_changeset_after_partial_word_accept(self, temp_xml):
        # A changeset spanned two paragraphs (revs 1-4, one author+date).
        # Word accepted rev 1 (unwrapped it to plain text), leaving revs
        # 2,3,4. The survivors reconstruct as a rump changeset — its first
        # group lost a member — that reject_changeset resolves as a unit.
        survivors = (
            f"<w:p><w:r><w:t>one </w:t></w:r>{_del_xml(2, 'two ')}</w:p>"
            f"<w:p>{_ins_xml(3, 'three ')}{_del_xml(4, 'four')}</w:p>"
        )
        manager = _make_manager(temp_xml(survivors))

        # Two groups (one per paragraph), one changeset (shared author+date).
        assert len(manager._groups) == 2
        (cs_id,) = manager._changesets
        assert manager._changeset_sources[cs_id] == "inferred"
        assert set(manager.changeset_groups(cs_id)) == set(manager._groups)
        # The rump resolves without error (3 survivors across two groups).
        assert manager.reject_changeset(cs_id) == 3
        assert manager.list_revisions() == []

    def test_partially_resolved_changeset_is_rump_tolerant(self, temp_xml):
        # Resolve one group of a two-group changeset individually, then the
        # changeset still resolves the survivors (skips the resolved pair).
        body = (
            f"<w:p>{_ins_xml(1, 'one ')}{_del_xml(2, 'two ')}</w:p>"
            f"<w:p>{_ins_xml(3, 'three ')}{_del_xml(4, 'four')}</w:p>"
        )
        manager = _make_manager(temp_xml(body))
        (cs_id,) = manager._changesets

        first_group = manager.group_id_of(1)
        assert first_group is not None
        assert manager.accept_group(first_group) == 2  # revs 1,2 resolved

        assert manager.accept_changeset(cs_id) == 2  # only revs 3,4 remain
        assert manager.list_revisions() == []

    def test_reconstruction_partitions_by_author_date(self, temp_xml):
        # Same (author, date) in two paragraphs → two groups, ONE changeset
        # (global equivalence class). A different date → a distinct changeset.
        body = (
            f"<w:p>{_ins_xml(1, 'a')}</w:p>"
            f"<w:p>{_ins_xml(2, 'b')}</w:p>"  # same author+date as rev 1
            f"<w:p>{_ins_xml(3, 'c', date=DATE_B)}</w:p>"  # different date
        )
        manager = _make_manager(temp_xml(body))

        assert manager._groups == {1: (1,), 2: (2,), 3: (3,)}
        # Groups 1 and 2 share a changeset; group 3 is its own.
        cs1 = manager.changeset_id_of(1)
        cs2 = manager.changeset_id_of(2)
        cs3 = manager.changeset_id_of(3)
        assert cs1 is not None and cs3 is not None
        assert cs1 == cs2 and cs1 != cs3
        assert set(manager.changeset_groups(cs1)) == {1, 2}
        assert manager.changeset_groups(cs3) == (3,)
        assert all(src == "inferred" for src in manager._changeset_sources.values())

    def test_ungrouped_revision_has_no_changeset(self, temp_xml):
        # A revision with no date is ungroupable → no group, no changeset.
        body = (
            f"<w:p>{_ins_xml(1, 'one ')}"
            f'<w:ins w:id="2" w:author="{AUTHOR_A}"><w:r><w:t>no date</w:t></w:r></w:ins></w:p>'
        )
        manager = _make_manager(temp_xml(body))
        by_id = {rev.id: rev for rev in manager.list_revisions()}

        assert by_id[1].changeset_id is not None
        assert by_id[1].changeset_source == "inferred"
        assert by_id[2].group_id is None
        assert by_id[2].changeset_id is None
        assert by_id[2].changeset_source is None
        assert "cs=" not in repr(by_id[2])

    def test_foreign_revisions_carry_inferred_changesets(self, test_data_dir, temp_dir):
        # Every foreign revision now carries an inferred changeset, and
        # changeset membership never crosses an (author, date) boundary
        # (same changeset ⇒ same author+date — the equivalence-class contract).
        fixture = test_data_dir / "OXML_TrackChanges_Test.docx"
        dest = temp_dir / "foreign.docx"
        shutil.copy(fixture, dest)
        with Document.open(dest) as doc:
            revisions = doc.list_revisions()
            assert revisions
            assert all(rev.changeset_id is not None for rev in revisions)
            assert all(rev.changeset_source == "inferred" for rev in revisions)
            assert all(f", cs={rev.changeset_id}" in repr(rev) for rev in revisions)

            by_cs: dict[int, set[tuple[str, object]]] = {}
            for rev in revisions:
                assert rev.changeset_id is not None
                by_cs.setdefault(rev.changeset_id, set()).add((rev.author, rev.date))
            assert all(len(keys) == 1 for keys in by_cs.values())


class TestAcceptPathIndex:
    """The group/changeset accept path builds one w:id->element index per call
    instead of scanning the whole document per member (ISSUES.md #57).

    Each resolution does exactly two Document-level walks (one w:ins scan, one
    w:del scan) to build the index, no matter how many revisions it spans.
    Pre-#57, accept_revision/reject_revision each did up to two full-document
    walks, repeated per member per pass (O(members x doc)) — a 240-revision
    accept_changeset paid ~960 walks.
    """

    # A host insertion whose content includes a nested deletion, both by the
    # same author+date so reconstruction bundles them into ONE inferred group.
    NESTED_BODY = (
        f'<w:p><w:ins w:id="1" w:author="{AUTHOR_A}" w:date="{DATE_A}">'
        "<w:r><w:t>alpha </w:t></w:r>"
        f'<w:del w:id="2" w:author="{AUTHOR_A}" w:date="{DATE_A}">'
        "<w:r><w:delText>beta </w:delText></w:r></w:del>"
        "<w:r><w:t>gamma</w:t></w:r></w:ins></w:p>"
    )

    # Same nesting, but the host id (5) is HIGHER than the nested id (2), so
    # reverse-id order processes the host FIRST — before the nested member is
    # ever resolved on its own. This is the branch _is_in_document is written
    # for (a member detaching together with its already-resolved host), which
    # NESTED_BODY (host id 1 < nested id 2) never reaches.
    NESTED_BODY_HOST_HIGH = (
        f'<w:p><w:ins w:id="5" w:author="{AUTHOR_A}" w:date="{DATE_A}">'
        "<w:r><w:t>alpha </w:t></w:r>"
        f'<w:del w:id="2" w:author="{AUTHOR_A}" w:date="{DATE_A}">'
        "<w:r><w:delText>beta </w:delText></w:r></w:del>"
        "<w:r><w:t>gamma</w:t></w:r></w:ins></w:p>"
    )

    @staticmethod
    def _accepted_text(manager: RevisionManager) -> str:
        """Accepted-view text of the manager's DOM (w:delText excluded)."""
        return "".join(
            node.firstChild.data
            for node in manager.editor.dom.getElementsByTagName("w:t")
            if node.firstChild is not None
        )

    @pytest.mark.parametrize("method", ["accept_changeset", "reject_changeset"])
    def test_changeset_resolution_builds_index_once(self, temp_docx, monkeypatch, method):
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            # One batch = one changeset over three groups = six revisions.
            doc.batch_edit([
                EditOperation.replace("quick", "speedy", paragraph=ref),
                EditOperation.replace("brown", "red", paragraph=ref),
                EditOperation.replace("lazy", "sleepy", paragraph=ref),
            ])
            changeset_id = doc.list_revisions()[0].changeset_id
            assert changeset_id is not None
            assert len(doc.list_revisions()) == 6

            walks = count_dom_walks(monkeypatch)
            assert getattr(doc, method)(changeset_id) == 6
            # One index build, not one scan per member per pass.
            assert walks == ["w:ins", "w:del"]
            assert doc.list_revisions() == []

    @pytest.mark.parametrize("method", ["accept_group", "reject_group"])
    def test_group_resolution_builds_index_once(self, temp_docx, monkeypatch, method):
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            result = doc.replace("quick", "speedy", paragraph=ref)  # one group, two revs

            walks = count_dom_walks(monkeypatch)
            assert getattr(doc, method)(result.group_id) == 2
            assert walks == ["w:ins", "w:del"]
            assert doc.list_revisions() == []

    def test_reject_group_with_nested_member_counts_once(self, temp_xml):
        # Rejecting the host insertion removes its whole subtree; the nested
        # deletion (id 2, resolved first in reverse-id order) detaches with the
        # host and must be counted exactly once. Connectivity — not a fresh
        # per-member scan — is what stops the detached member being re-resolved
        # (which would mutate a detached subtree and over-count).
        manager = _make_manager(temp_xml(self.NESTED_BODY))
        assert manager.group_revisions(1) == (1, 2)  # host ins + nested del

        assert manager.reject_group(1) == 2
        assert manager.list_revisions() == []
        assert self._accepted_text(manager) == ""

    def test_accept_group_with_nested_member_counts_once(self, temp_xml):
        # Accepting keeps the insertion (unwrapped) and applies the nested
        # deletion (removed): both members counted once, neither skipped.
        manager = _make_manager(temp_xml(self.NESTED_BODY))
        assert manager.group_revisions(1) == (1, 2)

        assert manager.accept_group(1) == 2
        assert manager.list_revisions() == []
        assert self._accepted_text(manager) == "alpha gamma"

    def test_reject_group_host_removed_before_nested_is_resolved(self, temp_xml):
        # host id 5 > nested id 2, so reverse-id order rejects the host w:ins
        # FIRST, removing its whole subtree — the nested w:del is gone before it
        # is ever resolved on its own. The connectivity guard skips it (count 1,
        # not 2, and no mutation of a detached subtree). This order-dependent
        # count is pre-existing and identical under the old fresh-scan path (a
        # fresh scan would likewise not find the removed nested del).
        manager = _make_manager(temp_xml(self.NESTED_BODY_HOST_HIGH))
        assert manager.group_revisions(1) == (5, 2)  # host ins first, nested del

        assert manager.reject_group(1) == 1  # host applied; nested detached with it
        assert manager.list_revisions() == []
        assert self._accepted_text(manager) == ""

    def test_accept_group_host_high_still_resolves_both(self, temp_xml):
        # host id 5 > nested id 2: accepting the host UNWRAPS it (content, incl.
        # the nested del, stays attached), so the nested member is still live
        # and gets resolved on its own — both counted, unlike the reject case.
        manager = _make_manager(temp_xml(self.NESTED_BODY_HOST_HIGH))
        assert manager.group_revisions(1) == (5, 2)

        assert manager.accept_group(1) == 2
        assert manager.list_revisions() == []
        assert self._accepted_text(manager) == "alpha gamma"
