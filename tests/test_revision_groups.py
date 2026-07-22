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
from datetime import datetime
from pathlib import Path

import pytest
from conftest import find_ref

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
            # Same author and same (frozen) second everywhere: same-paragraph
            # contiguity is what partitions the edits into three groups.
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

    def test_same_paragraph_same_second_over_merges(self, temp_docx, frozen_clock):
        # Accepted relaxation, pinned: w:date has second precision, so two
        # separate edits to the SAME paragraph in the same second are
        # indistinguishable and reconstruct as one merged group.
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            first = doc.replace("quick", "speedy", paragraph=ref)
            second = doc.replace("lazy", "sleepy", paragraph=first)
            assert first.group_id != second.group_id  # two live groups...
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            assert len(revisions) == 4
            gids = {rev.group_id for rev in revisions}
            assert len(gids) == 1 and None not in gids  # ...one after reopen

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
