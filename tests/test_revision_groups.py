"""Tests for revision grouping (ISSUES.md #37).

Every logical edit operation registers the revisions it creates as one
in-memory revision group, exposed via EditResult (a str subclass carrying
``group_id``/``revision_ids``) and resolvable as a unit with
``accept_group``/``reject_group``. The headline failure this prevents:
accepting only some of a ``rewrite_paragraph``'s revisions garbles the
paragraph, because each revision is a diff hunk, not a self-contained edit.

Groups are per-open-Document: foreign or pre-session revisions report
``group_id=None`` and unknown group ids raise RevisionError.
"""

import shutil
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

    def test_foreign_revisions_have_no_group(self, test_data_dir, temp_dir):
        fixture = test_data_dir / "OXML_TrackChanges_Test.docx"
        dest = temp_dir / "foreign.docx"
        shutil.copy(fixture, dest)
        with Document.open(dest) as doc:
            revisions = doc.list_revisions()
            assert revisions
            assert all(rev.group_id is None for rev in revisions)
            assert all("group=" not in repr(rev) for rev in revisions)


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
    def test_groups_do_not_survive_reopen(self, temp_docx):
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            result = doc.replace("quick", "speedy", paragraph=ref)
            group_id = result.group_id
            assert group_id is not None
            doc.save()

        with Document.open(temp_docx) as doc:
            revisions = doc.list_revisions()
            assert revisions  # the tracked change persisted in the file
            assert all(rev.group_id is None for rev in revisions)
            with pytest.raises(RevisionError):
                doc.accept_group(group_id)

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

        result = doc.replace(" ZE9", " QX7", paragraph=first)

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

    def test_presession_middle_delete_leaves_tail_ungrouped(self, temp_docx):
        # The origin insertion predates this session (reopen dropped its
        # group), so its split-off tail has no group to join — and must not
        # be claimed into a phantom group by the delete operation.
        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            doc.insert_after("fox", " AAA MID BBB", paragraph=ref)
            doc.save()

        with Document.open(temp_docx) as doc:
            ref = find_ref(doc, "quick brown fox")
            result = doc.delete("MID ", paragraph=ref)

            assert result.group_id is None
            assert doc._revision_manager._groups == {}
            revisions = doc.list_revisions()
            assert len(revisions) == 2  # truncated head + split-off tail
            assert all(rev.group_id is None for rev in revisions)

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
            # insertion — no trackable revisions, so no group at all.
            assert result.group_id is None
            assert doc._revision_manager._groups == {}
            assert all(rev.group_id is None for rev in doc.list_revisions())
            # Placement-agnostic content check: replacement position inside
            # an own insertion is the pre-existing ISSUES.md #31 ordering
            # bug, out of scope here.
            text = _paragraph_text(doc, 1)
            assert "BETA" in text and "beta" not in text
            assert "alpha" in text and "gamma" in text


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
