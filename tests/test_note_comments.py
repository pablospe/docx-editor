"""Tests for the ``note=`` rationale channel.

A ``note=`` on an edit method anchors a comment on the revisions that edit
creates, and that comment is deleted when the last of those revisions is
resolved — accepted or rejected alike.
"""

import warnings
import zipfile

import pytest
from conftest import NS, find_ref, replace_document_xml

from docx_editor import (
    BatchOperationError,
    Document,
    EditOperation,
    UnanchoredNoteWarning,
)


@pytest.fixture
def doc(temp_docx):
    """An open Document over the simple.docx copy, closed after the test."""
    document = Document.open(temp_docx)
    yield document
    document.close()


def _prev_element(node):
    """The previous sibling that is an element (skipping whitespace text)."""
    sib = node.previousSibling
    while sib is not None and sib.nodeType != sib.ELEMENT_NODE:
        sib = sib.previousSibling
    return sib


def _next_element(node):
    """The next sibling that is an element (skipping whitespace text)."""
    sib = node.nextSibling
    while sib is not None and sib.nodeType != sib.ELEMENT_NODE:
        sib = sib.nextSibling
    return sib


def _markers(doc, comment_id):
    """The (start, end) range markers of ``comment_id``, exactly one of each."""
    dom = doc._document_editor.dom

    def one(tag):
        found = [e for e in dom.getElementsByTagName(tag) if e.getAttribute("w:id") == str(comment_id)]
        assert len(found) == 1, f"expected one {tag} for comment {comment_id}, got {len(found)}"
        return found[0]

    return one("w:commentRangeStart"), one("w:commentRangeEnd")


def _marker_order(doc):
    """Every range marker and reference run in document order, as (tag, id)."""
    dom = doc._document_editor.dom
    tags = ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference")
    nodes = [(n, n.getAttribute("w:id")) for tag in tags for n in dom.getElementsByTagName(tag)]
    ordered = [n for n in dom.getElementsByTagName("*") if any(n is m for m, _ in nodes)]
    return [int(n.getAttribute("w:id")) for n in ordered]


def _marker_counts(doc):
    """(range starts, range ends, reference runs) left in document.xml."""
    dom = doc._document_editor.dom
    return tuple(
        len(dom.getElementsByTagName(tag)) for tag in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference")
    )


def _paragraph_index(doc, node):
    """1-based index of the paragraph containing ``node``."""
    dom = doc._document_editor.dom
    while node is not None and getattr(node, "tagName", "") != "w:p":
        node = node.parentNode
    return dom.getElementsByTagName("w:p").index(node) + 1


def _foreign_ins_docx(simple_docx, temp_dir):
    """A .docx whose only paragraph holds another author's pending insertion."""
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {NS}><w:body><w:p>"
        '<w:r><w:t xml:space="preserve">The term is </w:t></w:r>'
        '<w:ins w:id="90" w:author="Other" w:date="2024-01-01T00:00:00Z">'
        "<w:r><w:t>thirty (30) days</w:t></w:r></w:ins>"
        '<w:r><w:t xml:space="preserve"> from signing.</w:t></w:r>'
        "</w:p></w:body></w:document>"
    )
    src = temp_dir / "foreign.docx"
    replace_document_xml(simple_docx, src, doc_xml)
    return src


class TestAnchoring:
    """A note becomes a comment bracketing the edit's own revisions."""

    def test_replace_note_returns_a_live_comment_id(self, doc):
        result = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone: plainer word")

        assert isinstance(result.comment_id, int)
        assert [(c.id, c.text) for c in doc.list_comments()] == [(result.comment_id, "tone: plainer word")]

    def test_replace_markers_bracket_the_revisions(self, doc):
        """The markers are siblings *outside* the w:del/w:ins pair, not inside."""
        result = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")

        start, end = _markers(doc, result.comment_id)
        assert _next_element(start).tagName == "w:del"
        # The end marker's own next sibling is the reference run it came with;
        # what precedes it is the last revision of the edit.
        assert _prev_element(end).tagName == "w:ins"
        assert _next_element(end).getElementsByTagName("w:commentReference")

    def test_delete_note_anchors_on_the_deletion(self, doc):
        """A deletion's text is absent from the accepted text map, so only an
        element anchor can reach it at all."""
        result = doc.delete("lazy ", paragraph=find_ref(doc, "quick brown"), note="redundant")

        start, end = _markers(doc, result.comment_id)
        assert _next_element(start).tagName == "w:del"
        assert _prev_element(end).tagName == "w:del"

    @pytest.mark.parametrize("method", ["insert_after", "insert_before"])
    def test_insert_note_anchors_on_the_insertion(self, doc, method):
        result = getattr(doc, method)("fox", " (the animal)", paragraph=find_ref(doc, "quick brown"), note="clarify")

        start, end = _markers(doc, result.comment_id)
        assert _next_element(start).tagName == "w:ins"
        assert _prev_element(end).tagName == "w:ins"

    def test_rewrite_note_spans_first_through_last_revision(self, doc):
        """One comment covers the whole rewrite, not one per diff hunk."""
        result = doc.rewrite_paragraph(
            find_ref(doc, "quick brown"),
            "The swift brown fox leaps over the lazy dog.",
            note="two word swaps, one rationale",
        )

        assert len(result.revision_ids) > 2
        assert len(doc.list_comments()) == 1
        start, end = _markers(doc, result.comment_id)
        # Every revision of the group lies between the two markers.
        dom = doc._document_editor.dom
        order = [n for n in dom.getElementsByTagName("w:p")[1].childNodes if n.nodeType == n.ELEMENT_NODE]
        between = order[order.index(start) + 1 : order.index(end)]
        revisions = [n for n in between if n.tagName in ("w:ins", "w:del")]
        assert {int(n.getAttribute("w:id")) for n in revisions} == set(result.revision_ids)

    def test_edit_without_note_creates_no_comment(self, doc):
        """The feature is additive: no note, no comment, no comment_id."""
        result = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"))

        assert result.comment_id is None
        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    def test_note_anchors_outside_a_foreign_insertion(self, simple_docx, temp_dir):
        """Editing inside another author's pending insertion anchors *outside* it.

        Our ``w:del`` is nested in their ``w:ins``, but a marker left in there
        would be carried away when that author's proposal is rejected — so the
        span is hoisted to the outermost revision and both markers stay at
        run level.
        """
        doc = Document.open(_foreign_ins_docx(simple_docx, temp_dir), force_recreate=True)
        try:
            result = doc.replace("thirty (30)", "sixty (60)", paragraph=find_ref(doc, "term is"), note="extend")

            assert isinstance(result.comment_id, int)
            start, end = _markers(doc, result.comment_id)
            assert start.parentNode.tagName == "w:p"
            foreign = _next_element(start)
            assert foreign.tagName == "w:ins" and foreign.getAttribute("w:author") == "Other"
            assert _prev_element(end).tagName == "w:ins"

            # And it is still reaped when the edit is rejected.
            assert result.group_id is not None
            doc.reject_group(result.group_id)
            assert doc.list_comments() == []
        finally:
            doc.close()

    def test_rejecting_the_host_insertion_strands_no_marker(self, simple_docx, temp_dir):
        """The reason for hoisting: their reject must not break our range."""
        doc = Document.open(_foreign_ins_docx(simple_docx, temp_dir), force_recreate=True)
        try:
            result = doc.replace("thirty (30)", "sixty (60)", paragraph=find_ref(doc, "term is"), note="extend")

            doc.reject_all(author="Other")

            # Our own insertion outlives their sweep, so the note does too —
            # and it still has both of its range markers.
            assert [c.text for c in doc.list_comments()] == ["extend"]
            assert _marker_counts(doc) == (1, 1, 1)
            _markers(doc, result.comment_id)
        finally:
            doc.close()

    def test_note_spanning_a_tracked_split_brackets_both_paragraphs(self, doc):
        """A ``\\n`` replacement puts the markers in different paragraphs."""
        result = doc.replace(
            "jumps",
            "leaps\nAnd then it rests",
            paragraph=find_ref(doc, "quick brown"),
            note="break the sentence in two",
        )

        start, end = _markers(doc, result.comment_id)
        assert _paragraph_index(doc, start) == 2
        assert _paragraph_index(doc, end) == 3

        # delete_comment locates both markers by w:id, so the split does not
        # strand one of them.
        doc.reject_group(result.group_id)
        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)


class TestResolutionRemovesTheNote:
    """A note explains a proposal; resolving the proposal ends the note."""

    def test_reject_group_removes_comment_and_markers(self, doc):
        result = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")

        doc.reject_group(result.group_id)

        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    def test_accept_group_removes_comment_and_markers(self, doc):
        """Accepting ends the proposal too — a clean deliverable must not ship
        the agent's rationale as a live comment."""
        result = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")

        doc.accept_group(result.group_id)

        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    @pytest.mark.parametrize("verb", ["accept_changeset", "reject_changeset"])
    def test_changeset_resolution_removes_the_note(self, doc, verb):
        result = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")

        getattr(doc, verb)(result.changeset_id)

        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    @pytest.mark.parametrize("verb", ["accept_all", "reject_all"])
    def test_whole_document_sweep_removes_the_note(self, doc, verb):
        doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")

        getattr(doc, verb)()

        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    @pytest.mark.parametrize("verb", ["accept_revision", "reject_revision"])
    def test_single_revision_resolution_of_a_one_revision_group(self, doc, verb):
        """Resolving by revision id cleans up too — one invariant, not four."""
        result = doc.insert_after("fox", " (the animal)", paragraph=find_ref(doc, "quick brown"), note="clarify")
        assert len(result.revision_ids) == 1

        getattr(doc, verb)(result.revision_ids[0])

        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    def test_partial_group_resolution_keeps_the_note(self, doc):
        """A multi-revision edit keeps its note while any revision is pending."""
        result = doc.rewrite_paragraph(
            find_ref(doc, "quick brown"), "The swift brown fox leaps over the lazy dog.", note="two swaps"
        )

        doc.reject_revision(result.revision_ids[0])

        assert [c.text for c in doc.list_comments()] == ["two swaps"]

    def test_sweep_by_another_author_leaves_the_note(self, doc):
        """A foreign-author sweep touches neither our revisions nor our notes."""
        doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")

        doc.accept_all(author="Someone Else")

        assert [c.text for c in doc.list_comments()] == ["tone"]
        assert _marker_counts(doc) == (1, 1, 1)

    def test_sweep_by_our_own_author_removes_the_note(self, doc):
        doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")

        doc.accept_all(author=doc.author)

        assert doc.list_comments() == []

    def test_shared_note_survives_until_the_last_group_resolves(self, doc):
        """A deduped note outlives the first of the operations it explains."""
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]
        results = doc.batch_edit([
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="house style"),
            EditOperation.replace("sample", "example", paragraph=refs[1], note="house style"),
        ])
        assert results[0].comment_id == results[1].comment_id

        doc.reject_group(results[0].group_id)
        assert [c.text for c in doc.list_comments()] == ["house style"]

        doc.reject_group(results[1].group_id)
        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    def test_resolving_an_unrelated_group_keeps_the_note(self, doc):
        annotated = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")
        other = doc.replace("sample", "example", paragraph=find_ref(doc, "sample document"))

        doc.accept_group(other.group_id)

        assert [c.text for c in doc.list_comments()] == ["tone"]
        assert annotated.comment_id == doc.list_comments()[0].id

    def test_amending_the_annotated_insertion_away_removes_the_note(self, doc):
        """A later edit can amend the annotated insertion out of existence.

        No accept or reject call ever runs, so nothing but the next edit can
        notice; without cleanup there the rationale would ship in the saved
        file bracketing a document with no revisions at all.
        """
        annotated = doc.insert_after("fox", " ZE9", paragraph=find_ref(doc, "quick brown"), note="why inserted")
        assert annotated.comment_id is not None

        doc.delete(" ZE9", paragraph=str(annotated))

        assert doc.list_revisions() == []
        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    def test_batch_rewrite_amending_the_insertion_away_removes_the_note(self, doc):
        """batch_rewrite takes no note=, but its amendments empty groups too."""
        annotated = doc.insert_after("fox", " ZE9", paragraph=find_ref(doc, "quick brown"), note="why inserted")

        doc.batch_rewrite([(str(annotated), "The quick brown fox jumps over the lazy dog.")])

        assert doc.list_revisions() == []
        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    def test_a_foreign_reject_that_carries_our_revision_away_reaps_the_note(self, simple_docx, temp_dir):
        """Rejecting their insertion takes the w:del we nested inside it.

        Our group empties under a sweep filtered to *their* name, so narrowing
        cleanup to "groups this call resolved" would leave the rationale behind
        with nothing to explain. Liveness is the only safe authority.
        """
        doc = Document.open(_foreign_ins_docx(simple_docx, temp_dir), force_recreate=True)
        try:
            doc.delete("thirty (30)", paragraph=find_ref(doc, "term is"), note="why we cut it")

            doc.reject_all(author="Other")

            assert doc.list_revisions() == []
            assert doc.list_comments() == []
            assert _marker_counts(doc) == (0, 0, 0)
        finally:
            doc.close()

    def test_a_shared_note_moves_onto_a_redline_that_is_still_pending(self, doc):
        """Resolving the anchoring operation must not leave the rationale
        bracketing text nobody changed."""
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]
        results = doc.batch_edit([
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="house style"),
            EditOperation.replace("sample", "example", paragraph=refs[1], note="house style"),
        ])
        assert _paragraph_index(doc, _markers(doc, results[0].comment_id)[0]) == 2

        doc.reject_group(results[0].group_id)

        comment_id = results[1].comment_id
        start, end = _markers(doc, comment_id)
        assert _paragraph_index(doc, start) == 3  # with the surviving redline
        assert _next_element(start).tagName == "w:del"
        assert _prev_element(end).tagName == "w:ins"
        assert _marker_counts(doc) == (1, 1, 1)

        doc.reject_group(results[1].group_id)
        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    def test_a_deleted_note_is_never_resurrected(self, doc):
        """delete_comment() on a note ends it: re-placing markers for a comment
        with no body left would make the file unreadable."""
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]
        results = doc.batch_edit([
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="house style"),
            EditOperation.replace("sample", "example", paragraph=refs[1], note="house style"),
        ])
        assert doc.delete_comment(results[0].comment_id) is True
        assert _marker_counts(doc) == (0, 0, 0)

        doc.reject_group(results[0].group_id)

        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)
        doc.reject_group(results[1].group_id)
        assert _marker_counts(doc) == (0, 0, 0)

    def test_re_anchoring_moves_the_whole_thread(self, doc):
        """A reply's markers are seated on its parent's, so they move with it."""
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]
        results = doc.batch_edit([
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="house style"),
            EditOperation.replace("sample", "example", paragraph=refs[1], note="house style"),
        ])
        reply_id = doc.reply_to_comment(results[0].comment_id, "agreed")

        doc.reject_group(results[0].group_id)

        for comment_id in (results[0].comment_id, reply_id):
            start, end = _markers(doc, comment_id)
            assert _paragraph_index(doc, start) == 3
            assert _paragraph_index(doc, end) == 3
        assert [r.text for r in doc.list_comments()[0].replies] == ["agreed"]

    def test_a_paragraph_mark_only_op_shares_the_note_without_extending_its_life(self, doc):
        """A pure split cannot host a comment marker, so it inherits the shared
        id but is not registered as a group keeping the comment alive."""
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]
        results = doc.batch_edit([
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="house style"),
            EditOperation.insert_after("document", "\n", paragraph=refs[1], note="house style"),
        ])
        assert results[0].comment_id == results[1].comment_id

        doc.reject_group(results[0].group_id)

        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)
        # The split itself is untouched — only its rationale went.
        assert doc.list_revisions() != []

    def test_re_anchoring_reproduces_the_thread_layout(self, doc):
        """A moved thread must be seated exactly as reply_to_comment seats it:
        every reply on its own parent's markers, not flattened onto the root."""
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]
        results = doc.batch_edit([
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="house style"),
            EditOperation.replace("sample", "example", paragraph=refs[1], note="house style"),
        ])
        root = results[0].comment_id
        child = doc.reply_to_comment(root, "agreed")
        grandchild = doc.reply_to_comment(child, "and here too")
        sibling = doc.reply_to_comment(root, "one more")
        before = _marker_order(doc)

        doc.reject_group(results[0].group_id)

        assert _marker_order(doc) == before
        assert {root, child, grandchild, sibling} == set(before)
        for comment_id in (root, child, grandchild, sibling):
            start, end = _markers(doc, comment_id)
            assert _paragraph_index(doc, start) == 3
            assert _paragraph_index(doc, end) == 3

    def test_a_note_whose_own_anchor_is_amended_away_goes_with_it(self, doc):
        """The group survives — its paragraph mark is still pending — but a
        marker cannot bracket a paragraph mark, so an empty range would be all
        that was left of the rationale."""
        annotated = doc.insert_after(
            "document", "\nNew line", paragraph=find_ref(doc, "sample document"), note="split it"
        )
        assert annotated.comment_id is not None

        doc.delete("New line", paragraph=annotated.refs[1])

        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)
        assert doc.list_revisions() != []  # the split is still pending

    def test_deleting_a_note_comment_takes_its_replies(self, doc):
        """Deleting a thread parent must not orphan its replies: reply_ids
        walks paraIdParent, so a stranded descendant is unreachable after."""
        annotated = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")
        reply_id = doc.reply_to_comment(annotated.comment_id, "agreed")
        doc.reply_to_comment(reply_id, "and here too")

        assert doc.delete_comment(annotated.comment_id) is True

        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    def test_a_shared_note_dies_when_no_survivor_can_hold_the_anchor(self, doc):
        """A registered group can still be whittled down to a paragraph mark.

        The second operation is anchorable when the note lands, so it is
        registered; amending its inserted text away leaves it live but with
        nothing a marker can bracket. Resolving the anchor then has nowhere
        honest to move to.
        """
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]
        results = doc.batch_edit([
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="house style"),
            EditOperation.insert_after("document", "\nNew line", paragraph=refs[1], note="house style"),
        ])
        assert results[0].comment_id == results[1].comment_id

        # Amend the inserted text away: the group keeps only its paragraph mark.
        doc.delete("New line", paragraph=results[1].refs[1])
        assert [c.text for c in doc.list_comments()] == ["house style"]

        doc.reject_group(results[0].group_id)

        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)
        assert doc.list_revisions() != []  # the split is still pending

    def test_a_shared_note_stays_put_when_its_own_anchor_survives(self, doc):
        """Only a dead anchor moves: resolving the other operation is not a
        reason to churn the markers."""
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]
        results = doc.batch_edit([
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="house style"),
            EditOperation.replace("sample", "example", paragraph=refs[1], note="house style"),
        ])

        doc.reject_group(results[1].group_id)

        start, _ = _markers(doc, results[0].comment_id)
        assert _paragraph_index(doc, start) == 2
        assert _next_element(start).tagName == "w:del"

    def test_a_partly_amended_group_keeps_its_note(self, doc):
        """Amending one revision away is not resolution: the group's deletion
        is still pending, so the rationale still has something to explain."""
        annotated = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")

        doc.delete("swift", paragraph=str(annotated))

        assert [c.text for c in doc.list_comments()] == ["tone"]
        assert annotated.comment_id == doc.list_comments()[0].id


class TestBatchNotes:
    """One comment per rationale, per call — not per redline."""

    def test_same_note_across_ops_makes_one_comment(self, doc):
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document"), find_ref(doc, "well-structured")]
        results = doc.batch_edit([
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="house style"),
            EditOperation.replace("sample", "example", paragraph=refs[1], note="house style"),
            EditOperation.replace("well-structured", "well structured", paragraph=refs[2], note="house style"),
        ])

        assert len({r.comment_id for r in results}) == 1
        assert [c.text for c in doc.list_comments()] == ["house style"]

    def test_distinct_notes_make_distinct_comments(self, doc):
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document"), find_ref(doc, "well-structured")]
        results = doc.batch_edit([
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="tone"),
            EditOperation.replace("sample", "example", paragraph=refs[1], note="precision"),
            EditOperation.replace("well-structured", "well structured", paragraph=refs[2], note="hyphenation"),
        ])

        assert len({r.comment_id for r in results}) == 3
        assert sorted(c.text for c in doc.list_comments()) == ["hyphenation", "precision", "tone"]

    def test_ops_without_a_note_report_none(self, doc):
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]
        results = doc.batch_edit([
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="tone"),
            EditOperation.replace("sample", "example", paragraph=refs[1]),
        ])

        assert isinstance(results[0].comment_id, int)
        assert results[1].comment_id is None
        assert len(doc.list_comments()) == 1

    def test_two_calls_with_the_same_note_make_two_comments(self, doc):
        """Dedupe scope is one call: two calls are two proposals at two moments."""
        doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="house style")
        doc.replace("sample", "example", paragraph=find_ref(doc, "sample document"), note="house style")

        assert [c.text for c in doc.list_comments()] == ["house style", "house style"]

    def test_invalid_note_rolls_back_the_whole_batch(self, doc):
        """A bad note on any op applies nothing and creates no comment."""
        before = doc.get_visible_text()
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]
        ops = [
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="tone"),
            # The typed constructor would reject this, so build the raw form.
            EditOperation(action="replace", paragraph=refs[1], find="sample", replace_with="example", note="bad\nnote"),
        ]

        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit(ops)

        assert exc.value.operation_index == 1
        assert doc.get_visible_text() == before
        assert doc.list_revisions() == []
        assert doc.list_comments() == []

    def test_dry_run_reports_the_invalid_note_row(self, doc):
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]
        ops = [
            EditOperation.replace("quick", "swift", paragraph=refs[0], note="tone"),
            EditOperation(action="replace", paragraph=refs[1], find="sample", replace_with="example", note="bad\nnote"),
        ]

        rows = doc.batch_edit(ops, dry_run=True)

        assert rows[0].valid
        assert not rows[1].valid
        assert "'note'" in rows[1].error
        assert doc.list_comments() == []


class TestNothingToAnchor:
    """A dropped rationale is never silent — and a recorded one never warns."""

    def test_noop_replace_warns_and_records_nothing(self, doc):
        ref = find_ref(doc, "quick brown")

        with pytest.warns(UnanchoredNoteWarning, match="no-op"):
            result = doc.replace("quick", "quick", paragraph=ref, note="why I left it")

        assert result.comment_id is None
        assert doc.list_comments() == []
        assert doc.list_revisions() == []

    def test_own_insertion_amendment_warns(self, doc):
        """Amending our own pending insertion creates no new revision."""
        first = doc.insert_after("fox", " ZE9", paragraph=find_ref(doc, "quick brown"))

        with pytest.warns(UnanchoredNoteWarning, match="amended your own pending insertion"):
            result = doc.replace(" ZE9", "QX7", paragraph=first, note="tighten it")

        assert result.comment_id is None
        assert doc.list_comments() == []
        # The edit itself still applied.
        assert "foxQX7" in doc.get_visible_text()

    def test_unchanged_rewrite_warns(self, doc):
        ref = find_ref(doc, "quick brown")
        unchanged = "The quick brown fox jumps over the lazy dog."

        with pytest.warns(UnanchoredNoteWarning, match="rewrite created no revisions"):
            result = doc.rewrite_paragraph(ref, unchanged, note="reviewed, no change needed")

        assert result.comment_id is None
        assert doc.list_comments() == []

    def test_paragraph_split_only_warns(self, doc):
        """A bare ``\\n`` insert's only revision is the paragraph mark, which
        lives in w:pPr/w:rPr where a comment marker cannot go."""
        ref = find_ref(doc, "quick brown")

        with pytest.warns(UnanchoredNoteWarning, match="tracked paragraph mark"):
            result = doc.insert_before("jumps", "\n", paragraph=ref, note="split for readability")

        assert result.comment_id is None
        assert result.group_id is not None  # the split itself was recorded
        assert doc.list_comments() == []
        assert len(doc.list_paragraphs()) == 5

    def test_warning_points_at_the_callers_line(self, doc):
        """A wrong stacklevel would blame library internals for the caller's edit."""
        ref = find_ref(doc, "quick brown")

        with pytest.warns(UnanchoredNoteWarning) as record:
            doc.replace("quick", "quick", paragraph=ref, note="why")

        assert record[0].filename == __file__

    def test_anchored_note_emits_no_warning(self, doc):
        """A false positive here would train callers to filter the category out."""
        ref = find_ref(doc, "quick brown")

        with warnings.catch_warnings():
            warnings.simplefilter("error", UnanchoredNoteWarning)
            result = doc.replace("quick", "swift", paragraph=ref, note="tone")

        assert isinstance(result.comment_id, int)

    def test_edit_without_a_note_emits_no_warning(self, doc):
        ref = find_ref(doc, "quick brown")

        with warnings.catch_warnings():
            warnings.simplefilter("error", UnanchoredNoteWarning)
            doc.replace("quick", "quick", paragraph=ref)
            doc.insert_before("jumps", "\n", paragraph=find_ref(doc, "quick brown"))

    def test_an_op_whose_note_a_sibling_recorded_reports_it_and_does_not_warn(self, doc):
        """The note reached the document, so the rationale was not dropped —
        whatever this particular operation managed to create."""
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]
        with warnings.catch_warnings():
            warnings.simplefilter("error", UnanchoredNoteWarning)
            results = doc.batch_edit([
                # A no-op and a paragraph-mark-only split, both sharing their
                # note with the one operation that does create a revision.
                EditOperation.replace("quick", "quick", paragraph=refs[0], note="house style"),
                EditOperation.insert_before("jumps", "\n", paragraph=refs[0], note="house style"),
                EditOperation.replace("sample", "example", paragraph=refs[1], note="house style"),
            ])

        assert len({r.comment_id for r in results}) == 1
        assert results[0].comment_id is not None
        assert [c.text for c in doc.list_comments()] == ["house style"]

    def test_batch_warns_once_per_unanchorable_op(self, doc):
        refs = [find_ref(doc, "quick brown"), find_ref(doc, "sample document")]

        with pytest.warns(UnanchoredNoteWarning) as record:
            results = doc.batch_edit([
                EditOperation.replace("quick", "quick", paragraph=refs[0], note="left as is"),
                EditOperation.replace("sample", "example", paragraph=refs[1], note="precision"),
            ])

        assert len([w for w in record if w.category is UnanchoredNoteWarning]) == 1
        assert results[0].comment_id is None
        assert isinstance(results[1].comment_id, int)


class TestNoteValidation:
    """A bad note fails before the edit runs, never after it."""

    @pytest.mark.parametrize(
        "note",
        ["", "line\nbreak", "tab\there", 42, ["a note"]],
        ids=["empty", "newline", "tab", "int", "list"],
    )
    def test_bad_note_is_rejected_before_any_mutation(self, doc, note):
        ref = find_ref(doc, "quick brown")
        before = doc.get_visible_text()

        with pytest.raises(ValueError, match="'note'"):
            doc.replace("quick", "swift", paragraph=ref, note=note)

        assert doc.get_visible_text() == before
        assert doc.list_revisions() == []
        assert doc.list_comments() == []

    @pytest.mark.parametrize("method", ["delete", "insert_after", "insert_before"])
    def test_every_edit_method_validates_its_note(self, doc, method):
        ref = find_ref(doc, "quick brown")
        args = ("lazy ",) if method == "delete" else ("fox", " (the animal)")

        with pytest.raises(ValueError, match="'note'"):
            getattr(doc, method)(*args, paragraph=ref, note="bad\nnote")

        assert doc.list_revisions() == []

    def test_rewrite_paragraph_validates_its_note(self, doc):
        ref = find_ref(doc, "quick brown")

        with pytest.raises(ValueError, match="'note'"):
            doc.rewrite_paragraph(ref, "Something else entirely.", note="bad\nnote")

        assert doc.list_revisions() == []

    @pytest.mark.parametrize(
        ("constructor", "args"),
        [
            ("replace", ("quick", "swift")),
            ("delete", ("quick",)),
            ("insert_after", ("quick", " (fast)")),
            ("insert_before", ("quick", "very ")),
        ],
    )
    def test_edit_operation_constructors_validate_their_note(self, constructor, args):
        with pytest.raises(ValueError, match="'note'"):
            getattr(EditOperation, constructor)(*args, paragraph="P2#a7b2", note="bad\nnote")


class TestNotesAndOrdinaryComments:
    """Note comments and add_comment() comments share one id allocator."""

    def test_interleaved_ids_do_not_collide(self, doc):
        first = doc.add_comment("fox", "an ordinary comment")
        annotated = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="a note")
        last = doc.add_comment("lazy", "another ordinary comment")

        assert len({first, annotated.comment_id, last}) == 3
        assert sorted(c.id for c in doc.list_comments()) == sorted([first, annotated.comment_id, last])

    def test_resolving_the_edit_leaves_ordinary_comments_alone(self, doc):
        ordinary = doc.add_comment("fox", "an ordinary comment")
        annotated = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="a note")

        doc.accept_group(annotated.group_id)

        assert [(c.id, c.text) for c in doc.list_comments()] == [(ordinary, "an ordinary comment")]

    def test_a_note_comment_can_be_replied_to(self, doc):
        """comment_id is a live comment id, not an opaque token."""
        annotated = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")

        reply_id = doc.reply_to_comment(annotated.comment_id, "agreed")

        assert reply_id != annotated.comment_id
        assert [r.text for r in doc.list_comments()[0].replies] == ["agreed"]

    def test_reaping_a_note_takes_its_replies_with_it(self, doc):
        """A reply outliving its parent points at a paraId nothing still holds."""
        annotated = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")
        doc.reply_to_comment(annotated.comment_id, "agreed")

        doc.accept_group(annotated.group_id)

        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    def test_reaping_a_note_takes_a_nested_reply_too(self, doc):
        """A reply to a reply is still part of the thread the note started."""
        annotated = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")
        reply_id = doc.reply_to_comment(annotated.comment_id, "agreed")
        doc.reply_to_comment(reply_id, "and here too")

        doc.accept_group(annotated.group_id)

        assert doc.list_comments() == []
        assert _marker_counts(doc) == (0, 0, 0)

    def test_deleting_a_comment_removes_its_extensible_entry(self, doc, temp_dir):
        """Automatic reaping made the one part delete_comment skipped add up."""
        annotated = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")
        doc.accept_group(annotated.group_id)
        out = temp_dir / "reaped.docx"
        doc.save(out)

        with zipfile.ZipFile(out) as zf:
            names = set(zf.namelist())
            part = "word/commentsExtensible.xml"
            extensible = zf.read(part).decode() if part in names else ""
        assert "commentExtensible" not in extensible


class TestNoteLifetimeAcrossSave:
    """The comment persists; the note-to-group link is per-session."""

    def test_comment_survives_save_and_reopen(self, doc, temp_dir):
        result = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone: plainer word")
        out = temp_dir / "saved.docx"
        doc.save(out)

        reopened = Document.open(out, force_recreate=True)
        try:
            assert [c.text for c in reopened.list_comments()] == ["tone: plainer word"]
            start, end = _markers(reopened, result.comment_id)
            assert _next_element(start).tagName == "w:del"
            assert _prev_element(end).tagName == "w:ins"
        finally:
            reopened.close()

    def test_the_link_is_per_session(self, doc, temp_dir):
        """After reopen the note comment is an ordinary comment: rejecting the
        inferred group leaves it, delete_comment removes it."""
        doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="tone")
        out = temp_dir / "saved.docx"
        doc.save(out)

        reopened = Document.open(out, force_recreate=True)
        try:
            group_id = reopened.list_revisions()[0].group_id
            assert group_id is not None
            reopened.reject_group(group_id)

            assert [c.text for c in reopened.list_comments()] == ["tone"]
            assert reopened.delete_comment(reopened.list_comments()[0].id) is True
            assert reopened.list_comments() == []
        finally:
            reopened.close()
