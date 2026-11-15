"""Comments written without ``w14:paraId`` are replyable, resolvable and deletable.

Word 2007, LibreOffice and Google Docs write comments with no ``w14:paraId``
and no commentsExtended.xml, and often anchor them with a reference run alone.
``list_comments`` returns them, so every verb that takes a comment id must
accept theirs. Every case saves and reopens.
"""

import shutil
import zipfile
from pathlib import Path

import defusedxml.minidom
import pytest
from conftest import replace_document_xml
from test_comment_namespaces import (
    ONE_OLD_COMMENT,
    W14_NS,
    W_NS,
    XML_DECL,
    make_docx_with_comments,
    marker_count,
    read_part,
    save_and_reopen,
)
from test_corpus_harness import harness

from docx_editor import Document
from docx_editor.exceptions import CommentError, MultipleNodesFoundError, XMLError

W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
EXTENDED = "word/commentsExtended.xml"

QUICK = "<w:r><w:t>The quick brown fox jumps over the lazy dog.</w:t></w:r>"
SAMPLE = "<w:r><w:t>This is a sample document for testing the editing features.</w:t></w:r>"

SHAPES = ("range", "point", "unanchored")
ANCHORED_SHAPES = ("range", "point")


def old_comment(comment_id: int, paragraphs: str = "<w:p><w:r><w:t>old remark</w:t></w:r></w:p>") -> str:
    return f'<w:comment w:id="{comment_id}" w:author="Old Author" w:initials="OA">{paragraphs}</w:comment>'


def comments_part(*comments: str, root_attrs: str = "") -> str:
    return f'{XML_DECL}<w:comments xmlns:w="{W_NS}"{root_attrs}>{"".join(comments)}</w:comments>'


def extended_part(*rows: str) -> str:
    return f'{XML_DECL}<w15:commentsEx xmlns:w15="{W15_NS}">{"".join(rows)}</w15:commentsEx>'


def anchored(run: str, comment_id: int | str, shape: str) -> str:
    """``run`` carrying comment ``comment_id``: a range with its reference, a reference alone, or nothing."""
    reference = f'<w:r><w:commentReference w:id="{comment_id}"/></w:r>'
    if shape == "range":
        return f'<w:commentRangeStart w:id="{comment_id}"/>{run}<w:commentRangeEnd w:id="{comment_id}"/>{reference}'
    if shape == "point":
        return run + reference
    return run


def build(
    simple_docx: Path,
    tmp_path: Path,
    comments_xml: str,
    anchors: dict[str, tuple[int | str, str]],
    extended_xml: str | None = None,
) -> Path:
    """``simple_docx`` with ``comments_xml``, each run in ``anchors`` anchored, and an optional commentsExtended.xml."""
    path = make_docx_with_comments(simple_docx, tmp_path, comments_xml)
    document_xml = read_part(path, "word/document.xml")
    for run, (comment_id, shape) in anchors.items():
        assert run in document_xml
        document_xml = document_xml.replace(run, anchored(run, comment_id, shape))
    dest = tmp_path / "legacy.docx"
    replace_document_xml(path, dest, document_xml)
    if extended_xml is not None:
        with zipfile.ZipFile(dest, "a") as z:
            z.writestr(EXTENDED, extended_xml)
    return dest


@pytest.fixture
def open_legacy(simple_docx, tmp_path):
    """Open a document built by :func:`build`; closes on teardown."""
    opened: list[Document] = []

    def _open(
        comments_xml: str, anchors: dict[str, tuple[int | str, str]], extended_xml: str | None = None
    ) -> Document:
        doc = Document.open(build(simple_docx, tmp_path, comments_xml, anchors, extended_xml), author="Tester")
        opened.append(doc)
        return doc

    yield _open
    for doc in opened:
        doc.close()


def reopened_comments(doc: Document, tmp_path: Path):
    """``(saved path, list_comments() of the saved file)``."""
    out, reopened = save_and_reopen(doc, tmp_path)
    try:
        return out, reopened.list_comments()
    finally:
        reopened.close()


def elements(docx: Path, part: str, tag: str) -> list:
    return defusedxml.minidom.parseString(read_part(docx, part)).getElementsByTagName(tag)


def para_ids(docx: Path) -> list[str]:
    """``w14:paraId`` of every paragraph in comments.xml, ``""`` where it has none."""
    return [p.getAttribute("w14:paraId") for p in elements(docx, "word/comments.xml", "w:p")]


class TestOldCommentIsActionable:
    """The reported gap: a listed comment that no verb accepted."""

    @pytest.mark.parametrize("shape", ANCHORED_SHAPES)
    def test_reply_is_threaded_under_the_old_comment(self, open_legacy, tmp_path, shape):
        doc = open_legacy(ONE_OLD_COMMENT, {QUICK: (0, shape)})

        reply_id = doc.reply_to_comment(0, "agreed")

        _, comments = reopened_comments(doc, tmp_path)
        assert [c.id for c in comments] == [0]
        assert [(r.id, r.text) for r in comments[0].replies] == [(reply_id, "agreed")]

    @pytest.mark.parametrize("shape", SHAPES)
    def test_resolve_survives_reopen(self, open_legacy, tmp_path, shape):
        doc = open_legacy(ONE_OLD_COMMENT, {QUICK: (0, shape)})
        assert [c.resolved for c in doc.list_comments()] == [False]

        assert doc.resolve_comment(0) is True

        assert [c.resolved for c in doc.list_comments()] == [True]
        _, comments = reopened_comments(doc, tmp_path)
        assert [(c.id, c.resolved) for c in comments] == [(0, True)]

    @pytest.mark.parametrize("shape", SHAPES)
    def test_delete_removes_comment_and_markers(self, open_legacy, tmp_path, shape):
        doc = open_legacy(ONE_OLD_COMMENT, {QUICK: (0, shape)})

        assert doc.delete_comment(0) is True

        assert marker_count(doc) == 0
        out, comments = reopened_comments(doc, tmp_path)
        assert comments == []
        assert len(elements(out, "word/comments.xml", "w:comment")) == 0

    @pytest.mark.parametrize("shape", SHAPES)
    def test_every_listed_id_is_accepted_by_resolve_and_delete(self, open_legacy, shape):
        doc = open_legacy(comments_part(old_comment(0), old_comment(1)), {QUICK: (0, shape), SAMPLE: (1, shape)})
        listed = [c.id for c in doc.list_comments()]
        assert listed == [0, 1]

        assert [doc.resolve_comment(comment_id) for comment_id in listed] == [True, True]
        assert [doc.delete_comment(comment_id) for comment_id in listed] == [True, True]
        assert doc.list_comments() == []

    def test_resolve_and_delete_still_report_an_unknown_id(self, open_legacy):
        doc = open_legacy(ONE_OLD_COMMENT, {QUICK: (0, "range")})

        assert doc.resolve_comment(7) is False
        assert doc.delete_comment(7) is False


class TestReplyAnchor:
    def test_reply_to_a_point_comment_is_a_point_comment(self, open_legacy):
        doc = open_legacy(ONE_OLD_COMMENT, {QUICK: (0, "point")})

        reply_id = doc.reply_to_comment(0, "agreed")

        dom = doc._document_editor.dom
        assert len(dom.getElementsByTagName("w:commentRangeStart")) == 0
        assert len(dom.getElementsByTagName("w:commentRangeEnd")) == 0
        references = dom.getElementsByTagName("w:commentReference")
        assert [ref.getAttribute("w:id") for ref in references] == ["0", str(reply_id)]
        assert references[0].parentNode.nextSibling is references[1].parentNode

    def test_reply_to_a_reply_of_a_point_comment(self, open_legacy, tmp_path):
        doc = open_legacy(ONE_OLD_COMMENT, {QUICK: (0, "point")})

        first = doc.reply_to_comment(0, "agreed")
        second = doc.reply_to_comment(first, "so do I")

        _, comments = reopened_comments(doc, tmp_path)
        assert [r.id for r in comments[0].replies] == [first]
        assert [r.id for r in comments[0].replies[0].replies] == [second]

    @pytest.mark.parametrize("shape", ANCHORED_SHAPES)
    def test_reply_to_a_reference_outside_a_run_stays_in_the_paragraph(self, open_legacy, shape):
        """A bare reference's parent is the paragraph: the reply must not be seated after it, in w:body."""
        doc = open_legacy(ONE_OLD_COMMENT, {QUICK: (0, shape)})
        dom = doc._document_editor.dom
        reference = dom.getElementsByTagName("w:commentReference")[0]
        run, paragraph = reference.parentNode, reference.parentNode.parentNode
        paragraph.replaceChild(reference, run)

        reply_id = doc.reply_to_comment(0, "agreed")

        reply_reference = dom.getElementsByTagName("w:commentReference")[1]
        assert reply_reference.getAttribute("w:id") == str(reply_id)
        assert reply_reference.parentNode.parentNode is paragraph
        body = dom.getElementsByTagName("w:body")[0]
        assert not [child for child in body.childNodes if child.nodeName == "w:r"]

    def test_reply_to_an_unanchored_comment_is_refused_untouched(self, open_legacy, tmp_path):
        doc = open_legacy(ONE_OLD_COMMENT, {})
        document_before = doc._document_editor.dom.toxml()

        with pytest.raises(CommentError, match="no w:commentReference") as excinfo:
            doc.reply_to_comment(0, "agreed")

        assert excinfo.value.comment_id == 0
        assert doc._document_editor.dom.toxml() == document_before
        out, comments = reopened_comments(doc, tmp_path)
        assert [(c.id, c.replies) for c in comments] == [(0, [])]
        assert para_ids(out) == [""]
        assert EXTENDED not in zipfile.ZipFile(out).namelist()


class TestMinting:
    def test_reading_and_adding_leave_the_old_comment_alone(self, open_legacy, tmp_path):
        doc = open_legacy(ONE_OLD_COMMENT, {QUICK: (0, "range")})

        doc.list_comments()
        doc.add_comment("sample document", "check this")

        out, _ = reopened_comments(doc, tmp_path)
        assert para_ids(out)[0] == ""

    def test_para_id_is_minted_once_on_the_last_paragraph(self, open_legacy, tmp_path):
        two_paragraphs = "<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>"
        doc = open_legacy(comments_part(old_comment(0, two_paragraphs)), {QUICK: (0, "range")})

        doc.resolve_comment(0)
        minted = doc._comment_manager.existing_comments[0]["para_id"]
        doc.resolve_comment(0)
        doc.reply_to_comment(0, "agreed")

        out, _ = reopened_comments(doc, tmp_path)
        assert para_ids(out)[:2] == ["", minted]
        rows = elements(out, EXTENDED, "w15:commentEx")
        assert [row.getAttribute("w15:paraId") for row in rows][0] == minted
        assert [row.getAttribute("w15:paraIdParent") for row in rows] == ["", minted]

    def test_duplicate_id_mints_on_the_registered_comment(self, open_legacy, tmp_path):
        """Of two comments sharing a w:id the last is the listed one; the first keeps its paraId and row."""
        first = old_comment(0, '<w:p w14:paraId="0000AAAA"><w:r><w:t>first</w:t></w:r></w:p>')
        comments_xml = comments_part(first, old_comment(0), root_attrs=f' xmlns:w14="{W14_NS}"')
        doc = open_legacy(
            comments_xml, {QUICK: (0, "range")}, extended_part('<w15:commentEx w15:paraId="0000AAAA" w15:done="0"/>')
        )

        assert doc.resolve_comment(0) is True

        out, comments = reopened_comments(doc, tmp_path)
        assert [(c.id, c.text, c.resolved) for c in comments] == [(0, "old remark", True)]
        minted = doc._comment_manager.existing_comments[0]["para_id"]
        assert para_ids(out) == ["0000AAAA", minted]
        rows = elements(out, EXTENDED, "w15:commentEx")
        assert [(row.getAttribute("w15:paraId"), row.getAttribute("w15:done")) for row in rows] == [
            ("0000AAAA", "0"),
            (minted, "1"),
        ]

    def test_duplicate_id_deletes_the_registered_comment(self, open_legacy, tmp_path):
        """delete_comment removes the listed duplicate, not the first element carrying the id."""
        first = old_comment(0, '<w:p w14:paraId="0000AAAA"><w:r><w:t>first</w:t></w:r></w:p>')
        comments_xml = comments_part(first, old_comment(0), root_attrs=f' xmlns:w14="{W14_NS}"')
        doc = open_legacy(
            comments_xml, {QUICK: (0, "range")}, extended_part('<w15:commentEx w15:paraId="0000AAAA" w15:done="0"/>')
        )
        assert [c.text for c in doc.list_comments()] == ["old remark"]

        assert doc.delete_comment(0) is True

        out, comments = reopened_comments(doc, tmp_path)
        assert [(c.id, c.text) for c in comments] == [(0, "first")]
        assert para_ids(out) == ["0000AAAA"]
        assert [row.getAttribute("w15:paraId") for row in elements(out, EXTENDED, "w15:commentEx")] == ["0000AAAA"]

    def test_noncanonical_id_is_deleted_by_its_integer(self, open_legacy, tmp_path):
        """``w:id="05"`` is listed as 5; delete_comment(5) must remove the element it was listed from."""
        doc = open_legacy(comments_part(old_comment(5).replace('w:id="5"', 'w:id="05"')), {})
        assert [c.id for c in doc.list_comments()] == [5]

        assert doc.delete_comment(5) is True

        _, comments = reopened_comments(doc, tmp_path)
        assert comments == []

    @pytest.mark.parametrize("shape", ANCHORED_SHAPES)
    def test_noncanonical_marker_ids_are_matched_by_integer(self, open_legacy, tmp_path, shape):
        """Markers written ``w:id="05"`` belong to comment 5: a reply finds them and a delete removes them."""

        doc = open_legacy(comments_part(old_comment(5).replace('w:id="5"', 'w:id="05"')), {QUICK: ("05", shape)})
        assert marker_count(doc) == (3 if shape == "range" else 1)

        reply_id = doc.reply_to_comment(5, "agreed")
        for comment in doc.list_comments():
            assert doc.resolve_comment(comment.id) is True
        _, comments = reopened_comments(doc, tmp_path)
        assert [(c.id, c.resolved, [r.id for r in c.replies]) for c in comments] == [(5, True, [reply_id])]

        for comment in doc.list_comments():
            assert doc.delete_comment(comment.id) is True
        assert marker_count(doc) == 0
        assert doc.list_comments() == []

    def test_markers_differing_only_in_spelling_are_duplicates(self, open_legacy):
        """``w:id="5"`` and ``w:id="05"`` are two references of comment 5: the reply is refused untouched."""
        doc = open_legacy(comments_part(old_comment(5)), {QUICK: (5, "point")})
        editor = doc._document_editor
        reference_run = editor.dom.getElementsByTagName("w:commentReference")[0].parentNode
        editor.insert_after(reference_run, '<w:r><w:commentReference w:id="05"/></w:r>')
        document_before = editor.dom.toxml()

        with pytest.raises(MultipleNodesFoundError):
            doc.reply_to_comment(5, "agreed")

        assert editor.dom.toxml() == document_before
        assert doc._comment_manager.existing_comments[5]["para_id"] is None

    @pytest.mark.parametrize("survivor_has_para_id", [True, False])
    def test_surviving_duplicate_stays_actionable_after_a_delete(self, open_legacy, survivor_has_para_id):
        """Deleting the listed duplicate lists the earlier one at once; its id must still be accepted."""
        with_para_id = '<w:p w14:paraId="0000AAAA"><w:r><w:t>with paraId</w:t></w:r></w:p>'
        without = "<w:p><w:r><w:t>without paraId</w:t></w:r></w:p>"
        first, last = (with_para_id, without) if survivor_has_para_id else (without, with_para_id)
        doc = open_legacy(
            comments_part(old_comment(0, first), old_comment(0, last), root_attrs=f' xmlns:w14="{W14_NS}"'),
            {QUICK: (0, "range")},
            extended_part('<w15:commentEx w15:paraId="0000AAAA" w15:done="0"/>'),
        )
        survivor_text = "with paraId" if survivor_has_para_id else "without paraId"

        assert doc.delete_comment(0) is True

        assert [(c.id, c.text) for c in doc.list_comments()] == [(0, survivor_text)]
        expected_para_id = "0000AAAA" if survivor_has_para_id else None
        assert doc._comment_manager.existing_comments == {0: {"para_id": expected_para_id}}
        for comment in doc.list_comments():
            assert doc.resolve_comment(comment.id) is True
        assert [c.resolved for c in doc.list_comments()] == [True]
        for comment in doc.list_comments():
            assert doc.delete_comment(comment.id) is True
        assert doc.list_comments() == []
        assert doc._comment_manager.existing_comments == {}

    def test_comment_without_a_paragraph_gains_one(self, open_legacy, tmp_path):
        doc = open_legacy(comments_part(old_comment(0, "")), {QUICK: (0, "range")})

        assert doc.resolve_comment(0) is True

        out, comments = reopened_comments(doc, tmp_path)
        assert [(c.id, c.text, c.resolved) for c in comments] == [(0, "", True)]
        assert len(para_ids(out)) == 1
        assert para_ids(out)[0] != ""

    def test_minted_para_id_is_in_the_w14_namespace(self, open_legacy, tmp_path):
        doc = open_legacy(ONE_OLD_COMMENT, {QUICK: (0, "point")})

        doc.resolve_comment(0)

        out, _ = reopened_comments(doc, tmp_path)
        paragraph = elements(out, "word/comments.xml", "w:p")[0]
        assert paragraph.getAttributeNS(W14_NS, "paraId") == doc._comment_manager.existing_comments[0]["para_id"]

    @pytest.mark.parametrize("verb", ["resolve", "reply"])
    def test_refused_call_mints_nothing(self, open_legacy, tmp_path, verb):
        comments_xml = comments_part(old_comment(0), root_attrs=' xmlns:w14="urn:not-w14"')
        doc = open_legacy(comments_xml, {QUICK: (0, "range")})
        markers_before = marker_count(doc)

        with pytest.raises(XMLError):
            doc.resolve_comment(0) if verb == "resolve" else doc.reply_to_comment(0, "agreed")

        assert marker_count(doc) == markers_before
        assert doc._comment_manager.existing_comments[0]["para_id"] is None
        out, comments = reopened_comments(doc, tmp_path)
        assert [(c.id, c.resolved, c.replies) for c in comments] == [(0, False, [])]
        assert [p.attributes.length for p in elements(out, "word/comments.xml", "w:p")] == [0]
        assert EXTENDED not in zipfile.ZipFile(out).namelist()


class TestParaIdWithoutExtendedRow:
    """Word 2010 writes ``w14:paraId`` and no commentsExtended.xml."""

    COMMENTS = comments_part(
        old_comment(0, '<w:p w14:paraId="0A0B0C0D"><w:r><w:t>old remark</w:t></w:r></w:p>'),
        root_attrs=f' xmlns:w14="{W14_NS}"',
    )

    def test_resolve_creates_and_wires_the_extended_part(self, open_legacy, tmp_path):
        doc = open_legacy(self.COMMENTS, {QUICK: (0, "range")})

        assert doc.resolve_comment(0) is True

        out, comments = reopened_comments(doc, tmp_path)
        assert [(c.id, c.resolved) for c in comments] == [(0, True)]
        assert para_ids(out) == ["0A0B0C0D"]
        rows = elements(out, EXTENDED, "w15:commentEx")
        assert [(row.getAttribute("w15:paraId"), row.getAttribute("w15:done")) for row in rows] == [("0A0B0C0D", "1")]
        assert read_part(out, "word/_rels/document.xml.rels").count('Target="commentsExtended.xml"') == 1
        assert read_part(out, "[Content_Types].xml").count('PartName="/word/commentsExtended.xml"') == 1

    def test_reply_writes_a_parent_row_and_a_child_row(self, open_legacy, tmp_path):
        doc = open_legacy(self.COMMENTS, {QUICK: (0, "range")})

        reply_id = doc.reply_to_comment(0, "agreed")

        out, comments = reopened_comments(doc, tmp_path)
        assert [r.id for r in comments[0].replies] == [reply_id]
        rows = elements(out, EXTENDED, "w15:commentEx")
        assert [row.getAttribute("w15:paraId") for row in rows][0] == "0A0B0C0D"
        assert [row.getAttribute("w15:paraIdParent") for row in rows] == ["", "0A0B0C0D"]

    def test_existing_row_is_reused(self, open_legacy, tmp_path):
        extended = extended_part('<w15:commentEx w15:paraId="0A0B0C0D" w15:done="0"/>')
        doc = open_legacy(self.COMMENTS, {QUICK: (0, "range")}, extended)

        doc.resolve_comment(0)
        doc.reply_to_comment(0, "agreed")

        out, _ = reopened_comments(doc, tmp_path)
        rows = elements(out, EXTENDED, "w15:commentEx")
        assert [row.getAttribute("w15:paraId") for row in rows].count("0A0B0C0D") == 1
        assert len(rows) == 2


class TestLastParagraphRule:
    """commentsExtended.xml names a comment by the paraId of its last paragraph."""

    COMMENTS = comments_part(
        old_comment(
            0,
            '<w:p w14:paraId="AAAA0001"><w:r><w:t>first</w:t></w:r></w:p>'
            '<w:p w14:paraId="AAAA0002"><w:r><w:t>second</w:t></w:r></w:p>',
        ),
        root_attrs=f' xmlns:w14="{W14_NS}"',
    )

    def test_resolved_is_read_from_the_last_paragraphs_row(self, open_legacy):
        extended = extended_part('<w15:commentEx w15:paraId="AAAA0002" w15:done="1"/>')
        doc = open_legacy(self.COMMENTS, {QUICK: (0, "range")}, extended)

        assert [c.resolved for c in doc.list_comments()] == [True]

    def test_resolve_flips_the_last_paragraphs_row(self, open_legacy, tmp_path):
        extended = extended_part('<w15:commentEx w15:paraId="AAAA0002" w15:done="0"/>')
        doc = open_legacy(self.COMMENTS, {QUICK: (0, "range")}, extended)

        assert doc.resolve_comment(0) is True

        out, comments = reopened_comments(doc, tmp_path)
        assert [c.resolved for c in comments] == [True]
        rows = elements(out, EXTENDED, "w15:commentEx")
        assert [(row.getAttribute("w15:paraId"), row.getAttribute("w15:done")) for row in rows] == [("AAAA0002", "1")]
        assert para_ids(out) == ["AAAA0001", "AAAA0002"]


class TestBesideAWordThread:
    """Old comments next to a thread Word wrote with paraIds."""

    COMMENTS = comments_part(
        old_comment(0),
        old_comment(1, '<w:p w14:paraId="BBBB0001"><w:r><w:t>word remark</w:t></w:r></w:p>'),
        old_comment(2, '<w:p w14:paraId="BBBB0002"><w:r><w:t>word reply</w:t></w:r></w:p>'),
        old_comment(3),
        root_attrs=f' xmlns:w14="{W14_NS}"',
    )
    EXTENDED_XML = extended_part(
        '<w15:commentEx w15:paraId="BBBB0001" w15:done="0"/>',
        '<w15:commentEx w15:paraId="BBBB0002" w15:paraIdParent="BBBB0001" w15:done="0"/>',
    )

    def test_comments_without_para_id_are_listed_apart(self, open_legacy):
        doc = open_legacy(self.COMMENTS, {QUICK: (0, "range")}, self.EXTENDED_XML)

        comments = doc.list_comments()

        assert [(c.id, [r.id for r in c.replies]) for c in comments] == [(0, []), (1, [2]), (3, [])]
        assert doc._comment_manager.reply_ids(0) == []
        assert doc._comment_manager.reply_ids(1) == [2]

    def test_deleting_an_old_comment_leaves_the_thread_intact(self, open_legacy, tmp_path):
        doc = open_legacy(self.COMMENTS, {QUICK: (0, "range")}, self.EXTENDED_XML)

        assert doc.delete_comment(0) is True

        out, comments = reopened_comments(doc, tmp_path)
        assert [(c.id, [r.id for r in c.replies]) for c in comments] == [(1, [2]), (3, [])]
        assert len(elements(out, EXTENDED, "w15:commentEx")) == 2
        assert para_ids(out) == ["BBBB0001", "BBBB0002", ""]


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")
@pytest.mark.parametrize("shape", ANCHORED_SHAPES)
def test_libreoffice_opens_the_output(open_legacy, tmp_path, monkeypatch, shape):
    comments_xml = comments_part(old_comment(0), old_comment(1), old_comment(2))
    doc = open_legacy(comments_xml, {QUICK: (0, shape), SAMPLE: (1, shape)})
    doc.reply_to_comment(0, "agreed")
    doc.resolve_comment(0)
    doc.resolve_comment(2)
    doc.delete_comment(1)
    out = tmp_path / "out.docx"
    doc.save(out)

    lo_dir = tmp_path / "lo"
    lo_dir.mkdir()
    monkeypatch.setattr(harness, "LO_DIR", lo_dir)
    monkeypatch.setattr(harness, "LO_PROFILE", tmp_path / "profile")
    record = harness.run_lo_roundtrip(out, False, shutil.which("soffice"), tmp_path / "work")

    assert record["status"] == "pass", record
