"""Comment writes into a pre-existing ``comments.xml`` that is namespace-poor.

Word 2007, LibreOffice and Google Docs exports ship ``word/comments.xml``
alone, with a root that declares ``xmlns:w`` and nothing else. New comments
carry ``w14:paraId``, so the part must gain the declaration first, and the
three sibling parts created on the way must be wired into the package.
"""

import re
import shutil
import zipfile
from pathlib import Path

import defusedxml.minidom
import pytest
from conftest import find_ref, replace_document_xml
from test_corpus_harness import harness

from docx_editor import Document
from docx_editor.exceptions import XMLError

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
COMMENT_PARTS = ("comments.xml", "commentsExtended.xml", "commentsIds.xml", "commentsExtensible.xml")

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
EMPTY = f'{XML_DECL}<w:comments xmlns:w="{W_NS}"/>'
# One comment as a pre-2010 producer writes it: no w14:paraId on its paragraph.
ONE_OLD_COMMENT = (
    f'{XML_DECL}<w:comments xmlns:w="{W_NS}">'
    '<w:comment w:id="0" w:author="Old Author" w:date="2020-01-01T00:00:00Z" w:initials="OA">'
    "<w:p><w:r><w:t>old remark</w:t></w:r></w:p>"
    "</w:comment></w:comments>"
)
DEFAULT_NS_ROOT = f'{XML_DECL}<comments xmlns="{W_NS}"/>'
NESTED_ROOT = f'{XML_DECL}<wrapper xmlns:w="{W_NS}"><w:comments/></wrapper>'
W14_BOUND_ELSEWHERE = f'{XML_DECL}<w:comments xmlns:w="{W_NS}" xmlns:w14="urn:not-w14"/>'


def make_docx_with_comments(simple_docx: Path, tmp_path: Path, comments_xml: str) -> Path:
    """Copy ``simple_docx`` adding ``word/comments.xml`` with its override and relationship."""
    dest = tmp_path / "with_comments.docx"
    with zipfile.ZipFile(simple_docx) as z_in, zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z_out:
        for item in z_in.infolist():
            data = z_in.read(item.filename)
            if item.filename == "[Content_Types].xml":
                override = f'<Override PartName="/word/comments.xml" ContentType="{COMMENTS_CONTENT_TYPE}"/>'
                data = data.decode("utf-8").replace("</Types>", override + "</Types>").encode("utf-8")
            elif item.filename == "word/_rels/document.xml.rels":
                rel = f'<Relationship Id="rId900" Type="{COMMENTS_REL_TYPE}" Target="comments.xml"/>'
                data = data.decode("utf-8").replace("</Relationships>", rel + "</Relationships>").encode("utf-8")
            z_out.writestr(item, data)
        z_out.writestr("word/comments.xml", comments_xml)
    return dest


def read_part(docx: Path, name: str) -> str:
    with zipfile.ZipFile(docx) as z:
        return z.read(name).decode("utf-8")


def marker_count(doc: Document) -> int:
    dom = doc._document_editor.dom
    return sum(
        len(dom.getElementsByTagName(tag)) for tag in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference")
    )


@pytest.fixture
def open_doc(simple_docx, tmp_path):
    """Open ``simple_docx`` with the given ``comments.xml``; closes on teardown."""
    opened: list[Document] = []

    def _open(comments_xml: str) -> Document:
        doc = Document.open(make_docx_with_comments(simple_docx, tmp_path, comments_xml), author="Tester")
        opened.append(doc)
        return doc

    yield _open
    for doc in opened:
        doc.close()


def save_and_reopen(doc: Document, tmp_path: Path) -> tuple[Path, Document]:
    out = tmp_path / "out.docx"
    doc.save(out)
    return out, Document.open(out, author="Tester")


class TestNamespacePoorPart:
    """The reported crash: ``xmlns:w`` is the root's only declaration."""

    @pytest.mark.parametrize("comments_xml", [EMPTY, ONE_OLD_COMMENT], ids=["empty", "one-old-comment"])
    def test_add_comment_declares_w14_and_survives_reopen(self, open_doc, tmp_path, comments_xml):
        doc = open_doc(comments_xml)
        before = [(c.id, c.text) for c in doc.list_comments()]

        comment_id = doc.add_comment("quick brown", "check this")

        out, reopened = save_and_reopen(doc, tmp_path)
        try:
            assert [(c.id, c.text) for c in reopened.list_comments()] == [*before, (comment_id, "check this")]
        finally:
            reopened.close()
        root = defusedxml.minidom.parseString(read_part(out, "word/comments.xml")).documentElement
        assert root.getAttribute("xmlns:w14") == W14_NS

    @pytest.mark.parametrize("comments_xml", [EMPTY, ONE_OLD_COMMENT], ids=["empty", "one-old-comment"])
    def test_reply_is_threaded_after_reopen(self, open_doc, tmp_path, comments_xml):
        doc = open_doc(comments_xml)
        parent_id = doc.add_comment("quick brown", "check this")

        reply_id = doc.reply_to_comment(parent_id, "agreed")

        out, reopened = save_and_reopen(doc, tmp_path)
        try:
            parent = next(c for c in reopened.list_comments() if c.id == parent_id)
            assert [(r.id, r.text) for r in parent.replies] == [(reply_id, "agreed")]
        finally:
            reopened.close()
        assert "w15:paraIdParent" in read_part(out, "word/commentsExtended.xml")

    @pytest.mark.parametrize("comments_xml", [EMPTY, ONE_OLD_COMMENT], ids=["empty", "one-old-comment"])
    def test_note_edit_attaches_a_comment(self, open_doc, tmp_path, comments_xml):
        doc = open_doc(comments_xml)

        result = doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="plainer word")

        assert isinstance(result.comment_id, int)
        _, reopened = save_and_reopen(doc, tmp_path)
        try:
            assert (result.comment_id, "plainer word") in [(c.id, c.text) for c in reopened.list_comments()]
        finally:
            reopened.close()

    def test_existing_comment_is_left_untouched(self, open_doc, tmp_path):
        doc = open_doc(ONE_OLD_COMMENT)

        new_id = doc.add_comment("quick brown", "check this")

        assert new_id != 0
        out, _reopened = save_and_reopen(doc, tmp_path)
        _reopened.close()
        old = re.search(r'<w:comment w:id="0".*?</w:comment>', read_part(out, "word/comments.xml"), re.S)
        assert old is not None
        assert "w14:paraId" not in old.group(0)
        assert "old remark" in old.group(0)


class TestSiblingPartWiring:
    """Parts created beside a pre-existing ``comments.xml`` are wired into the package."""

    def test_each_part_has_one_relationship_and_one_override(self, open_doc, tmp_path):
        doc = open_doc(EMPTY)
        doc.add_comment("quick brown", "check this")

        out, reopened = save_and_reopen(doc, tmp_path)
        reopened.close()

        rels = defusedxml.minidom.parseString(read_part(out, "word/_rels/document.xml.rels"))
        rel_elems = rels.getElementsByTagName("Relationship")
        targets = [r.getAttribute("Target") for r in rel_elems]
        assert {part: targets.count(part) for part in COMMENT_PARTS} == dict.fromkeys(COMMENT_PARTS, 1)
        ids = [r.getAttribute("Id") for r in rel_elems]
        assert len(ids) == len(set(ids))

        types = defusedxml.minidom.parseString(read_part(out, "[Content_Types].xml"))
        part_names = [o.getAttribute("PartName") for o in types.getElementsByTagName("Override")]
        assert {part: part_names.count(f"/word/{part}") for part in COMMENT_PARTS} == dict.fromkeys(COMMENT_PARTS, 1)

    def test_absolute_target_is_not_wired_twice(self, simple_docx, tmp_path):
        src = make_docx_with_comments(simple_docx, tmp_path, EMPTY)
        absolute = tmp_path / "absolute.docx"
        with zipfile.ZipFile(src) as z_in, zipfile.ZipFile(absolute, "w", zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                data = z_in.read(item.filename)
                if item.filename == "word/_rels/document.xml.rels":
                    data = data.replace(b'Target="comments.xml"', b'Target="/word/comments.xml"')
                z_out.writestr(item, data)
        out = tmp_path / "out.docx"
        with Document.open(absolute, author="Tester") as doc:
            doc.save(out)

        assert read_part(out, "word/_rels/document.xml.rels").count(COMMENTS_REL_TYPE) == 1

    def test_second_save_adds_nothing(self, open_doc, tmp_path):
        doc = open_doc(EMPTY)
        doc.add_comment("quick brown", "check this")
        out, reopened = save_and_reopen(doc, tmp_path)
        try:
            again = tmp_path / "again.docx"
            reopened.add_comment("lazy dog", "and this")
            reopened.save(again)
        finally:
            reopened.close()

        assert read_part(again, "word/_rels/document.xml.rels").count("Target=") == read_part(
            out, "word/_rels/document.xml.rels"
        ).count("Target=")
        assert read_part(again, "[Content_Types].xml").count("<Override") == read_part(
            out, "[Content_Types].xml"
        ).count("<Override")


class TestTypedFailureLeavesNoMarkers:
    """A part that cannot take the comment fails typed and leaves no comment markers."""

    @pytest.mark.parametrize(
        "comments_xml",
        [DEFAULT_NS_ROOT, NESTED_ROOT, W14_BOUND_ELSEWHERE],
        ids=["default-ns-root", "nested-root", "w14-bound-elsewhere"],
    )
    def test_add_comment(self, open_doc, comments_xml):
        doc = open_doc(comments_xml)

        with pytest.raises(XMLError):
            doc.add_comment("quick brown", "check this")

        assert marker_count(doc) == 0
        assert doc.list_comments() == []
        # No sibling part was created on the way to the failure.
        assert sorted(p.name for p in doc._workspace.word_path.glob("comments*.xml")) == ["comments.xml"]

    @pytest.mark.parametrize(
        "comments_xml",
        [DEFAULT_NS_ROOT, NESTED_ROOT, W14_BOUND_ELSEWHERE],
        ids=["default-ns-root", "nested-root", "w14-bound-elsewhere"],
    )
    def test_note_edit(self, open_doc, comments_xml):
        """A note is attached after its edit: the edit stays applied, the note leaves no markers."""
        doc = open_doc(comments_xml)

        with pytest.raises(XMLError):
            doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"), note="plainer word")

        assert marker_count(doc) == 0
        assert len(doc._document_editor.dom.getElementsByTagName("w:ins")) == 1

    def test_sibling_part_binding_its_prefix_elsewhere(self, simple_docx, tmp_path):
        """``commentsExtensible.xml`` is the last part written; it is still checked first."""
        src = make_docx_with_comments(simple_docx, tmp_path, EMPTY)
        with zipfile.ZipFile(src, "a") as z:
            z.writestr(
                "word/commentsExtensible.xml",
                f'{XML_DECL}<w16cex:commentsExtensible xmlns:w16cex="urn:not-w16cex"/>',
            )
        with Document.open(src, author="Tester") as doc:
            with pytest.raises(XMLError, match=r"commentsExtensible\.xml: prefix 'w16cex' is bound"):
                doc.add_comment("quick brown", "check this")

            assert marker_count(doc) == 0
            assert doc.list_comments() == []
            # comments.xml passed its own check but is left as it was.
            assert not doc._comment_manager._get_editor(
                doc._comment_manager.comments_path
            ).dom.documentElement.hasAttribute("xmlns:w14")

    def test_wrong_w14_binding_is_named(self, open_doc):
        doc = open_doc(W14_BOUND_ELSEWHERE)

        with pytest.raises(XMLError, match=r"comments\.xml: prefix 'w14' is bound to 'urn:not-w14'"):
            doc.add_comment("quick brown", "check this")


class TestPrefixDeclaredOnNonRootAncestor:
    """A prefix declared on a paragraph, not the root, is out of scope for re-serialized runs.

    Not supported yet (ROADMAP.md #93); what is pinned here is that it fails
    typed and leaves no comment markers behind.
    """

    @pytest.fixture
    def doc(self, simple_docx, tmp_path):
        doc_xml = (
            f'{XML_DECL}<w:document xmlns:w="{W_NS}"><w:body>'
            '<w:p xmlns:foo="urn:foo"><w:r><foo:mark/><w:t>The quick brown fox.</w:t></w:r></w:p>'
            "</w:body></w:document>"
        )
        src = tmp_path / "ancestor_ns.docx"
        replace_document_xml(simple_docx, src, doc_xml)
        doc = Document.open(src, author="Tester")
        yield doc
        doc.close()

    def test_replace_raises_xml_error(self, doc):
        with pytest.raises(XMLError, match="cannot parse XML fragment for document.xml"):
            doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"))

    def test_add_comment_raises_xml_error_without_markers(self, doc):
        with pytest.raises(XMLError, match="cannot parse XML fragment for document.xml"):
            doc.add_comment("quick", "check this")

        assert marker_count(doc) == 0

    def test_failed_add_comment_creates_no_comment_parts(self, doc, tmp_path):
        with pytest.raises(XMLError):
            doc.add_comment("quick", "check this")

        out = tmp_path / "out.docx"
        doc.save(out)
        with zipfile.ZipFile(out) as z:
            assert [n for n in z.namelist() if "comments" in n.lower()] == []
        assert "comments" not in read_part(out, "word/_rels/document.xml.rels")


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")
def test_libreoffice_opens_the_output(open_doc, tmp_path, monkeypatch):
    doc = open_doc(ONE_OLD_COMMENT)
    parent_id = doc.add_comment("quick brown", "check this")
    doc.reply_to_comment(parent_id, "agreed")
    doc.replace("lazy", "idle", paragraph=find_ref(doc, "quick brown"), note="plainer word")
    out = tmp_path / "out.docx"
    doc.save(out)

    lo_dir = tmp_path / "lo"
    lo_dir.mkdir()
    monkeypatch.setattr(harness, "LO_DIR", lo_dir)
    monkeypatch.setattr(harness, "LO_PROFILE", tmp_path / "profile")
    record = harness.run_lo_roundtrip(out, False, shutil.which("soffice"), tmp_path / "work")

    assert record["status"] == "pass", record
