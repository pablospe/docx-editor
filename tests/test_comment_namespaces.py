"""Comment writes into a pre-existing ``comments.xml`` that is namespace-poor.

Word 2007, LibreOffice and Google Docs exports ship ``word/comments.xml``
alone, with a root that declares ``xmlns:w`` and nothing else. New comments
carry ``w14:paraId``, so the part must gain the declaration first, and the
three sibling parts created on the way must be wired into the package.

Also covered: which file is a comment part at all. It is the one the
relationship of its type names, which may not be ``word/comments.xml``.
"""

import re
import shutil
import zipfile
from pathlib import Path

import defusedxml.minidom
import pytest
from conftest import find_ref, replace_document_xml, replace_docx_parts
from test_corpus_harness import harness

from docx_editor import Document
from docx_editor.comments import CommentManager, _is_inside
from docx_editor.exceptions import XMLError

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
COMMENT_PART_NAMES = ("comments.xml", "commentsExtended.xml", "commentsIds.xml", "commentsExtensible.xml")

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


def make_docx_with_comments(
    simple_docx: Path, tmp_path: Path, comments_xml: str | None, part_name: str = "comments.xml"
) -> Path:
    """Copy ``simple_docx`` adding ``word/{part_name}`` with its override and relationship.

    ``comments_xml=None`` adds the override and the relationship but no part.
    """
    dest = tmp_path / "with_comments.docx"
    with zipfile.ZipFile(simple_docx) as z_in, zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z_out:
        for item in z_in.infolist():
            data = z_in.read(item.filename)
            if item.filename == "[Content_Types].xml":
                override = f'<Override PartName="/word/{part_name}" ContentType="{COMMENTS_CONTENT_TYPE}"/>'
                data = data.decode("utf-8").replace("</Types>", override + "</Types>").encode("utf-8")
            elif item.filename == "word/_rels/document.xml.rels":
                rel = f'<Relationship Id="rId900" Type="{COMMENTS_REL_TYPE}" Target="{part_name}"/>'
                data = data.decode("utf-8").replace("</Relationships>", rel + "</Relationships>").encode("utf-8")
            z_out.writestr(item, data)
        if comments_xml is not None:
            z_out.writestr(f"word/{part_name}", comments_xml)
    return dest


def read_part(docx: Path, name: str) -> str:
    with zipfile.ZipFile(docx) as z:
        return z.read(name).decode("utf-8")


RELS = "word/_rels/document.xml.rels"


def rewrite_rels(src: Path, dest: Path, old: str, new: str) -> Path:
    """Copy ``src`` to ``dest`` with ``old`` replaced by ``new`` in ``document.xml.rels``."""
    rels = read_part(src, RELS)
    assert old in rels
    replace_docx_parts(src, dest, {RELS: rels.replace(old, new)})
    return dest


def add_relationship(src: Path, dest: Path, rel_type: str, target: str, extra: str = "") -> Path:
    rel = f'<Relationship Id="rId901" Type="{rel_type}" Target="{target}"{extra}/>'
    return rewrite_rels(src, dest, "</Relationships>", rel + "</Relationships>")


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


def test_delete_comment_when_comments_xml_is_the_only_part(simple_docx, tmp_path):
    """Word 2010 shape: ``w14:paraId`` on the comment, none of the three sibling parts."""
    comments_xml = ONE_OLD_COMMENT.replace("<w:p>", f'<w:p xmlns:w14="{W14_NS}" w14:paraId="00000001">')
    src = make_docx_with_comments(simple_docx, tmp_path, comments_xml)
    out = tmp_path / "out.docx"
    with Document.open(src, author="Tester") as doc:
        assert doc.delete_comment(0) is True
        assert doc.list_comments() == []
        doc.save(out)

    assert "old remark" not in read_part(out, "word/comments.xml")
    with zipfile.ZipFile(out) as z:
        assert [n for n in z.namelist() if n.startswith("word/comments")] == ["word/comments.xml"]


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
        assert {part: targets.count(part) for part in COMMENT_PART_NAMES} == dict.fromkeys(COMMENT_PART_NAMES, 1)
        ids = [r.getAttribute("Id") for r in rel_elems]
        assert len(ids) == len(set(ids))

        types = defusedxml.minidom.parseString(read_part(out, "[Content_Types].xml"))
        part_names = [o.getAttribute("PartName") for o in types.getElementsByTagName("Override")]
        assert {part: part_names.count(f"/word/{part}") for part in COMMENT_PART_NAMES} == dict.fromkeys(
            COMMENT_PART_NAMES, 1
        )

    def test_absolute_target_is_not_wired_twice(self, simple_docx, tmp_path):
        src = make_docx_with_comments(simple_docx, tmp_path, EMPTY)
        absolute = tmp_path / "absolute.docx"
        rewrite_rels(src, absolute, 'Target="comments.xml"', 'Target="/word/comments.xml"')
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


class TestCommentsPartUnderAnotherName:
    """The comments part is the one the comments relationship names, whatever its file name.

    Shape: ``word/comments1.xml`` holding the document's comments, wired by a
    comments-type relationship with ``Target="comments1.xml"``, and no
    ``word/comments.xml``.
    """

    def test_existing_comment_is_listed(self, simple_docx, tmp_path):
        src = make_docx_with_comments(simple_docx, tmp_path, ONE_OLD_COMMENT, part_name="comments1.xml")
        with Document.open(src, author="Tester") as doc:
            assert [(c.id, c.text) for c in doc.list_comments()] == [(0, "old remark")]

    def test_new_comment_survives_save(self, simple_docx, tmp_path):
        src = make_docx_with_comments(simple_docx, tmp_path, ONE_OLD_COMMENT, part_name="comments1.xml")
        out = tmp_path / "out.docx"
        with Document.open(src, author="Tester") as doc:
            new_id = doc.add_comment("quick brown", "check this")
            doc.save(out)

        assert new_id == 1
        with Document.open(out, author="Tester") as reopened:
            assert [(c.id, c.text) for c in reopened.list_comments()] == [(0, "old remark"), (1, "check this")]
        # The comment went into the part the package already uses; no second
        # comments part and no second comments relationship appeared.
        with zipfile.ZipFile(out) as z:
            assert "word/comments.xml" not in z.namelist()
        assert "check this" in read_part(out, "word/comments1.xml")
        rels = read_part(out, "word/_rels/document.xml.rels")
        assert rels.count(f'Type="{COMMENTS_REL_TYPE}"') == 1
        assert 'Target="comments1.xml"' in rels

    def test_absolute_target_names_the_same_part(self, simple_docx, tmp_path):
        src = make_docx_with_comments(simple_docx, tmp_path, ONE_OLD_COMMENT, part_name="comments1.xml")
        absolute = tmp_path / "absolute.docx"
        rewrite_rels(src, absolute, 'Target="comments1.xml"', 'Target="/word/comments1.xml"')
        out = tmp_path / "out.docx"
        with Document.open(absolute, author="Tester") as doc:
            doc.add_comment("quick brown", "check this")
            doc.save(out)

        with Document.open(out, author="Tester") as reopened:
            assert [c.text for c in reopened.list_comments()] == ["old remark", "check this"]

    def test_relationship_to_a_missing_part_refuses_the_comment(self, simple_docx, tmp_path):
        """Nothing can be written where the relationship points, so nothing is written at all."""
        src = make_docx_with_comments(simple_docx, tmp_path, None, part_name="comments1.xml")
        out = tmp_path / "out.docx"
        with Document.open(src, author="Tester") as doc:
            with pytest.raises(XMLError, match=r"comments1\.xml"):
                doc.add_comment("quick brown", "check this")

            assert marker_count(doc) == 0
            doc.save(out)
        with zipfile.ZipFile(out) as z:
            assert [n for n in z.namelist() if n.startswith("word/comments")] == []

    def test_external_target_refuses_the_comment(self, simple_docx, tmp_path):
        src = make_docx_with_comments(simple_docx, tmp_path, None, part_name="comments1.xml")
        external = tmp_path / "external.docx"
        rewrite_rels(
            src, external, 'Target="comments1.xml"', 'Target="https://example.invalid/c.xml" TargetMode="External"'
        )
        with Document.open(external, author="Tester") as doc:
            with pytest.raises(XMLError, match="external target"):
                doc.add_comment("quick brown", "check this")

            assert marker_count(doc) == 0

    def test_relationship_to_a_missing_default_part_is_filled(self, simple_docx, tmp_path):
        """A relationship to ``comments.xml`` with no such part: the part is created where it points."""
        src = make_docx_with_comments(simple_docx, tmp_path, None)
        out = tmp_path / "out.docx"
        with Document.open(src, author="Tester") as doc:
            doc.add_comment("quick brown", "check this")
            doc.save(out)

        with Document.open(out, author="Tester") as reopened:
            assert [c.text for c in reopened.list_comments()] == ["check this"]
        assert read_part(out, "word/_rels/document.xml.rels").count(f'Type="{COMMENTS_REL_TYPE}"') == 1

    def test_two_parts_of_one_type_refuse_the_comment(self, simple_docx, tmp_path):
        src = make_docx_with_comments(simple_docx, tmp_path, ONE_OLD_COMMENT, part_name="comments1.xml")
        with zipfile.ZipFile(src, "a") as z:
            z.writestr("word/comments2.xml", EMPTY)
        two = add_relationship(src, tmp_path / "two.docx", COMMENTS_REL_TYPE, "comments2.xml")
        with Document.open(two, author="Tester") as doc:
            with pytest.raises(XMLError, match="more than one part"):
                doc.add_comment("quick brown", "check this")

            assert marker_count(doc) == 0

    def test_adopted_part_without_an_override_gains_one(self, simple_docx, tmp_path):
        """Word finds a comments part by its content type; a comment must not go into an untyped part."""
        src = make_docx_with_comments(simple_docx, tmp_path, ONE_OLD_COMMENT, part_name="comments1.xml")
        override = f'<Override PartName="/word/comments1.xml" ContentType="{COMMENTS_CONTENT_TYPE}"/>'
        untyped = tmp_path / "untyped.docx"
        replace_docx_parts(
            src, untyped, {"[Content_Types].xml": read_part(src, "[Content_Types].xml").replace(override, "")}
        )
        out = tmp_path / "out.docx"
        with Document.open(untyped, author="Tester") as doc:
            doc.add_comment("quick brown", "check this")
            doc.save(out)

        assert "check this" in read_part(out, "word/comments1.xml")
        assert read_part(out, "[Content_Types].xml").count(override) == 1

    def test_override_differing_only_by_case_is_not_duplicated(self, simple_docx, tmp_path):
        """Part names compare case-insensitively: the part is already typed."""
        src = make_docx_with_comments(simple_docx, tmp_path, ONE_OLD_COMMENT, part_name="comments1.xml")
        cased = tmp_path / "cased.docx"
        types = read_part(src, "[Content_Types].xml")
        assert 'PartName="/word/comments1.xml"' in types
        replace_docx_parts(
            src, cased, {"[Content_Types].xml": types.replace("/word/comments1.xml", "/word/Comments1.xml")}
        )
        out = tmp_path / "out.docx"
        with Document.open(cased, author="Tester") as doc:
            doc.save(out)

        assert read_part(out, "[Content_Types].xml").lower().count("/word/comments1.xml") == 1

    def test_overlong_target_refuses_the_comment(self, simple_docx, tmp_path):
        src = add_relationship(simple_docx, tmp_path / "long.docx", COMMENTS_REL_TYPE, "a" * 300 + ".xml")
        with Document.open(src, author="Tester") as doc:
            with pytest.raises(XMLError, match="is not a comment part"):
                doc.add_comment("quick brown", "check this")

            assert marker_count(doc) == 0

    def test_non_integer_comment_id_is_skipped(self, simple_docx, tmp_path):
        """A comment with no usable id cannot be addressed: it is not listed, and it blocks nothing."""
        comments_xml = ONE_OLD_COMMENT.replace('w:id="0"', 'w:id="x"').replace(
            "<w:p>", f'<w:p xmlns:w14="{W14_NS}" w14:paraId="00000001">'
        )
        src = make_docx_with_comments(simple_docx, tmp_path, comments_xml, part_name="comments1.xml")
        out = tmp_path / "out.docx"
        with Document.open(src, author="Tester") as doc:
            assert doc.list_comments() == []
            assert doc.add_comment("quick brown", "check this") == 0
            assert [(c.id, c.text) for c in doc.list_comments()] == [(0, "check this")]
            doc.save(out)

        # Skipped, not dropped: the comment is still in the part.
        assert 'w:id="x"' in read_part(out, "word/comments1.xml")


class TestRelationshipNamingAnotherPart:
    """A comment relationship is honoured only when its target is a comment part.

    Adopting any other file would put a second editor on a part the document
    already edits, and the stale copy would overwrite it at save.
    """

    EXTENDED_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"

    @pytest.mark.parametrize("rel_type", [COMMENTS_REL_TYPE, EXTENDED_REL_TYPE], ids=["comments", "commentsExtended"])
    @pytest.mark.parametrize("target", ["document.xml", "_rels/document.xml.rels", "../[Content_Types].xml"])
    def test_tracked_edit_survives_save(self, simple_docx, tmp_path, rel_type, target):
        src = add_relationship(simple_docx, tmp_path / "aliased.docx", rel_type, target)
        out = tmp_path / "out.docx"
        with Document.open(src, author="Tester") as doc:
            doc.replace("quick", "swift", paragraph=find_ref(doc, "quick brown"))
            with pytest.raises(XMLError, match="is not a comment part"):
                doc.add_comment("brown", "check this")
            assert marker_count(doc) == 0
            doc.save(out)

        assert "swift" in read_part(out, "word/document.xml")
        with Document.open(out, author="Tester") as reopened:
            assert len(reopened.list_revisions()) == 2

    @pytest.mark.parametrize(
        ("rel_type", "target", "root"),
        [
            (COMMENTS_REL_TYPE, "commentsExtended.xml", f'<w:comments xmlns:w="{W_NS}"/>'),
            (
                EXTENDED_REL_TYPE,
                "comments.xml",
                '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>',
            ),
        ],
        ids=["comments-in-commentsExtended.xml", "commentsEx-in-comments.xml"],
    )
    def test_part_stored_under_another_types_default_name(self, simple_docx, tmp_path, rel_type, target, root):
        """Such a file is one type's part by relationship and another's by name: neither is wired or written."""
        src = tmp_path / "crossed.docx"
        add_relationship(simple_docx, src, rel_type, target)
        with zipfile.ZipFile(src, "a") as z:
            z.writestr(f"word/{target}", XML_DECL + root)
        out = tmp_path / "out.docx"
        with Document.open(src, author="Tester") as doc:
            with pytest.raises(XMLError, match=r"document\.xml\.rels: .* also resolves to"):
                doc.add_comment("quick brown", "check this")
            assert marker_count(doc) == 0
            doc.save(out)

        saved = read_part(out, RELS)
        assert saved.count("relationships/comments") == 1
        assert saved.count(f'Target="{target}"') == 1

    def test_external_target_with_a_local_looking_name_refuses_the_comment(self, simple_docx, tmp_path):
        src = make_docx_with_comments(simple_docx, tmp_path, EMPTY)
        external = rewrite_rels(
            src, tmp_path / "external.docx", 'Target="comments.xml"', 'Target="comments.xml" TargetMode="External"'
        )
        with Document.open(external, author="Tester") as doc:
            with pytest.raises(XMLError, match="external"):
                doc.add_comment("quick brown", "check this")

            assert marker_count(doc) == 0

    def test_another_type_targeting_comments_xml_does_not_count_as_wired(self, simple_docx, tmp_path):
        src = make_docx_with_comments(simple_docx, tmp_path, EMPTY)
        other = rewrite_rels(src, tmp_path / "other.docx", COMMENTS_REL_TYPE, "urn:not-the-comments-type")
        out = tmp_path / "out.docx"
        with Document.open(other, author="Tester") as doc:
            doc.add_comment("quick brown", "check this")
            doc.save(out)

        rels = read_part(out, RELS)
        assert rels.count(f'Type="{COMMENTS_REL_TYPE}"') == 1


class TestTargetLeavingThePackage:
    """A relationship target is honoured only when it stays inside the package."""

    OUTSIDE = f'{XML_DECL}<w:comments xmlns:w="{W_NS}"><!-- outside the package --></w:comments>'

    def test_backslash_target_is_refused(self, simple_docx, tmp_path):
        """``..\\..\\outside.xml`` is one file name on POSIX and a climb out of the package on Windows."""
        src = add_relationship(simple_docx, tmp_path / "backslash.docx", COMMENTS_REL_TYPE, "..\\..\\outside.xml")
        with Document.open(src, author="Tester", workspace_dir=tmp_path / "ws") as doc:
            workspace = doc._workspace.workspace_path
        # Where the target would land if backslashes were path separators.
        outside = workspace.parent / "outside.xml"
        outside.write_text(self.OUTSIDE, encoding="utf-8")
        before = (outside.read_bytes(), outside.stat().st_mtime_ns)

        out = tmp_path / "out.docx"
        with Document.open(src, author="Tester", workspace_dir=tmp_path / "ws") as doc:
            assert doc._workspace.workspace_path == workspace
            assert doc._comment_manager.comments_path == workspace / "word" / "comments.xml"
            with pytest.raises(XMLError, match="backslash"):
                doc.add_comment("quick brown", "check this")
            assert marker_count(doc) == 0
            doc.save(out)

        assert (outside.read_bytes(), outside.stat().st_mtime_ns) == before
        with zipfile.ZipFile(out) as z:
            assert [n for n in z.namelist() if "comments" in n.lower() or "outside" in n] == []

    def test_symlinked_part_pointing_outside_is_refused(self, simple_docx, tmp_path):
        outside = tmp_path / "outside.xml"
        outside.write_text(self.OUTSIDE, encoding="utf-8")
        before = (outside.read_bytes(), outside.stat().st_mtime_ns)
        src = add_relationship(simple_docx, tmp_path / "linked.docx", COMMENTS_REL_TYPE, "comments1.xml")

        with Document.open(src, author="Tester", workspace_dir=tmp_path / "ws") as doc:
            workspace = doc._workspace.workspace_path
            # A zip cannot carry the link (unpack refuses symlink entries), so
            # plant it in the workspace and resolve the parts again.
            (workspace / "word" / "comments1.xml").symlink_to(outside)
            manager = CommentManager(workspace, doc._document_editor, "Tester", "T")

            assert manager.comments_path == workspace / "word" / "comments.xml"
            assert manager.list_comments() == []
            with pytest.raises(XMLError, match="leaves the package"):
                manager.add_comment("quick brown", "check this")
            assert marker_count(doc) == 0
            manager.save_all()

        assert (outside.read_bytes(), outside.stat().st_mtime_ns) == before

    @pytest.mark.parametrize("with_relationship", [True, False], ids=["named-by-relationship", "no-relationship"])
    def test_symlinked_default_named_part_is_refused(self, simple_docx, tmp_path, with_relationship):
        """``word/comments.xml`` itself is a link out of the workspace: never read, never written."""
        outside = tmp_path / "outside.xml"
        outside.write_text(ONE_OLD_COMMENT, encoding="utf-8")
        before = (outside.read_bytes(), outside.stat().st_mtime_ns)
        src = simple_docx
        if with_relationship:
            src = add_relationship(simple_docx, tmp_path / "linked.docx", COMMENTS_REL_TYPE, "comments.xml")

        with Document.open(src, author="Tester", workspace_dir=tmp_path / "ws") as doc:
            workspace = doc._workspace.workspace_path
            (workspace / "word" / "comments.xml").symlink_to(outside)
            manager = CommentManager(workspace, doc._document_editor, "Tester", "T")

            # The comment in the outside file is not read into the listing.
            assert manager.existing_comments == {}
            assert manager.list_comments() == []
            assert not manager.part_exists(manager.comments_path)
            with pytest.raises(XMLError, match="leaves the package"):
                manager.add_comment("quick brown", "check this")
            assert marker_count(doc) == 0
            manager.save_all()

        assert (outside.read_bytes(), outside.stat().st_mtime_ns) == before

    def test_link_that_cannot_be_followed_is_refused(self, simple_docx, tmp_path):
        src = add_relationship(simple_docx, tmp_path / "loop.docx", COMMENTS_REL_TYPE, "comments1.xml")
        with Document.open(src, author="Tester", workspace_dir=tmp_path / "ws") as doc:
            word = doc._workspace.workspace_path / "word"
            (word / "comments1.xml").symlink_to(word / "comments2.xml")
            (word / "comments2.xml").symlink_to(word / "comments1.xml")
            manager = CommentManager(doc._workspace.workspace_path, doc._document_editor, "Tester", "T")

            with pytest.raises(XMLError):
                manager.add_comment("quick brown", "check this")

    def test_unresolvable_path_counts_as_outside(self, tmp_path, monkeypatch):
        def refuse(self, strict=False):
            raise OSError("cannot resolve")

        monkeypatch.setattr(Path, "resolve", refuse)
        assert not _is_inside(tmp_path / "word" / "comments.xml", tmp_path)


def test_package_without_document_rels_has_no_comments(simple_docx, tmp_path):
    src = tmp_path / "norels.docx"
    replace_docx_parts(simple_docx, src, {RELS: None})
    with Document.open(src, author="Tester") as doc:
        assert doc._comment_manager.part_paths == tuple(doc._workspace.word_path / name for name in COMMENT_PART_NAMES)
        assert doc.list_comments() == []


class TestWiringAtSaveChecksTheRoot:
    """A default-named file is wired at save only if it is the comment part its name says."""

    def test_default_named_file_with_another_root_is_not_wired(self, simple_docx, tmp_path):
        src = tmp_path / "stray.docx"
        shutil.copy(simple_docx, src)
        with zipfile.ZipFile(src, "a") as z:
            z.writestr("word/comments.xml", f'{XML_DECL}<w:settings xmlns:w="{W_NS}"/>')
        out = tmp_path / "out.docx"
        with Document.open(src, author="Tester") as doc:
            doc.save(out)

        assert COMMENTS_REL_TYPE not in read_part(out, RELS)
        assert "/word/comments.xml" not in read_part(out, "[Content_Types].xml")

    def test_unwired_comments_part_is_still_wired(self, simple_docx, tmp_path):
        src = tmp_path / "unwired.docx"
        shutil.copy(simple_docx, src)
        with zipfile.ZipFile(src, "a") as z:
            z.writestr("word/comments.xml", EMPTY)
        out = tmp_path / "out.docx"
        with Document.open(src, author="Tester") as doc:
            doc.save(out)

        assert read_part(out, RELS).count(f'Type="{COMMENTS_REL_TYPE}"') == 1
        assert read_part(out, "[Content_Types].xml").count('PartName="/word/comments.xml"') == 1

    def test_wired_parts_are_not_parsed_at_save(self, open_doc, tmp_path, monkeypatch):
        """An ordinary document, every part already wired, pays no root check on save."""
        doc = open_doc(EMPTY)
        doc.add_comment("quick brown", "check this")
        _, reopened = save_and_reopen(doc, tmp_path)
        try:
            checked: list[Path] = []
            monkeypatch.setattr("docx_editor.document.has_root", lambda path, root_tag: checked.append(path) or True)
            reopened.save(tmp_path / "again.docx")
        finally:
            reopened.close()

        assert checked == []


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
        """``commentsExtensible.xml`` is the last part written; it is still checked before anything is."""
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
