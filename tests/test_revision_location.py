"""Tests for revision location and nested-state reporting (ISSUES.md #33).

``list_revisions`` populates four location/nesting fields on ``Revision``:

- ``paragraph_ref`` — hash-anchored ref of the containing paragraph, matching
  ``list_paragraphs()`` output; None outside any ``<w:p>``.
- ``occurrence`` — 0-based index of the revision's text within its paragraph,
  in the view where that text lives (accepted for insertions, original for
  deletions), so it plugs into the ``occurrence=`` parameter of the anchor
  APIs. None whenever targeting-by-text does not apply (empty text, nested
  deletions, host insertions partially consumed by a nested deletion).
- ``nested_under`` / ``contains_ids`` — revision nesting, e.g. a foreign
  deletion inside another author's pending insertion (produced by this
  library's author-aware edits and by Word itself).

Also covers ``list_revisions(paragraph=...)`` filtering and the
``get_markup_text()`` verification view.
"""

from pathlib import Path

import pytest
from conftest import NS, find_ref, replace_document_xml

from docx_editor import Document, HashMismatchError, ParagraphIndexError, Revision
from docx_editor.track_changes import RevisionManager
from docx_editor.xml_editor import DocxXMLEditor, compute_paragraph_hash

AUTHOR_A = "Reviewer A"
AUTHOR_B = "Reviewer B"
DATE_A = "2024-01-01T00:00:00Z"
DATE_B = "2024-01-02T00:00:00Z"


@pytest.fixture
def temp_xml(tmp_path):
    """Fixture that returns a function to create temp XML files."""

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


def _ins(content: str, ins_id: int = 1, author: str = AUTHOR_A) -> str:
    return f'<w:ins w:id="{ins_id}" w:author="{author}" w:date="{DATE_A}">{content}</w:ins>'


def _del(content: str, del_id: int = 2, author: str = AUTHOR_B) -> str:
    return f'<w:del w:id="{del_id}" w:author="{author}" w:date="{DATE_B}">{content}</w:del>'


def _rev(manager_or_doc, rev_id: int) -> Revision:
    return next(r for r in manager_or_doc.list_revisions() if r.id == rev_id)


class TestParagraphRef:
    """paragraph_ref matches list_paragraphs() refs, including table cells."""

    def test_refs_match_list_paragraphs_across_paragraphs(self, temp_docx):
        with Document.open(temp_docx, author="Test Editor") as doc:
            refs = [entry.split("|")[0] for entry in doc.list_paragraphs()]
            first_ref, second_ref = refs[0], refs[1]
            first_ref = doc.replace("Sample", "Example", paragraph=first_ref)
            doc.delete("quick ", paragraph=second_ref)

            revisions = doc.list_revisions()
            assert revisions, "expected tracked changes"
            current_refs = [entry.split("|")[0] for entry in doc.list_paragraphs()]
            # replace() produces a del+ins pair in P1; delete() a del in P2.
            p1_revs = [r for r in revisions if r.paragraph_ref == current_refs[0]]
            p2_revs = [r for r in revisions if r.paragraph_ref == current_refs[1]]
            assert {r.type for r in p1_revs} == {"insertion", "deletion"}
            assert [r.type for r in p2_revs] == ["deletion"]
            assert len(p1_revs) + len(p2_revs) == len(revisions)
            assert first_ref == current_refs[0]

    def test_refs_in_table_cells(self, simple_docx, tmp_path):
        body = (
            "<w:p><w:r><w:t>Body para</w:t></w:r></w:p>"
            "<w:tbl><w:tr>"
            "<w:tc><w:p><w:r><w:t>Cell alpha text</w:t></w:r></w:p></w:tc>"
            "<w:tc><w:p><w:r><w:t>Cell beta text</w:t></w:r></w:p></w:tc>"
            "</w:tr></w:tbl>"
        )
        docx_path = tmp_path / "tables.docx"
        replace_document_xml(
            simple_docx, docx_path, f'<?xml version="1.0"?><w:document {NS}><w:body>{body}</w:body></w:document>'
        )
        with Document.open(docx_path, author="Test Editor") as doc:
            doc.delete("alpha ", paragraph=find_ref(doc, "Cell alpha"))
            doc.replace("beta", "BETA", paragraph=find_ref(doc, "Cell beta"))

            del_alpha = next(r for r in doc.list_revisions() if r.text == "alpha ")
            assert del_alpha.paragraph_ref == find_ref(doc, "Cell text")
            ins_beta = next(r for r in doc.list_revisions() if r.text == "BETA")
            assert ins_beta.paragraph_ref == find_ref(doc, "Cell BETA text")

    def test_revision_outside_paragraph_has_no_ref(self, temp_xml):
        """w:trPr row-insertion markers sit outside any w:p → paragraph_ref None."""
        body = (
            "<w:tbl><w:tr>"
            f'<w:trPr><w:ins w:id="9" w:author="{AUTHOR_A}" w:date="{DATE_A}"/></w:trPr>'
            "<w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc>"
            "</w:tr></w:tbl>"
        )
        manager = _make_manager(temp_xml(body))
        rev = _rev(manager, 9)
        assert rev.paragraph_ref is None
        assert rev.occurrence is None
        assert rev.text == ""


class TestOccurrence:
    """0-based occurrence in the view where the revision's text lives."""

    def test_twin_insertions_get_distinct_occurrences(self, temp_xml):
        body = (
            "<w:p><w:r><w:t>start </w:t></w:r>"
            f"{_ins('<w:r><w:t>dup</w:t></w:r>', ins_id=11)}"
            "<w:r><w:t> mid </w:t></w:r>"
            f"{_ins('<w:r><w:t>dup</w:t></w:r>', ins_id=12)}"
            "<w:r><w:t> end</w:t></w:r></w:p>"
        )
        manager = _make_manager(temp_xml(body))
        assert _rev(manager, 11).occurrence == 0
        assert _rev(manager, 12).occurrence == 1
        assert _rev(manager, 11).paragraph_ref == _rev(manager, 12).paragraph_ref

    def test_deletion_occurrence_counts_in_original_view(self, temp_xml):
        """'foo bar foo' with the second 'foo' deleted → occurrence 1."""
        body = f"<w:p><w:r><w:t>foo bar </w:t></w:r>{_del('<w:r><w:delText>foo</w:delText></w:r>', del_id=5)}</w:p>"
        manager = _make_manager(temp_xml(body))
        rev = _rev(manager, 5)
        assert rev.text == "foo"
        assert rev.occurrence == 1

    def test_occurrence_targets_intended_span_via_add_comment(self, simple_docx, tmp_path):
        """rev.occurrence feeds add_comment(...) and hits the revision's own span."""
        body = (
            "<w:p><w:r><w:t>start </w:t></w:r>"
            f"{_ins('<w:r><w:t>dup</w:t></w:r>', ins_id=11)}"
            "<w:r><w:t> mid </w:t></w:r>"
            f"{_ins('<w:r><w:t>dup</w:t></w:r>', ins_id=12)}"
            "<w:r><w:t> end</w:t></w:r></w:p>"
        )
        docx_path = tmp_path / "twins.docx"
        replace_document_xml(
            simple_docx, docx_path, f'<?xml version="1.0"?><w:document {NS}><w:body>{body}</w:body></w:document>'
        )
        with Document.open(docx_path, author="Test Editor") as doc:
            rev = _rev(doc, 12)
            assert rev.paragraph_ref is not None
            assert rev.occurrence is not None
            doc.add_comment(rev.text, "second twin", paragraph=rev.paragraph_ref, occurrence=rev.occurrence)

            dom = doc._document_editor.dom
            before: list[str] = []
            inside: list[str] = []
            state = "before"

            def walk(node):
                nonlocal state
                for child in node.childNodes:
                    if child.nodeType != child.ELEMENT_NODE:
                        continue
                    if child.tagName == "w:commentRangeStart":
                        state = "inside"
                    elif child.tagName == "w:commentRangeEnd":
                        state = "after"
                    elif child.tagName == "w:t":
                        text = "".join(c.data for c in child.childNodes if c.nodeType == c.TEXT_NODE)
                        if state == "before":
                            before.append(text)
                        elif state == "inside":
                            inside.append(text)
                    else:
                        walk(child)

            walk(dom.documentElement)
            assert "".join(inside) == "dup"
            # The anchor is the SECOND dup: the first one lies before the range.
            assert "".join(before) == "start dup mid "


class TestNestedRevisions:
    """Foreign-deletion-inside-pending-insertion state is fully reported."""

    def test_nested_full_consume(self, temp_xml):
        """B deleted ALL of A's insertion: host keeps full text, both unlocatable."""
        body = (
            "<w:p><w:r><w:t>Before </w:t></w:r>"
            f"{_ins(_del('<w:r><w:delText>Hello</w:delText></w:r>', del_id=2), ins_id=1)}"
            "<w:r><w:t> after</w:t></w:r></w:p>"
        )
        manager = _make_manager(temp_xml(body))
        host = _rev(manager, 1)
        nested = _rev(manager, 2)

        assert host.text == "Hello"  # full original insertion, not ''
        assert host.contains_ids == (2,)
        assert host.nested_under is None
        assert host.occurrence is None  # visible span no longer matches its text

        assert nested.text == "Hello"
        assert nested.nested_under == 1
        assert nested.contains_ids == ()
        assert nested.occurrence is None  # never existed in the original view

        assert host.paragraph_ref is not None
        assert host.paragraph_ref == nested.paragraph_ref

    def test_nested_partial_consume(self, temp_xml):
        """B deleted part of A's insertion: host reports the full original text."""
        ins_content = "<w:r><w:t>Hello </w:t></w:r>" + _del("<w:r><w:delText>world</w:delText></w:r>", del_id=2)
        body = f"<w:p>{_ins(ins_content, ins_id=1)}</w:p>"
        manager = _make_manager(temp_xml(body))
        host = _rev(manager, 1)
        assert host.text == "Hello world"
        assert host.contains_ids == (2,)
        assert host.occurrence is None
        assert _rev(manager, 2).nested_under == 1

    def test_partial_consume_host_unlocatable_even_when_following_text_matches(self, temp_xml):
        """Following visible text spelling the deleted suffix must not fake a match.

        The host's visible span is only "Hello "; the "world" completing
        "Hello world" belongs to the run AFTER the insertion. An occurrence
        here would anchor across the revision boundary — must be None.
        """
        ins_content = "<w:r><w:t>Hello </w:t></w:r>" + _del("<w:r><w:delText>world</w:delText></w:r>", del_id=2)
        body = f"<w:p>{_ins(ins_content, ins_id=1)}<w:r><w:t>world peace</w:t></w:r></w:p>"
        manager = _make_manager(temp_xml(body))
        host = _rev(manager, 1)
        assert host.text == "Hello world"
        assert host.occurrence is None

    def test_unnested_revisions_have_default_nesting_fields(self, temp_xml):
        body = (
            "<w:p><w:r><w:t>keep </w:t></w:r>"
            f"{_ins('<w:r><w:t>new</w:t></w:r>', ins_id=1)}"
            f"{_del('<w:r><w:delText>old</w:delText></w:r>', del_id=2)}"
            "</w:p>"
        )
        manager = _make_manager(temp_xml(body))
        for rev_id in (1, 2):
            rev = _rev(manager, rev_id)
            assert rev.nested_under is None
            assert rev.contains_ids == ()
            assert rev.occurrence == 0

    def test_library_produced_nesting_is_reported(self, temp_xml):
        """Deleting inside a foreign ins (PR #43 path) round-trips through list_revisions."""
        xml_path = temp_xml(f"<w:p>{_ins('<w:r><w:t>Hello world</w:t></w:r>', ins_id=1)}</w:p>")
        manager = _make_manager(xml_path, author=AUTHOR_B)
        p = manager.editor.dom.getElementsByTagName("w:p")[0]
        manager.suggest_deletion("world", paragraph=f"P1#{compute_paragraph_hash(p)}")

        revisions = manager.list_revisions()
        nested_del = next(r for r in revisions if r.type == "deletion")
        hosts = [r for r in revisions if r.contains_ids]
        assert len(hosts) == 1
        assert nested_del.nested_under == hosts[0].id
        assert nested_del.id in hosts[0].contains_ids


class TestWithLocationFlag:
    """with_location=False skips location work for id-only callers."""

    def test_with_location_false_leaves_location_unset(self, temp_xml):
        body = f"<w:p><w:r><w:t>keep </w:t></w:r>{_ins('<w:r><w:t>new</w:t></w:r>', ins_id=1)}</w:p>"
        manager = _make_manager(temp_xml(body))
        (rev,) = manager.list_revisions(with_location=False)
        assert rev.paragraph_ref is None
        assert rev.occurrence is None
        # Nesting state is cheap and still reported.
        assert rev.nested_under is None
        assert rev.contains_ids == ()

    def test_paragraph_filter_overrides_with_location_false(self, temp_xml):
        body = f"<w:p><w:r><w:t>keep </w:t></w:r>{_ins('<w:r><w:t>new</w:t></w:r>', ins_id=1)}</w:p>"
        manager = _make_manager(temp_xml(body))
        ref = _rev(manager, 1).paragraph_ref
        assert ref is not None
        (rev,) = manager.list_revisions(paragraph=ref, with_location=False)
        assert rev.paragraph_ref == ref


class TestParagraphFilter:
    """list_revisions(paragraph=...) scoping and validation."""

    def test_filters_to_single_paragraph(self, temp_docx):
        with Document.open(temp_docx, author="Test Editor") as doc:
            refs = [entry.split("|")[0] for entry in doc.list_paragraphs()]
            ref1 = doc.replace("Sample", "Example", paragraph=refs[0])
            doc.delete("quick ", paragraph=refs[1])

            current_refs = [entry.split("|")[0] for entry in doc.list_paragraphs()]
            p1_revs = doc.list_revisions(paragraph=current_refs[0])
            p2_revs = doc.list_revisions(paragraph=current_refs[1])
            assert {r.type for r in p1_revs} == {"insertion", "deletion"}
            assert [r.type for r in p2_revs] == ["deletion"]
            assert all(r.paragraph_ref == current_refs[0] for r in p1_revs)
            assert len(p1_revs) + len(p2_revs) == len(doc.list_revisions())
            assert ref1 == current_refs[0]

    def test_composes_with_author_filter(self, temp_xml):
        body = (
            "<w:p>"
            f"{_ins('<w:r><w:t>one</w:t></w:r>', ins_id=1, author=AUTHOR_A)}"
            f"{_ins('<w:r><w:t>two</w:t></w:r>', ins_id=2, author=AUTHOR_B)}"
            "</w:p>"
        )
        manager = _make_manager(temp_xml(body))
        ref = _rev(manager, 1).paragraph_ref
        assert ref is not None
        both = manager.list_revisions(paragraph=ref)
        assert [r.id for r in both] == [1, 2]
        only_a = manager.list_revisions(author=AUTHOR_A, paragraph=ref)
        assert [r.id for r in only_a] == [1]

    def test_invalid_ref_format_raises_value_error(self, temp_docx):
        with Document.open(temp_docx, author="Test Editor") as doc:
            with pytest.raises(ValueError, match="Invalid paragraph reference"):
                doc.list_revisions(paragraph="P#bad")

    def test_out_of_range_index_raises(self, temp_docx):
        with Document.open(temp_docx, author="Test Editor") as doc:
            with pytest.raises(ParagraphIndexError):
                doc.list_revisions(paragraph="P999#abcd")

    def test_stale_hash_raises(self, temp_docx):
        with Document.open(temp_docx, author="Test Editor") as doc:
            ref = doc.list_paragraphs()[0].split("|")[0]
            index, actual_hash = ref.split("#")
            stale_hash = "0000" if actual_hash != "0000" else "ffff"
            with pytest.raises(HashMismatchError):
                doc.list_revisions(paragraph=f"{index}#{stale_hash}")


class TestRevisionRepr:
    """repr spells the type, omits fake ellipses, and shows location/nesting."""

    def test_deletion_id_zero_renders_cleanly(self):
        rev = Revision(id=0, type="deletion", author="Bob", date=None, text="gone")
        assert repr(rev) == "Revision(del 0: 'gone' by Bob)"

    def test_short_text_has_no_ellipsis(self):
        rev = Revision(id=1, type="insertion", author="Ann", date=None, text="short")
        assert "..." not in repr(rev)

    def test_long_text_is_truncated_with_ellipsis(self):
        rev = Revision(id=1, type="insertion", author="Ann", date=None, text="x" * 40)
        assert f"'{'x' * 30}...'" in repr(rev)

    def test_location_and_nesting_shown(self):
        rev = Revision(
            id=7,
            type="deletion",
            author="Bob",
            date=None,
            text="gone",
            paragraph_ref="P3#a7b2",
            nested_under=5,
        )
        assert repr(rev) == "Revision(del 7 @P3#a7b2: 'gone' by Bob, nested_under=5)"

    def test_contains_shown_when_nonempty(self):
        rev = Revision(
            id=5,
            type="insertion",
            author="Ann",
            date=None,
            text="Hello",
            paragraph_ref="P3#a7b2",
            contains_ids=(7,),
        )
        assert repr(rev) == "Revision(ins 5 @P3#a7b2: 'Hello' by Ann, contains=[7])"


class TestGetMarkupText:
    """Inline [ins#id:author]/[del#id:author] verification view."""

    def test_markup_renders_plain_ins_del_and_nested(self, temp_xml):
        body = (
            "<w:p><w:r><w:t>Plain</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>Keep </w:t></w:r>"
            f"{_ins('<w:r><w:t>added</w:t></w:r>', ins_id=1, author='A')}"
            f"{_del('<w:r><w:delText>removed</w:delText></w:r>', del_id=2, author='B')}"
            "</w:p>"
            "<w:p>"
            f'<w:ins w:id="3" w:author="A" w:date="{DATE_A}">'
            "<w:r><w:t>kept </w:t></w:r>"
            f'<w:del w:id="4" w:author="B" w:date="{DATE_B}"><w:r><w:delText>gone</w:delText></w:r></w:del>'
            "</w:ins>"
            "</w:p>"
        )
        manager = _make_manager(temp_xml(body))
        assert manager.get_markup_text() == (
            "Plain\nKeep [ins#1:A]added[/ins][del#2:B]removed[/del]\n[ins#3:A]kept [del#4:B]gone[/del][/ins]"
        )

    def test_document_wrapper_shows_tracked_edit(self, temp_docx):
        with Document.open(temp_docx, author="Test Editor") as doc:
            ref = doc.list_paragraphs()[0].split("|")[0]
            doc.replace("Sample", "Example", paragraph=ref)
            markup = doc.get_markup_text()
            assert "[del#" in markup
            assert "]Sample[/del]" in markup
            assert "]Example[/ins]" in markup
            assert ":Test Editor]" in markup
