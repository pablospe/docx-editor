"""Text-box content is excluded from the document's addressable surface.

Word normally stores one text box twice — once under ``mc:Choice``
(``wps:txbx``) and once under ``mc:Fallback`` (``v:textbox``) — so before this
exclusion such a box leaked its text four times: twice inline in the host
paragraph's text map and twice more as paragraphs of its own. Those extra
paragraphs were addressable, which made each copy independently editable and
let one edit desynchronize the pair. The exclusion does not depend on the copy
count: a box stored once (see ``TestSingleCopyBox``) is excluded the same way.

Text boxes are therefore not an editing surface at all: their paragraphs are
absent from ``paragraph_count``/``list_paragraphs``/``find_all``, their text is
absent from every text view and from paragraph hashes, and their content is
carried through the save unchanged (ISSUES.md #65).
"""

import re
import zipfile
from pathlib import Path

import pytest
from conftest import replace_document_xml

from docx_editor import Document, TextNotFoundError

# Namespaces a Word-shaped text box needs. simple.docx declares all of these on
# its own <w:document>; a swapped-in body has to redeclare them.
BOX_NS = " ".join([
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
    'xmlns:v="urn:schemas-microsoft-com:vml"',
])


def word_box(inner: str = "<w:p><w:r><w:t>BOXED</w:t></w:r></w:p>", fallback_inner: str | None = None) -> str:
    """A Word-shaped text box: the ``mc:Choice``/``mc:Fallback`` pair.

    Both twins hold the same text, which is exactly why the box is not
    addressable — an edit through a ref could only ever reach one of them.
    ``fallback_inner`` overrides the fallback copy's content, for fixtures whose
    twins must carry distinct ``w:id`` attributes.
    """
    return (
        "<mc:AlternateContent>"
        '<mc:Choice Requires="wps"><w:drawing><wp:anchor><wps:wsp><wps:txbx>'
        f"<w:txbxContent>{inner}</w:txbxContent>"
        "</wps:txbx></wps:wsp></wp:anchor></w:drawing></mc:Choice>"
        "<mc:Fallback><w:pict><v:shape><v:textbox>"
        f"<w:txbxContent>{fallback_inner if fallback_inner is not None else inner}</w:txbxContent>"
        "</v:textbox></v:shape></w:pict></mc:Fallback>"
        "</mc:AlternateContent>"
    )


def _boxed_ins(rev_id: int) -> str:
    """One tracked insertion, as a text box's whole paragraph content."""
    return (
        "<w:p>"
        f'<w:ins w:id="{rev_id}" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">'
        "<w:r><w:t>BOXADD</w:t></w:r></w:ins>"
        "</w:p>"
    )


def make_docx(simple_docx: Path, dest: Path, body: str) -> Path:
    """simple.docx with its body replaced by ``body`` (a ``<w:body>`` fragment)."""
    replace_document_xml(simple_docx, dest, f"<w:document {BOX_NS}><w:body>{body}</w:body></w:document>")
    return dest


HOST_BODY = (
    f"<w:p><w:r><w:t>Host before </w:t>{word_box()}<w:t> host after</w:t></w:r></w:p>"
    "<w:p><w:r><w:t>Second body paragraph</w:t></w:r></w:p>"
)

HOST_TEXT = "Host before  host after"


@pytest.fixture
def box_docx(simple_docx, temp_dir) -> Path:
    """Two body paragraphs; the first anchors one Word-shaped text box."""
    return make_docx(simple_docx, temp_dir / "textbox.docx", HOST_BODY)


@pytest.fixture
def box_doc(box_docx):
    doc = Document.open(box_docx)
    yield doc
    doc.close()


def saved_document_xml(path: Path) -> str:
    return zipfile.ZipFile(path).read("word/document.xml").decode()


def txbx_contents(xml: str) -> list[str]:
    """Every ``w:txbxContent`` subtree in ``xml``, verbatim."""
    return re.findall(r"<w:txbxContent>.*?</w:txbxContent>", xml, flags=re.DOTALL)


class TestHostParagraphText:
    """The host paragraph's text stops at the box."""

    def test_host_text_excludes_both_box_copies(self, box_doc):
        assert box_doc.list_paragraphs()[0].split("| ", 1)[1] == HOST_TEXT

    def test_paragraph_count_ignores_box_paragraphs(self, box_doc):
        assert box_doc.paragraph_count() == 2

    def test_list_paragraphs_has_no_box_row(self, box_doc):
        entries = box_doc.list_paragraphs()
        assert len(entries) == 2
        assert not any("BOXED" in entry for entry in entries)

    def test_visible_text_has_no_box_text(self, box_doc):
        assert box_doc.get_visible_text() == f"{HOST_TEXT}\nSecond body paragraph"

    def test_original_text_has_no_box_text(self, box_doc):
        assert box_doc.get_original_text() == f"{HOST_TEXT}\nSecond body paragraph"

    def test_markup_text_has_no_box_text(self, box_doc):
        assert box_doc.get_markup_text() == f"{HOST_TEXT}\nSecond body paragraph"

    def test_get_paragraph_matches_the_listing(self, box_doc):
        assert box_doc.get_paragraph(1).text == HOST_TEXT
        assert box_doc.get_paragraph(2).text == "Second body paragraph"


class TestSearchSkipsBoxes:
    """Box text is not findable — there is no ref that could act on a match."""

    def test_find_all_returns_nothing(self, box_doc):
        assert box_doc.find_all("BOXED") == []

    def test_count_matches_is_zero(self, box_doc):
        assert box_doc.count_matches("BOXED") == 0

    def test_find_text_returns_none(self, box_doc):
        assert box_doc.find_text("BOXED") is None

    def test_body_text_is_still_found(self, box_doc):
        match = box_doc.find_text("host after")
        assert match is not None
        assert match.paragraph_index == 1


class TestHostHashIsBoxIndependent:
    """The host paragraph hashes its own text, so the box cannot shift it."""

    def test_hash_equals_the_same_paragraph_without_the_drawing(self, simple_docx, temp_dir):
        stripped = "<w:p><w:r><w:t>Host before </w:t><w:t> host after</w:t></w:r></w:p>"
        with_box = make_docx(simple_docx, temp_dir / "with_box.docx", HOST_BODY)
        without_box = make_docx(simple_docx, temp_dir / "without_box.docx", stripped)

        boxed = Document.open(with_box)
        plain = Document.open(without_box)
        try:
            assert boxed.list_paragraphs()[0] == plain.list_paragraphs()[0]
        finally:
            boxed.close()
            plain.close()

    def test_hash_survives_save_and_reopen(self, box_docx):
        doc = Document.open(box_docx)
        try:
            before = doc.list_paragraphs()
            doc.save()
        finally:
            doc.close()

        reopened = Document.open(box_docx)
        try:
            assert reopened.list_paragraphs() == before
        finally:
            reopened.close()


class TestBoxXmlIsCarriedThrough:
    """Excluded means carried through, not dropped and not rewritten.

    An edit to the host paragraph re-serializes the run that carries the box,
    which lets the rsid-injection pass stamp ``w:rsidR``/``w14:paraId`` onto the
    box's own elements — pre-existing behavior for any re-serialized subtree,
    and symmetric across the twins. What must not change is the content: both
    copies survive, holding the same text, still exactly two of them.
    """

    def test_replace_on_the_host_carries_both_copies_through(self, box_docx):
        assert len(txbx_contents(saved_document_xml(box_docx))) == 2

        doc = Document.open(box_docx)
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            doc.replace("Host before", "HOST BEFORE", paragraph=ref)
            doc.save()
        finally:
            doc.close()

        after_xml = saved_document_xml(box_docx)
        after = txbx_contents(after_xml)
        assert len(after) == 2
        # Neither copy lost its text, and neither gained the host's edit.
        assert all(copy.count("<w:t>BOXED</w:t>") == 1 for copy in after)
        assert all("HOST BEFORE" not in copy for copy in after)
        assert after_xml.count("BOXED") == 2
        assert "HOST BEFORE" in after_xml

    def test_the_box_is_not_re_enumerated_after_the_edit(self, box_docx):
        doc = Document.open(box_docx)
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            doc.replace("Host before", "HOST BEFORE", paragraph=ref)
            assert doc.paragraph_count() == 2
            assert doc.get_visible_text() == "HOST BEFORE  host after\nSecond body paragraph"
            assert doc.find_all("BOXED") == []
        finally:
            doc.close()

    def test_rewrite_carries_the_box_through(self, box_docx):
        """The diff path edits the host's text map only.

        ``rewrite_paragraph`` applies difflib hunks in reverse over the whole
        paragraph — the widest text-map consumer there is — so a box that
        leaked back into the map would be diffed against and rewritten.
        """
        doc = Document.open(box_docx)
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            doc.rewrite_paragraph(ref, "Host first  host later")
            assert doc.get_visible_text() == "Host first  host later\nSecond body paragraph"
            doc.save()
        finally:
            doc.close()

        after = txbx_contents(saved_document_xml(box_docx))
        assert len(after) == 2
        assert all(copy.count("<w:t>BOXED</w:t>") == 1 for copy in after)

    def test_split_carries_the_box_through(self, box_docx):
        """Splitting the host moves nodes; the box must move once, intact."""
        doc = Document.open(box_docx)
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            doc.split_paragraph(ref, before="host after")
            assert doc.paragraph_count() == 3
            assert doc.find_all("BOXED") == []
            doc.save()
        finally:
            doc.close()

        after = txbx_contents(saved_document_xml(box_docx))
        assert len(after) == 2
        assert all(copy.count("<w:t>BOXED</w:t>") == 1 for copy in after)


class TestCommentAnchorsSkipBoxes:
    """Comment anchoring shares the exclusion — both its scoped and its
    document-wide lookup enumerate the same body paragraphs."""

    def test_box_text_is_not_an_anchor(self, box_doc):
        with pytest.raises(TextNotFoundError):
            box_doc.add_comment("BOXED", "should not anchor")

    def test_box_text_is_not_an_anchor_in_the_host_paragraph(self, box_doc):
        ref = box_doc.list_paragraphs()[0].split("|")[0]
        with pytest.raises(TextNotFoundError):
            box_doc.add_comment("BOXED", "should not anchor", paragraph=ref)

    def test_host_text_still_anchors(self, box_doc):
        comment_id = box_doc.add_comment("host after", "a real anchor")
        assert comment_id >= 0
        assert [c.text for c in box_doc.list_comments()] == ["a real anchor"]


class TestIndexSpacesAgree:
    """Every enumeration must share one index space — a partial migration
    would have index N mean different paragraphs in different methods."""

    def test_count_listing_search_and_lookup_agree(self, box_doc):
        count = box_doc.paragraph_count()
        entries = box_doc.list_paragraphs(limit=None)
        assert len(entries) == count

        for i, entry in enumerate(entries, start=1):
            ref = entry.split("|")[0]
            info = box_doc.get_paragraph(i)
            assert info.ref == ref
            # The ref resolves without HashMismatchError, to this same text.
            assert box_doc.find_all(info.text, paragraph=ref)[0].paragraph_index == i

    def test_structured_listing_and_locations_agree(self, box_doc):
        structured = box_doc.list_paragraphs_structured(limit=None)
        locations = box_doc.list_paragraph_locations()
        assert len(structured) == len(locations) == box_doc.paragraph_count()
        assert [info.ref for info in structured] == [ref for ref, _ in locations]


class TestRevisionsInsideABox:
    """A revision inside a box still lists and still resolves by id — only its
    location is unavailable, because its paragraph is not addressable."""

    @pytest.fixture
    def boxed_revision_docx(self, simple_docx, temp_dir) -> Path:
        body = (
            f"<w:p><w:r><w:t>Host </w:t>{word_box(_boxed_ins(90), _boxed_ins(91))}<w:t> tail</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>Second body paragraph</w:t></w:r></w:p>"
        )
        return make_docx(simple_docx, temp_dir / "boxed_revision.docx", body)

    def test_listed_without_a_location(self, boxed_revision_docx):
        doc = Document.open(boxed_revision_docx)
        try:
            revisions = [r for r in doc.list_revisions() if r.id in (90, 91)]
            assert len(revisions) == 2  # mc:Choice + mc:Fallback copies
            for rev in revisions:
                assert rev.type == "insertion"
                assert rev.text == "BOXADD"
                assert rev.paragraph_ref is None
                # Half a location is worse than none: an occurrence with no
                # ref cannot be passed to any edit method.
                assert rev.occurrence is None
        finally:
            doc.close()

    def test_the_host_paragraph_is_unaffected(self, boxed_revision_docx):
        doc = Document.open(boxed_revision_docx)
        try:
            assert doc.paragraph_count() == 2
            assert doc.get_visible_text() == "Host  tail\nSecond body paragraph"
        finally:
            doc.close()

    def test_accept_revision_resolves_only_the_copy_it_lands_on(self, boxed_revision_docx):
        """Pins the per-copy limitation: one id-keyed call, one copy.

        A box is stored twice, so one logical insertion is two w:ins elements.
        Resolving one by id leaves the other pending — the twins go out of
        step, which is the same desynchronization that makes box paragraphs
        unaddressable in the first place.
        """
        doc = Document.open(boxed_revision_docx)
        try:
            assert doc.accept_revision(90) is True
            assert [r.id for r in doc.list_revisions()] == [91]
            doc.save()
        finally:
            doc.close()

        assert saved_document_xml(boxed_revision_docx).count("<w:ins ") == 1

    def test_accept_changeset_resolves_both_copies_when_they_are_groupable(self, simple_docx, temp_dir):
        """Twins with distinct ids join one inferred changeset.

        They share a ``(w:author, w:date)``, so resolving the changeset takes
        both while another author's edit survives — which ``accept_all()``
        would not. Note the scope: an inferred changeset is a global class
        over the author and the *identical raw* ``w:date`` string, so it also
        takes anything that author stamped with that exact string, box or
        not — and leaves a revision whose date only means the same instant
        (``.000Z`` for ``Z``) in a changeset of its own.
        """
        body = (
            f"<w:p><w:r>{word_box(_boxed_ins(90), _boxed_ins(91))}</w:r></w:p>"
            # Same author and the identical raw w:date as the twins, so it
            # joins their changeset: the scope the docstring warns about.
            '<w:p><w:ins w:id="93" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>SWEPT</w:t></w:r></w:ins></w:p>"
            # Same author and the same *instant*, but a different raw string —
            # the key is the string, so this one is a changeset of its own.
            '<w:p><w:ins w:id="94" w:author="Reviewer" w:date="2024-01-01T00:00:00.000Z">'
            "<w:r><w:t>KEPT</w:t></w:r></w:ins></w:p>"
            '<w:p><w:ins w:id="92" w:author="Someone Else" w:date="2024-06-01T00:00:00Z">'
            "<w:r><w:t>OTHER</w:t></w:r></w:ins></w:p>"
        )
        docx = make_docx(simple_docx, temp_dir / "groupable_twins.docx", body)
        doc = Document.open(docx)
        try:
            twins = [r for r in doc.list_revisions() if r.id in (90, 91)]
            (changeset_id,) = {r.changeset_id for r in twins}
            assert changeset_id is not None
            # Both copies plus the unrelated insertion with the identical
            # raw date — but not the one whose date merely parses the same.
            assert doc.accept_changeset(changeset_id) == 3
            assert [r.id for r in doc.list_revisions()] == [92, 94]
        finally:
            doc.close()

    def test_twins_sharing_an_id_are_ungroupable(self, simple_docx, temp_dir):
        """Why accept_all is the unconditional path.

        A producer that copies the mc:Choice content into mc:Fallback
        verbatim carries the w:id across with it. A duplicated id is
        ungroupable, so group_id and changeset_id are both None and no group-
        or changeset-keyed call can reach either copy.
        """
        body = f"<w:p><w:r>{word_box(_boxed_ins(90))}</w:r></w:p>"
        docx = make_docx(simple_docx, temp_dir / "same_id_twins.docx", body)
        doc = Document.open(docx)
        try:
            revisions = doc.list_revisions()
            assert [(r.id, r.group_id, r.changeset_id) for r in revisions] == [(90, None, None)] * 2
            # Repeating the call happens to reach the twin here, but that is
            # a property of the duplicated id, not a technique: any third
            # element sharing the id would absorb one of the calls. accept_all
            # is what the docs recommend, and it is checked below.
            assert [doc.accept_revision(90) for _ in range(3)] == [True, True, False]
            assert doc.list_revisions() == []
        finally:
            doc.close()

        for name, resolve in (("accept", "accept_all"), ("reject", "reject_all")):
            doc = Document.open(make_docx(simple_docx, temp_dir / f"same_id_{name}.docx", body))
            try:
                assert getattr(doc, resolve)() == 2
                assert doc.list_revisions() == []
            finally:
                doc.close()

    def test_accept_all_still_resolves_it(self, boxed_revision_docx):
        doc = Document.open(boxed_revision_docx)
        try:
            assert doc.accept_all() == 2
            doc.save()
        finally:
            doc.close()

        xml = saved_document_xml(boxed_revision_docx)
        assert "<w:ins " not in xml
        assert xml.count("BOXADD") == 2


class TestMarksWithNoParagraphOfTheirOwn:
    """A boxed mark must not borrow the host paragraph's ref.

    ``w:trPr`` row markers and ``w:tblPrChange`` have no ``w:p`` between them
    and the box's edge, so the plain ancestor walk climbs out of the box and
    lands on the host — which would attribute the box's content to a paragraph
    whose own text excludes it, and return it from a ``paragraph=`` filter on
    that ref.
    """

    ROW_MARK = (
        "<w:tbl><w:tr><w:trPr>"
        '<w:ins w:id="77" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"/>'
        "</w:trPr><w:tc><w:p><w:r><w:t>CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    )
    TBL_CHANGE = (
        "<w:tbl><w:tblPr>"
        '<w:tblPrChange w:id="55" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:tblPr/></w:tblPrChange>'
        "</w:tblPr><w:tr><w:tc><w:p><w:r><w:t>CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    )

    def _doc(self, simple_docx, temp_dir, name: str, inner: str) -> Path:
        return make_docx(
            simple_docx,
            temp_dir / name,
            f"<w:p><w:r><w:t>Host</w:t>{word_box(inner)}</w:r></w:p>",
        )

    def test_boxed_row_marker_has_no_location(self, simple_docx, temp_dir):
        doc = Document.open(self._doc(simple_docx, temp_dir, "boxed_row.docx", self.ROW_MARK))
        try:
            host_ref = doc.list_paragraphs()[0].split("|")[0]
            assert [(r.id, r.paragraph_ref) for r in doc.list_revisions()] == [(77, None)] * 2
            # …and the host's own filter does not pick them up.
            assert doc.list_revisions(paragraph=host_ref) == []
        finally:
            doc.close()

    def test_boxed_unhandled_mark_has_no_location(self, simple_docx, temp_dir):
        doc = Document.open(self._doc(simple_docx, temp_dir, "boxed_tblpr.docx", self.TBL_CHANGE))
        try:
            rows = doc.list_unhandled_revisions()
            assert [(r.tag, r.paragraph_ref) for r in rows] == [("w:tblPrChange", None)] * 2
        finally:
            doc.close()

    def test_a_body_row_marker_is_unchanged(self, simple_docx, temp_dir):
        """The same mark outside a box already reported None — this is the
        behaviour boxed marks now match, not a new one."""
        docx = make_docx(
            simple_docx, temp_dir / "body_row.docx", f"<w:p><w:r><w:t>Host</w:t></w:r></w:p>{self.ROW_MARK}"
        )
        doc = Document.open(docx)
        try:
            assert [(r.id, r.paragraph_ref) for r in doc.list_revisions()] == [(77, None)]
        finally:
            doc.close()


class TestSingleCopyBox:
    """A box stored in one form only is listed once and behaves normally.

    The twin caveats are about how a box is stored, not about boxes: a bare
    VML ``w:pict`` has no ``mc:Fallback`` copy, so nothing is duplicated. The
    exclusion applies all the same.
    """

    @pytest.fixture
    def vml_box_docx(self, simple_docx, temp_dir) -> Path:
        vml = (
            "<w:pict><v:shape><v:textbox>"
            f"<w:txbxContent>{_boxed_ins(90)}</w:txbxContent>"
            "</v:textbox></v:shape></w:pict>"
        )
        return make_docx(simple_docx, temp_dir / "vml_box.docx", f"<w:p><w:r>{vml}</w:r></w:p>")

    def test_text_is_still_excluded(self, vml_box_docx):
        doc = Document.open(vml_box_docx)
        try:
            assert doc.paragraph_count() == 1
            assert doc.get_visible_text() == ""
            assert doc.has_textbox_content is True
        finally:
            doc.close()

    def test_its_revision_is_listed_once_and_is_groupable(self, vml_box_docx):
        doc = Document.open(vml_box_docx)
        try:
            (revision,) = doc.list_revisions()
            assert revision.id == 90
            assert revision.paragraph_ref is None
            assert revision.group_id is not None
            assert doc.accept_revision(90) is True
            assert doc.list_revisions() == []
        finally:
            doc.close()


class TestHostRevisionsIgnoreBoxText:
    """A w:ins/w:del wrapping a run that carries a box reports the text it
    changed, not the box's content."""

    def test_insertion_text_stops_at_the_box(self, simple_docx, temp_dir):
        body = (
            "<w:p>"
            '<w:ins w:id="7" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">'
            f"<w:r><w:t>added </w:t>{word_box()}<w:t>tail</w:t></w:r></w:ins>"
            "</w:p>"
        )
        docx = make_docx(simple_docx, temp_dir / "ins_around_box.docx", body)
        doc = Document.open(docx)
        try:
            (revision,) = [r for r in doc.list_revisions() if r.id == 7]
            assert revision.text == "added tail"
            assert revision.paragraph_ref is not None
        finally:
            doc.close()

    def test_deletion_text_stops_at_the_box(self, simple_docx, temp_dir):
        body = (
            "<w:p>"
            '<w:del w:id="8" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">'
            f"<w:r><w:delText>gone </w:delText>{word_box()}<w:delText>too</w:delText></w:r></w:del>"
            "</w:p>"
        )
        docx = make_docx(simple_docx, temp_dir / "del_around_box.docx", body)
        doc = Document.open(docx)
        try:
            (revision,) = [r for r in doc.list_revisions() if r.id == 8]
            assert revision.text == "gone too"
        finally:
            doc.close()


class TestTableIndexSpace:
    """Tables inside a box share the paragraphs' fate: they are not counted.

    Otherwise the paragraph index space would close while the table index
    space stayed open, and ``TableCell.index`` would name tables no ref can
    reach (ISSUES.md #65).
    """

    @pytest.fixture
    def boxed_table_docx(self, simple_docx, temp_dir) -> Path:
        """A body table preceded by a table living inside a text box."""
        boxed_table = "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>BOXCELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        body = (
            f"<w:p><w:r><w:t>Host</w:t>{word_box(boxed_table)}</w:r></w:p>"
            "<w:tbl><w:tr><w:tc>"
            "<w:p><w:r><w:t>Real cell</w:t></w:r></w:p>"
            "</w:tc></w:tr></w:tbl>"
        )
        return make_docx(simple_docx, temp_dir / "boxed_table.docx", body)

    def test_the_only_addressable_table_is_index_one(self, boxed_table_docx):
        doc = Document.open(boxed_table_docx)
        try:
            assert doc.paragraph_count() == 2
            location = doc.get_paragraph_location(doc.list_paragraphs()[1].split("|")[0])
            assert location.table is not None
            assert location.table.index == 1
        finally:
            doc.close()

    def test_batch_and_per_call_index_agree(self, boxed_table_docx):
        """list_paragraph_locations() precomputes the map; the per-ref path
        rescans. The two must not drift apart — and both must land on 1.

        Filtering only one of them would make index N mean different tables in
        different methods; filtering neither would agree on the wrong number,
        which is why the expected value is asserted here too.
        """
        doc = Document.open(boxed_table_docx)
        try:
            batch = [loc for _, loc in doc.list_paragraph_locations()]
            per_call = [doc.get_paragraph_location(doc.get_paragraph(i).ref) for i in (1, 2)]
            assert [loc.table for loc in batch] == [loc.table for loc in per_call]
            assert [loc.table and loc.table.index for loc in batch] == [None, 1]
        finally:
            doc.close()


class TestHasTextboxContent:
    """The one signal that an all-text-box document is not simply empty."""

    def test_true_when_a_box_is_present(self, box_doc):
        assert box_doc.has_textbox_content is True

    def test_false_for_an_empty_box(self, simple_docx, temp_dir):
        """A box with no paragraph in it hides nothing, so the flag stays False.

        The flag is the complement of what ``body_paragraphs`` drops, not a
        "is a drawing present" check — reporting hidden text where there is
        none is the false positive it exists to avoid.
        """
        docx = make_docx(simple_docx, temp_dir / "empty_box_flag.docx", f"<w:p><w:r>{word_box('')}</w:r></w:p>")
        doc = Document.open(docx)
        try:
            assert doc.has_textbox_content is False
        finally:
            doc.close()

    def test_false_for_a_plain_document(self, simple_docx, temp_dir):
        docx = make_docx(simple_docx, temp_dir / "plain.docx", "<w:p><w:r><w:t>Plain</w:t></w:r></w:p>")
        doc = Document.open(docx)
        try:
            assert doc.has_textbox_content is False
        finally:
            doc.close()

    def test_distinguishes_an_all_box_document_from_an_empty_one(self, simple_docx, temp_dir):
        r"""Both read as blank text; only the flag tells them apart.

        Each host paragraph still contributes its (empty) line, so three boxes
        in three host paragraphs give "\n\n" while three in one give "" — the
        visible text is never guaranteed to be exactly "". That is why the
        documented idiom tests ``.strip()``.
        """
        lines = ["ACME REPORT", "Revenue grew 12 percent.", "Q3 2024"]
        spread = make_docx(
            simple_docx,
            temp_dir / "poster_spread.docx",
            "".join(f"<w:p><w:r>{word_box(f'<w:p><w:r><w:t>{line}</w:t></w:r></w:p>')}</w:r></w:p>" for line in lines),
        )
        shared = make_docx(
            simple_docx,
            temp_dir / "poster_shared.docx",
            "<w:p><w:r>"
            + "".join(word_box(f"<w:p><w:r><w:t>{line}</w:t></w:r></w:p>") for line in lines)
            + "</w:r></w:p>",
        )
        empty = make_docx(simple_docx, temp_dir / "empty.docx", "<w:p/>")
        one_per_paragraph = Document.open(spread)
        all_in_one = Document.open(shared)
        blank = Document.open(empty)
        try:
            assert one_per_paragraph.get_visible_text() == "\n\n"
            assert one_per_paragraph.paragraph_count() == 3
            assert all_in_one.get_visible_text() == ""
            assert all_in_one.paragraph_count() == 1
            assert blank.get_visible_text() == ""
            # The documented idiom: .strip() covers both box layouts, and the
            # flag is what separates them from a genuinely empty document.
            for doc in (one_per_paragraph, all_in_one):
                assert not doc.get_visible_text().strip()
                assert doc.has_textbox_content is True
            assert not blank.get_visible_text().strip()
            assert blank.has_textbox_content is False
        finally:
            one_per_paragraph.close()
            all_in_one.close()
            blank.close()


class TestBoxPlacementVariants:
    """The exclusion is about ``w:txbxContent``, not about where the box sits."""

    def test_box_in_a_table_cell_keeps_the_cell_paragraph(self, simple_docx, temp_dir):
        body = (
            "<w:tbl><w:tr><w:tc>"
            f"<w:p><w:r><w:t>Cell </w:t>{word_box()}<w:t>text</w:t></w:r></w:p>"
            "</w:tc></w:tr></w:tbl>"
            "<w:p><w:r><w:t>After table</w:t></w:r></w:p>"
        )
        docx = make_docx(simple_docx, temp_dir / "box_in_cell.docx", body)
        doc = Document.open(docx)
        try:
            assert doc.paragraph_count() == 2
            assert doc.get_visible_text() == "Cell text\nAfter table"
            assert doc.get_paragraph(1).in_table is True
        finally:
            doc.close()

    def test_box_inside_an_insertion_is_still_excluded(self, simple_docx, temp_dir):
        body = (
            "<w:p>"
            '<w:ins w:id="11" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">'
            f"<w:r><w:t>ins </w:t>{word_box()}</w:r></w:ins>"
            "<w:r><w:t>plain</w:t></w:r></w:p>"
        )
        docx = make_docx(simple_docx, temp_dir / "box_in_ins.docx", body)
        doc = Document.open(docx)
        try:
            assert doc.paragraph_count() == 1
            assert doc.get_visible_text() == "ins plain"
            assert doc.find_all("BOXED") == []
        finally:
            doc.close()

    def test_box_only_paragraph_stays_enumerated_with_empty_text(self, simple_docx, temp_dir):
        body = f"<w:p><w:r>{word_box()}</w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>"
        docx = make_docx(simple_docx, temp_dir / "box_only.docx", body)
        doc = Document.open(docx)
        try:
            assert doc.paragraph_count() == 2
            assert doc.get_paragraph(1).text == ""
            assert doc.get_visible_text() == "\nAfter"
        finally:
            doc.close()

    def test_nested_boxes_are_all_excluded(self, simple_docx, temp_dir):
        inner_box = word_box("<w:p><w:r><w:t>INNER</w:t></w:r></w:p>")
        outer = word_box(f"<w:p><w:r><w:t>OUTER</w:t>{inner_box}</w:r></w:p>")
        body = f"<w:p><w:r><w:t>Host</w:t>{outer}</w:r></w:p>"
        docx = make_docx(simple_docx, temp_dir / "nested_boxes.docx", body)
        doc = Document.open(docx)
        try:
            assert doc.paragraph_count() == 1
            assert doc.get_visible_text() == "Host"
            assert doc.find_all("INNER") == []
            assert doc.find_all("OUTER") == []
        finally:
            doc.close()

    def test_empty_box_changes_nothing(self, simple_docx, temp_dir):
        body = f"<w:p><w:r><w:t>Host</w:t>{word_box('')}</w:r></w:p>"
        docx = make_docx(simple_docx, temp_dir / "empty_box.docx", body)
        doc = Document.open(docx)
        try:
            assert doc.paragraph_count() == 1
            assert doc.get_visible_text() == "Host"
        finally:
            doc.close()
