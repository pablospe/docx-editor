"""Edits landing inside a field result are refused (ROADMAP.md #63c).

Word regenerates a field's result (TOC entries, page numbers, cross-reference
text) whenever fields update, so a tracked edit there would be silently
discarded. Every edit verb refuses a target with any character in a result,
before mutating; text outside the field stays editable, reads are unchanged,
and results Word keeps (HYPERLINK display text, locked fields) are exempt.
"""

from pathlib import Path

import defusedxml.minidom
import pytest
from conftest import NS, count_dom_walks, find_ref, parse_paragraph, replace_document_xml

from docx_editor import BatchOperationError, Document, DocxEditError, EditOperation, FieldResultError
from docx_editor.xml_editor import build_text_map, field_result_index, get_text_node_data

AUTHOR = "Test Author"


def _r(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def _fld(kind: str, lock: bool = False) -> str:
    attr = ' w:fldLock="true"' if lock else ""
    return f'<w:r><w:fldChar w:fldCharType="{kind}"{attr}/></w:r>'


def _instr(code: str) -> str:
    return f'<w:r><w:instrText xml:space="preserve"> {code} </w:instrText></w:r>'


def _complex(code: str, result: str, lock: bool = False) -> str:
    return _fld("begin", lock) + _instr(code) + _fld("separate") + result + _fld("end")


def _p(*parts: str) -> str:
    return "<w:p>" + "".join(parts) + "</w:p>"


# A two-entry TOC spanning three paragraphs, each entry with a nested PAGEREF.
TOC = (
    _p(_r("Contents")),
    _p(_fld("begin"), _instr('TOC \\o "1-3" \\h'), _fld("separate"), _r("Introduction"), _r("\t"),
       _complex("PAGEREF _Toc1 \\h", _r("2"))),
    _p(_r("Scope"), _complex("PAGEREF _Toc2 \\h", _r("5")), _fld("end")),
    _p(_r("After the table.")),
)  # fmt: skip
REF = _p(_r("See section "), _complex("REF _Ref1 \\h", _r("Overview")), _r(" for details."))
PAGE = _p(_r("Page "), '<w:fldSimple w:instr=" PAGE "><w:r><w:t>3</w:t></w:r></w:fldSimple>', _r(" of the report"))
LOCKED = _p(_r("Dated "), _complex("DATE", _r("January 2020"), lock=True), _r(" here."))
LINK = _p(_r("Visit "), _complex('HYPERLINK "https://example.com"', _r("our site")), _r(" today."))
BODY = "".join(TOC) + REF + PAGE + LOCKED + LINK


def _open(simple_docx: Path, tmp_path: Path, body_xml: str = BODY) -> Document:
    path = tmp_path / "fields.docx"
    replace_document_xml(
        simple_docx, path, f'<?xml version="1.0"?><w:document {NS}><w:body>{body_xml}</w:body></w:document>'
    )
    return Document.open(path, author=AUTHOR)


def _xml(doc: Document) -> str:
    return doc._document_editor.dom.toxml()


@pytest.fixture
def doc(simple_docx, tmp_path):
    with _open(simple_docx, tmp_path) as d:
        yield d


# (anchor paragraph text, target, expected field_type)
TARGETS = [
    ("Introduction", "Introduction", "TOC"),
    ("Scope", "5", "TOC"),  # nested PAGEREF reports the outer TOC
    ("See section", "Overview", "REF"),
    ("Page ", "3", "PAGE"),
    ("See section", "section Overview", "REF"),  # spanning the field edge refuses too
]


class TestRefusal:
    @pytest.mark.parametrize(("para", "target", "field_type"), TARGETS)
    @pytest.mark.parametrize("verb", ["replace", "delete", "insert_after", "insert_before"])
    def test_scoped_edit_inside_result_refuses(self, doc, verb, para, target, field_type):
        ref = find_ref(doc, para)
        before = _xml(doc)
        call = {
            "replace": lambda: doc.replace(target, "X", paragraph=ref),
            "delete": lambda: doc.delete(target, paragraph=ref),
            "insert_after": lambda: doc.insert_after(target, "X", paragraph=ref),
            "insert_before": lambda: doc.insert_before(target, "X", paragraph=ref),
        }[verb]
        with pytest.raises(FieldResultError) as exc:
            call()
        e = exc.value
        assert isinstance(e, ValueError) and isinstance(e, DocxEditError)
        assert e.field_type == field_type
        assert e.field_code.split()[0] == field_type
        assert e.paragraph_ref == ref
        assert e.text == target
        assert e.next is None
        assert str(e).startswith(f"{verb}(): ")
        assert field_type in str(e) and "Narrow the target" in str(e)
        assert _xml(doc) == before

    @pytest.mark.parametrize(("para", "target", "field_type"), TARGETS)
    def test_document_wide_edit_refuses(self, doc, para, target, field_type):
        before = _xml(doc)
        with pytest.raises(FieldResultError) as exc:
            doc._revision_manager.replace_text(target, "X")
        assert exc.value.field_type == field_type
        assert exc.value.paragraph_ref == find_ref(doc, para)  # the located paragraph
        assert _xml(doc) == before

    @pytest.mark.parametrize(
        "call",
        [
            lambda d, ref: d.insert_after("section ", "X\nY", paragraph=ref),
            lambda d, ref: d.replace("section ", "sec\ntion ", paragraph=ref),
        ],
        ids=["insert_after", "replace"],
    )
    def test_split_right_before_a_field_refuses(self, doc, call):
        # The anchor is outside, but segment 0 and the split land before the
        # field's first result character, i.e. inside the field.
        ref = find_ref(doc, "See section")
        before = _xml(doc)
        with pytest.raises(FieldResultError) as exc:
            call(doc, ref)
        assert exc.value.field_type == "REF"
        assert str(exc.value).startswith(("insert_after(): ", "replace(): "))
        assert _xml(doc) == before

    @pytest.mark.parametrize(
        "op",
        [
            lambda ref: EditOperation.insert_after("section ", "X\nY", paragraph=ref),
            lambda ref: EditOperation.replace("section ", "sec\ntion ", paragraph=ref),
        ],
        ids=["insert_after", "replace"],
    )
    def test_dry_run_reports_a_split_before_a_field(self, doc, op):
        ref = find_ref(doc, "See section")
        [result] = doc.batch_edit([op(ref)], dry_run=True)
        assert not result.valid and "REF field" in result.error
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit([op(ref)])
        assert isinstance(exc.value.original, FieldResultError)

    @pytest.mark.parametrize(
        "call",
        [
            lambda d, ref: d.insert_after("Contents", "\nNew", paragraph=ref),
            lambda d, ref: d.replace("Contents", "Contents\nNew", paragraph=ref),
            lambda d, ref: d.rewrite_paragraph(ref, "Contents\nNew"),
        ],
        ids=["insert_after", "replace", "rewrite"],
    )
    @pytest.mark.parametrize(
        "after_p1",
        [
            _p(_r("Intro")) + _p(_r("Scope"), _fld("end")),
            _p() + _p(_r("Intro"), _fld("end")),  # empty next paragraph
            _p(_fld("end"), _r("After")),  # next paragraph opens with the field's end
            "<w:tbl><w:tr><w:tc>" + _p(_r("Cell")) + "</w:tc></w:tr></w:tbl>" + _p(_fld("end")),  # table in the result
        ],
        ids=["entries", "empty_next", "next_starts_with_end", "table_next"],
    )
    def test_split_at_the_end_of_a_paragraph_opening_a_field_refuses(self, simple_docx, tmp_path, call, after_p1):
        # The TOC opens at the end of P1, so a new paragraph after P1 would sit
        # inside its result.
        body = _p(_r("Contents"), _fld("begin"), _instr("TOC"), _fld("separate")) + after_p1
        with _open(simple_docx, tmp_path, body) as d:
            ref = find_ref(d, "Contents")
            before = _xml(d)
            with pytest.raises(FieldResultError):
                call(d, ref)
            assert _xml(d) == before
            [result] = d.batch_edit([EditOperation.insert_after("Contents", "\nNew", paragraph=ref)], dry_run=True)
            assert not result.valid

    @pytest.mark.parametrize(
        "call",
        [
            lambda d, ref: d.insert_after("Contents", "\nNew", paragraph=ref),
            lambda d, ref: d.rewrite_paragraph(ref, "Contents\nNew"),
        ],
        ids=["insert_after", "rewrite"],
    )
    def test_split_after_a_heading_before_a_field_applies(self, simple_docx, tmp_path, call):
        # The TOC begins in the next paragraph: the new paragraph lands before
        # its begin, outside the field.
        body = _p(_r("Contents")) + _p(_complex("REF _Ref1 \\h", _r("Overview")), _r(" here."))
        with _open(simple_docx, tmp_path, body) as d:
            call(d, find_ref(d, "Contents"))
            assert any(t.endswith("New") for t in d.list_paragraphs(limit=None))

    def test_batch_rewrite_split_check_survives_an_earlier_rewrite(self, simple_docx, tmp_path):
        # P2 is rewritten first (descending order) and re-creates its "Intro"
        # node; P1's end-of-paragraph check must not depend on that node.
        run = '<w:r><w:t>Intro</w:t><w:fldChar w:fldCharType="end"/><w:t xml:space="preserve"> tail</w:t></w:r>'
        body = _p(_r("Contents"), _fld("begin"), _instr("TOC"), _fld("separate")) + _p(run)
        with _open(simple_docx, tmp_path, body) as d:
            rewrites = [(find_ref(d, "Contents"), "Contents\nNew"), (find_ref(d, "Intro"), "Intro TAIL")]
            before = _xml(d)
            with pytest.raises(BatchOperationError) as exc:
                d.batch_rewrite(rewrites)
            assert isinstance(exc.value.original, FieldResultError)
            assert _xml(d) == before

    def test_split_replace_after_a_closing_field_refuses(self, doc):
        # Replacing the paragraph's tail after the field: segment 0 would be
        # appended after the last character before the match — the REF result.
        ref = find_ref(doc, "See section")
        before = _xml(doc)
        with pytest.raises(FieldResultError) as exc:
            doc.replace(" for details.", ", below.\nNew paragraph", paragraph=ref)
        assert exc.value.paragraph_ref == ref
        assert _xml(doc) == before
        op = EditOperation.replace(" for details.", ", below.\nNew paragraph", paragraph=ref)
        [result] = doc.batch_edit([op], dry_run=True)
        assert not result.valid

    def test_new_paragraph_after_a_trailing_field_applies(self, simple_docx, tmp_path):
        body = _p(_r("As shown in Figure "), _complex("SEQ Figure", _r("3")))
        with _open(simple_docx, tmp_path, body) as d:
            d.rewrite_paragraph(find_ref(d, "As shown"), "As shown in Figure 3\nNew paragraph")
            assert d.list_paragraphs(limit=None)[-1].endswith("New paragraph")

    def test_rewrite_into_an_empty_paragraph_inside_a_field_refuses(self, simple_docx, tmp_path):
        body = _p(_r("Contents"), _fld("begin"), _instr("TOC"), _fld("separate")) + _p() + _p(_r("Intro"), _fld("end"))
        with _open(simple_docx, tmp_path, body) as d:
            empty = d.list_paragraphs(limit=None)[1].split("|")[0]
            before = _xml(d)
            with pytest.raises(FieldResultError):
                d.rewrite_paragraph(empty, "Injected")
            assert _xml(d) == before

    def test_found_then_scoped_edit_refuses(self, doc):
        # The locate-document-wide-then-scoped route: the guard lives in
        # _locate_in_paragraph, which every scoped edit passes through.
        hit = doc.find_text("Overview")
        assert hit is not None
        before = _xml(doc)
        with pytest.raises(FieldResultError):
            doc.replace("Overview", "Summary", paragraph=hit.paragraph_ref)
        assert _xml(doc) == before

    def test_split_paragraph_inside_result_refuses(self, doc):
        ref = find_ref(doc, "See section")
        before = _xml(doc)
        with pytest.raises(FieldResultError):
            doc.split_paragraph(ref, before="Overview")
        assert _xml(doc) == before

    def test_field_code_is_whitespace_normalized(self, doc):
        with pytest.raises(FieldResultError) as exc:
            doc.delete("Overview", paragraph=find_ref(doc, "See section"))
        assert exc.value.field_code == "REF _Ref1 \\h"


class TestOutsideTheFieldStillEdits:
    def test_text_beside_a_field(self, doc):
        doc.replace("details", "context", paragraph=find_ref(doc, "See section"))
        doc.insert_before("of the report", "end ", paragraph=find_ref(doc, "Page "))
        doc.insert_after("Visit", " now", paragraph=find_ref(doc, "Visit"))
        texts = doc.list_paragraphs(limit=None)
        assert any("for context." in t for t in texts)
        assert any("Page 3 end of the report" in t for t in texts)

    def test_text_after_the_toc_end(self, doc):
        doc.replace("After", "Below", paragraph=find_ref(doc, "After the table"))
        assert any("Below the table." in t for t in doc.list_paragraphs(limit=None))

    def test_text_before_the_toc(self, doc):
        doc.replace("Contents", "Table of contents", paragraph=find_ref(doc, "Contents"))

    def test_locked_field_and_hyperlink_are_editable(self, doc):
        doc.replace("January 2020", "March 2021", paragraph=find_ref(doc, "Dated"))
        doc.replace("our site", "the site", paragraph=find_ref(doc, "Visit"))
        texts = doc.list_paragraphs(limit=None)
        assert any("March 2021" in t for t in texts) and any("the site" in t for t in texts)

    def test_unterminated_field_does_not_lock_the_rest(self, simple_docx, tmp_path):
        body = (
            _p(_fld("begin"), _instr("TOC"), _fld("separate"), _r("Entry one")) + _p(_r("Plain text.")) + _p(_r("End."))
        )
        with _open(simple_docx, tmp_path, body) as d:
            d.replace("Entry", "Item", paragraph=find_ref(d, "Entry"))
            d.replace("Plain", "Simple", paragraph=find_ref(d, "Plain"))
            d.insert_after("text.", "\nNew", paragraph=find_ref(d, "text."))  # a split at a paragraph end

    def test_reads_still_find_field_results(self, doc):
        assert doc.find_text("Overview") is not None
        assert len(doc.find_all("Introduction")) == 1


class TestBatch:
    def test_batch_edit_refuses_atomically(self, doc):
        ok = EditOperation.replace("details", "context", paragraph=find_ref(doc, "See section"))
        bad = EditOperation.replace("3", "4", paragraph=find_ref(doc, "Page "))
        before = _xml(doc)
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_edit([ok, bad])
        assert exc.value.operation_index == 1
        assert isinstance(exc.value.original, FieldResultError)
        assert _xml(doc) == before

    def test_batch_rebuilds_the_index_after_a_run_split(self, simple_docx, tmp_path):
        # A fldChar sharing a run with text: editing " tail" re-creates the
        # run's other w:t nodes, including the result "Res".
        run = '<w:r><w:t>Res</w:t><w:fldChar w:fldCharType="end"/><w:t xml:space="preserve"> tail</w:t></w:r>'
        body = _p(_fld("begin"), _instr("REF _Ref1 \\h"), _fld("separate"), run)
        with _open(simple_docx, tmp_path, body) as d:
            ref = find_ref(d, "Res")
            ops = [
                EditOperation.replace("tail", "TAIL", paragraph=ref),
                EditOperation.replace("Res", "XX", paragraph=ref),
            ]
            before = _xml(d)
            with pytest.raises(BatchOperationError) as exc:
                d.batch_edit(ops)
            assert exc.value.operation_index == 1
            assert isinstance(exc.value.original, FieldResultError)
            assert _xml(d) == before

    def test_dry_run_reports_refusal(self, doc):
        ok = EditOperation.replace("details", "context", paragraph=find_ref(doc, "See section"))
        bad = EditOperation.insert_after("Overview", "!", paragraph=find_ref(doc, "See section"))
        results = doc.batch_edit([ok, bad], dry_run=True)
        assert results[0].valid
        assert not results[1].valid and "REF field" in results[1].error

    def test_batch_builds_the_field_index_once(self, doc, monkeypatch):
        import docx_editor.track_changes.batch as batch_mod

        calls = []
        real = batch_mod.field_result_index
        monkeypatch.setattr(batch_mod, "field_result_index", lambda dom: calls.append(1) or real(dom))
        # Rebuilt only before a second op on a paragraph holding a fldChar:
        # PAGE is a fldSimple, and one op per fldChar paragraph needs no rebuild.
        ops = [EditOperation.replace("table", "list", paragraph=find_ref(doc, "After the table"))]
        ops += [EditOperation.insert_before("Page", f"{i} ", paragraph=find_ref(doc, "Page ")) for i in range(3)]
        ops += [EditOperation.replace("details", "context", paragraph=find_ref(doc, "See section"))]
        ops += [EditOperation.replace("Visit", "See", paragraph=find_ref(doc, "Visit"))]
        doc.batch_edit(ops)
        assert calls == [1]
        calls.clear()
        ref = find_ref(doc, "See section")
        doc.batch_edit([
            EditOperation.replace("See", "Read", paragraph=ref),
            EditOperation.replace("context", "it", paragraph=ref),
        ])
        assert calls == [1, 1]

    def test_field_index_adds_no_document_walk(self, doc, monkeypatch):
        walks = count_dom_walks(monkeypatch)
        field_result_index(doc._document_editor.dom)
        assert walks == []


class TestRewrite:
    def test_rewrite_outside_the_field_works(self, doc):
        ref = find_ref(doc, "See section")
        doc.rewrite_paragraph(ref, "See section Overview for context.")
        assert any("for context." in t for t in doc.list_paragraphs(limit=None))

    @pytest.mark.parametrize(
        "new_text",
        [
            "See section Summary for context.",  # one hunk in the field, one outside
            "See section Overview! for details.",  # a changed word touching the field (word-level hunk)
        ],
    )
    def test_rewrite_into_the_field_refuses_untouched(self, doc, new_text):
        ref = find_ref(doc, "See section")
        before = _xml(doc)
        with pytest.raises(FieldResultError) as exc:
            doc.rewrite_paragraph(ref, new_text)
        assert exc.value.field_type == "REF"
        assert str(exc.value).startswith("rewrite_paragraph(): ")
        assert _xml(doc) == before

    def test_insert_hunk_at_a_leading_field_refuses(self, simple_docx, tmp_path):
        # insert at 0 attaches before the result run, inside the field.
        body = _p(_complex("REF _Ref1 \\h", _r("Overview")), _r(" for details."))
        with _open(simple_docx, tmp_path, body) as d:
            before = _xml(d)
            with pytest.raises(FieldResultError):
                d.rewrite_paragraph(find_ref(d, "Overview"), "The Overview for details.")
            assert _xml(d) == before

    def test_append_after_a_trailing_field_refuses(self, simple_docx, tmp_path):
        # The insert attaches after the last result run, before the field's end.
        body = _p(_r("Page "), _complex("PAGE", _r("3")))
        with _open(simple_docx, tmp_path, body) as d:
            before = _xml(d)
            with pytest.raises(FieldResultError):
                d.rewrite_paragraph(find_ref(d, "Page"), "Page 3 of 10")
            assert _xml(d) == before

    def test_split_before_a_field_refuses(self, doc):
        ref = find_ref(doc, "See section")
        before = _xml(doc)
        with pytest.raises(FieldResultError):
            doc.rewrite_paragraph(ref, "See part\nOverview for details.")
        assert _xml(doc) == before

    def test_batch_rewrite_refuses(self, doc):
        rewrites = [
            (find_ref(doc, "After the table"), "After the tables."),
            (find_ref(doc, "Page "), "Page 4 of the report"),
        ]
        before = _xml(doc)
        with pytest.raises(BatchOperationError) as exc:
            doc.batch_rewrite(rewrites)
        assert isinstance(exc.value.original, FieldResultError)
        assert _xml(doc) == before


def _index_texts(body_xml: str) -> dict[str, str]:
    """``field_result_index`` over ``body_xml``, keyed by each indexed node's text."""
    dom = defusedxml.minidom.parseString(f"<w:document {NS}><w:body>{body_xml}</w:body></w:document>")
    out: dict[str, str] = {}
    for node, instr in field_result_index(dom).items():
        if getattr(node, "tagName", "") != "w:p":  # paragraph-end keys: see test_paragraph_end_inside_a_field_is_keyed
            out[get_text_node_data(node) or "\t"] = instr
    return out


class TestFieldResultIndex:
    def test_nested_and_cross_paragraph(self):
        got = _index_texts("".join(TOC))
        assert set(got) == {"Introduction", "\t", "2", "Scope", "5"}
        assert all(v.startswith("TOC") for v in got.values())

    def test_simple_field(self):
        assert _index_texts(PAGE) == {"3": "PAGE"}

    def test_locked_and_hyperlink_exempt(self):
        assert _index_texts(LOCKED + LINK) == {}

    def test_stray_end_is_ignored(self):
        assert _index_texts(_p(_fld("end"), _r("free")) + REF) == {"Overview": "REF _Ref1 \\h"}

    def test_field_in_instruction_reports_inner(self):
        inner = _complex("MERGEFIELD Name", _r("Alice"))
        body = _p(_fld("begin"), _instr("IF"), inner, _instr('= "Alice" "yes" "no"'), _fld("separate"),
                  _r("yes"), _fld("end"))  # fmt: skip
        got = _index_texts(body)
        assert got["Alice"] == "MERGEFIELD Name"
        assert got["yes"].startswith("IF")

    def test_split_instruction_is_joined_without_a_separator(self):
        # Word splits instrText anywhere, even mid-word.
        split = "".join(f"<w:r><w:instrText>{t}</w:instrText></w:r>" for t in ("PAGE", "REF _To", "c1 \\h"))
        body = _p(_fld("begin"), split, _fld("separate"), _r("7"), _fld("end"))
        assert _index_texts(body) == {"7": "PAGEREF _Toc1 \\h"}

    def test_split_hyperlink_instruction_stays_exempt(self):
        split = '<w:r><w:instrText>HYPER</w:instrText></w:r><w:r><w:instrText>LINK "x"</w:instrText></w:r>'
        assert _index_texts(_p(_fld("begin"), split, _fld("separate"), _r("site"), _fld("end"))) == {}

    def test_field_nested_in_hyperlink_result_is_indexed(self):
        simple = '<w:fldSimple w:instr="PAGE"><w:r><w:t>4</w:t></w:r></w:fldSimple>'
        body = _p(_fld("begin"), _instr('HYPERLINK "x"'), _fld("separate"), _r("see "),
                  _complex("PAGEREF _Toc1 \\h", _r("2")), simple, _fld("end"))  # fmt: skip
        assert _index_texts(body) == {"2": "PAGEREF _Toc1 \\h", "4": "PAGE"}

    def test_unterminated_outer_field_keeps_later_fields_indexed(self):
        open_field = _p(_fld("begin"), _instr("TOC"), _fld("separate"), _r("Entry"))
        assert set(_index_texts(open_field + REF + PAGE)) == {"Overview", "3"}

    def test_deleted_field_marks_are_ignored(self):
        deleted_end = '<w:del w:id="1" w:author="A"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:del>'
        inserted_end = '<w:ins w:id="2" w:author="A"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:ins>'
        body = _p(_fld("begin"), _instr("REF _Ref1"), _fld("separate"), _r("one"), deleted_end, _r(" two"),
                  inserted_end, _r(" after"))  # fmt: skip
        assert _index_texts(body) == {"one": "REF _Ref1", " two": "REF _Ref1"}

    def test_formtext_is_exempt(self):
        assert _index_texts(_p(_complex("FORMTEXT", _r("enter name")))) == {}

    def test_nested_simple_fields_report_the_outer(self):
        body = _p(
            '<w:fldSimple w:instr="TOC"><w:fldSimple w:instr="PAGE"><w:r><w:t>4</w:t></w:r></w:fldSimple></w:fldSimple>'
        )
        assert _index_texts(body) == {"4": "TOC"}

    def test_paragraph_end_inside_a_field_is_keyed(self):
        dom = defusedxml.minidom.parseString(f"<w:document {NS}><w:body>{''.join(TOC)}</w:body></w:document>")
        index = field_result_index(dom)
        paragraphs = dom.getElementsByTagName("w:p")
        # TOC opens in P2 and closes in P3: only P2's end lies inside it.
        assert [index.get(p) for p in paragraphs] == [None, 'TOC \\o "1-3" \\h', None, None]

    def test_unterminated_field_keys_no_paragraph(self):
        body = _p(_fld("begin"), _instr("TOC"), _fld("separate"), _r("Entry")) + _p(_r("Plain")) + _p(_r("End"))
        dom = defusedxml.minidom.parseString(f"<w:document {NS}><w:body>{body}</w:body></w:document>")
        assert field_result_index(dom) == {}

    def test_text_box_paragraphs_are_skipped(self):
        box = "<w:r><w:pict><w:txbxContent>" + _p(_r("Box")) + "</w:txbxContent></w:pict></w:r>"
        body = _p(box, _r("Contents"), _fld("begin"), _instr("TOC"), _fld("separate")) + _p(_r("Entry"), _fld("end"))
        dom = defusedxml.minidom.parseString(f"<w:document {NS}><w:body>{body}</w:body></w:document>")
        index = field_result_index(dom)
        outer = [p for p in dom.getElementsByTagName("w:p") if p.parentNode.tagName == "w:body"]
        assert index.get(outer[0]) == "TOC" and set(_index_texts(body)) == {"Entry"}

    def test_no_fields(self):
        assert _index_texts(_p(_r("plain"))) == {}

    def test_text_map_unchanged(self):
        p = parse_paragraph(REF)
        assert build_text_map(p).text == "See section Overview for details."


# The REF fixture plus the same word outside any field, after it in document
# order: occurrence 0 of "Overview" is the REF result, occurrence 1 is free.
OVERVIEW_TWICE = BODY + _p(_r("Overview of the plan."))


class TestUnscopedRoutes:
    """The public verbs without ``paragraph=`` (ROADMAP.md #95) and ``replace_all``."""

    @pytest.mark.parametrize(
        ("verb", "call"),
        [
            ("replace", lambda d: d.replace("Overview", "X")),
            ("delete", lambda d: d.delete("Overview")),
            ("insert_after", lambda d: d.insert_after("Overview", "X")),
            ("insert_before", lambda d: d.insert_before("Overview", "X")),
        ],
        ids=["replace", "delete", "insert_after", "insert_before"],
    )
    def test_unscoped_edit_inside_a_result_refuses(self, doc, verb, call):
        before = _xml(doc)
        with pytest.raises(FieldResultError) as exc:
            call(doc)
        assert exc.value.field_type == "REF"
        assert str(exc.value).startswith(f"{verb}(): ")
        assert exc.value.paragraph_ref == find_ref(doc, "See section")  # the located paragraph
        assert _xml(doc) == before

    @pytest.mark.parametrize(
        "call",
        [
            lambda d: d.replace("Overview", "X", occurrence=0),
            lambda d: d.delete("Overview", occurrence=0),
            lambda d: d.insert_after("Overview", "X", occurrence=0),
            lambda d: d.insert_before("Overview", "X", occurrence=0),
        ],
        ids=["replace", "delete", "insert_after", "insert_before"],
    )
    def test_unscoped_occurrence_inside_a_result_refuses(self, simple_docx, tmp_path, call):
        with _open(simple_docx, tmp_path, OVERVIEW_TWICE) as d:
            before = _xml(d)
            with pytest.raises(FieldResultError):
                call(d)
            assert _xml(d) == before

    def test_unscoped_occurrence_outside_the_field_applies(self, simple_docx, tmp_path):
        with _open(simple_docx, tmp_path, OVERVIEW_TWICE) as d:
            d.replace("Overview", "Summary", occurrence=1)
            assert "Summary of the plan." in d.list_paragraphs(limit=None)[-1]

    @pytest.mark.parametrize(
        "call",
        [
            lambda d: d.insert_after("section ", "X\nY"),
            lambda d: d.replace(" for details.", ", below.\nNew paragraph"),
        ],
        ids=["insert_after", "replace"],
    )
    def test_unscoped_split_into_a_field_refuses(self, doc, call):
        before = _xml(doc)
        with pytest.raises(FieldResultError) as exc:
            call(doc)
        assert exc.value.paragraph_ref == find_ref(doc, "See section")
        assert _xml(doc) == before

    def test_replace_all_with_one_match_in_a_field_writes_nothing(self, simple_docx, tmp_path):
        with _open(simple_docx, tmp_path, OVERVIEW_TWICE) as d:
            before = _xml(d)
            with pytest.raises(BatchOperationError) as exc:
                d.replace_all("Overview", "Summary", expected_count=2)
            assert isinstance(exc.value.original, FieldResultError)
            assert exc.value.operation_index == 0  # document order: the REF match
            assert _xml(d) == before  # not even the match outside the field
            assert len(d.find_all("Overview", expected_count=2)) == 2  # reads still count it

    @pytest.mark.parametrize(
        "body",
        [
            # "Contents" shares its run with a TOC's begin/separate; the TOC ends in P2.
            _p(
                '<w:r><w:t>Contents</w:t><w:fldChar w:fldCharType="begin"/>'
                '<w:instrText xml:space="preserve"> TOC </w:instrText><w:fldChar w:fldCharType="separate"/></w:r>'
            )
            + _p(_r("Intro"), _fld("end")),
            # An empty-result REF opens in the last character's run and closes in P1.
            _p(
                '<w:r><w:t>Contents</w:t><w:fldChar w:fldCharType="begin"/>'
                '<w:instrText xml:space="preserve"> REF _Ref1 </w:instrText>'
                '<w:fldChar w:fldCharType="separate"/></w:r>',
                _fld("end"),
            ),
        ],
        ids=["toc_continues", "empty_ref_closes"],
    )
    def test_rewrite_append_before_fldchars_sharing_the_last_run(self, simple_docx, tmp_path, body):
        # The run is split after the last character, so the new text lands
        # before the fldChars, outside the field.
        with _open(simple_docx, tmp_path, body) as d:
            d.rewrite_paragraph(find_ref(d, "Contents"), "Contents extra")
            dom = d._document_editor.dom
            index = field_result_index(dom)
            inserted = [t for ins in dom.getElementsByTagName("w:ins") for t in ins.getElementsByTagName("w:t")]
            assert [get_text_node_data(t) for t in inserted] == [" extra"]
            assert not any(t in index for t in inserted)

    @pytest.mark.parametrize(
        "head",
        [
            '<w:r><w:t>Contents</w:t><w:fldChar w:fldCharType="begin"/>'
            '<w:instrText xml:space="preserve"> TOC </w:instrText><w:fldChar w:fldCharType="separate"/></w:r>',
            _r("Contents") + _fld("begin") + _instr("TOC") + _fld("separate"),
        ],
        ids=["same_run", "separate_runs"],
    )
    @pytest.mark.parametrize(
        "call",
        [
            lambda d, ref: d.rewrite_paragraph(ref, "Contents extra"),
            lambda d, ref: d.batch_rewrite([(ref, "Contents extra")]),
            lambda d, ref: d.insert_after("Contents", " extra", paragraph=ref),
            lambda d, ref: d.insert_after("Contents", " extra"),
        ],
        ids=["rewrite", "batch_rewrite", "insert_after", "unscoped_insert_after"],
    )
    def test_append_inside_a_foreign_insertion_opening_a_field(self, simple_docx, tmp_path, head, call):
        # Another author inserted "Contents" and the TOC's opening marks; our
        # text must land before those marks, not after the whole insertion.
        foreign = f'<w:ins w:id="5" w:author="Other" w:date="2024-01-01T00:00:00Z">{head}</w:ins>'
        with _open(simple_docx, tmp_path, _p(foreign) + _p(_r("Intro"), _fld("end"))) as d:
            call(d, find_ref(d, "Contents"))
            self._assert_own_text_outside_fields(d, [" extra"])

    @pytest.mark.parametrize(
        "call",
        [
            lambda d: d.insert_before("tail", "X ", paragraph=find_ref(d, "tail")),
            lambda d: d.insert_before("tail", "X "),
        ],
        ids=["scoped", "unscoped"],
    )
    def test_insert_at_a_foreign_insertion_closing_a_field(self, simple_docx, tmp_path, call):
        # The foreign insertion starts with the PAGEREF's end mark: our text
        # must land after it, not before the whole insertion.
        foreign = f'<w:ins w:id="5" w:author="Other" w:date="2024-01-01T00:00:00Z">{_fld("end")}{_r("tail")}</w:ins>'
        body = _p(_r("See "), _fld("begin"), _instr("PAGEREF x"), _fld("separate"), _r("5"), foreign)
        with _open(simple_docx, tmp_path, body) as d:
            call(d)
            self._assert_own_text_outside_fields(d, ["X "])

    def test_rewrite_append_after_a_trailing_break_stays_after_it(self, simple_docx, tmp_path):
        # No fldChar in the last run: the append still follows the whole run,
        # as before (after the page break).
        body = _p('<w:r><w:t>Body.</w:t><w:br w:type="page"/></w:r>')
        with _open(simple_docx, tmp_path, body) as d:
            d.rewrite_paragraph(find_ref(d, "Body."), "Body. Next")
            p = d._document_editor.dom.getElementsByTagName("w:p")[0]
            tags = [c.tagName for c in p.childNodes if c.nodeType == c.ELEMENT_NODE]
            assert tags[-2:] == ["w:r", "w:ins"]
            assert p.getElementsByTagName("w:r")[0].getElementsByTagName("w:br")

    @staticmethod
    def _assert_own_text_outside_fields(d, expected):
        dom = d._document_editor.dom
        index = field_result_index(dom)
        own = [
            t
            for ins in dom.getElementsByTagName("w:ins")
            if ins.getAttribute("w:author") == AUTHOR
            for t in ins.getElementsByTagName("w:t")
        ]
        assert [get_text_node_data(t) for t in own] == expected
        assert not any(t in index for t in own)

    def test_unscoped_edit_outside_any_field_applies(self, doc):
        doc.replace("details", "context")
        assert any("for context." in t for t in doc.list_paragraphs(limit=None))
