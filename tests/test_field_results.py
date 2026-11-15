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
from docx_editor.xml_editor import build_text_map, field_result_index

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
        assert exc.value.paragraph_ref is None
        assert _xml(doc) == before

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
        body = _p(_fld("begin"), _instr("TOC"), _fld("separate"), _r("Entry one")) + _p(_r("Plain text."))
        with _open(simple_docx, tmp_path, body) as d:
            d.replace("Entry", "Item", paragraph=find_ref(d, "Entry"))
            d.replace("Plain", "Simple", paragraph=find_ref(d, "Plain"))

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
        ref = find_ref(doc, "See section")
        ops = [EditOperation.replace("details", "context", paragraph=ref)]
        ops += [EditOperation.insert_before("Page", f"{i} ", paragraph=find_ref(doc, "Page ")) for i in range(3)]
        doc.batch_edit(ops)
        assert calls == [1]

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
            "See section Overview! for details.",  # insert at the field's last character
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
        out[node.firstChild.data if node.firstChild else "\t"] = instr
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

    def test_no_fields(self):
        assert _index_texts(_p(_r("plain"))) == {}

    def test_text_map_unchanged(self):
        p = parse_paragraph(REF)
        assert build_text_map(p).text == "See section Overview for details."
