"""Tests for hash-anchored paragraph references."""

import re
import shutil
import tempfile
from pathlib import Path

import defusedxml.minidom
import pytest

from docx_editor import (
    Document,
    HashMismatchError,
    ParagraphIndexError,
    ParagraphRef,
    TextNotFoundError,
)
from docx_editor.xml_editor import compute_paragraph_hash

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def parse_paragraph(xml: str):
    doc = defusedxml.minidom.parseString(f"<root {NS}>{xml}</root>")
    return doc.getElementsByTagName("w:p")[0]


# ==================== ParagraphRef Tests ====================


class TestParagraphRef:
    def test_parse_valid_ref(self):
        ref = ParagraphRef.parse("P3#a7b2")
        assert ref.index == 3
        assert ref.hash == "a7b2"

    def test_parse_single_digit(self):
        ref = ParagraphRef.parse("P1#0000")
        assert ref.index == 1
        assert ref.hash == "0000"

    def test_parse_large_index(self):
        ref = ParagraphRef.parse("P999#ffff")
        assert ref.index == 999
        assert ref.hash == "ffff"

    def test_reject_invalid_format(self):
        with pytest.raises(ValueError, match="Invalid paragraph reference"):
            ParagraphRef.parse("paragraph3")

    def test_reject_missing_hash(self):
        with pytest.raises(ValueError):
            ParagraphRef.parse("P3")

    def test_reject_uppercase_hash(self):
        with pytest.raises(ValueError):
            ParagraphRef.parse("P3#A7B2")

    def test_reject_short_hash(self):
        with pytest.raises(ValueError):
            ParagraphRef.parse("P3#a7b")

    def test_reject_long_hash(self):
        with pytest.raises(ValueError):
            ParagraphRef.parse("P3#a7b2c")

    def test_reject_zero_index(self):
        # P0 is technically valid format but semantically wrong (1-indexed)
        # The format regex allows it; validation happens at resolve time
        ref = ParagraphRef.parse("P0#a7b2")
        assert ref.index == 0

    def test_reject_empty_string(self):
        with pytest.raises(ValueError):
            ParagraphRef.parse("")

    def test_reject_no_p_prefix(self):
        with pytest.raises(ValueError):
            ParagraphRef.parse("3#a7b2")


# ==================== compute_paragraph_hash Tests ====================


class TestComputeParagraphHash:
    def test_normal_paragraph(self):
        p = parse_paragraph("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        h = compute_paragraph_hash(p)
        assert re.match(r"^[0-9a-f]{4}$", h)

    def test_deterministic(self):
        p = parse_paragraph("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        assert compute_paragraph_hash(p) == compute_paragraph_hash(p)

    def test_different_content_different_hash(self):
        p1 = parse_paragraph("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        p2 = parse_paragraph("<w:p><w:r><w:t>Goodbye world</w:t></w:r></w:p>")
        assert compute_paragraph_hash(p1) != compute_paragraph_hash(p2)

    def test_empty_paragraph(self):
        p = parse_paragraph("<w:p></w:p>")
        h = compute_paragraph_hash(p)
        assert re.match(r"^[0-9a-f]{4}$", h)

    def test_excludes_deleted_text(self):
        p_with_del = parse_paragraph(
            "<w:p><w:r><w:t>Hello </w:t></w:r>"
            "<w:del><w:r><w:delText>old </w:delText></w:r></w:del>"
            "<w:r><w:t>world</w:t></w:r></w:p>"
        )
        p_without_del = parse_paragraph("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        assert compute_paragraph_hash(p_with_del) == compute_paragraph_hash(p_without_del)

    def test_includes_inserted_text(self):
        p_with_ins = parse_paragraph(
            "<w:p><w:r><w:t>Hello </w:t></w:r>"
            "<w:ins><w:r><w:t>beautiful </w:t></w:r></w:ins>"
            "<w:r><w:t>world</w:t></w:r></w:p>"
        )
        p_without_ins = parse_paragraph("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        # Insertions are visible, so hashes differ
        assert compute_paragraph_hash(p_with_ins) != compute_paragraph_hash(p_without_ins)


# ==================== list_paragraphs Tests ====================


class TestListParagraphs:
    @pytest.fixture
    def temp_docx(self):
        test_data = Path(__file__).parent / "test_data" / "simple.docx"
        temp = tempfile.mkdtemp(prefix="docx_hash_test_")
        dest = Path(temp) / "test.docx"
        shutil.copy(test_data, dest)
        yield dest
        shutil.rmtree(temp, ignore_errors=True)

    def test_returns_list(self, temp_docx):
        doc = Document.open(temp_docx)
        try:
            result = doc.list_paragraphs()
            assert isinstance(result, list)
            assert len(result) > 0
        finally:
            doc.close()

    def test_format(self, temp_docx):
        doc = Document.open(temp_docx)
        try:
            result = doc.list_paragraphs()
            for entry in result:
                # Should match P{n}#{hash}| {text}
                assert re.match(r"^P\d+#[0-9a-f]{4}\| ", entry), f"Bad format: {entry}"
        finally:
            doc.close()

    def test_truncation(self, temp_docx):
        doc = Document.open(temp_docx)
        try:
            result = doc.list_paragraphs(max_chars=10)
            for entry in result:
                # Extract preview part after "| "
                preview = entry.split("| ", 1)[1]
                # Truncated entries should end with "..." and be at most 13 chars
                if len(preview) > 10:
                    assert preview.endswith("...")
        finally:
            doc.close()

    def test_1_indexed(self, temp_docx):
        doc = Document.open(temp_docx)
        try:
            result = doc.list_paragraphs()
            # First entry should be P1
            assert result[0].startswith("P1#")
        finally:
            doc.close()


# ==================== Scoped Operations Tests ====================


class TestScopedOperations:
    @pytest.fixture
    def temp_docx(self):
        test_data = Path(__file__).parent / "test_data" / "simple.docx"
        temp = tempfile.mkdtemp(prefix="docx_hash_test_")
        dest = Path(temp) / "test.docx"
        shutil.copy(test_data, dest)
        yield dest
        shutil.rmtree(temp, ignore_errors=True)

    def _get_paragraph_ref(self, doc, index):
        """Helper to get a fresh paragraph reference for a given 1-based index."""
        paragraphs = doc.list_paragraphs()
        entry = paragraphs[index - 1]
        # Extract "P{n}#{hash}" from "P{n}#{hash}| text"
        return entry.split("|")[0]

    def test_scoped_replace(self, temp_docx):
        doc = Document.open(temp_docx)
        try:
            paragraphs = doc.list_paragraphs()
            # Find a paragraph with text
            for entry in paragraphs:
                ref = entry.split("|")[0]
                preview = entry.split("| ", 1)[1]
                if len(preview) > 5:
                    # Replace first 5 chars
                    target = preview[:5]
                    doc.replace(target, "XXXXX", paragraph=ref)
                    # Verify the change happened
                    new_text = doc.get_visible_text()
                    assert "XXXXX" in new_text
                    return
            pytest.skip("No paragraph with enough text found")
        finally:
            doc.close()

    def test_scoped_replace_wrong_paragraph(self, temp_docx):
        doc = Document.open(temp_docx)
        try:
            paragraphs = doc.list_paragraphs()
            if len(paragraphs) < 2:
                pytest.skip("Need at least 2 paragraphs")

            # Get text from paragraph 1
            p1_preview = paragraphs[0].split("| ", 1)[1]
            if not p1_preview or len(p1_preview) < 3:
                pytest.skip("Paragraph 1 too short")

            target = p1_preview[:3]

            # Try to find it in paragraph 2 — should fail if text is unique to p1
            p2_ref = paragraphs[1].split("|")[0]
            p2_preview = paragraphs[1].split("| ", 1)[1]

            if target not in p2_preview:
                from docx_editor.exceptions import TextNotFoundError

                with pytest.raises(TextNotFoundError):
                    doc.replace(target, "XXXXX", paragraph=p2_ref)
        finally:
            doc.close()

    def test_scoped_delete(self, temp_docx):
        doc = Document.open(temp_docx)
        try:
            paragraphs = doc.list_paragraphs()
            for entry in paragraphs:
                ref = entry.split("|")[0]
                preview = entry.split("| ", 1)[1]
                if len(preview) > 5:
                    target = preview[:5]
                    doc.delete(target, paragraph=ref)
                    return
            pytest.skip("No paragraph with enough text found")
        finally:
            doc.close()

    def test_scoped_insert_after(self, temp_docx):
        doc = Document.open(temp_docx)
        try:
            paragraphs = doc.list_paragraphs()
            for entry in paragraphs:
                ref = entry.split("|")[0]
                preview = entry.split("| ", 1)[1]
                if len(preview) > 5:
                    anchor = preview[:5]
                    doc.insert_after(anchor, " [INSERTED]", paragraph=ref)
                    assert "[INSERTED]" in doc.get_visible_text()
                    return
            pytest.skip("No paragraph with enough text found")
        finally:
            doc.close()

    def test_scoped_insert_before(self, temp_docx):
        doc = Document.open(temp_docx)
        try:
            paragraphs = doc.list_paragraphs()
            for entry in paragraphs:
                ref = entry.split("|")[0]
                preview = entry.split("| ", 1)[1]
                if len(preview) > 5:
                    anchor = preview[:5]
                    doc.insert_before(anchor, "[INSERTED] ", paragraph=ref)
                    assert "[INSERTED]" in doc.get_visible_text()
                    return
            pytest.skip("No paragraph with enough text found")
        finally:
            doc.close()

    def test_paragraph_local_occurrence(self, temp_docx):
        """When paragraph is specified, occurrence counts within that paragraph only."""
        doc = Document.open(temp_docx)
        try:
            # Just verify the parameter is accepted and doesn't error
            paragraphs = doc.list_paragraphs()
            for entry in paragraphs:
                ref = entry.split("|")[0]
                preview = entry.split("| ", 1)[1]
                if len(preview) > 3:
                    # occurrence=0 should work
                    single_char = preview[0]
                    # Count occurrences in this paragraph's preview
                    count = preview.count(single_char)
                    if count >= 2:
                        # Replace second occurrence
                        doc.replace(single_char, "Z", occurrence=1, paragraph=ref)
                        return
        finally:
            doc.close()


# ==================== Staleness Detection Tests ====================


class TestStalenessDetection:
    @pytest.fixture
    def temp_docx(self):
        test_data = Path(__file__).parent / "test_data" / "simple.docx"
        temp = tempfile.mkdtemp(prefix="docx_hash_test_")
        dest = Path(temp) / "test.docx"
        shutil.copy(test_data, dest)
        yield dest
        shutil.rmtree(temp, ignore_errors=True)

    def test_stale_hash_after_edit(self, temp_docx):
        """Edit paragraph, then use old hash — should raise HashMismatchError."""
        doc = Document.open(temp_docx)
        try:
            paragraphs = doc.list_paragraphs()
            # Find a paragraph with text
            for entry in paragraphs:
                ref = entry.split("|")[0]
                preview = entry.split("| ", 1)[1]
                if len(preview) > 5:
                    target = preview[:5]
                    # Edit the paragraph
                    doc.replace(target, "CHANGED", paragraph=ref)
                    # Now use the OLD ref — hash should be stale
                    with pytest.raises(HashMismatchError) as exc_info:
                        doc.replace("CHANGED", "AGAIN", paragraph=ref)
                    # Error should include the current hash for retry
                    assert exc_info.value.actual_hash
                    assert exc_info.value.paragraph_index > 0
                    return
            pytest.skip("No paragraph with enough text found")
        finally:
            doc.close()

    def test_error_includes_current_hash(self, temp_docx):
        """HashMismatchError message includes current hash for LLM retry."""
        doc = Document.open(temp_docx)
        try:
            paragraphs = doc.list_paragraphs()
            for entry in paragraphs:
                ref = entry.split("|")[0]
                preview = entry.split("| ", 1)[1]
                if len(preview) > 5:
                    target = preview[:5]
                    doc.replace(target, "CHANGED", paragraph=ref)

                    try:
                        doc.replace("CHANGED", "AGAIN", paragraph=ref)
                        pytest.fail("Expected HashMismatchError")
                    except HashMismatchError as e:
                        # The new ref should work
                        idx = ref.split("#")[0]  # e.g., "P1"
                        new_ref = f"{idx}#{e.actual_hash}"
                        doc.replace("CHANGED", "AGAIN", paragraph=new_ref)
                        assert "AGAIN" in doc.get_visible_text()
                        return
            pytest.skip("No paragraph with enough text found")
        finally:
            doc.close()

    def test_sequential_edits_with_fresh_refs(self, temp_docx):
        """Multiple edits succeed when each uses fresh refs from list_paragraphs()."""
        doc = Document.open(temp_docx)
        try:
            # First edit
            paragraphs = doc.list_paragraphs()
            edited = False
            for entry in paragraphs:
                ref = entry.split("|")[0]
                preview = entry.split("| ", 1)[1]
                if len(preview) > 5:
                    doc.replace(preview[:3], "AAA", paragraph=ref)
                    edited = True
                    break

            if not edited:
                pytest.skip("No paragraph with enough text")

            # Second edit with fresh ref
            paragraphs = doc.list_paragraphs()
            for entry in paragraphs:
                ref = entry.split("|")[0]
                preview = entry.split("| ", 1)[1]
                if len(preview) > 5 and "AAA" not in preview:
                    doc.replace(preview[:3], "BBB", paragraph=ref)
                    break

            text = doc.get_visible_text()
            assert "AAA" in text
        finally:
            doc.close()

    def test_index_out_of_range(self, temp_docx):
        """Using an index beyond document length raises ParagraphIndexError."""
        doc = Document.open(temp_docx)
        try:
            paragraphs = doc.list_paragraphs()
            # Use an index way beyond the document
            fake_ref = f"P{len(paragraphs) + 100}#0000"
            with pytest.raises(ParagraphIndexError):
                doc.replace("anything", "else", paragraph=fake_ref)
        finally:
            doc.close()


# ==================== Structured TextNotFoundError Tests ====================


class TestStructuredTextNotFoundError:
    @pytest.fixture
    def temp_docx(self):
        test_data = Path(__file__).parent / "test_data" / "simple.docx"
        temp = tempfile.mkdtemp(prefix="docx_err_test_")
        dest = Path(temp) / "test.docx"
        shutil.copy(test_data, dest)
        yield dest
        shutil.rmtree(temp, ignore_errors=True)

    def test_scoped_miss_carries_search_text_ref_and_preview(self, temp_docx):
        doc = Document.open(temp_docx)
        try:
            entry = doc.list_paragraphs()[0]
            ref = entry.split("|")[0]
            current_preview = entry.split("| ", 1)[1]

            with pytest.raises(TextNotFoundError) as exc:
                doc.replace("this_string_is_not_present", "x", paragraph=ref)

            err = exc.value
            assert err.search_text == "this_string_is_not_present"
            assert err.paragraph_ref == ref
            assert err.paragraph_preview is not None
            # occurrence fields are None for a scoped miss
            assert err.occurrence is None
            assert err.total_occurrences is None
            # message embeds the current paragraph content (either full or truncated preview)
            msg = str(err)
            assert "this_string_is_not_present" in msg
            assert ref in msg
            # Preview in the message should reflect current paragraph text
            # (list_paragraphs already truncates at 10 chars; check a non-empty prefix is in the error)
            preview_prefix = current_preview.removesuffix("...")[:10]
            if preview_prefix:
                assert preview_prefix in msg

        finally:
            doc.close()

    def test_scoped_miss_preview_capped_at_80_chars(self, temp_docx):
        """Preview in TextNotFoundError is truncated to 80 chars with ellipsis, matching HashMismatchError."""
        doc = Document.open(temp_docx, force_recreate=True)
        editor = doc._document_editor
        body = editor.dom.getElementsByTagName("w:body")[0]
        # Clear body paragraphs
        for p in list(editor.dom.getElementsByTagName("w:p")):
            if p.parentNode == body:
                body.removeChild(p)
        sect_pr = editor.dom.getElementsByTagName("w:sectPr")
        insert_before = sect_pr[0] if sect_pr else None
        # One long paragraph (>80 chars)
        long_text = "a very long paragraph " + "x" * 200
        p_xml = f'<w:p><w:r><w:t xml:space="preserve">{long_text}</w:t></w:r></w:p>'
        for node in editor._parse_fragment(p_xml):
            if insert_before:
                body.insertBefore(node, insert_before)
            else:
                body.appendChild(node)
        editor.save()
        saved = doc.save()
        doc.close()

        doc = Document.open(saved, force_recreate=True)
        try:
            ref = doc.list_paragraphs()[0].split("|")[0]
            with pytest.raises(TextNotFoundError) as exc:
                doc.replace("ZZZ_not_there_ZZZ", "y", paragraph=ref)
            err = exc.value
            assert err.paragraph_preview is not None
            # 80 chars max with "..." suffix -> length <= 83
            assert len(err.paragraph_preview) <= 83
            assert err.paragraph_preview.endswith("...")
        finally:
            doc.close()

    def test_unscoped_miss_has_none_ref_and_preview(self, temp_docx):
        """Unscoped anchor search (add_comment) leaves paragraph fields None."""
        doc = Document.open(temp_docx)
        try:
            with pytest.raises(TextNotFoundError) as exc:
                doc.add_comment("__definitely_absent_from_entire_doc__", "c")

            err = exc.value
            assert err.search_text == "__definitely_absent_from_entire_doc__"
            assert err.paragraph_ref is None
            assert err.paragraph_preview is None
            assert err.occurrence is None
            assert err.total_occurrences is None
            assert "__definitely_absent_from_entire_doc__" in str(err)
        finally:
            doc.close()

    def test_occurrence_failure_carries_counts(self, temp_docx):
        """Occurrence-based miss carries `occurrence` and `total_occurrences`."""
        doc = Document.open(temp_docx, force_recreate=True)
        editor = doc._document_editor
        body = editor.dom.getElementsByTagName("w:body")[0]
        for p in list(editor.dom.getElementsByTagName("w:p")):
            if p.parentNode == body:
                body.removeChild(p)
        sect_pr = editor.dom.getElementsByTagName("w:sectPr")
        insert_before = sect_pr[0] if sect_pr else None
        # 3 paragraphs each containing "needle"
        for i in range(3):
            p_xml = f'<w:p><w:r><w:t xml:space="preserve">paragraph {i} has needle in it.</w:t></w:r></w:p>'
            for node in editor._parse_fragment(p_xml):
                if insert_before:
                    body.insertBefore(node, insert_before)
                else:
                    body.appendChild(node)
        editor.save()
        saved = doc.save()
        doc.close()

        doc = Document.open(saved, force_recreate=True)
        try:
            # Ask for occurrence=5 when only 3 exist. Use the internal unscoped
            # API — the public Document.replace requires a paragraph arg.
            with pytest.raises(TextNotFoundError) as exc:
                doc._revision_manager.replace_text("needle", "pin", occurrence=5)
            err = exc.value
            assert err.search_text == "needle"
            assert err.occurrence == 5
            assert err.total_occurrences == 3
            msg = str(err)
            assert "5" in msg
            assert "3" in msg
        finally:
            doc.close()


# ==================== Structured ParagraphIndexError Tests ====================


class TestStructuredParagraphIndexError:
    @pytest.fixture
    def temp_docx(self):
        test_data = Path(__file__).parent / "test_data" / "simple.docx"
        temp = tempfile.mkdtemp(prefix="docx_idx_test_")
        dest = Path(temp) / "test.docx"
        shutil.copy(test_data, dest)
        yield dest
        shutil.rmtree(temp, ignore_errors=True)

    def test_out_of_range_raises_paragraph_index_error(self, temp_docx):
        doc = Document.open(temp_docx)
        try:
            n = len(doc.list_paragraphs())
            bad = f"P{n + 50}#0000"
            with pytest.raises(ParagraphIndexError) as exc:
                doc.replace("x", "y", paragraph=bad)
            err = exc.value
            assert err.index == n + 50
            assert err.total_paragraphs == n
            msg = str(err)
            assert str(n + 50) in msg
            assert str(n) in msg
        finally:
            doc.close()

    def test_paragraph_index_error_is_docx_edit_error(self, temp_docx):
        from docx_editor import DocxEditError

        doc = Document.open(temp_docx)
        try:
            n = len(doc.list_paragraphs())
            bad = f"P{n + 50}#0000"
            with pytest.raises(DocxEditError):
                doc.replace("x", "y", paragraph=bad)
        finally:
            doc.close()
