"""Tests for batch_edit() with reverse-order application."""

import shutil
import tempfile
from pathlib import Path

import pytest

from docx_editor import Document, EditOperation, HashMismatchError, TextNotFoundError


@pytest.fixture
def multi_para_doc():
    """Build a document with 10 paragraphs, each containing repeated phrases."""
    test_data = Path(__file__).parent / "test_data" / "simple.docx"
    tmp = tempfile.mkdtemp(prefix="batch_test_")
    dest = Path(tmp) / "test.docx"
    shutil.copy(test_data, dest)

    doc = Document.open(dest)
    doc.accept_all()
    doc.save()
    doc.close()

    # Inject paragraphs directly
    doc = Document.open(dest, force_recreate=True)
    editor = doc._document_editor
    body = editor.dom.getElementsByTagName("w:body")[0]

    # Remove existing paragraphs
    for p in list(editor.dom.getElementsByTagName("w:p")):
        if p.parentNode == body:
            body.removeChild(p)

    sect_pr = editor.dom.getElementsByTagName("w:sectPr")
    insert_before = sect_pr[0] if sect_pr else None

    for i in range(1, 11):
        text = f"[P{i:02d}] The committee shall review item {i}. The report includes findings from the committee."
        p_xml = f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        nodes = editor._parse_fragment(p_xml)
        for node in nodes:
            if insert_before:
                body.insertBefore(node, insert_before)
            else:
                body.appendChild(node)

    editor.save()
    save_path = doc.save()
    doc.close()

    doc = Document.open(save_path, force_recreate=True)
    yield doc, Path(tmp)
    doc.close()
    shutil.rmtree(tmp, ignore_errors=True)


class TestBatchEdit:
    def test_batch_multiple_paragraphs(self, multi_para_doc):
        """Batch of edits to different paragraphs all succeed."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        ops = [
            EditOperation(
                action="replace",
                find="item 3",
                replace_with="ITEM_THREE",
                paragraph=refs[2].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="item 7",
                replace_with="ITEM_SEVEN",
                paragraph=refs[6].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="item 9",
                replace_with="ITEM_NINE",
                paragraph=refs[8].split("|")[0],
            ),
        ]

        result = doc.batch_edit(ops)
        assert len(result) == 3

        vis = doc.get_visible_text()
        assert "ITEM_THREE" in vis
        assert "ITEM_SEVEN" in vis
        assert "ITEM_NINE" in vis

    def test_batch_rejected_on_stale_hash(self, multi_para_doc):
        """Entire batch rejected if any hash is stale."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        # Edit P5 to make its hash stale
        p5_ref = refs[4].split("|")[0]
        doc.replace("item 5", "CHANGED", paragraph=p5_ref)

        # Now try a batch using OLD refs (P5 hash is stale)
        ops = [
            EditOperation(
                action="replace",
                find="item 3",
                replace_with="EDIT_3",
                paragraph=refs[2].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="CHANGED",
                replace_with="EDIT_5",
                paragraph=p5_ref,  # STALE hash
            ),
            EditOperation(
                action="replace",
                find="item 8",
                replace_with="EDIT_8",
                paragraph=refs[7].split("|")[0],
            ),
        ]

        with pytest.raises(HashMismatchError):
            doc.batch_edit(ops)

        # Verify NO edits were applied
        vis = doc.get_visible_text()
        assert "EDIT_3" not in vis
        assert "EDIT_5" not in vis
        assert "EDIT_8" not in vis

    def test_batch_reverse_order_preserves_hashes(self, multi_para_doc):
        """Edits applied in reverse order — P10 edit doesn't invalidate P2 hash."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        # Take ONE snapshot, edit both ends
        ops = [
            EditOperation(
                action="replace",
                find="item 2",
                replace_with="SECOND",
                paragraph=refs[1].split("|")[0],
            ),
            EditOperation(
                action="replace",
                find="item 10",
                replace_with="TENTH",
                paragraph=refs[9].split("|")[0],
            ),
        ]

        result = doc.batch_edit(ops)
        assert len(result) == 2

        vis = doc.get_visible_text()
        assert "SECOND" in vis
        assert "TENTH" in vis

    def test_single_snapshot_suffices(self, multi_para_doc):
        """One list_paragraphs() call works for all 10 edits."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        ops = []
        for i in range(10):
            ops.append(
                EditOperation(
                    action="replace",
                    find=f"item {i + 1}",
                    replace_with=f"BATCH_{i + 1}",
                    paragraph=refs[i].split("|")[0],
                )
            )

        result = doc.batch_edit(ops)
        assert len(result) == 10

        vis = doc.get_visible_text()
        for i in range(10):
            assert f"BATCH_{i + 1}" in vis

    def test_duplicate_paragraph_targets(self, multi_para_doc):
        """Two edits to the same paragraph both apply."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        p3_ref = refs[2].split("|")[0]
        ops = [
            EditOperation(
                action="replace",
                find="item 3",
                replace_with="FIRST_EDIT",
                paragraph=p3_ref,
            ),
            EditOperation(
                action="replace",
                find="committee",
                replace_with="BOARD",
                paragraph=p3_ref,
                occurrence=0,
            ),
        ]

        result = doc.batch_edit(ops)
        assert len(result) == 2

        vis = doc.get_visible_text()
        assert "FIRST_EDIT" in vis
        assert "BOARD" in vis

    def test_empty_batch(self, multi_para_doc):
        """Empty batch returns empty list."""
        doc, _ = multi_para_doc
        result = doc.batch_edit([])
        assert result == []

    def test_batch_mixed_actions(self, multi_para_doc):
        """Batch with replace, delete, and insert actions."""
        doc, _ = multi_para_doc
        refs = doc.list_paragraphs()

        ops = [
            EditOperation(
                action="replace",
                find="item 2",
                replace_with="REPLACED",
                paragraph=refs[1].split("|")[0],
            ),
            EditOperation(
                action="delete",
                text="item 5",
                paragraph=refs[4].split("|")[0],
            ),
            EditOperation(
                action="insert_after",
                anchor="item 8",
                text=" [APPENDED]",
                paragraph=refs[7].split("|")[0],
            ),
        ]

        result = doc.batch_edit(ops)
        assert len(result) == 3

        vis = doc.get_visible_text()
        assert "REPLACED" in vis
        # Verify "item 5" was deleted (but [P05] marker remains in the paragraph)
        assert "[P05]" in vis
        p05_line = [line for line in vis.split("\n") if "[P05]" in line][0]
        assert "item 5" not in p05_line
        assert "[APPENDED]" in vis

    def test_batch_missing_paragraph_raises(self, multi_para_doc):
        """Batch with missing paragraph field raises ValueError."""
        doc, _ = multi_para_doc

        ops = [
            EditOperation(action="replace", find="a", replace_with="b", paragraph=""),
        ]

        with pytest.raises(ValueError, match="paragraph reference is required"):
            doc.batch_edit(ops)

    def test_batch_replace_missing_find(self, multi_para_doc):
        """Replace without find raises ValueError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="replace", find=None, replace_with="x", paragraph=ref)]
        with pytest.raises(ValueError, match="replace requires"):
            doc.batch_edit(ops)

    def test_batch_replace_text_not_found(self, multi_para_doc):
        """Replace with non-existent text raises TextNotFoundError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="replace", find="NONEXISTENT", replace_with="x", paragraph=ref)]
        with pytest.raises(TextNotFoundError):
            doc.batch_edit(ops)

    def test_batch_delete_missing_text(self, multi_para_doc):
        """Delete without text raises ValueError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="delete", text=None, paragraph=ref)]
        with pytest.raises(ValueError, match="delete requires"):
            doc.batch_edit(ops)

    def test_batch_delete_text_not_found(self, multi_para_doc):
        """Delete with non-existent text raises TextNotFoundError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="delete", text="NONEXISTENT", paragraph=ref)]
        with pytest.raises(TextNotFoundError):
            doc.batch_edit(ops)

    def test_batch_insert_missing_anchor(self, multi_para_doc):
        """Insert without anchor raises ValueError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="insert_after", anchor=None, text="x", paragraph=ref)]
        with pytest.raises(ValueError, match="insert_after requires"):
            doc.batch_edit(ops)

    def test_batch_insert_anchor_not_found(self, multi_para_doc):
        """Insert with non-existent anchor raises TextNotFoundError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="insert_after", anchor="NONEXISTENT", text="x", paragraph=ref)]
        with pytest.raises(TextNotFoundError):
            doc.batch_edit(ops)

    def test_batch_unknown_action(self, multi_para_doc):
        """Unknown action raises ValueError."""
        doc, _ = multi_para_doc
        ref = doc.list_paragraphs()[0].split("|")[0]
        ops = [EditOperation(action="unknown", paragraph=ref)]
        with pytest.raises(ValueError, match="Unknown action"):
            doc.batch_edit(ops)
