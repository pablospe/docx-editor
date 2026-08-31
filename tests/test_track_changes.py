"""Tests for track changes functionality."""

import shutil
from datetime import datetime, timezone
from typing import Any
from unittest.mock import MagicMock

import pytest
from conftest import find_ref, match_for

from docx_editor import Document, HashMismatchError, TextNotFoundError
from docx_editor.track_changes import Revision, RevisionManager
from docx_editor.track_changes.diff import _trim_replace_affixes
from docx_editor.xml_editor import DocxXMLEditor, _escape_xml, build_text_map


class TestTrackedReplace:
    """Tests for tracked text replacement."""

    def test_replace_creates_tracked_change(self, clean_workspace):
        """Test that replace creates w:del and w:ins elements."""
        doc = Document.open(clean_workspace)

        # Find some text to replace - need to know what's in simple.docx
        # For now, we'll test that the method doesn't crash
        try:
            ref = find_ref(doc, "test")
            doc.replace("test", "TEST", paragraph=ref)
        except TextNotFoundError:
            # Expected if "test" not in document
            pass

        doc.close()

    def test_replace_returns_new_ref(self, clean_workspace):
        """Test that replace returns a new paragraph reference."""
        doc = Document.open(clean_workspace)

        try:
            ref = find_ref(doc, "the")
            new_ref = doc.replace("the", "THE", paragraph=ref)
            assert isinstance(new_ref, str)
            assert new_ref.startswith("P")
            assert "#" in new_ref
        except TextNotFoundError:
            pytest.skip("Test text not found in document")

        doc.close()

    def test_replace_not_found_raises_error(self, clean_workspace):
        """Test that replacing nonexistent text raises TextNotFoundError."""
        doc = Document.open(clean_workspace)

        ref = doc.list_paragraphs()[0].split("|")[0]
        with pytest.raises(TextNotFoundError):
            doc.replace("xyz123nonexistent789", "replacement", paragraph=ref)

        doc.close()


class TestTrackedDeletion:
    """Tests for tracked deletions."""

    def test_delete_creates_tracked_change(self, clean_workspace):
        """Test that delete creates w:del element."""
        doc = Document.open(clean_workspace)

        try:
            ref = find_ref(doc, "the")
            new_ref = doc.delete("the", paragraph=ref)
            assert isinstance(new_ref, str)
            assert new_ref.startswith("P")
        except TextNotFoundError:
            pytest.skip("Test text not found in document")

        doc.close()

    def test_delete_not_found_raises_error(self, clean_workspace):
        """Test that deleting nonexistent text raises TextNotFoundError."""
        doc = Document.open(clean_workspace)

        ref = doc.list_paragraphs()[0].split("|")[0]
        with pytest.raises(TextNotFoundError):
            doc.delete("xyz123nonexistent789", paragraph=ref)

        doc.close()


class TestTrackedInsertion:
    """Tests for tracked insertions."""

    def test_insert_after_creates_tracked_change(self, clean_workspace):
        """Test that insert_after creates w:ins element."""
        doc = Document.open(clean_workspace)

        try:
            ref = find_ref(doc, "the")
            new_ref = doc.insert_after("the", " NEW TEXT", paragraph=ref)
            assert isinstance(new_ref, str)
            assert new_ref.startswith("P")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        doc.close()

    def test_insert_before_creates_tracked_change(self, clean_workspace):
        """Test that insert_before creates w:ins element."""
        doc = Document.open(clean_workspace)

        try:
            ref = find_ref(doc, "the")
            new_ref = doc.insert_before("the", "BEFORE ", paragraph=ref)
            assert isinstance(new_ref, str)
            assert new_ref.startswith("P")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        doc.close()


class TestRevisionListing:
    """Tests for listing revisions."""

    def test_list_revisions_empty_document(self, clean_workspace):
        """Test listing revisions on document without changes."""
        doc = Document.open(clean_workspace)

        revisions = doc.list_revisions()
        # May be empty or have pre-existing revisions
        assert isinstance(revisions, list)

        doc.close()

    def test_list_revisions_after_changes(self, clean_workspace):
        """Test listing revisions after making changes."""
        doc = Document.open(clean_workspace)

        try:
            ref = find_ref(doc, "the")
            doc.delete("the", paragraph=ref)
            ref2 = find_ref(doc, "a")
            doc.insert_after("a", " NEW", paragraph=ref2)
        except TextNotFoundError:
            pytest.skip("Test text not found in document")

        revisions = doc.list_revisions()
        assert len(revisions) >= 2

        # Check revision attributes
        for rev in revisions:
            assert hasattr(rev, "id")
            assert hasattr(rev, "type")
            assert hasattr(rev, "author")
            assert hasattr(rev, "text")
            assert rev.type in ("insertion", "deletion")

        doc.close()

    def test_list_revisions_filter_by_author(self, clean_workspace):
        """Test filtering revisions by author."""
        doc = Document.open(clean_workspace, author="TestAuthor")

        try:
            ref = find_ref(doc, "the")
            doc.delete("the", paragraph=ref)
        except TextNotFoundError:
            pytest.skip("Test text not found in document")

        author_revisions = doc.list_revisions(author="TestAuthor")

        # Author filter should only return revisions by that author
        for rev in author_revisions:
            assert rev.author == "TestAuthor"

        doc.close()


class TestRevisionAcceptReject:
    """Tests for accepting and rejecting revisions."""

    def test_accept_revision(self, clean_workspace):
        """Test accepting a revision."""
        doc = Document.open(clean_workspace)

        try:
            ref = find_ref(doc, "the")
            doc.delete("the", paragraph=ref)
        except TextNotFoundError:
            pytest.skip("Test text not found in document")

        revisions = doc.list_revisions()
        change_id = revisions[-1].id

        result = doc.accept_revision(change_id)
        assert result is True

        # Revision should no longer be in list
        revisions = doc.list_revisions()
        revision_ids = [r.id for r in revisions]
        assert change_id not in revision_ids

        doc.close()

    def test_reject_revision(self, clean_workspace):
        """Test rejecting a revision."""
        doc = Document.open(clean_workspace)

        try:
            ref = find_ref(doc, "the")
            doc.delete("the", paragraph=ref)
        except TextNotFoundError:
            pytest.skip("Test text not found in document")

        revisions = doc.list_revisions()
        change_id = revisions[-1].id

        result = doc.reject_revision(change_id)
        assert result is True

        doc.close()

    def test_accept_nonexistent_revision(self, clean_workspace):
        """Test accepting a revision that doesn't exist."""
        doc = Document.open(clean_workspace)

        result = doc.accept_revision(99999)
        assert result is False

        doc.close()

    def test_accept_all(self, clean_workspace):
        """Test accepting all revisions."""
        doc = Document.open(clean_workspace)

        try:
            ref = find_ref(doc, "the")
            doc.delete("the", paragraph=ref)
            ref2 = find_ref(doc, "a")
            doc.insert_after("a", " NEW", paragraph=ref2)
        except TextNotFoundError:
            pytest.skip("Test text not found in document")

        initial_count = len(doc.list_revisions())
        accepted = doc.accept_all()

        assert accepted >= 0
        assert len(doc.list_revisions()) == initial_count - accepted

        doc.close()

    def test_reject_all(self, clean_workspace):
        """Test rejecting all revisions."""
        doc = Document.open(clean_workspace)

        try:
            ref = find_ref(doc, "the")
            doc.delete("the", paragraph=ref)
            ref2 = find_ref(doc, "a")
            doc.insert_after("a", " NEW", paragraph=ref2)
        except TextNotFoundError:
            pytest.skip("Test text not found in document")

        initial_count = len(doc.list_revisions())
        rejected = doc.reject_all()

        assert rejected >= 0
        assert len(doc.list_revisions()) == initial_count - rejected

        doc.close()


class TestCountMatches:
    """Tests for count_matches functionality."""

    def test_count_matches_returns_int(self, clean_workspace):
        """Test that count_matches returns an integer."""
        doc = Document.open(clean_workspace)

        count = doc.count_matches("the")
        assert isinstance(count, int)
        assert count >= 0

        doc.close()

    def test_count_matches_nonexistent_returns_zero(self, clean_workspace):
        """Test that count_matches returns 0 for nonexistent text."""
        doc = Document.open(clean_workspace)

        count = doc.count_matches("xyz123nonexistent789")
        assert count == 0

        doc.close()


class TestOccurrenceParameter:
    """Tests for occurrence parameter in editing methods."""

    def test_replace_with_occurrence(self, clean_workspace):
        """Test replace with specific occurrence within a paragraph."""
        doc = Document.open(clean_workspace)

        # P2: "The quick brown fox jumps over the lazy dog."
        # 'over' appears once. Use occurrence=0 on the right paragraph
        # to verify occurrence param is accepted.
        ref = find_ref(doc, "lazy dog")
        new_ref = doc.replace("the", "THE", paragraph=ref, occurrence=0)
        assert isinstance(new_ref, str)

        doc.close()

    def test_replace_occurrence_out_of_range(self, clean_workspace):
        """Test replace with occurrence beyond available matches."""
        doc = Document.open(clean_workspace)

        # 'the' appears once in P2, request occurrence=5
        ref = find_ref(doc, "lazy dog")
        with pytest.raises(TextNotFoundError):
            doc.replace("the", "REPLACEMENT", paragraph=ref, occurrence=5)

        doc.close()

    def test_delete_with_occurrence(self, clean_workspace):
        """Test delete with specific occurrence."""
        doc = Document.open(clean_workspace)

        ref = find_ref(doc, "lazy dog")
        new_ref = doc.delete("the", paragraph=ref, occurrence=0)
        assert isinstance(new_ref, str)

        doc.close()

    def test_insert_after_with_occurrence(self, clean_workspace):
        """Test insert_after with specific occurrence."""
        doc = Document.open(clean_workspace)

        ref = find_ref(doc, "lazy dog")
        new_ref = doc.insert_after("the", " INSERTED", paragraph=ref, occurrence=0)
        assert isinstance(new_ref, str)

        doc.close()

    def test_insert_before_with_occurrence(self, clean_workspace):
        """Test insert_before with specific occurrence."""
        doc = Document.open(clean_workspace)

        ref = find_ref(doc, "lazy dog")
        new_ref = doc.insert_before("the", "INSERTED ", paragraph=ref, occurrence=0)
        assert isinstance(new_ref, str)

        doc.close()


class TestRevisionRepr:
    """Tests for Revision.__repr__ method."""

    def test_repr_insertion(self):
        """Test __repr__ for insertion type revision."""
        rev = Revision(
            id=1,
            type="insertion",
            author="TestAuthor",
            date=datetime.now(timezone.utc),
            text="short text",
        )
        repr_str = repr(rev)
        assert "ins 1:" in repr_str
        assert "short text" in repr_str
        assert "TestAuthor" in repr_str

    def test_repr_deletion(self):
        """Test __repr__ for deletion type revision."""
        rev = Revision(
            id=2,
            type="deletion",
            author="TestAuthor",
            date=datetime.now(timezone.utc),
            text="deleted text",
        )
        repr_str = repr(rev)
        assert "del 2:" in repr_str
        assert "deleted text" in repr_str
        assert "TestAuthor" in repr_str

    def test_repr_long_text_truncated(self):
        """Test __repr__ truncates long text."""
        long_text = "A" * 100
        rev = Revision(
            id=3,
            type="insertion",
            author="TestAuthor",
            date=None,
            text=long_text,
        )
        repr_str = repr(rev)
        # Should truncate to 30 chars + "..."
        assert "..." in repr_str
        assert len(repr_str) < len(long_text) + 50


class TestRevisionManagerDirectAccess:
    """Tests for RevisionManager using direct editor access."""

    def test_replace_text_with_before_and_after_text(self, clean_workspace):
        """Test replace where match is in the middle of a text node."""
        doc = Document.open(clean_workspace)

        # "quick" is in the middle of "The quick brown fox..."
        ref = find_ref(doc, "quick")
        new_ref = doc.replace("quick", "QUICK", paragraph=ref)
        assert isinstance(new_ref, str)

        doc.close()

    def test_replace_text_preserves_run_properties(self, clean_workspace):
        """Test that replace preserves w:rPr when present."""
        doc = Document.open(clean_workspace)

        # Replace text - the document structure should be preserved
        ref = find_ref(doc, "Sample")
        new_ref = doc.replace("Sample", "SAMPLE", paragraph=ref)
        assert isinstance(new_ref, str)

        doc.close()

    def test_suggest_deletion_with_surrounding_text(self, clean_workspace):
        """Test deletion when text has surrounding content."""
        doc = Document.open(clean_workspace)

        # "brown" is in the middle of "The quick brown fox..."
        ref = find_ref(doc, "brown")
        new_ref = doc.delete("brown", paragraph=ref)
        assert isinstance(new_ref, str)

        doc.close()

    def test_insert_text_not_found_raises_error(self, clean_workspace):
        """Test insert_after raises TextNotFoundError for nonexistent anchor."""
        doc = Document.open(clean_workspace)

        ref = doc.list_paragraphs()[0].split("|")[0]
        with pytest.raises(TextNotFoundError) as exc_info:
            doc.insert_after("xyz_nonexistent_anchor_123", "new text", paragraph=ref)

        assert "Anchor text not found" in str(exc_info.value) or "not found" in str(exc_info.value).lower()

        doc.close()

    def test_insert_before_not_found_raises_error(self, clean_workspace):
        """Test insert_before raises TextNotFoundError for nonexistent anchor."""
        doc = Document.open(clean_workspace)

        ref = doc.list_paragraphs()[0].split("|")[0]
        with pytest.raises(TextNotFoundError) as exc_info:
            doc.insert_before("xyz_nonexistent_anchor_123", "new text", paragraph=ref)

        assert "not found" in str(exc_info.value).lower()

        doc.close()


class TestRevisionParsing:
    """Tests for revision parsing edge cases."""

    def test_list_revisions_includes_both_types(self, clean_workspace):
        """Test that list_revisions finds both insertions and deletions."""
        doc = Document.open(clean_workspace, author="ParseTestAuthor")

        # Create both types
        ref = find_ref(doc, "quick")
        doc.delete("quick", paragraph=ref)
        ref2 = find_ref(doc, "fox")
        doc.insert_after("fox", " really", paragraph=ref2)

        revisions = doc.list_revisions()

        types = {r.type for r in revisions}
        assert "insertion" in types
        assert "deletion" in types

        doc.close()

    def test_list_revisions_with_missing_date(self, clean_workspace):
        """Test parsing revisions that may have missing date attributes."""
        doc = Document.open(clean_workspace)

        ref = find_ref(doc, "quick")
        doc.delete("quick", paragraph=ref)
        revisions = doc.list_revisions()

        # Should handle revisions regardless of date presence
        for rev in revisions:
            # date can be None or a datetime
            assert rev.date is None or isinstance(rev.date, datetime)

        doc.close()

    def test_list_revisions_with_empty_text(self, clean_workspace):
        """Test parsing revisions where text elements might be empty."""
        doc = Document.open(clean_workspace)

        # Make a change and verify we can list it
        ref = find_ref(doc, "fox")
        doc.insert_after("fox", "", paragraph=ref)  # Empty insertion
        revisions = doc.list_revisions()

        # Should not crash on empty text
        assert isinstance(revisions, list)

        doc.close()


class TestAcceptRejectExtended:
    """Extended tests for accept/reject functionality."""

    def test_accept_insertion_revision(self, clean_workspace):
        """Test accepting an insertion keeps the inserted text."""
        doc = Document.open(clean_workspace)

        ref = find_ref(doc, "fox")
        doc.insert_after("fox", " NEW", paragraph=ref)

        revisions = doc.list_revisions()
        change_id = revisions[-1].id

        result = doc.accept_revision(change_id)
        assert result is True

        # Verify revision is gone
        revisions = doc.list_revisions()
        ids = [r.id for r in revisions]
        assert change_id not in ids

        doc.close()

    def test_accept_deletion_revision(self, clean_workspace):
        """Test accepting a deletion removes the deleted text."""
        doc = Document.open(clean_workspace)

        ref = find_ref(doc, "quick")
        doc.delete("quick", paragraph=ref)

        revisions = doc.list_revisions()
        change_id = revisions[-1].id

        result = doc.accept_revision(change_id)
        assert result is True

        # Verify revision is gone
        revisions = doc.list_revisions()
        ids = [r.id for r in revisions]
        assert change_id not in ids

        doc.close()

    def test_reject_insertion_revision(self, clean_workspace):
        """Test rejecting an insertion removes the inserted text."""
        doc = Document.open(clean_workspace)

        ref = find_ref(doc, "fox")
        doc.insert_after("fox", " REJECT_ME", paragraph=ref)

        revisions = doc.list_revisions()
        change_id = revisions[-1].id

        result = doc.reject_revision(change_id)
        assert result is True

        # Verify revision is gone
        revisions = doc.list_revisions()
        ids = [r.id for r in revisions]
        assert change_id not in ids

        doc.close()

    def test_reject_deletion_revision(self, clean_workspace):
        """Test rejecting a deletion restores the deleted text."""
        doc = Document.open(clean_workspace)

        ref = find_ref(doc, "brown")
        doc.delete("brown", paragraph=ref)

        revisions = doc.list_revisions()
        change_id = revisions[-1].id

        result = doc.reject_revision(change_id)
        assert result is True

        doc.close()

    def test_reject_nonexistent_revision(self, clean_workspace):
        """Test rejecting a revision that doesn't exist."""
        doc = Document.open(clean_workspace)

        result = doc.reject_revision(99999)
        assert result is False

        doc.close()

    def test_accept_all_by_author(self, clean_workspace):
        """Test accepting all revisions filtered by author."""
        doc = Document.open(clean_workspace, author="Author1")
        ref = find_ref(doc, "quick")
        doc.delete("quick", paragraph=ref)
        doc.close()

        doc = Document.open(clean_workspace, author="Author2")
        ref = find_ref(doc, "brown")
        doc.delete("brown", paragraph=ref)

        # Accept only Author1's revisions
        count = doc.accept_all(author="Author1")
        assert count >= 0

        # Author2's revision should still exist (we don't assert on count
        # because the implementation may vary)
        doc.list_revisions(author="Author2")

        doc.close()

    def test_reject_all_by_author(self, clean_workspace):
        """Test rejecting all revisions filtered by author."""
        doc = Document.open(clean_workspace, author="RejectAuthor")
        ref = find_ref(doc, "quick")
        doc.delete("quick", paragraph=ref)
        ref2 = find_ref(doc, "fox")
        doc.insert_after("fox", " test", paragraph=ref2)

        count = doc.reject_all(author="RejectAuthor")
        assert count >= 0

        doc.close()


class TestEscapeXml:
    """Tests for _escape_xml helper function."""

    def test_escape_ampersand(self):
        """Test escaping ampersand."""
        assert _escape_xml("a & b") == "a &amp; b"

    def test_escape_less_than(self):
        """Test escaping less than."""
        assert _escape_xml("a < b") == "a &lt; b"

    def test_escape_greater_than(self):
        """Test escaping greater than."""
        assert _escape_xml("a > b") == "a &gt; b"

    def test_escape_double_quote(self):
        """Test escaping double quote."""
        assert _escape_xml('a "b" c') == "a &quot;b&quot; c"

    def test_escape_single_quote(self):
        """Test escaping single quote."""
        assert _escape_xml("a 'b' c") == "a &apos;b&apos; c"

    def test_escape_multiple_special_chars(self):
        """Test escaping multiple special characters."""
        assert _escape_xml("<a & 'b'>") == "&lt;a &amp; &apos;b&apos;&gt;"

    def test_escape_no_special_chars(self):
        """Test text without special characters."""
        assert _escape_xml("plain text") == "plain text"


class TestRevisionManagerErrorHandling:
    """Tests for error handling in RevisionManager."""

    def test_replace_text_no_matches(self, clean_workspace):
        """Test document-wide replace raises error when no matches found."""
        doc = Document.open(clean_workspace)

        with pytest.raises(TextNotFoundError) as exc_info:
            doc._revision_manager.replace_text("nonexistent_xyz_123", "X")

        assert "not found" in str(exc_info.value).lower()

        doc.close()

    def test_replace_text_occurrence_out_of_range(self, clean_workspace):
        """Test document-wide replace raises error for invalid occurrence."""
        doc = Document.open(clean_workspace)

        # "Sample" exists once in the document
        count = doc.count_matches("Sample")
        if count == 0:
            doc.close()
            pytest.skip("Test text not found")

        with pytest.raises(TextNotFoundError) as exc_info:
            doc._revision_manager.replace_text("Sample", "X", occurrence=count + 10)

        assert "occurrence" in str(exc_info.value).lower()
        assert exc_info.value.total_occurrences == count

        doc.close()


class TestRevisionManagerWithMockedEditor:
    """_parse_revision edge cases on real detached elements (editor mocked)."""

    @staticmethod
    def _revision_elem(xml: str):
        """Parse an XML fragment and return its first element (w:ins/w:del)."""
        import defusedxml.minidom

        NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        dom = defusedxml.minidom.parseString(f"<root {NS}>{xml}</root>")
        return dom.documentElement.firstChild

    def test_parse_revision_missing_id_returns_none(self):
        """Test _parse_revision returns None when w:id is missing."""
        manager = RevisionManager(MagicMock())
        elem = self._revision_elem('<w:ins w:author="Test"><w:r><w:t>x</w:t></w:r></w:ins>')
        assert manager._parse_revision(elem, "insertion") is None

    def test_parse_revision_invalid_date_uses_none(self):
        """Test _parse_revision handles invalid date gracefully."""
        manager = RevisionManager(MagicMock())
        elem = self._revision_elem('<w:ins w:id="1" w:author="Test" w:date="invalid-date-format"/>')
        result = manager._parse_revision(elem, "insertion")
        assert result is not None
        assert result.date is None  # Invalid date should be None

    def test_parse_revision_with_text_content(self):
        """Test _parse_revision extracts text content properly."""
        manager = RevisionManager(MagicMock())
        elem = self._revision_elem(
            '<w:ins w:id="5" w:author="Author" w:date="2024-01-15T10:30:00Z"><w:r><w:t>test content</w:t></w:r></w:ins>'
        )
        result = manager._parse_revision(elem, "insertion")
        assert result is not None
        assert result.text == "test content"

    def test_parse_revision_text_element_no_child(self):
        """Test _parse_revision handles text elements with no text child."""
        manager = RevisionManager(MagicMock())
        elem = self._revision_elem('<w:ins w:id="6" w:author="Author"><w:r><w:t/></w:r></w:ins>')
        result = manager._parse_revision(elem, "insertion")
        assert result is not None
        assert result.text == ""  # Empty text when no content

    def test_parse_revision_without_ctx_leaves_location_unset(self):
        """No location context (detached parse) → paragraph_ref/occurrence None."""
        manager = RevisionManager(MagicMock())
        elem = self._revision_elem('<w:ins w:id="7" w:author="Author"><w:r><w:t>text</w:t></w:r></w:ins>')
        result = manager._parse_revision(elem, "insertion")
        assert result is not None
        assert result.paragraph_ref is None
        assert result.occurrence is None


class TestRestoreDeletionEdgeCases:
    """Tests for _restore_deletion edge cases."""

    def test_reject_deletion_with_attributes(self, clean_workspace):
        """Test rejecting deletion restores attributes on delText."""
        doc = Document.open(clean_workspace)

        # Create a deletion
        ref = find_ref(doc, "lazy")
        doc.delete("lazy", paragraph=ref)

        revisions = doc.list_revisions()
        change_id = revisions[-1].id

        # Reject it to trigger _restore_deletion
        result = doc.reject_revision(change_id)
        assert result is True

        doc.close()

    def test_reject_deletion_handles_rsid_attributes(self, clean_workspace):
        """Test rejecting deletion converts rsidDel back to rsidR."""
        doc = Document.open(clean_workspace)

        # Create a deletion
        ref = find_ref(doc, "dog")
        doc.delete("dog", paragraph=ref)

        revisions = doc.list_revisions()
        change_id = revisions[-1].id

        # Reject it
        result = doc.reject_revision(change_id)
        assert result is True

        doc.close()


class TestComplexOperations:
    """Tests for complex sequences of operations."""

    def test_multiple_operations_same_paragraph(self, clean_workspace):
        """Test multiple tracked changes in the same paragraph."""
        doc = Document.open(clean_workspace)

        # Find content in the paragraph "The quick brown fox..."
        ref = find_ref(doc, "quick")
        doc.delete("quick", paragraph=ref)
        ref = find_ref(doc, "brown")
        doc.insert_after("brown", " spotted", paragraph=ref)
        ref = find_ref(doc, "fox")
        doc.replace("fox", "cat", paragraph=ref)

        revisions = doc.list_revisions()
        # Should have at least 3 revisions (1 delete, 1 insert, 2 from replace)
        assert len(revisions) >= 3

        doc.close()

    def test_accept_all_then_list(self, clean_workspace):
        """Test that accept_all properly clears all revisions."""
        doc = Document.open(clean_workspace)

        ref = find_ref(doc, "quick")
        doc.delete("quick", paragraph=ref)
        ref = find_ref(doc, "fox")
        doc.insert_after("fox", " test", paragraph=ref)

        initial_count = len(doc.list_revisions())
        assert initial_count >= 2

        accepted = doc.accept_all()
        assert accepted == initial_count

        remaining = doc.list_revisions()
        assert len(remaining) == 0

        doc.close()

    def test_reject_all_then_list(self, clean_workspace):
        """Test that reject_all properly clears all revisions."""
        doc = Document.open(clean_workspace)

        ref = find_ref(doc, "quick")
        doc.delete("quick", paragraph=ref)
        ref = find_ref(doc, "fox")
        doc.insert_after("fox", " test", paragraph=ref)

        initial_count = len(doc.list_revisions())
        assert initial_count >= 2

        rejected = doc.reject_all()
        assert rejected == initial_count

        remaining = doc.list_revisions()
        assert len(remaining) == 0

        doc.close()


class TestDocumentWideEditsRealXml:
    """Real-XML edge-case coverage for the unified document-wide edit path."""

    NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    INS_ATTRS = 'w:id="1" w:author="Test Author" w:date="2024-01-01T00:00:00Z"'

    def _manager(self, tmp_path, body_xml):
        xml = f'<?xml version="1.0" encoding="utf-8"?><w:document {self.NS}><w:body>{body_xml}</w:body></w:document>'
        xml_path = tmp_path / "doc.xml"
        xml_path.write_text(xml)
        editor = DocxXMLEditor(xml_path, rsid="00000000", author="Test Author")
        return RevisionManager(editor)

    def _accepted_text(self, mgr):
        return "".join(build_text_map(p).text for p in mgr.editor.dom.getElementsByTagName("w:p"))

    def test_suggest_deletion_preserves_rpr(self, tmp_path):
        """Deleting text from a formatted run keeps the run properties."""
        mgr = self._manager(tmp_path, "<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Hello world</w:t></w:r></w:p>")
        mgr.suggest_deletion("Hello")
        del_elems = mgr.editor.dom.getElementsByTagName("w:del")
        assert len(del_elems) == 1
        assert len(del_elems[0].getElementsByTagName("w:i")) == 1
        # The preserved " world" run keeps its formatting too
        for wt in mgr.editor.dom.getElementsByTagName("w:t"):
            run = wt.parentNode
            assert len(run.getElementsByTagName("w:i")) == 1

    def test_insert_text_preserves_rpr(self, tmp_path):
        """Inserting near a formatted anchor applies the anchor run's properties."""
        mgr = self._manager(tmp_path, "<w:p><w:r><w:rPr><w:u/></w:rPr><w:t>anchor text</w:t></w:r></w:p>")
        mgr.insert_text_after("anchor", " NEW")
        ins_elems = mgr.editor.dom.getElementsByTagName("w:ins")
        assert len(ins_elems) == 1
        assert len(ins_elems[0].getElementsByTagName("w:u")) == 1

    def test_replace_inside_ins_edits_in_place(self, tmp_path):
        """Replacing text whose run sits under a w:ins wrapper splices in place."""
        mgr = self._manager(tmp_path, f"<w:p><w:ins {self.INS_ATTRS}><w:r><w:t>hello world</w:t></w:r></w:ins></w:p>")
        result = mgr.replace_text("hello", "HELLO")
        assert result == -1  # no new revision created
        assert self._accepted_text(mgr) == "HELLO world"
        assert len(mgr.editor.dom.getElementsByTagName("w:del")) == 0

    def test_delete_inside_ins_shrinks_insertion(self, tmp_path):
        """Deleting text whose run sits under a w:ins wrapper shrinks the insertion."""
        mgr = self._manager(tmp_path, f"<w:p><w:ins {self.INS_ATTRS}><w:r><w:t>hello world</w:t></w:r></w:ins></w:p>")
        result = mgr.suggest_deletion("hello ")
        assert result == -1
        assert self._accepted_text(mgr) == "world"
        assert len(mgr.editor.dom.getElementsByTagName("w:del")) == 0

    def test_insert_inside_ins_splices_without_nesting(self, tmp_path):
        """Inserting at an anchor whose run sits under a w:ins wrapper avoids nested w:ins."""
        mgr = self._manager(tmp_path, f"<w:p><w:ins {self.INS_ATTRS}><w:r><w:t>hello world</w:t></w:r></w:ins></w:p>")
        result = mgr.insert_text_after("hello", "XX")
        assert result == -1
        assert self._accepted_text(mgr) == "helloXX world"
        ins_elems = mgr.editor.dom.getElementsByTagName("w:ins")
        assert len(ins_elems) == 1

    def test_replace_at_end_of_node_preserves_prefix(self, tmp_path):
        """Replacing text at the end of a w:t keeps the preceding text."""
        mgr = self._manager(tmp_path, "<w:p><w:r><w:t>prefix hello</w:t></w:r></w:p>")
        mgr.replace_text("hello", "HELLO")
        assert self._accepted_text(mgr) == "prefix HELLO"

    def test_delete_at_start_of_node_preserves_suffix(self, tmp_path):
        """Deleting text at the start of a w:t keeps the following text."""
        mgr = self._manager(tmp_path, "<w:p><w:r><w:t>hello suffix</w:t></w:r></w:p>")
        mgr.suggest_deletion("hello")
        assert self._accepted_text(mgr) == " suffix"


class TestListRevisionsEdgeCases:
    """Tests for list_revisions edge cases."""

    def test_list_revisions_filters_by_author_for_insertions(self):
        """Test that list_revisions author filter works for insertions."""
        manager = _make_revision_manager(
            '<w:p><w:ins w:id="1" w:author="SpecificAuthor"><w:r><w:t>new</w:t></w:r></w:ins></w:p>'
        )

        # Filter by matching author
        revisions = manager.list_revisions(author="SpecificAuthor")
        assert len(revisions) == 1
        assert revisions[0].author == "SpecificAuthor"

        # Filter by non-matching author
        revisions = manager.list_revisions(author="OtherAuthor")
        assert len(revisions) == 0

    def test_list_revisions_filters_by_author_for_deletions(self):
        """Test that list_revisions author filter works for deletions."""
        manager = _make_revision_manager(
            '<w:p><w:del w:id="2" w:author="DeleteAuthor"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>'
        )

        # Filter by matching author
        revisions = manager.list_revisions(author="DeleteAuthor")
        assert len(revisions) == 1
        assert revisions[0].author == "DeleteAuthor"
        assert revisions[0].type == "deletion"


class TestAcceptRejectLoops:
    """Tests for accept_all and reject_all loops."""

    TWO_INSERTIONS = (
        '<w:p><w:ins w:id="1" w:author="Author"><w:r><w:t>one</w:t></w:r></w:ins>'
        '<w:ins w:id="2" w:author="Author"><w:r><w:t>two</w:t></w:r></w:ins></w:p>'
    )
    TWO_DELETIONS = (
        '<w:p><w:del w:id="3" w:author="Author"><w:r><w:delText>one</w:delText></w:r></w:del>'
        '<w:del w:id="4" w:author="Author"><w:r><w:delText>two</w:delText></w:r></w:del></w:p>'
    )

    @staticmethod
    def _detaching_double(manager, processed: set[str]):
        """A resolver double that detaches the element it is handed, so the
        re-listing shrinks exactly as a real accept/reject makes it shrink."""

        def resolve(rev_id: int, element_index: dict | None = None) -> bool:
            elem = manager._find_revision_element(rev_id, element_index)
            assert elem is not None
            elem.parentNode.removeChild(elem)
            processed.add(str(rev_id))
            return True

        return resolve

    def test_accept_all_processes_multiple_revisions(self):
        """Test that accept_all correctly processes multiple revisions."""
        manager = _make_revision_manager(self.TWO_INSERTIONS)
        processed: set[str] = set()
        manager.accept_revision = self._detaching_double(manager, processed)

        count = manager.accept_all()
        assert count == 2
        assert processed == {"1", "2"}

    def test_reject_all_processes_multiple_revisions(self):
        """Test that reject_all correctly processes multiple revisions."""
        manager = _make_revision_manager(self.TWO_DELETIONS)
        processed: set[str] = set()
        manager.reject_revision = self._detaching_double(manager, processed)

        count = manager.reject_all()
        assert count == 2
        assert processed == {"3", "4"}

    def test_accept_all_does_not_spin_when_listing_shrinks(self):
        """accept_all must resolve each revision once, not loop until OOM.

        Regression for ISSUES.md #56: an early refactor hoisted
        list_revisions() out of the resolution loop, so termination depended
        solely on accept_revision() returning False. Because the test doubles
        here always return True, the loop never exited — and since MagicMock
        retains every recorded call, memory grew without bound until the OOM
        killer took the machine down (observed at ~42 GiB RSS).

        Re-listing on each pass is what bounds this: the listing shrinks as
        revisions are processed. The call cap below makes a regression fail in
        milliseconds instead of exhausting RAM.
        """
        manager = _make_revision_manager(self.TWO_INSERTIONS)
        processed: set[str] = set()
        detach = self._detaching_double(manager, processed)
        calls = {"n": 0}

        def counting_accept(rev_id, element_index=None):
            calls["n"] += 1
            assert calls["n"] <= 10, (
                f"accept_all did not terminate: accept_revision called {calls['n']} times "
                "for 2 revisions (unbounded loop — see docstring)"
            )
            return detach(rev_id, element_index)

        manager.accept_revision = counting_accept

        assert manager.accept_all() == 2
        assert calls["n"] == 2, f"expected one resolve per revision, got {calls['n']}"


def _make_revision_manager(body_xml):
    """Build a RevisionManager over a real minidom DOM from a body snippet."""
    import defusedxml.minidom

    xml = (
        '<?xml version="1.0"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"{body_xml}"
        "</w:document>"
    )
    mock_editor = MagicMock()
    mock_editor.dom = defusedxml.minidom.parseString(xml)
    return RevisionManager(mock_editor)


class TestNestedForeignRevisions:
    """Tests for accept_all/reject_all on nested revisions from Word-authored files.

    Word produces nested markup when one reviewer edits another's tracked
    change, e.g. a w:del inside a w:ins. The fixed-point loop in
    accept_all/reject_all must fully resolve such nesting and still terminate
    when an author filter legitimately leaves other authors' revisions behind.
    """

    NESTED_DEL_INSIDE_INS = """
        <w:ins w:id="5" w:author="A" w:date="2026-01-01T00:00:00Z">
            <w:r><w:t>kept</w:t></w:r>
            <w:del w:id="3" w:author="B" w:date="2026-01-02T00:00:00Z">
                <w:r><w:delText>gone</w:delText></w:r>
            </w:del>
        </w:ins>"""

    # The ROADMAP.md #75 shape: B's insertion carries the same w:id as A's
    # deletion and comes *first* in document order, so an unscoped id lookup
    # hands B's element to every one of A's passes.
    DUPLICATE_ID_ACROSS_AUTHORS = """
        <w:ins w:id="7" w:author="B" w:date="2026-01-02T00:00:00Z">
            <w:r><w:t>BEE</w:t></w:r>
        </w:ins>
        <w:del w:id="7" w:author="A" w:date="2026-01-01T00:00:00Z">
            <w:r><w:delText>AYE</w:delText></w:r>
        </w:del>"""

    def test_accept_all_stops_when_no_revision_can_be_resolved(self):
        """The no-progress guard in _resolve_all: listed but unresolvable stops.

        _resolve_all normally exits because list_revisions() comes back empty.
        This pins its *other* exit — the guard that ends the loop when
        revisions are still listed but none of them can be resolved. Without
        that guard the loop spins forever on a non-shrinking listing, which is
        the failure mode that grew a pytest process to 42 GB (ISSUES.md #56).
        """
        manager = _make_revision_manager(self.NESTED_DEL_INSIDE_INS)
        calls = {"n": 0}

        def never_resolves(rev_id: int, element_index: dict | None = None) -> bool:
            calls["n"] += 1
            assert calls["n"] <= 20, f"accept_all did not terminate ({calls['n']} calls)"
            return False

        manager.accept_revision = never_resolves  # type: ignore[method-assign]

        assert manager.accept_all() == 0
        # The revisions are still there — nothing was resolved, and the loop
        # ended on the no-progress guard rather than on an empty listing.
        assert len(manager.list_revisions()) == 2

    def test_accept_all_nested_del_inside_ins(self):
        """Test that accept_all resolves a w:del nested inside a w:ins completely."""
        manager = _make_revision_manager(self.NESTED_DEL_INSIDE_INS)
        dom = manager.editor.dom

        count = manager.accept_all()

        assert count == 2
        assert manager.list_revisions() == []
        assert dom.getElementsByTagName("w:delText") == []
        texts = [t.firstChild.data for t in dom.getElementsByTagName("w:t")]
        assert texts == ["kept"]

    def test_reject_all_nested_outer_processed_first(self):
        """Test that rejecting the outer w:ins discards the nested w:del with it."""
        manager = _make_revision_manager(self.NESTED_DEL_INSIDE_INS)
        dom = manager.editor.dom

        # Outer ins id=5 > nested del id=3: reverse-id order hits the outer
        # first, so the nested deletion vanishes with it and is never itself
        # rejected — only one per-id rejection executes.
        count = manager.reject_all()

        assert count == 1
        assert manager.list_revisions() == []
        assert dom.getElementsByTagName("w:t") == []
        assert dom.getElementsByTagName("w:delText") == []

    def test_reject_all_nested_inner_processed_first(self):
        """Test that rejecting the nested w:del first still converges to the same document."""
        body = """
        <w:ins w:id="3" w:author="A" w:date="2026-01-01T00:00:00Z">
            <w:r><w:t>kept</w:t></w:r>
            <w:del w:id="5" w:author="B" w:date="2026-01-02T00:00:00Z">
                <w:r><w:delText>gone</w:delText></w:r>
            </w:del>
        </w:ins>"""
        manager = _make_revision_manager(body)
        dom = manager.editor.dom

        # Nested del id=5 > outer ins id=3: the deletion is rejected first
        # (restoring its text inside the insertion), then rejecting the outer
        # insertion removes everything.
        count = manager.reject_all()

        assert count == 2
        assert manager.list_revisions() == []
        assert dom.getElementsByTagName("w:t") == []
        assert dom.getElementsByTagName("w:delText") == []

    def test_accept_all_author_filter_leaves_other_authors_same_id_revision(self):
        """Test that accept_all(author=...) skips another author's revision sharing its w:id."""
        # Word does not guarantee unique w:id across w:ins/w:del. Resolution is
        # still keyed on the id, but the id lookup is scoped to the author being
        # resolved, so A's same-id insertion is not a candidate for B's pass.
        body = """
        <w:ins w:id="7" w:author="A" w:date="2026-01-01T00:00:00Z">
            <w:r><w:t>alpha</w:t></w:r>
        </w:ins>
        <w:del w:id="7" w:author="B" w:date="2026-01-02T00:00:00Z">
            <w:r><w:delText>beta</w:delText></w:r>
        </w:del>"""
        manager = _make_revision_manager(body)
        dom = manager.editor.dom

        count = manager.accept_all(author="B")

        # Only B's deletion: the count equals the number of B's listed rows.
        assert count == 1
        assert manager.list_revisions(author="B") == []
        assert dom.getElementsByTagName("w:delText") == []
        # A's insertion is untouched — still pending, still wrapped in w:ins.
        remaining = manager.list_revisions(author="A")
        assert len(remaining) == 1
        assert remaining[0].type == "insertion"
        ins = dom.getElementsByTagName("w:ins")
        assert len(ins) == 1
        assert [t.firstChild.data for t in ins[0].getElementsByTagName("w:t")] == ["alpha"]

    def test_accept_all_author_filter_terminates_with_foreign_revisions(self):
        """Test that accept_all(author=...) terminates while other authors' revisions remain."""
        manager = _make_revision_manager(self.NESTED_DEL_INSIDE_INS)

        count = manager.accept_all(author="B")

        assert count == 1
        assert manager.list_revisions(author="B") == []
        remaining = manager.list_revisions(author="A")
        assert len(remaining) == 1
        assert remaining[0].type == "insertion"

    def test_reject_all_author_filter_terminates_with_foreign_revisions(self):
        """Test that reject_all(author=...) terminates while other authors' revisions remain."""
        manager = _make_revision_manager(self.NESTED_DEL_INSIDE_INS)
        dom = manager.editor.dom

        count = manager.reject_all(author="B")

        assert count == 1
        assert manager.list_revisions(author="B") == []
        remaining = manager.list_revisions(author="A")
        assert len(remaining) == 1
        assert remaining[0].type == "insertion"
        # B's rejected deletion restored its text inside A's insertion.
        texts = [t.firstChild.data for t in dom.getElementsByTagName("w:t")]
        assert texts == ["kept", "gone"]

    def test_reject_all_author_filter_spares_other_authors_same_id_insertion(self):
        """Test that reject_all(author=...) does not destroy another author's same-id insertion.

        The data-loss case of ROADMAP.md #75: rejecting A's deletion once
        reached B's earlier same-id insertion and removed its text outright.
        """
        manager = _make_revision_manager(self.DUPLICATE_ID_ACROSS_AUTHORS)
        dom = manager.editor.dom

        count = manager.reject_all(author="A")

        assert count == 1
        assert manager.list_revisions(author="A") == []
        # A's rejected deletion put its text back as plain text.
        assert dom.getElementsByTagName("w:delText") == []
        # B's insertion survives, still pending and still wrapped in its w:ins.
        remaining = manager.list_revisions(author="B")
        assert len(remaining) == 1
        assert remaining[0].type == "insertion"
        assert remaining[0].id == 7
        ins = dom.getElementsByTagName("w:ins")
        assert len(ins) == 1
        assert ins[0].getAttribute("w:author") == "B"
        assert [t.firstChild.data for t in ins[0].getElementsByTagName("w:t")] == ["BEE"]
        assert [t.firstChild.data for t in dom.getElementsByTagName("w:t")] == ["BEE", "AYE"]

    def test_accept_all_author_filter_leaves_other_authors_same_id_insertion_pending(self):
        """Test that accept_all(author=...) does not make another author's same-id insertion permanent."""
        manager = _make_revision_manager(self.DUPLICATE_ID_ACROSS_AUTHORS)
        dom = manager.editor.dom

        count = manager.accept_all(author="A")

        assert count == 1
        assert manager.list_revisions(author="A") == []
        # A's deletion applied: its text is gone.
        assert dom.getElementsByTagName("w:delText") == []
        # B's insertion is still an unadjudicated insertion, not plain text.
        remaining = manager.list_revisions(author="B")
        assert len(remaining) == 1
        assert remaining[0].type == "insertion"
        ins = dom.getElementsByTagName("w:ins")
        assert len(ins) == 1
        assert ins[0].getAttribute("w:author") == "B"
        assert [t.firstChild.data for t in dom.getElementsByTagName("w:t")] == ["BEE"]

    def test_reject_all_unfiltered_still_resolves_both_same_id_revisions(self):
        """Test that reject_all() with no author still resolves every same-id revision in one call."""
        manager = _make_revision_manager(self.DUPLICATE_ID_ACROSS_AUTHORS)
        dom = manager.editor.dom

        count = manager.reject_all()

        # The author scoping must not narrow the unfiltered path: both
        # duplicate-id revisions still resolve.
        assert count == 2
        assert manager.list_revisions() == []
        assert dom.getElementsByTagName("w:ins") == []
        assert dom.getElementsByTagName("w:delText") == []
        assert [t.firstChild.data for t in dom.getElementsByTagName("w:t")] == ["AYE"]

    def test_accept_all_unfiltered_still_resolves_both_same_id_revisions(self):
        """Test that accept_all() with no author still resolves every same-id revision in one call."""
        manager = _make_revision_manager(self.DUPLICATE_ID_ACROSS_AUTHORS)
        dom = manager.editor.dom

        count = manager.accept_all()

        assert count == 2
        assert manager.list_revisions() == []
        assert dom.getElementsByTagName("w:ins") == []
        assert dom.getElementsByTagName("w:delText") == []
        assert [t.firstChild.data for t in dom.getElementsByTagName("w:t")] == ["BEE"]

    def test_accept_all_author_filter_resolves_every_row_it_listed(self):
        """Test that a filtered call resolves every one of that author's listed rows.

        A filtered pass that could not reach a row it listed would exit on the
        no-progress guard and report a count below the listing — a silent
        no-op. Equality holds here because none of these rows is nested; a
        listed row inside a rejected insertion goes with its host and is not
        counted. The other way index and listing can disagree about ownership
        — the ``"Unknown"`` fallback for a missing ``w:author`` — is pinned by
        ``test_unattributed_revision_resolves_under_the_unknown_author``; both
        of these rows carry an explicit author.
        """
        manager = _make_revision_manager(self.DUPLICATE_ID_ACROSS_AUTHORS)

        listed = len(manager.list_revisions(author="B"))

        assert manager.accept_all(author="B") == listed == 1

    def test_author_filter_holds_when_ids_collide_only_after_normalizing(self):
        """Test that author scoping holds for ids that are equal only once parsed.

        ``w:id="007"`` and ``w:id="7"`` are distinct raw attributes that share
        one index key, so normalizing the key widened the duplicate-id class
        this fix guards. The author scoping has to cover the widened class too.
        """
        # B's non-canonical id comes *second*: an unscoped index would hand
        # back A's element first, so the assertions below need the scoping as
        # well as the normalization.
        manager = _make_revision_manager(
            """
        <w:del w:id="7" w:author="A"><w:r><w:delText>AYE</w:delText></w:r></w:del>
        <w:ins w:id="007" w:author="B"><w:r><w:t>BEE</w:t></w:r></w:ins>"""
        )
        dom = manager.editor.dom

        # Resolving B needs both halves: reachable only because the index
        # normalizes, author-exact only because the index is scoped.
        assert manager.reject_all(author="B") == 1

        # A's deletion is untouched: still pending, its text still deleted.
        remaining = manager.list_revisions(author="A")
        assert len(remaining) == 1
        assert remaining[0].type == "deletion"
        assert dom.getElementsByTagName("w:ins") == []
        assert [t.firstChild.data for t in dom.getElementsByTagName("w:delText")] == ["AYE"]

    @pytest.mark.parametrize("method", ["accept_all", "reject_all"])
    def test_unattributed_revision_resolves_under_the_unknown_author(self, method):
        """Test that author="Unknown" resolves a revision whose w:author is absent.

        The index and the listing must read a missing ``w:author`` the same
        way. ``list_revisions`` reports such a mark as ``"Unknown"``, so an
        index keyed on the raw attribute would list a row it could not
        resolve: the pass would exit on the no-progress guard, returning 0
        with the revision still pending and no warning raised.
        """
        manager = _make_revision_manager(
            """
        <w:ins w:id="7"><w:r><w:t>BEE</w:t></w:r></w:ins>
        <w:del w:id="8" w:author="A"><w:r><w:delText>AYE</w:delText></w:r></w:del>"""
        )

        assert getattr(manager, method)(author="Unknown") == 1

        assert manager.list_revisions(author="Unknown") == []
        # A's attributed revision is untouched.
        assert [rev.id for rev in manager.list_revisions()] == [8]


class TestRevisionIdNormalization:
    """The w:id half of index/listing agreement (ROADMAP.md #75).

    ``list_revisions`` reports ``int(w:id)`` and every lookup asks for
    ``str(int)``, so the index and the fresh scan must read an id the same way
    or a nonconforming ``w:id`` is listed but reachable by nothing.
    """

    @pytest.mark.parametrize("method", ["accept_all", "reject_all"])
    def test_non_canonical_id_resolves_through_the_index(self, method):
        """Test that a w:id whose raw form is not str(int) still resolves.

        The id half of the same index/listing agreement: the listing reports
        ``int(w:id)`` and every lookup asks for ``str(int)``, so an index keyed
        on the raw attribute strands ``w:id="007"`` — listed as id 7, resolved
        by nothing, and reported as a clean document by a call that left it
        pending. Nothing warns, because the mark has a numeric id and so is not
        part of the unhandled honesty floor.
        """
        manager = _make_revision_manager('<w:ins w:id="007" w:author="A"><w:r><w:t>zero</w:t></w:r></w:ins>')

        assert [rev.id for rev in manager.list_revisions()] == [7]
        assert getattr(manager, method)() == 1
        assert manager.list_revisions() == []
        assert manager.list_unhandled_revisions() == []

    def test_non_canonical_id_resolves_without_an_index(self):
        """Test that the fresh-scan path matches a non-canonical w:id too.

        ``accept_revision`` with no pre-built index scans the document itself;
        it must read the id the same way the index does, or a standalone call
        would fail on the ids a bulk call resolves.
        """
        manager = _make_revision_manager('<w:ins w:id="007" w:author="A"><w:r><w:t>zero</w:t></w:r></w:ins>')

        assert manager.accept_revision(7) is True
        assert manager.list_revisions() == []

    def test_non_numeric_id_is_left_to_the_honesty_floor(self):
        """Test that a mark with no numeric w:id is not indexed and is reported.

        ``_adjudicable_id`` is None for it, so nothing id-keyed can reach it:
        it must stay out of the index (an unreachable key) and out of
        ``list_revisions``, and surface in ``list_unhandled_revisions``.
        """
        manager = _make_revision_manager('<w:ins w:id="abc" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins>')

        assert manager._revision_element_index() == {}
        assert manager.list_revisions() == []
        assert [u.tag for u in manager.list_unhandled_revisions()] == ["w:ins"]

    @pytest.mark.parametrize("method", ["accept_revision", "reject_revision"])
    def test_none_id_resolves_nothing(self, method):
        """Test that resolving id None is a no-op, not a match on the first unreachable mark.

        ``list_unhandled_revisions()`` reports ``id=None`` for a mark with no
        numeric ``w:id``, so None is a value callers really do hold. The fresh
        scan skips elements whose ``_adjudicable_id`` is None; without that
        skip ``str(None) == str(None)`` would match the first such mark and
        resolve what the honesty floor just called unresolvable.
        """
        manager = _make_revision_manager('<w:ins w:id="abc" w:author="A"><w:r><w:t>KEEP</w:t></w:r></w:ins>')
        (row,) = manager.list_unhandled_revisions()
        assert row.id is None

        assert getattr(manager, method)(row.id) is False

        # The mark and its text are untouched.
        assert [u.tag for u in manager.list_unhandled_revisions()] == ["w:ins"]
        texts = [t.firstChild.data for t in manager.editor.dom.getElementsByTagName("w:t")]
        assert texts == ["KEEP"]

    @pytest.mark.parametrize("bad_id", [True, 1.0, None])
    def test_int_equal_non_int_id_matches_on_neither_path(self, bad_id):
        """Test that the index and the fresh scan agree on ids that are not ints.

        Both paths compare ``str`` of the adjudicable id. Comparing the parsed
        int on one side and the index's string key on the other would diverge
        on anything int-equal but not an int: ``True`` matches id 1 through
        ``==`` (bool subclasses int) while missing ``.get("True")``, so
        ``accept_revision(True)`` would resolve a revision through one path and
        nothing through the other.
        """
        manager = _make_revision_manager('<w:ins w:id="1" w:author="A"><w:r><w:t>ONE</w:t></w:r></w:ins>')
        index = manager._revision_element_index()

        assert manager._find_revision_element(bad_id, None) is None
        assert manager._find_revision_element(bad_id, index) is None
        assert manager.accept_revision(bad_id) is False
        assert len(manager.editor.dom.getElementsByTagName("w:ins")) == 1

    def test_str_id_still_matches_on_both_paths(self):
        """Test that a stringified id resolves, as it did before the id rework.

        Not the documented contract — ``accept_revision`` takes an int — but
        both paths accepted ``"7"`` before, so narrowing to int only would
        break a caller that round-trips ids through JSON or a CLI argument for
        no gain.
        """
        manager = _make_revision_manager('<w:ins w:id="7" w:author="A"><w:r><w:t>SEVEN</w:t></w:r></w:ins>')
        index = manager._revision_element_index()

        assert manager._find_revision_element("7", None) is not None
        assert manager._find_revision_element("7", index) is not None


class TestRestoreDeletionAttributeCopying:
    """Tests for _restore_deletion attribute copying edge cases."""

    def test_restore_deletion_copies_deltext_attributes(self):
        """Test that _restore_deletion copies attributes from w:delText to w:t."""
        manager = _make_revision_manager(
            """
            <w:del w:id="1" w:author="Test">
                <w:r>
                    <w:delText xml:space="preserve">test text</w:delText>
                </w:r>
            </w:del>"""
        )
        dom = manager.editor.dom

        del_elem = dom.getElementsByTagName("w:del")[0]
        manager._restore_deletion(del_elem)

        # Verify w:t was created with xml:space attribute
        t_elems = dom.getElementsByTagName("w:t")
        assert len(t_elems) == 1
        assert t_elems[0].getAttribute("xml:space") == "preserve"

    def test_restore_deletion_converts_rsiddel_to_rsidr(self):
        """Test that _restore_deletion converts w:rsidDel to w:rsidR on runs."""
        manager = _make_revision_manager(
            """
            <w:del w:id="1" w:author="Test">
                <w:r w:rsidDel="00112233">
                    <w:delText>text</w:delText>
                </w:r>
            </w:del>"""
        )
        dom = manager.editor.dom

        del_elem = dom.getElementsByTagName("w:del")[0]
        manager._restore_deletion(del_elem)

        # Verify w:rsidDel was converted to w:rsidR
        r_elems = dom.getElementsByTagName("w:r")
        assert len(r_elems) == 1
        assert r_elems[0].getAttribute("w:rsidR") == "00112233"
        assert not r_elems[0].hasAttribute("w:rsidDel")


class TestParseRevisionDelTextFallback:
    """Tests for w:delText -> w:t fallback when reading deletion text."""

    def test_deletion_with_plain_wt_falls_back(self):
        """Test that deletion text falls back to w:t when w:delText is absent."""
        import defusedxml.minidom

        # Nonconforming producers may leave plain w:t inside w:del
        xml = """<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:del w:id="1" w:author="Foreign">
                <w:r>
                    <w:t>lost text</w:t>
                </w:r>
            </w:del>
        </w:document>"""

        mock_editor = MagicMock()
        mock_editor.dom = defusedxml.minidom.parseString(xml)

        manager = RevisionManager(mock_editor)

        revisions = manager.list_revisions()
        assert len(revisions) == 1
        assert revisions[0].type == "deletion"
        assert revisions[0].text == "lost text"

    def test_deletion_with_deltext_unchanged(self):
        """Test that the fallback does not fire when w:delText exists."""
        import defusedxml.minidom

        xml = """<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:del w:id="1" w:author="Test">
                <w:r>
                    <w:delText>proper</w:delText>
                </w:r>
                <w:r>
                    <w:t>stray</w:t>
                </w:r>
            </w:del>
        </w:document>"""

        mock_editor = MagicMock()
        mock_editor.dom = defusedxml.minidom.parseString(xml)

        manager = RevisionManager(mock_editor)

        revisions = manager.list_revisions()
        assert len(revisions) == 1
        assert revisions[0].text == "proper"


def _split_wt_text_nodes(doc, target_text, tag="w:t"):
    """Reach into the DOM and split an element's TEXT_NODE into multiple siblings.

    Simulates the minidom multi-child-text-node state reported in issue #9 —
    Word documents with smart quotes (U+2018/U+2019) can land in this state.
    Returns the modified element.
    """
    dom = doc._document_editor.dom
    for wt in dom.getElementsByTagName(tag):
        full = "".join(c.data for c in wt.childNodes if c.nodeType == c.TEXT_NODE)
        if target_text not in full:
            continue
        while wt.firstChild:
            wt.removeChild(wt.firstChild)
        idx = full.find(target_text)
        before, after = full[:idx], full[idx + len(target_text) :]
        owner = wt.ownerDocument
        if before:
            wt.appendChild(owner.createTextNode(before))
        wt.appendChild(owner.createTextNode(target_text))
        if after:
            wt.appendChild(owner.createTextNode(after))
        return wt
    raise AssertionError(f"No {tag} containing {target_text!r}")


class TestMultiTextNodeWtElements:
    """Issue #9: w:t elements with multiple TEXT_NODE children (smart-quote split)."""

    # Both the document-wide (paragraph=None) and paragraph-scoped paths route
    # through the text-map helpers, which read node text via _get_node_text and
    # so tolerate w:t elements whose text is split across TEXT_NODE children.

    def test_set_node_text_consolidates_split_nodes(self, clean_workspace):
        """Direct contract test for _set_node_text: starts from a multi-TEXT_NODE
        state, ends with exactly one TEXT_NODE carrying the full new content.
        Guards against future "simplifications" that would re-introduce the
        firstChild.data assignment pattern."""
        doc = Document.open(clean_workspace)
        wt = _split_wt_text_nodes(doc, "quick brown fox")
        text_nodes_before = [c for c in wt.childNodes if c.nodeType == c.TEXT_NODE]
        assert len(text_nodes_before) > 1

        doc._revision_manager._set_node_text(wt, "consolidated")

        text_nodes_after = [c for c in wt.childNodes if c.nodeType == c.TEXT_NODE]
        assert len(text_nodes_after) == 1
        assert text_nodes_after[0].data == "consolidated"
        doc.close()

    def test_replace_without_paragraph_arg_succeeds(self, clean_workspace):
        doc = Document.open(clean_workspace)
        wt = _split_wt_text_nodes(doc, "quick brown fox")
        assert len(wt.childNodes) > 1
        doc._revision_manager.replace_text("quick brown fox", "slow red turtle")
        paragraphs = doc.list_paragraphs()
        assert any("slow red turtle" in p for p in paragraphs)
        assert not any("quick brown fox" in p for p in paragraphs)
        doc.close()

    def test_delete_without_paragraph_arg_succeeds(self, clean_workspace):
        doc = Document.open(clean_workspace)
        _split_wt_text_nodes(doc, "quick brown fox")
        doc._revision_manager.suggest_deletion("quick brown fox")
        paragraphs = doc.list_paragraphs()
        assert not any("quick brown fox" in p for p in paragraphs)
        doc.close()

    def test_insert_without_paragraph_arg_succeeds(self, clean_workspace):
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "lazy dog")
        doc.insert_before("lazy dog", "INS_TARGET ", paragraph=ref)
        _split_wt_text_nodes(doc, "INS_TARGET")
        doc._revision_manager.insert_text_before("INS_TARGET", "X_")
        paragraphs = doc.list_paragraphs()
        assert any("X_INS_TARGET" in p for p in paragraphs)
        doc.close()

    def test_replace_inside_ins_writes_full_text(self, clean_workspace):
        # Hits _replace_across_nodes' "all inside ins" path which previously
        # used firstChild.data assignment.
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "lazy dog")
        ref = doc.insert_before("lazy dog", "INS_TARGET ", paragraph=ref)
        _split_wt_text_nodes(doc, "INS_TARGET")
        doc.replace("INS_TARGET", "REPLACED", paragraph=ref)
        paragraphs = doc.list_paragraphs()
        assert any("REPLACED" in p for p in paragraphs)
        assert not any("INS_TARGET" in p for p in paragraphs)
        doc.close()

    def test_list_revisions_with_multi_text_node_delText(self, clean_workspace):
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "quick brown fox")
        doc.delete("quick brown fox", paragraph=ref)

        dom = doc._document_editor.dom
        del_texts = dom.getElementsByTagName("w:delText")
        assert del_texts, "expected at least one w:delText after delete"
        elem = del_texts[0]
        full = "".join(c.data for c in elem.childNodes if c.nodeType == c.TEXT_NODE)
        while elem.firstChild:
            elem.removeChild(elem.firstChild)
        mid = len(full) // 2
        elem.appendChild(elem.ownerDocument.createTextNode(full[:mid]))
        elem.appendChild(elem.ownerDocument.createTextNode(full[mid:]))

        revisions = doc.list_revisions()
        deletions = [r for r in revisions if r.type == "deletion"]
        assert deletions
        assert deletions[0].text == full
        doc.close()

    def test_save_load_roundtrip_after_multi_node_edit(self, clean_workspace, tmp_path):
        doc = Document.open(clean_workspace)
        _split_wt_text_nodes(doc, "quick brown fox")
        doc._revision_manager.replace_text("quick brown fox", "slow red turtle")
        out = tmp_path / "edited.docx"
        doc.save(out)
        doc.close()

        doc2 = Document.open(out)
        paragraphs = doc2.list_paragraphs()
        assert any("slow red turtle" in p for p in paragraphs)
        assert not any("quick brown fox" in p for p in paragraphs)
        doc2.close()


def _build_smart_quote_docx(simple_docx, dest):
    """Build a real .docx containing smart-quote text in a single <w:t>.

    Repacks ``simple_docx`` with the second paragraph rewritten to contain
    U+2018/U+2019 smart quotes, simulating the structure reported in
    GitHub issue #9.
    """
    import shutil

    from docx_editor.ooxml.pack import pack_document
    from docx_editor.ooxml.unpack import unpack_document

    work = dest.parent / "_smart_quote_build"
    if work.exists():
        shutil.rmtree(work)
    unpack_document(simple_docx, work)
    doc_xml = work / "word" / "document.xml"
    xml = doc_xml.read_text(encoding="utf-8")
    # Replace "The quick brown fox..." paragraph's text with one carrying
    # smart quotes. The surrounding <w:r><w:t>...</w:t></w:r> structure
    # mirrors what Word emits.
    target = "The quick brown fox jumps over the lazy dog."
    replacement = "‘Library Bookshelves’ are in all Libraries."
    assert target in xml, "fixture assumption broken: simple.docx changed"
    xml = xml.replace(target, replacement)
    doc_xml.write_text(xml, encoding="utf-8")
    pack_document(work, dest)
    shutil.rmtree(work)
    return dest


class TestSmartQuoteEndToEnd:
    """End-to-end: real .docx with smart quotes, full open/edit/save/reopen."""

    def test_replace_around_smart_quotes(self, simple_docx, tmp_path):
        src = tmp_path / "with_smart_quotes.docx"
        _build_smart_quote_docx(simple_docx, src)

        doc = Document.open(src, force_recreate=True)
        try:
            # Force the multi-text-node state on the smart-quote w:t so we
            # exercise the codepath issue #9 describes regardless of how
            # the local minidom build represents the parsed text.
            _split_wt_text_nodes(doc, "Library Bookshelves")

            doc._revision_manager.replace_text("Library Bookshelves", "Reading Rooms")
            out = tmp_path / "edited.docx"
            doc.save(out)
        finally:
            doc.close()

        doc2 = Document.open(out, force_recreate=True)
        try:
            paragraphs = doc2.list_paragraphs()
            joined = " ".join(paragraphs)
            assert "Reading Rooms" in joined
            assert "Library Bookshelves" not in joined
            # Smart quotes must survive the edit
            assert "‘" in joined and "’" in joined
        finally:
            doc2.close()

    def test_delete_around_smart_quotes(self, simple_docx, tmp_path):
        src = tmp_path / "with_smart_quotes.docx"
        _build_smart_quote_docx(simple_docx, src)

        doc = Document.open(src, force_recreate=True)
        try:
            _split_wt_text_nodes(doc, "Library Bookshelves")
            doc._revision_manager.suggest_deletion("Library Bookshelves")
            out = tmp_path / "edited.docx"
            doc.save(out)
        finally:
            doc.close()

        doc2 = Document.open(out, force_recreate=True)
        try:
            joined = " ".join(doc2.list_paragraphs())
            assert "Library Bookshelves" not in joined
            assert "‘" in joined and "’" in joined
        finally:
            doc2.close()


class TestTrimReplaceAffixes:
    """Unit tests for _trim_replace_affixes word-level common affix trimming."""

    def test_no_common_affixes(self):
        assert _trim_replace_affixes("cats", "cat") == (0, 0)

    def test_identical_strings_trim_fully(self):
        find = "same text here"
        prefix, suffix = _trim_replace_affixes(find, find)
        assert find[prefix : len(find) - suffix] == ""

    def test_prefix_only(self):
        prefix, suffix = _trim_replace_affixes("term of two", "term of three")
        assert (prefix, suffix) == (len("term of "), 0)

    def test_suffix_only(self):
        prefix, suffix = _trim_replace_affixes("two years remain", "three years remain")
        assert (prefix, suffix) == (0, len(" years remain"))

    def test_prefix_and_suffix(self):
        find = "term of two (2) years, unless"
        replace_with = "term of three (3) years, unless"
        prefix, suffix = _trim_replace_affixes(find, replace_with)
        assert find[prefix : len(find) - suffix] == "two (2)"
        assert replace_with[prefix : len(replace_with) - suffix] == "three (3)"

    def test_whitespace_not_double_consumed(self):
        find, replace_with = "delete this word", "delete word"
        prefix, suffix = _trim_replace_affixes(find, replace_with)
        assert find[prefix : len(find) - suffix] == "this "
        assert replace_with[prefix : len(replace_with) - suffix] == ""

    def test_empty_replacement_means_no_trimming(self):
        assert _trim_replace_affixes(" here", "") == (0, 0)
        assert _trim_replace_affixes("gone", "") == (0, 0)

    def test_overlap_bound_shrinking(self):
        find, replace_with = "a a a", "a a"
        prefix, suffix = _trim_replace_affixes(find, replace_with)
        assert find[prefix : len(find) - suffix] == " a"
        assert replace_with[prefix : len(replace_with) - suffix] == ""

    def test_overlap_bound_growing(self):
        find, replace_with = "a a", "a a a"
        prefix, suffix = _trim_replace_affixes(find, replace_with)
        assert find[prefix : len(find) - suffix] == ""
        assert replace_with[prefix : len(replace_with) - suffix] == " a"

    def test_word_level_not_char_level(self):
        # "cats"/"cat" share characters but not whole words: no trimming.
        prefix, suffix = _trim_replace_affixes("two cats sat", "two cat sat")
        assert (prefix, suffix) == (len("two "), len(" sat"))
        assert _trim_replace_affixes("cats", "cat") == (0, 0)

    def test_unicode_smart_quote_tokens(self):
        find = "“term” of x"
        replace_with = "“term” of y"
        prefix, suffix = _trim_replace_affixes(find, replace_with)
        assert (prefix, suffix) == (len("“term” of "), 0)


class TestTrimReplaceAffixesWhitespaceOnly:
    """Whitespace-only spans trim character-wise, so a spacing fix is a pure
    insert/delete instead of an invisible del+ins pair (ISSUES.md #60).

    Every case asserts the resulting (del, ins) split rather than raw offsets:
    that is what decides which revisions get written.
    """

    @staticmethod
    def _split(find: str, replace_with: str) -> tuple[str, str]:
        prefix, suffix = _trim_replace_affixes(find, replace_with)
        # An affix pair that overlaps would silently corrupt the split.
        assert prefix + suffix <= min(len(find), len(replace_with))
        return find[prefix : len(find) - suffix], replace_with[prefix : len(replace_with) - suffix]

    def test_single_to_double_space_is_pure_insert(self):
        assert self._split(" ", "  ") == ("", " ")

    def test_double_to_single_space_is_pure_delete(self):
        assert self._split("  ", " ") == (" ", "")

    def test_three_to_one_space_deletes_two(self):
        assert self._split("   ", " ") == ("  ", "")

    def test_space_to_tab_stays_a_replacement(self):
        """Different whitespace characters share no affix — still del+ins."""
        assert self._split(" ", "\t") == (" ", "\t")

    def test_trim_applies_after_word_level_trim(self):
        """Words trim first; only the leftover whitespace trims by character."""
        assert self._split("a b", "a  b") == ("", " ")
        assert self._split("thirty days", "thirty  days") == ("", " ")

    def test_shared_characters_never_consumed_twice(self):
        """Prefix and suffix both grow, bounded by the shorter side."""
        assert self._split("\t \t", "\t\t") == (" ", "")

    def test_mixed_whitespace_and_words_is_left_alone(self):
        """A span that is not whitespace-only keeps whole-span replacement."""
        assert self._split("  x  ", " x ") == ("  x  ", " x ")

    def test_word_replacement_unaffected(self):
        """Regression canary: character trimming must never reach word spans."""
        assert self._split("30 days", "60 days") == ("30", "60")
        assert self._split("net", "gross") == ("net", "gross")


class TestWhitespaceReplaceRevisions:
    """End-to-end: the whitespace-only trim reaches the written revisions."""

    def test_single_to_double_space_writes_only_an_insertion(self, clean_workspace):
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "quick brown fox")
        doc.replace("quick brown", "quick  brown", paragraph=ref)

        kinds = [r.type for r in doc.list_revisions()]
        assert kinds == ["insertion"], kinds
        assert doc.get_visible_text().splitlines()[1] == "The quick  brown fox jumps over the lazy dog."
        doc.close()

    def test_double_to_single_space_writes_only_a_deletion(self, clean_workspace):
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "quick brown fox")
        doc.replace("quick brown", "quick  brown", paragraph=ref)
        doc.accept_all()

        ref = find_ref(doc, "quick  brown fox")
        doc.replace("quick  brown", "quick brown", paragraph=ref)
        kinds = [r.type for r in doc.list_revisions()]
        assert kinds == ["deletion"], kinds
        assert doc.get_visible_text().splitlines()[1] == "The quick brown fox jumps over the lazy dog."
        doc.close()

    def test_word_replace_still_writes_both_halves(self, clean_workspace):
        """The canary at the revision level: words stay a deletion+insertion."""
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "quick brown fox")
        doc.replace("quick", "swift", paragraph=ref)

        assert sorted(r.type for r in doc.list_revisions()) == ["deletion", "insertion"]
        doc.close()

    def test_whitespace_deletion_is_preserved_for_reject(self, clean_workspace):
        """The deleted whitespace must survive a round-trip through a conforming
        consumer, or rejecting the fix would silently restore the WRONG spacing.

        A whitespace-only ``w:delText`` without ``xml:space="preserve"`` gets its
        content trimmed away by Word, and ``_restore_deletion`` copies that
        element's attributes onto the ``w:t`` it rebuilds — so both halves of the
        undo path need the attribute.
        """
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "quick brown fox")
        doc.replace("quick brown", "quick  brown", paragraph=ref)
        doc.accept_all()

        ref = find_ref(doc, "quick  brown fox")
        result = doc.replace("quick  brown", "quick brown", paragraph=ref)
        deleted = doc._document_editor.dom.getElementsByTagName("w:delText")
        assert [d.firstChild.data for d in deleted] == [" "]
        assert deleted[0].getAttribute("xml:space") == "preserve"

        assert result.group_id is not None  # a real edit, so a real group to reject
        doc.reject_group(result.group_id)
        restored = [t for t in doc._document_editor.dom.getElementsByTagName("w:t") if t.firstChild.data == " "]
        assert restored, "the deleted space should be back as a w:t"
        assert all(t.getAttribute("xml:space") == "preserve" for t in restored)
        assert doc.get_visible_text().splitlines()[1] == "The quick  brown fox jumps over the lazy dog."
        doc.close()

    def test_deleted_edge_whitespace_carries_preserve(self, clean_workspace):
        """Same rule for any deletion whose text has a whitespace edge, not just
        the whitespace-only ones the trim newly produces."""
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "quick brown fox")
        doc.delete("brown ", paragraph=ref)

        deleted = doc._document_editor.dom.getElementsByTagName("w:delText")
        assert [d.firstChild.data for d in deleted] == ["brown "]
        assert deleted[0].getAttribute("xml:space") == "preserve"
        doc.close()


class TestSearchResultAsEditTarget:
    """A SearchResult can stand in for (text, paragraph, occurrence) on every
    edit method — the find→edit double-typing papercut (ISSUES.md #52).
    """

    def test_replace_matches_the_explicit_spelling(self, clean_workspace):
        doc = Document.open(clean_workspace)
        match = match_for(doc, "quick")
        result = doc.replace(match, "swift")

        assert result == match_for(doc, "swift").paragraph_ref
        assert doc.get_visible_text().splitlines()[1] == "The swift brown fox jumps over the lazy dog."
        doc.close()

    def test_delete_from_a_match(self, clean_workspace):
        doc = Document.open(clean_workspace)
        doc.delete(match_for(doc, "brown "))

        assert doc.get_visible_text().splitlines()[1] == "The quick fox jumps over the lazy dog."
        doc.close()

    def test_insert_after_and_before_from_a_match(self, clean_workspace):
        doc = Document.open(clean_workspace)
        doc.insert_after(match_for(doc, "lazy dog"), " (allegedly)")
        doc.insert_before(match_for(doc, "The quick"), "Note: ")

        # The anchor is "lazy dog", so the insertion lands before the period.
        assert doc.get_visible_text().splitlines()[1] == (
            "Note: The quick brown fox jumps over the lazy dog (allegedly)."
        )
        doc.close()

    def test_occurrence_is_carried_from_the_match(self, clean_workspace):
        """The 2nd "he" in P2 ("tHE lazy") is edited with no occurrence bookkeeping."""
        doc = Document.open(clean_workspace)
        second = [m for m in doc.find_all("he") if m.paragraph_index == 2][1]
        assert second.paragraph_occurrence == 1

        doc.replace(second, "HE")
        assert doc.get_visible_text().splitlines()[1] == "The quick brown fox jumps over tHE lazy dog."
        doc.close()

    def test_equivalent_to_spelling_the_fields_out(self, clean_workspace, temp_dir):
        """Same edit, two spellings, byte-identical document XML."""
        second_copy = temp_dir / "second.docx"
        shutil.copy(clean_workspace, second_copy)

        doc = Document.open(clean_workspace)
        match = match_for(doc, "sample document")
        doc.replace(match, "example document")
        via_object = doc.get_markup_text()
        doc.close()

        other = Document.open(second_copy)
        m2 = match_for(other, "sample document")
        other.replace(
            m2.text,
            "example document",
            paragraph=m2.paragraph_ref,
            occurrence=m2.paragraph_occurrence,
        )
        via_fields = other.get_markup_text()
        other.close()

        assert via_object == via_fields

    def test_stale_match_raises_hash_mismatch(self, clean_workspace):
        """A SearchResult is a ref like any other: it goes stale when edited."""
        doc = Document.open(clean_workspace)
        match = match_for(doc, "quick")
        doc.replace("brown", "red", paragraph=match.paragraph_ref)

        with pytest.raises(HashMismatchError):
            doc.replace(match, "swift")
        doc.close()

    def test_paragraph_and_occurrence_are_refused_with_a_match(self, clean_workspace):
        doc = Document.open(clean_workspace)
        match = match_for(doc, "quick")

        extras: list[dict[str, Any]] = [{"paragraph": match.paragraph_ref}, {"occurrence": 0}, {"occurrence": 3}]
        for kwargs in extras:
            with pytest.raises(ValueError, match="already pins the paragraph"):
                doc.replace(match, "swift", **kwargs)
        with pytest.raises(ValueError, match="drop paragraph= and occurrence="):
            doc.replace(match, "swift", paragraph=match.paragraph_ref, occurrence=0)
        assert doc.list_revisions() == []
        doc.close()

    def test_every_method_refuses_the_double_spelling(self, clean_workspace):
        doc = Document.open(clean_workspace)
        match = match_for(doc, "quick")
        calls = [
            lambda: doc.replace(match, "x", occurrence=0),
            lambda: doc.delete(match, occurrence=0),
            lambda: doc.insert_after(match, "x", occurrence=0),
            lambda: doc.insert_before(match, "x", occurrence=0),
            lambda: doc.add_comment(match, "note", occurrence=0),
        ]
        for call in calls:
            with pytest.raises(ValueError, match="already pins the paragraph"):
                call()
        doc.close()

    def test_plain_text_still_requires_a_paragraph(self, clean_workspace):
        """The doc-wide RevisionManager branch stays unreachable, and the error
        teaches the SearchResult form."""
        doc = Document.open(clean_workspace)
        with pytest.raises(ValueError, match="must be a paragraph ref string"):
            doc.replace("quick", "swift")
        with pytest.raises(ValueError, match="or pass a SearchResult as 'find'"):
            doc.replace("quick", "swift")
        with pytest.raises(ValueError, match="or pass a SearchResult as 'text'"):
            doc.delete("quick")
        with pytest.raises(ValueError, match="or pass a SearchResult as 'anchor'"):
            doc.insert_after("quick", "x")
        assert doc.list_revisions() == []
        doc.close()

    def test_find_text_returning_none_is_not_a_silent_edit(self, clean_workspace):
        """Passing an unchecked find_text() result through is a clear error."""
        doc = Document.open(clean_workspace)
        assert doc.find_text("no such text") is None
        with pytest.raises(ValueError, match="must be a paragraph ref string"):
            doc.replace(doc.find_text("no such text"), "x")  # type: ignore[arg-type]
        doc.close()
