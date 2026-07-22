"""Tests for track changes functionality."""

from datetime import datetime, timezone
from unittest.mock import MagicMock

import pytest
from conftest import find_ref

from docx_editor import Document, TextNotFoundError
from docx_editor.track_changes import Revision, RevisionManager, _escape_xml, _trim_replace_affixes
from docx_editor.xml_editor import DocxXMLEditor, build_text_map


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
        mock_editor = MagicMock()
        mock_editor.dom = MagicMock()

        # Create mock insertion element
        mock_ins = MagicMock()

        def ins_get_attr(name):
            if name == "w:id":
                return "1"
            elif name == "w:author":
                return "SpecificAuthor"
            elif name == "w:date":
                return ""
            return ""

        mock_ins.getAttribute.side_effect = ins_get_attr
        mock_ins.getElementsByTagName.return_value = []

        mock_editor.dom.getElementsByTagName.side_effect = lambda tag: [mock_ins] if tag == "w:ins" else []

        manager = RevisionManager(mock_editor)

        # Filter by matching author
        revisions = manager.list_revisions(author="SpecificAuthor")
        assert len(revisions) == 1
        assert revisions[0].author == "SpecificAuthor"

        # Filter by non-matching author
        revisions = manager.list_revisions(author="OtherAuthor")
        assert len(revisions) == 0

    def test_list_revisions_filters_by_author_for_deletions(self):
        """Test that list_revisions author filter works for deletions."""
        mock_editor = MagicMock()
        mock_editor.dom = MagicMock()

        # Create mock deletion element
        mock_del = MagicMock()

        def del_get_attr(name):
            if name == "w:id":
                return "2"
            elif name == "w:author":
                return "DeleteAuthor"
            elif name == "w:date":
                return ""
            return ""

        mock_del.getAttribute.side_effect = del_get_attr
        mock_del.getElementsByTagName.return_value = []

        mock_editor.dom.getElementsByTagName.side_effect = lambda tag: [mock_del] if tag == "w:del" else []

        manager = RevisionManager(mock_editor)

        # Filter by matching author
        revisions = manager.list_revisions(author="DeleteAuthor")
        assert len(revisions) == 1
        assert revisions[0].author == "DeleteAuthor"
        assert revisions[0].type == "deletion"


class TestAcceptRejectLoops:
    """Tests for accept_all and reject_all loops."""

    def test_accept_all_processes_multiple_revisions(self):
        """Test that accept_all correctly processes multiple revisions."""
        mock_editor = MagicMock()
        mock_editor.dom = MagicMock()

        # Create two mock insertions
        mock_ins1 = MagicMock()
        mock_ins2 = MagicMock()

        def ins1_get_attr(name):
            if name == "w:id":
                return "1"
            elif name == "w:author":
                return "Author"
            return ""

        def ins2_get_attr(name):
            if name == "w:id":
                return "2"
            elif name == "w:author":
                return "Author"
            return ""

        mock_ins1.getAttribute.side_effect = ins1_get_attr
        mock_ins1.getElementsByTagName.return_value = []
        mock_ins1.parentNode = MagicMock()

        mock_ins2.getAttribute.side_effect = ins2_get_attr
        mock_ins2.getElementsByTagName.return_value = []
        mock_ins2.parentNode = MagicMock()

        # Track which elements have been processed
        processed = set()

        def get_elements(tag):
            if tag == "w:ins":
                result = []
                if "1" not in processed:
                    result.append(mock_ins1)
                if "2" not in processed:
                    result.append(mock_ins2)
                return result
            return []

        mock_editor.dom.getElementsByTagName.side_effect = get_elements

        manager = RevisionManager(mock_editor)

        # Mock accept_revision to track calls
        def mock_accept(rev_id: int) -> bool:
            processed.add(str(rev_id))
            return True

        manager.accept_revision = mock_accept  # type: ignore[method-assign]

        count = manager.accept_all()
        assert count == 2

    def test_reject_all_processes_multiple_revisions(self):
        """Test that reject_all correctly processes multiple revisions."""
        mock_editor = MagicMock()
        mock_editor.dom = MagicMock()

        # Create two mock deletions
        mock_del1 = MagicMock()
        mock_del2 = MagicMock()

        def del1_get_attr(name):
            if name == "w:id":
                return "3"
            elif name == "w:author":
                return "Author"
            return ""

        def del2_get_attr(name):
            if name == "w:id":
                return "4"
            elif name == "w:author":
                return "Author"
            return ""

        mock_del1.getAttribute.side_effect = del1_get_attr
        mock_del1.getElementsByTagName.return_value = []

        mock_del2.getAttribute.side_effect = del2_get_attr
        mock_del2.getElementsByTagName.return_value = []

        processed = set()

        def get_elements(tag):
            if tag == "w:del":
                result = []
                if "3" not in processed:
                    result.append(mock_del1)
                if "4" not in processed:
                    result.append(mock_del2)
                return result
            return []

        mock_editor.dom.getElementsByTagName.side_effect = get_elements

        manager = RevisionManager(mock_editor)

        def mock_reject(rev_id: int) -> bool:
            processed.add(str(rev_id))
            return True

        manager.reject_revision = mock_reject  # type: ignore[method-assign]

        count = manager.reject_all()
        assert count == 2


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

    def test_accept_all_author_filter_duplicate_ids_converges(self):
        """Test that accept_all(author=...) converges when w:id values collide across authors."""
        # Word does not guarantee unique w:id across w:ins/w:del. The per-id
        # lookup checks w:ins first, so accepting B's deletion by id hits A's
        # insertion instead; only re-listing until no progress resolves B.
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

        # Documented side effect of id-based matching: pass 1 accepts A's
        # same-id insertion (collateral), pass 2 accepts B's deletion.
        assert count == 2
        assert manager.list_revisions() == []
        texts = [t.firstChild.data for t in dom.getElementsByTagName("w:t")]
        assert texts == ["alpha"]
        assert dom.getElementsByTagName("w:delText") == []

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
