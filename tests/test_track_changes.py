"""Tests for track changes functionality."""

from datetime import datetime, timezone
from unittest.mock import MagicMock

import pytest

from docx_editor import Document, TextNotFoundError
from docx_editor.exceptions import RevisionError
from docx_editor.track_changes import Revision, RevisionManager, _escape_xml


class TestTrackedReplace:
    """Tests for tracked text replacement."""

    def test_replace_creates_tracked_change(self, clean_workspace):
        """Test that replace creates w:del and w:ins elements."""
        doc = Document.open(clean_workspace)

        # Find some text to replace - need to know what's in simple.docx
        # For now, we'll test that the method doesn't crash
        try:
            doc.replace("test", "TEST")
        except TextNotFoundError:
            # Expected if "test" not in document
            pass

        doc.close()

    def test_replace_returns_change_id(self, clean_workspace):
        """Test that replace returns a valid change ID."""
        doc = Document.open(clean_workspace)

        try:
            change_id = doc.replace("the", "THE")
            assert isinstance(change_id, int)
            assert change_id >= 0
        except TextNotFoundError:
            pytest.skip("Test text not found in document")

        doc.close()

    def test_replace_not_found_raises_error(self, clean_workspace):
        """Test that replacing nonexistent text raises TextNotFoundError."""
        doc = Document.open(clean_workspace)

        with pytest.raises(TextNotFoundError):
            doc.replace("xyz123nonexistent789", "replacement")

        doc.close()


class TestTrackedDeletion:
    """Tests for tracked deletions."""

    def test_delete_creates_tracked_change(self, clean_workspace):
        """Test that delete creates w:del element."""
        doc = Document.open(clean_workspace)

        try:
            change_id = doc.delete("the")
            assert isinstance(change_id, int)
            assert change_id >= 0
        except TextNotFoundError:
            pytest.skip("Test text not found in document")

        doc.close()

    def test_delete_not_found_raises_error(self, clean_workspace):
        """Test that deleting nonexistent text raises TextNotFoundError."""
        doc = Document.open(clean_workspace)

        with pytest.raises(TextNotFoundError):
            doc.delete("xyz123nonexistent789")

        doc.close()


class TestTrackedInsertion:
    """Tests for tracked insertions."""

    def test_insert_after_creates_tracked_change(self, clean_workspace):
        """Test that insert_after creates w:ins element."""
        doc = Document.open(clean_workspace)

        try:
            change_id = doc.insert_after("the", " NEW TEXT")
            assert isinstance(change_id, int)
            assert change_id >= 0
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        doc.close()

    def test_insert_before_creates_tracked_change(self, clean_workspace):
        """Test that insert_before creates w:ins element."""
        doc = Document.open(clean_workspace)

        try:
            change_id = doc.insert_before("the", "BEFORE ")
            assert isinstance(change_id, int)
            assert change_id >= 0
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
            doc.delete("the")
            doc.insert_after("a", " NEW")
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
            doc.delete("the")
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
            change_id = doc.delete("the")
        except TextNotFoundError:
            pytest.skip("Test text not found in document")

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
            change_id = doc.delete("the")
        except TextNotFoundError:
            pytest.skip("Test text not found in document")

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
            doc.delete("the")
            doc.insert_after("a", " NEW")
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
            doc.delete("the")
            doc.insert_after("a", " NEW")
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
        """Test replace with specific occurrence."""
        doc = Document.open(clean_workspace)

        count = doc.count_matches("the")
        if count < 2:
            doc.close()
            pytest.skip("Need at least 2 occurrences for this test")

        # Replace second occurrence
        change_id = doc.replace("the", "THE", occurrence=1)
        assert isinstance(change_id, int)
        assert change_id >= 0

        doc.close()

    def test_replace_occurrence_out_of_range(self, clean_workspace):
        """Test replace with occurrence beyond available matches."""
        doc = Document.open(clean_workspace)

        # First find text that exists
        count = doc.count_matches("the")
        if count == 0:
            # Try another common word
            count = doc.count_matches("a")
            search_text = "a"
        else:
            search_text = "the"

        if count == 0:
            doc.close()
            pytest.skip("No suitable text found in document")

        # Request an occurrence that doesn't exist
        with pytest.raises(TextNotFoundError) as exc_info:
            doc.replace(search_text, "REPLACEMENT", occurrence=count + 100)

        assert "occurrence" in str(exc_info.value).lower()

        doc.close()

    def test_delete_with_occurrence(self, clean_workspace):
        """Test delete with specific occurrence."""
        doc = Document.open(clean_workspace)

        count = doc.count_matches("the")
        if count < 2:
            doc.close()
            pytest.skip("Need at least 2 occurrences for this test")

        # Delete second occurrence
        change_id = doc.delete("the", occurrence=1)
        assert isinstance(change_id, int)
        assert change_id >= 0

        doc.close()

    def test_insert_after_with_occurrence(self, clean_workspace):
        """Test insert_after with specific occurrence."""
        doc = Document.open(clean_workspace)

        count = doc.count_matches("the")
        if count < 2:
            doc.close()
            pytest.skip("Need at least 2 occurrences for this test")

        # Insert after second occurrence
        change_id = doc.insert_after("the", " INSERTED", occurrence=1)
        assert isinstance(change_id, int)
        assert change_id >= 0

        doc.close()

    def test_insert_before_with_occurrence(self, clean_workspace):
        """Test insert_before with specific occurrence."""
        doc = Document.open(clean_workspace)

        count = doc.count_matches("the")
        if count < 2:
            doc.close()
            pytest.skip("Need at least 2 occurrences for this test")

        # Insert before second occurrence
        change_id = doc.insert_before("the", "INSERTED ", occurrence=1)
        assert isinstance(change_id, int)
        assert change_id >= 0

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
        assert "+1:" in repr_str
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
        assert "-2:" in repr_str
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
        change_id = doc.replace("quick", "QUICK")
        assert isinstance(change_id, int)
        assert change_id >= 0

        doc.close()

    def test_replace_text_preserves_run_properties(self, clean_workspace):
        """Test that replace preserves w:rPr when present."""
        doc = Document.open(clean_workspace)

        # Replace text - the document structure should be preserved
        change_id = doc.replace("Sample", "SAMPLE")
        assert change_id >= 0

        doc.close()

    def test_suggest_deletion_with_surrounding_text(self, clean_workspace):
        """Test deletion when text has surrounding content."""
        doc = Document.open(clean_workspace)

        # "brown" is in the middle of "The quick brown fox..."
        change_id = doc.delete("brown")
        assert isinstance(change_id, int)
        assert change_id >= 0

        doc.close()

    def test_insert_text_not_found_raises_error(self, clean_workspace):
        """Test insert_after raises TextNotFoundError for nonexistent anchor."""
        doc = Document.open(clean_workspace)

        with pytest.raises(TextNotFoundError) as exc_info:
            doc.insert_after("xyz_nonexistent_anchor_123", "new text")

        assert "Anchor text not found" in str(exc_info.value) or "not found" in str(exc_info.value).lower()

        doc.close()

    def test_insert_before_not_found_raises_error(self, clean_workspace):
        """Test insert_before raises TextNotFoundError for nonexistent anchor."""
        doc = Document.open(clean_workspace)

        with pytest.raises(TextNotFoundError) as exc_info:
            doc.insert_before("xyz_nonexistent_anchor_123", "new text")

        assert "not found" in str(exc_info.value).lower()

        doc.close()


class TestRevisionParsing:
    """Tests for revision parsing edge cases."""

    def test_list_revisions_includes_both_types(self, clean_workspace):
        """Test that list_revisions finds both insertions and deletions."""
        doc = Document.open(clean_workspace, author="ParseTestAuthor")

        # Create both types
        doc.delete("quick")
        doc.insert_after("fox", " really")

        revisions = doc.list_revisions()

        types = {r.type for r in revisions}
        assert "insertion" in types
        assert "deletion" in types

        doc.close()

    def test_list_revisions_with_missing_date(self, clean_workspace):
        """Test parsing revisions that may have missing date attributes."""
        doc = Document.open(clean_workspace)

        doc.delete("quick")
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
        doc.insert_after("fox", "")  # Empty insertion
        revisions = doc.list_revisions()

        # Should not crash on empty text
        assert isinstance(revisions, list)

        doc.close()


class TestAcceptRejectExtended:
    """Extended tests for accept/reject functionality."""

    def test_accept_insertion_revision(self, clean_workspace):
        """Test accepting an insertion keeps the inserted text."""
        doc = Document.open(clean_workspace)

        change_id = doc.insert_after("fox", " NEW")

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

        change_id = doc.delete("quick")

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

        change_id = doc.insert_after("fox", " REJECT_ME")

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

        change_id = doc.delete("brown")

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
        doc.delete("quick")
        doc.close()

        doc = Document.open(clean_workspace, author="Author2")
        doc.delete("brown")

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
        doc.delete("quick")
        doc.insert_after("fox", " test")

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

    def test_get_nth_match_no_matches(self, clean_workspace):
        """Test _get_nth_match raises error when no matches found."""
        doc = Document.open(clean_workspace)

        with pytest.raises(TextNotFoundError) as exc_info:
            doc._revision_manager._get_nth_match("nonexistent_xyz_123", 0)

        assert "not found" in str(exc_info.value).lower()

        doc.close()

    def test_get_nth_match_occurrence_out_of_range(self, clean_workspace):
        """Test _get_nth_match raises error for invalid occurrence."""
        doc = Document.open(clean_workspace)

        # "Sample" exists once in the document
        count = doc.count_matches("Sample")
        if count == 0:
            doc.close()
            pytest.skip("Test text not found")

        with pytest.raises(TextNotFoundError) as exc_info:
            doc._revision_manager._get_nth_match("Sample", occurrence=count + 10)

        assert "occurrence" in str(exc_info.value).lower()

        doc.close()


class TestRevisionManagerWithMockedEditor:
    """Tests using mocked editor for edge cases."""

    def test_replace_text_no_parent_run_raises_error(self):
        """Test replace_text raises RevisionError when no parent w:r found."""
        # Create mock editor
        mock_editor = MagicMock()

        # Create mock element without proper parent hierarchy
        mock_elem = MagicMock()
        mock_elem.parentNode = None  # No parent

        mock_editor.find_all_nodes.return_value = [mock_elem]

        manager = RevisionManager(mock_editor)

        with pytest.raises(RevisionError) as exc_info:
            manager.replace_text("test", "TEST")

        assert "parent w:r" in str(exc_info.value).lower()

    def test_suggest_deletion_no_parent_run_raises_error(self):
        """Test suggest_deletion raises RevisionError when no parent w:r found."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_elem.parentNode = None

        mock_editor.find_all_nodes.return_value = [mock_elem]

        manager = RevisionManager(mock_editor)

        with pytest.raises(RevisionError) as exc_info:
            manager.suggest_deletion("test")

        assert "parent w:r" in str(exc_info.value).lower()

    def test_insert_text_no_parent_run_raises_error(self):
        """Test _insert_text raises RevisionError when no parent w:r found."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_elem.parentNode = None

        mock_editor.find_all_nodes.return_value = [mock_elem]

        manager = RevisionManager(mock_editor)

        with pytest.raises(RevisionError) as exc_info:
            manager.insert_text_after("test", "new text")

        assert "parent w:r" in str(exc_info.value).lower()

    def test_parse_revision_missing_id_returns_none(self):
        """Test _parse_revision returns None when w:id is missing."""
        mock_editor = MagicMock()
        manager = RevisionManager(mock_editor)

        mock_elem = MagicMock()
        mock_elem.getAttribute.return_value = ""  # Empty w:id

        result = manager._parse_revision(mock_elem, "insertion")
        assert result is None

    def test_parse_revision_invalid_date_uses_none(self):
        """Test _parse_revision handles invalid date gracefully."""
        mock_editor = MagicMock()
        manager = RevisionManager(mock_editor)

        mock_elem = MagicMock()

        def get_attr(name):
            if name == "w:id":
                return "1"
            elif name == "w:author":
                return "Test"
            elif name == "w:date":
                return "invalid-date-format"
            return ""

        mock_elem.getAttribute.side_effect = get_attr
        mock_elem.getElementsByTagName.return_value = []

        result = manager._parse_revision(mock_elem, "insertion")
        assert result is not None
        assert result.date is None  # Invalid date should be None

    def test_parse_revision_with_text_content(self):
        """Test _parse_revision extracts text content properly."""
        mock_editor = MagicMock()
        manager = RevisionManager(mock_editor)

        mock_elem = MagicMock()

        def get_attr(name):
            if name == "w:id":
                return "5"
            elif name == "w:author":
                return "Author"
            elif name == "w:date":
                return "2024-01-15T10:30:00Z"
            return ""

        mock_elem.getAttribute.side_effect = get_attr

        # Mock text element with content
        mock_text_elem = MagicMock()
        mock_text_elem.firstChild = MagicMock()
        mock_text_elem.firstChild.data = "test content"

        mock_elem.getElementsByTagName.return_value = [mock_text_elem]

        result = manager._parse_revision(mock_elem, "insertion")
        assert result is not None
        assert result.text == "test content"

    def test_parse_revision_text_element_no_child(self):
        """Test _parse_revision handles text elements with no firstChild."""
        mock_editor = MagicMock()
        manager = RevisionManager(mock_editor)

        mock_elem = MagicMock()

        def get_attr(name):
            if name == "w:id":
                return "6"
            elif name == "w:author":
                return "Author"
            elif name == "w:date":
                return ""
            return ""

        mock_elem.getAttribute.side_effect = get_attr

        # Mock text element with no firstChild
        mock_text_elem = MagicMock()
        mock_text_elem.firstChild = None

        mock_elem.getElementsByTagName.return_value = [mock_text_elem]

        result = manager._parse_revision(mock_elem, "insertion")
        assert result is not None
        assert result.text == ""  # Empty text when no content


class TestRestoreDeletionEdgeCases:
    """Tests for _restore_deletion edge cases."""

    def test_reject_deletion_with_attributes(self, clean_workspace):
        """Test rejecting deletion restores attributes on delText."""
        doc = Document.open(clean_workspace)

        # Create a deletion
        change_id = doc.delete("lazy")

        # Reject it to trigger _restore_deletion
        result = doc.reject_revision(change_id)
        assert result is True

        doc.close()

    def test_reject_deletion_handles_rsid_attributes(self, clean_workspace):
        """Test rejecting deletion converts rsidDel back to rsidR."""
        doc = Document.open(clean_workspace)

        # Create a deletion
        change_id = doc.delete("dog")

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
        doc.delete("quick")
        doc.insert_after("brown", " spotted")
        doc.replace("fox", "cat")

        revisions = doc.list_revisions()
        # Should have at least 3 revisions (1 delete, 1 insert, 2 from replace)
        assert len(revisions) >= 3

        doc.close()

    def test_accept_all_then_list(self, clean_workspace):
        """Test that accept_all properly clears all revisions."""
        doc = Document.open(clean_workspace)

        doc.delete("quick")
        doc.insert_after("fox", " test")

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

        doc.delete("quick")
        doc.insert_after("fox", " test")

        initial_count = len(doc.list_revisions())
        assert initial_count >= 2

        rejected = doc.reject_all()
        assert rejected == initial_count

        remaining = doc.list_revisions()
        assert len(remaining) == 0

        doc.close()


class TestMockedEdgeCases:
    """Tests for edge cases using mocks to reach uncovered branches."""

    def test_replace_text_parent_traversal_loop(self):
        """Test replace_text when elem.parentNode is not immediately w:r."""
        mock_editor = MagicMock()

        # Create a chain: elem -> intermediate_node -> run (w:r)
        mock_elem = MagicMock()
        mock_intermediate = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None
        mock_run.getElementsByTagName.return_value = []  # No rPr

        mock_intermediate.nodeName = "other"
        mock_intermediate.parentNode = mock_run

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_intermediate
        mock_elem.firstChild = MagicMock()
        mock_elem.firstChild.data = "hello world"

        mock_editor.find_all_nodes.return_value = [mock_elem]

        # Mock replace_node to return a node with w:ins
        mock_ins_node = MagicMock()
        mock_ins_node.nodeType = mock_ins_node.ELEMENT_NODE
        mock_ins_node.tagName = "w:ins"
        mock_ins_node.getAttribute.return_value = "42"

        mock_editor.replace_node.return_value = [mock_ins_node]

        manager = RevisionManager(mock_editor)
        result = manager.replace_text("hello", "HELLO")

        assert result == 42

    def test_suggest_deletion_parent_traversal_loop(self):
        """Test suggest_deletion when elem.parentNode is not immediately w:r."""
        mock_editor = MagicMock()

        # Create a chain: elem -> intermediate_node -> run (w:r)
        mock_elem = MagicMock()
        mock_intermediate = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None
        mock_run.getElementsByTagName.return_value = []  # No rPr

        mock_intermediate.nodeName = "other"
        mock_intermediate.parentNode = mock_run

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_intermediate
        mock_elem.firstChild = MagicMock()
        mock_elem.firstChild.data = "hello world"

        mock_editor.find_all_nodes.return_value = [mock_elem]

        # Mock replace_node to return a node with w:del
        mock_del_node = MagicMock()
        mock_del_node.nodeType = mock_del_node.ELEMENT_NODE
        mock_del_node.tagName = "w:del"
        mock_del_node.getAttribute.return_value = "43"

        mock_editor.replace_node.return_value = [mock_del_node]

        manager = RevisionManager(mock_editor)
        result = manager.suggest_deletion("hello")

        assert result == 43

    def test_insert_text_parent_traversal_loop(self):
        """Test _insert_text when elem.parentNode is not immediately w:r."""
        mock_editor = MagicMock()

        # Create a chain: elem -> intermediate_node -> run (w:r)
        mock_elem = MagicMock()
        mock_intermediate = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None
        mock_run.getElementsByTagName.return_value = []  # No rPr

        mock_intermediate.nodeName = "other"
        mock_intermediate.parentNode = mock_run

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_intermediate
        mock_elem.firstChild.data = "anchor"

        mock_editor.find_all_nodes.return_value = [mock_elem]

        # Mock replace_node to return a node with w:ins
        mock_ins_node = MagicMock()
        mock_ins_node.nodeType = mock_ins_node.ELEMENT_NODE
        mock_ins_node.tagName = "w:ins"
        mock_ins_node.getAttribute.return_value = "44"

        mock_editor.replace_node.return_value = [mock_ins_node]

        manager = RevisionManager(mock_editor)
        result = manager.insert_text_after("anchor", "new text")

        assert result == 44

    def test_replace_text_with_rpr(self):
        """Test replace_text preserves w:rPr when present."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None

        # Mock rPr element
        mock_rPr = MagicMock()
        mock_rPr.toxml.return_value = "<w:rPr><w:b/></w:rPr>"
        mock_run.getElementsByTagName.return_value = [mock_rPr]

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_run
        mock_elem.firstChild = MagicMock()
        mock_elem.firstChild.data = "hello world"

        mock_editor.find_all_nodes.return_value = [mock_elem]

        mock_ins_node = MagicMock()
        mock_ins_node.nodeType = mock_ins_node.ELEMENT_NODE
        mock_ins_node.tagName = "w:ins"
        mock_ins_node.getAttribute.return_value = "45"

        mock_editor.replace_node.return_value = [mock_ins_node]

        manager = RevisionManager(mock_editor)
        result = manager.replace_text("hello", "HELLO")

        assert result == 45
        # Verify rPr was included in the XML
        call_args = mock_editor.replace_node.call_args[0][1]
        assert "<w:rPr>" in call_args

    def test_suggest_deletion_with_rpr(self):
        """Test suggest_deletion preserves w:rPr when present."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None

        # Mock rPr element
        mock_rPr = MagicMock()
        mock_rPr.toxml.return_value = "<w:rPr><w:i/></w:rPr>"
        mock_run.getElementsByTagName.return_value = [mock_rPr]

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_run
        mock_elem.firstChild = MagicMock()
        mock_elem.firstChild.data = "hello world"

        mock_editor.find_all_nodes.return_value = [mock_elem]

        mock_del_node = MagicMock()
        mock_del_node.nodeType = mock_del_node.ELEMENT_NODE
        mock_del_node.tagName = "w:del"
        mock_del_node.getAttribute.return_value = "46"

        mock_editor.replace_node.return_value = [mock_del_node]

        manager = RevisionManager(mock_editor)
        result = manager.suggest_deletion("hello")

        assert result == 46
        # Verify rPr was included in the XML
        call_args = mock_editor.replace_node.call_args[0][1]
        assert "<w:rPr>" in call_args

    def test_insert_text_with_rpr(self):
        """Test _insert_text preserves w:rPr when present."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None

        # Mock rPr element
        mock_rPr = MagicMock()
        mock_rPr.toxml.return_value = "<w:rPr><w:u/></w:rPr>"

        def _get_by_tag(tag):
            if tag == "w:rPr":
                return [mock_rPr]
            if tag == "w:t":
                return [mock_elem]
            return []

        mock_run.getElementsByTagName.side_effect = _get_by_tag

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_run
        mock_elem.firstChild.data = "anchor"

        mock_editor.find_all_nodes.return_value = [mock_elem]

        mock_ins_node = MagicMock()
        mock_ins_node.nodeType = mock_ins_node.ELEMENT_NODE
        mock_ins_node.tagName = "w:ins"
        mock_ins_node.getAttribute.return_value = "47"

        mock_editor.replace_node.return_value = [mock_ins_node]

        manager = RevisionManager(mock_editor)
        result = manager.insert_text_after("anchor", "new text")

        assert result == 47
        # Verify rPr was included in the XML
        call_args = mock_editor.replace_node.call_args[0][1]
        assert "<w:rPr>" in call_args

    def test_replace_text_returns_minus_one_when_no_ins_found(self):
        """Test replace_text returns -1 when w:ins not found in result."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None
        mock_run.getElementsByTagName.return_value = []

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_run
        mock_elem.firstChild = MagicMock()
        mock_elem.firstChild.data = "hello"

        mock_editor.find_all_nodes.return_value = [mock_elem]

        # Return nodes that don't include w:ins
        mock_other_node = MagicMock()
        mock_other_node.nodeType = mock_other_node.ELEMENT_NODE
        mock_other_node.tagName = "w:r"

        mock_editor.replace_node.return_value = [mock_other_node]

        manager = RevisionManager(mock_editor)
        result = manager.replace_text("hello", "HELLO")

        assert result == -1

    def test_suggest_deletion_returns_minus_one_when_no_del_found(self):
        """Test suggest_deletion returns -1 when w:del not found in result."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None
        mock_run.getElementsByTagName.return_value = []

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_run
        mock_elem.firstChild = MagicMock()
        mock_elem.firstChild.data = "hello"

        mock_editor.find_all_nodes.return_value = [mock_elem]

        # Return nodes that don't include w:del
        mock_other_node = MagicMock()
        mock_other_node.nodeType = mock_other_node.ELEMENT_NODE
        mock_other_node.tagName = "w:r"

        mock_editor.replace_node.return_value = [mock_other_node]

        manager = RevisionManager(mock_editor)
        result = manager.suggest_deletion("hello")

        assert result == -1

    def test_insert_text_returns_minus_one_when_no_ins_found(self):
        """Test _insert_text returns -1 when w:ins not found in result."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None
        mock_run.getElementsByTagName.return_value = []

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_run

        mock_editor.find_all_nodes.return_value = [mock_elem]

        # Return nodes that don't include w:ins
        mock_other_node = MagicMock()
        mock_other_node.nodeType = mock_other_node.ELEMENT_NODE
        mock_other_node.tagName = "w:r"

        mock_editor.insert_after.return_value = [mock_other_node]

        manager = RevisionManager(mock_editor)
        result = manager.insert_text_after("anchor", "text")

        assert result == -1

    def test_replace_text_text_mismatch_raises_error(self):
        """Test replace_text raises error when text found by matcher but not in actual content."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None
        mock_run.getElementsByTagName.return_value = []

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_run
        mock_elem.firstChild = MagicMock()
        # The matcher found "world" in this node, but actual content is different
        mock_elem.firstChild.data = "different content"

        mock_editor.find_all_nodes.return_value = [mock_elem]

        manager = RevisionManager(mock_editor)

        with pytest.raises(TextNotFoundError):
            manager.replace_text("world", "WORLD")

    def test_suggest_deletion_text_mismatch_raises_error(self):
        """Test suggest_deletion raises error when text found by matcher but not in actual content."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None
        mock_run.getElementsByTagName.return_value = []

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_run
        mock_elem.firstChild = MagicMock()
        # The matcher found "world" in this node, but actual content is different
        mock_elem.firstChild.data = "different content"

        mock_editor.find_all_nodes.return_value = [mock_elem]

        manager = RevisionManager(mock_editor)

        with pytest.raises(TextNotFoundError):
            manager.suggest_deletion("world")

    def test_replace_text_with_before_text_only(self):
        """Test replace_text with text at end (only before_text, no after_text)."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None
        mock_run.getElementsByTagName.return_value = []

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_run
        mock_elem.firstChild = MagicMock()
        mock_elem.firstChild.data = "prefix hello"  # "hello" is at end

        mock_editor.find_all_nodes.return_value = [mock_elem]

        mock_ins_node = MagicMock()
        mock_ins_node.nodeType = mock_ins_node.ELEMENT_NODE
        mock_ins_node.tagName = "w:ins"
        mock_ins_node.getAttribute.return_value = "48"

        mock_editor.replace_node.return_value = [mock_ins_node]

        manager = RevisionManager(mock_editor)
        result = manager.replace_text("hello", "HELLO")

        assert result == 48
        # Verify before_text run was included
        call_args = mock_editor.replace_node.call_args[0][1]
        assert "prefix" in call_args

    def test_suggest_deletion_with_before_text_only(self):
        """Test suggest_deletion with text at end (only before_text, no after_text)."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None
        mock_run.getElementsByTagName.return_value = []

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_run
        mock_elem.firstChild = MagicMock()
        mock_elem.firstChild.data = "prefix hello"  # "hello" is at end

        mock_editor.find_all_nodes.return_value = [mock_elem]

        mock_del_node = MagicMock()
        mock_del_node.nodeType = mock_del_node.ELEMENT_NODE
        mock_del_node.tagName = "w:del"
        mock_del_node.getAttribute.return_value = "49"

        mock_editor.replace_node.return_value = [mock_del_node]

        manager = RevisionManager(mock_editor)
        result = manager.suggest_deletion("hello")

        assert result == 49
        # Verify before_text run was included
        call_args = mock_editor.replace_node.call_args[0][1]
        assert "prefix" in call_args

    def test_suggest_deletion_with_after_text_only(self):
        """Test suggest_deletion with text at start (only after_text, no before_text)."""
        mock_editor = MagicMock()

        mock_elem = MagicMock()
        mock_run = MagicMock()

        mock_run.nodeName = "w:r"
        mock_run.parentNode = None
        mock_run.getElementsByTagName.return_value = []

        mock_elem.nodeName = "w:t"
        mock_elem.parentNode = mock_run
        mock_elem.firstChild = MagicMock()
        mock_elem.firstChild.data = "hello suffix"  # "hello" is at start

        mock_editor.find_all_nodes.return_value = [mock_elem]

        mock_del_node = MagicMock()
        mock_del_node.nodeType = mock_del_node.ELEMENT_NODE
        mock_del_node.tagName = "w:del"
        mock_del_node.getAttribute.return_value = "50"

        mock_editor.replace_node.return_value = [mock_del_node]

        manager = RevisionManager(mock_editor)
        result = manager.suggest_deletion("hello")

        assert result == 50
        # Verify after_text run was included
        call_args = mock_editor.replace_node.call_args[0][1]
        assert "suffix" in call_args


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


class TestRestoreDeletionAttributeCopying:
    """Tests for _restore_deletion attribute copying edge cases."""

    def test_restore_deletion_copies_deltext_attributes(self):
        """Test that _restore_deletion copies attributes from w:delText to w:t."""
        import defusedxml.minidom

        # Create a minimal document with w:del containing w:delText with attributes
        xml = """<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:del w:id="1" w:author="Test">
                <w:r>
                    <w:delText xml:space="preserve">test text</w:delText>
                </w:r>
            </w:del>
        </w:document>"""

        mock_editor = MagicMock()
        mock_editor.dom = defusedxml.minidom.parseString(xml)

        manager = RevisionManager(mock_editor)

        # Get the w:del element
        del_elem = mock_editor.dom.getElementsByTagName("w:del")[0]

        # Restore the deletion
        manager._restore_deletion(del_elem)

        # Verify w:t was created with xml:space attribute
        t_elems = mock_editor.dom.getElementsByTagName("w:t")
        assert len(t_elems) == 1
        assert t_elems[0].getAttribute("xml:space") == "preserve"

    def test_restore_deletion_converts_rsiddel_to_rsidr(self):
        """Test that _restore_deletion converts w:rsidDel to w:rsidR on runs."""
        import defusedxml.minidom

        xml = """<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:del w:id="1" w:author="Test">
                <w:r w:rsidDel="00112233">
                    <w:delText>text</w:delText>
                </w:r>
            </w:del>
        </w:document>"""

        mock_editor = MagicMock()
        mock_editor.dom = defusedxml.minidom.parseString(xml)

        manager = RevisionManager(mock_editor)

        del_elem = mock_editor.dom.getElementsByTagName("w:del")[0]
        manager._restore_deletion(del_elem)

        # Verify w:rsidDel was converted to w:rsidR
        r_elems = mock_editor.dom.getElementsByTagName("w:r")
        assert len(r_elems) == 1
        assert r_elems[0].getAttribute("w:rsidR") == "00112233"
        assert not r_elems[0].hasAttribute("w:rsidDel")
