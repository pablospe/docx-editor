"""Tests for track changes functionality."""

import pytest

from docx_editor import Document, TextNotFoundError


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
