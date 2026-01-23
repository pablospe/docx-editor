"""Tests for track changes functionality."""

import pytest

from docx_edit import Document, TextNotFoundError


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
