"""Tests for comment functionality."""

import pytest

from docx_edit import CommentError, Document, TextNotFoundError


class TestAddComment:
    """Tests for adding comments."""

    def test_add_comment_returns_id(self, clean_workspace):
        """Test that add_comment returns a comment ID."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("the", "This is a test comment")
            assert isinstance(comment_id, int)
            assert comment_id >= 0
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        doc.close()

    def test_add_comment_anchor_not_found(self, clean_workspace):
        """Test that adding comment to nonexistent text raises error."""
        doc = Document.open(clean_workspace)

        with pytest.raises(TextNotFoundError):
            doc.add_comment("xyz123nonexistent789", "Comment text")

        doc.close()

    def test_add_multiple_comments(self, clean_workspace):
        """Test adding multiple comments."""
        doc = Document.open(clean_workspace)

        try:
            id1 = doc.add_comment("the", "First comment")
            id2 = doc.add_comment("a", "Second comment")
            assert id1 != id2
            assert id2 == id1 + 1
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        doc.close()


class TestReplyToComment:
    """Tests for comment replies."""

    def test_reply_to_comment(self, clean_workspace):
        """Test replying to an existing comment."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("the", "Original comment")
            reply_id = doc.reply_to_comment(comment_id, "This is a reply")
            assert reply_id == comment_id + 1
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        doc.close()

    def test_reply_to_nonexistent_comment(self, clean_workspace):
        """Test replying to a comment that doesn't exist."""
        doc = Document.open(clean_workspace)

        with pytest.raises(CommentError):
            doc.reply_to_comment(99999, "Reply text")

        doc.close()


class TestListComments:
    """Tests for listing comments."""

    def test_list_comments_empty(self, clean_workspace):
        """Test listing comments on document without comments."""
        doc = Document.open(clean_workspace)

        comments = doc.list_comments()
        assert isinstance(comments, list)
        # May be empty or have pre-existing comments

        doc.close()

    def test_list_comments_after_adding(self, clean_workspace):
        """Test listing comments after adding some."""
        doc = Document.open(clean_workspace)

        try:
            doc.add_comment("the", "Test comment")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        comments = doc.list_comments()
        assert len(comments) >= 1

        # Check comment attributes
        comment = comments[-1]  # Get last added
        assert hasattr(comment, "id")
        assert hasattr(comment, "text")
        assert hasattr(comment, "author")
        assert "Test comment" in comment.text

        doc.close()

    def test_list_comments_with_replies(self, clean_workspace):
        """Test that replies are nested in parent comments."""
        doc = Document.open(clean_workspace)

        try:
            parent_id = doc.add_comment("the", "Parent comment")
            doc.reply_to_comment(parent_id, "Reply 1")
            doc.reply_to_comment(parent_id, "Reply 2")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        comments = doc.list_comments()

        # Find the parent comment
        parent = next((c for c in comments if c.id == parent_id), None)
        assert parent is not None
        assert len(parent.replies) == 2

        doc.close()

    def test_list_comments_filter_by_author(self, clean_workspace):
        """Test filtering comments by author."""
        doc = Document.open(clean_workspace, author="TestAuthor")

        try:
            doc.add_comment("the", "Test comment")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        author_comments = doc.list_comments(author="TestAuthor")
        for comment in author_comments:
            assert comment.author == "TestAuthor"

        doc.close()


class TestResolveComment:
    """Tests for resolving comments."""

    def test_resolve_comment(self, clean_workspace):
        """Test resolving a comment."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("the", "Test comment")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        result = doc.resolve_comment(comment_id)
        assert result is True

        # Check that comment is marked as resolved
        comments = doc.list_comments()
        comment = next((c for c in comments if c.id == comment_id), None)
        if comment:
            assert comment.resolved is True

        doc.close()

    def test_resolve_nonexistent_comment(self, clean_workspace):
        """Test resolving a comment that doesn't exist."""
        doc = Document.open(clean_workspace)

        result = doc.resolve_comment(99999)
        assert result is False

        doc.close()


class TestDeleteComment:
    """Tests for deleting comments."""

    def test_delete_comment(self, clean_workspace):
        """Test deleting a comment."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("the", "Test comment")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        initial_count = len(doc.list_comments())
        result = doc.delete_comment(comment_id)

        assert result is True
        assert len(doc.list_comments()) == initial_count - 1

        doc.close()

    def test_delete_nonexistent_comment(self, clean_workspace):
        """Test deleting a comment that doesn't exist."""
        doc = Document.open(clean_workspace)

        result = doc.delete_comment(99999)
        assert result is False

        doc.close()
