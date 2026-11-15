"""Tests for comment functionality."""

from typing import Any

import pytest
from conftest import count_dom_walks, find_ref, match_for

from docx_editor import (
    CommentError,
    Document,
    HashMismatchError,
    ParagraphIndexError,
    TextNotFoundError,
)


class TestAddComment:
    """Tests for adding comments."""

    def test_add_comment_returns_id(self, clean_workspace):
        """Test that add_comment returns a comment ID."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "This is a test comment")
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
            id1 = doc.add_comment("fox", "First comment")
            id2 = doc.add_comment("lazy", "Second comment")
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
            comment_id = doc.add_comment("fox", "Original comment")
            reply_id = doc.reply_to_comment(comment_id, "This is a reply")
            assert reply_id == comment_id + 1
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        doc.close()

    def test_reply_markers_seat_end_before_reference(self, clean_workspace):
        """A reply's range end precedes its reference run, as the parent's does.

        Two insert_after calls on the same anchor land in reverse order, which
        put the reply's reference run inside its own range (ROADMAP.md #76).
        """
        doc = Document.open(clean_workspace)
        parent = doc.add_comment("fox", "Original comment")
        reply = doc.reply_to_comment(parent, "This is a reply")

        markers = ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference")
        dom = doc._document_editor.dom
        order = [
            (n.tagName, int(n.getAttribute("w:id"))) for n in dom.getElementsByTagName("*") if n.tagName in markers
        ]
        assert order == [
            ("w:commentRangeStart", parent),
            ("w:commentRangeStart", reply),
            ("w:commentRangeEnd", parent),
            ("w:commentReference", parent),
            ("w:commentRangeEnd", reply),
            ("w:commentReference", reply),
        ]
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
            doc.add_comment("fox", "Test comment")
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
            parent_id = doc.add_comment("fox", "Parent comment")
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
            doc.add_comment("fox", "Test comment")
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
            comment_id = doc.add_comment("fox", "Test comment")
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
            comment_id = doc.add_comment("fox", "Test comment")
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


class TestCommentRepr:
    """Tests for Comment.__repr__ method."""

    def test_comment_repr_basic(self):
        """Test basic comment repr."""
        from docx_editor.comments import Comment

        comment = Comment(id=1, text="This is a test comment", author="TestUser", date=None)
        result = repr(comment)
        assert "Comment(1:" in result
        assert "This is a test comment" in result
        assert "TestUser" in result

    def test_comment_repr_resolved(self):
        """Test repr for resolved comment."""
        from docx_editor.comments import Comment

        comment = Comment(id=1, text="Resolved comment text here", author="Author", date=None, resolved=True)
        result = repr(comment)
        assert "[RESOLVED]" in result

    def test_comment_repr_with_replies(self):
        """Test repr showing reply count."""
        from docx_editor.comments import Comment

        reply1 = Comment(id=2, text="Reply 1", author="Author", date=None)
        reply2 = Comment(id=3, text="Reply 2", author="Author", date=None)
        comment = Comment(id=1, text="Parent comment text here", author="Author", date=None, replies=[reply1, reply2])
        result = repr(comment)
        assert "(2 replies)" in result

    def test_comment_repr_long_text_truncated(self):
        """Test that long text is truncated in repr."""
        from docx_editor.comments import Comment

        long_text = "A" * 100
        comment = Comment(id=1, text=long_text, author="Author", date=None)
        result = repr(comment)
        # Text should be truncated to 30 chars + "..."
        assert "..." in result
        assert "A" * 30 in result


class TestCommentManagerEdgeCases:
    """Tests for CommentManager edge cases and error handling."""

    def test_list_comments_with_resolved_status(self, clean_workspace):
        """Test that list_comments shows resolved status."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "Test comment for resolve")
            doc.resolve_comment(comment_id)
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        comments = doc.list_comments()
        comment = next((c for c in comments if c.id == comment_id), None)
        assert comment is not None
        assert comment.resolved is True

        doc.close()

    def test_resolve_comment_when_extended_file_missing(self, clean_workspace):
        """resolve_comment recreates a missing commentsExtended.xml and resolves the comment."""
        doc = Document.open(clean_workspace)

        # Add a comment then delete the commentsExtended.xml file
        try:
            comment_id = doc.add_comment("fox", "Test comment")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        # Delete the commentsExtended.xml file to simulate missing file
        import os

        comments_extended_path = doc._workspace.workspace_path / "word" / "commentsExtended.xml"
        if comments_extended_path.exists():
            os.remove(comments_extended_path)

        # Clear the editor cache so it doesn't use the old cached version
        doc._comment_manager._editors.clear()

        result = doc.resolve_comment(comment_id)
        assert result is True
        manager = doc._comment_manager
        row = manager._extended_row(manager.existing_comments[comment_id]["para_id"])
        assert row is not None
        assert row.getAttribute("w15:done") == "1"

        doc.close()

    def test_delete_comment_removes_from_all_files(self, clean_workspace):
        """Test that delete_comment removes entries from all XML files."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "Comment to delete")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        # Verify comment exists
        comments_before = doc.list_comments()
        assert any(c.id == comment_id for c in comments_before)

        # Delete the comment
        result = doc.delete_comment(comment_id)
        assert result is True

        # Verify comment is gone
        comments_after = doc.list_comments()
        assert not any(c.id == comment_id for c in comments_after)

        doc.close()

    def test_list_comments_empty_when_no_comments_file(self, clean_workspace):
        """Test list_comments returns empty list when comments.xml doesn't exist."""
        doc = Document.open(clean_workspace)

        # Before adding any comments, there should be no comments.xml
        comments = doc.list_comments()
        assert isinstance(comments, list)

        doc.close()


class TestCommentParsingEdgeCases:
    """Tests for edge cases in comment parsing."""

    def test_comment_with_invalid_date(self, clean_workspace):
        """Test that invalid date is handled gracefully."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "Test comment with date")
        except TextNotFoundError:
            doc.close()
            pytest.skip("Anchor text not found in document")

        # Save the document first to ensure comments.xml exists
        doc.save()

        # Manually corrupt the date in the XML
        comments_path = doc._workspace.workspace_path / "word" / "comments.xml"
        from docx_editor.xml_editor import DocxXMLEditor

        editor = DocxXMLEditor(comments_path, rsid="00000000", author="Test", initials="T")
        for comment_elem in editor.dom.getElementsByTagName("w:comment"):
            if comment_elem.getAttribute("w:id") == str(comment_id):
                comment_elem.setAttribute("w:date", "invalid-date-format")
        editor.save()

        # Close to flush the old cached editors and reopen
        doc.close(cleanup=False)
        doc = Document.open(clean_workspace)

        # List comments should handle invalid date
        comments = doc.list_comments()
        comment = next((c for c in comments if c.id == comment_id), None)
        assert comment is not None
        assert comment.date is None  # Invalid date should result in None

        doc.close()

    def test_comment_with_empty_text(self, clean_workspace):
        """Test parsing comment with no text content."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "")  # Empty comment text
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        comments = doc.list_comments()
        comment = next((c for c in comments if c.id == comment_id), None)
        assert comment is not None
        assert comment.text == ""

        doc.close()

    def test_comment_with_special_characters(self, clean_workspace):
        """Test comment with HTML special characters."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "Test <>&\"' special chars")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        comments = doc.list_comments()
        comment = next((c for c in comments if c.id == comment_id), None)
        assert comment is not None
        # The text should be properly escaped and then unescaped when read
        assert "<" in comment.text or "&lt;" in comment.text

        doc.close()


class TestCommentManagerInternalMethods:
    """Tests for internal CommentManager methods."""

    def test_get_next_comment_id_with_no_existing_comments(self, clean_workspace):
        """Test _get_next_comment_id when no comments exist."""
        doc = Document.open(clean_workspace)

        # Before adding comments, next_comment_id should be 0
        assert doc._comment_manager.next_comment_id == 0

        doc.close()

    def test_ensure_comment_file_without_write_hook(self, clean_workspace):
        """on_write is optional: a manager built without the hook must still
        create comment files from templates (None skips the write-ahead call)."""
        from docx_editor.comments import CommentManager

        doc = Document.open(clean_workspace)
        try:
            manager = CommentManager(
                doc._workspace.workspace_path,
                doc._document_editor,
                doc.author,
                "T",
            )
            assert not manager.comments_path.exists()
            manager._ensure_comment_file(manager.comments_path, "comments.xml")
            assert manager.comments_path.exists()
        finally:
            doc.close()

    def test_get_next_comment_id_after_adding_comments(self, clean_workspace):
        """Test _get_next_comment_id after adding comments."""
        doc = Document.open(clean_workspace)

        try:
            doc.add_comment("fox", "First comment")
            doc.add_comment("lazy", "Second comment")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        assert doc._comment_manager.next_comment_id == 2

        doc.close()

    def test_load_existing_comments_with_malformed_data(self, clean_workspace):
        """Test _load_existing_comments handles malformed comment data."""
        doc = Document.open(clean_workspace)

        try:
            doc.add_comment("fox", "Test comment")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        # Manually add a malformed comment without id attribute
        comments_path = doc._workspace.workspace_path / "word" / "comments.xml"
        from docx_editor.xml_editor import DocxXMLEditor

        editor = DocxXMLEditor(comments_path, rsid="00000000", author="Test", initials="T")
        root = editor.get_node(tag="w:comments")

        # Add comment without w:id attribute
        malformed_xml = """<w:comment>
          <w:p w14:paraId="12345678" w14:textId="77777777">
            <w:r><w:t>Malformed comment</w:t></w:r>
          </w:p>
        </w:comment>"""
        editor.append_to(root, malformed_xml)
        editor.save()

        # Reload the document to re-parse comments
        doc.close()
        doc = Document.open(clean_workspace)

        # Should still be able to list comments without error
        comments = doc.list_comments()
        assert isinstance(comments, list)

        doc.close()

    def test_comment_without_para_id_skipped(self, clean_workspace):
        """Test that comments without para_id are skipped in loading."""
        doc = Document.open(clean_workspace)

        try:
            doc.add_comment("fox", "Test comment")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        # Add a comment with id but without para_id
        comments_path = doc._workspace.workspace_path / "word" / "comments.xml"
        from docx_editor.xml_editor import DocxXMLEditor

        editor = DocxXMLEditor(comments_path, rsid="00000000", author="Test", initials="T")
        root = editor.get_node(tag="w:comments")

        # Add comment with w:id but w:p without w14:paraId
        malformed_xml = """<w:comment w:id="999">
          <w:p w14:textId="77777777">
            <w:r><w:t>Comment without paraId</w:t></w:r>
          </w:p>
        </w:comment>"""
        editor.append_to(root, malformed_xml)
        editor.save()

        # Reload to re-parse
        doc.close()
        doc = Document.open(clean_workspace)

        # Comment 999 should not be in existing_comments
        assert 999 not in doc._comment_manager.existing_comments

        doc.close()


class TestDeleteCommentErrorHandling:
    """Tests for delete_comment error handling paths."""

    def test_delete_comment_with_missing_range_start(self, clean_workspace):
        """Test delete_comment handles missing commentRangeStart gracefully."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "Comment to partially delete")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        # Manually remove commentRangeStart from document.xml
        document_path = doc._workspace.workspace_path / "word" / "document.xml"
        from docx_editor.xml_editor import DocxXMLEditor

        editor = DocxXMLEditor(document_path, rsid="00000000", author="Test", initials="T")
        for elem in editor.dom.getElementsByTagName("w:commentRangeStart"):
            if elem.getAttribute("w:id") == str(comment_id):
                elem.parentNode.removeChild(elem)
                break
        editor.save()

        # Flush the new comment, then clear the editor cache
        doc._comment_manager.save_all()
        doc._comment_manager._editors.clear()
        doc._comment_manager.document_editor = DocxXMLEditor(
            document_path, rsid="00000000", author="Test", initials="T"
        )

        # delete_comment should still succeed (handles missing elements)
        result = doc.delete_comment(comment_id)
        assert result is True

        doc.close()

    def test_delete_comment_with_missing_range_end(self, clean_workspace):
        """Test delete_comment handles missing commentRangeEnd gracefully."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "Comment to partially delete")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        # Manually remove commentRangeEnd from document.xml
        document_path = doc._workspace.workspace_path / "word" / "document.xml"
        from docx_editor.xml_editor import DocxXMLEditor

        editor = DocxXMLEditor(document_path, rsid="00000000", author="Test", initials="T")
        for elem in editor.dom.getElementsByTagName("w:commentRangeEnd"):
            if elem.getAttribute("w:id") == str(comment_id):
                elem.parentNode.removeChild(elem)
                break
        editor.save()

        # Flush the new comment, then clear the editor cache
        doc._comment_manager.save_all()
        doc._comment_manager._editors.clear()
        doc._comment_manager.document_editor = DocxXMLEditor(
            document_path, rsid="00000000", author="Test", initials="T"
        )

        # delete_comment should still succeed
        result = doc.delete_comment(comment_id)
        assert result is True

        doc.close()

    def test_delete_comment_with_missing_reference(self, clean_workspace):
        """Test delete_comment handles missing commentReference gracefully."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "Comment to partially delete")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        # Manually remove commentReference from document.xml
        document_path = doc._workspace.workspace_path / "word" / "document.xml"
        from docx_editor.xml_editor import DocxXMLEditor

        editor = DocxXMLEditor(document_path, rsid="00000000", author="Test", initials="T")
        for elem in editor.dom.getElementsByTagName("w:commentReference"):
            if elem.getAttribute("w:id") == str(comment_id):
                elem.parentNode.removeChild(elem)
                break
        editor.save()

        # Flush the new comment, then clear the editor cache
        doc._comment_manager.save_all()
        doc._comment_manager._editors.clear()
        doc._comment_manager.document_editor = DocxXMLEditor(
            document_path, rsid="00000000", author="Test", initials="T"
        )

        # delete_comment should still succeed
        result = doc.delete_comment(comment_id)
        assert result is True

        doc.close()

    def test_delete_comment_reference_not_in_run(self, clean_workspace):
        """Test delete_comment when commentReference parent is not w:r."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "Comment to test reference removal")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        # The delete should work normally - this tests the branch where ref.parentNode
        # might not be a w:r element
        result = doc.delete_comment(comment_id)
        assert result is True

        doc.close()


class TestResolveCommentParaIdNotFound:
    """Tests for resolve_comment when para_id is not found in commentsExtended."""

    def test_resolve_comment_para_id_not_in_extended(self, clean_workspace):
        """resolve_comment adds the missing commentEx row and resolves the comment."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "Test comment")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        # Manually remove the commentEx entry from commentsExtended.xml
        comments_extended_path = doc._workspace.workspace_path / "word" / "commentsExtended.xml"
        from docx_editor.xml_editor import DocxXMLEditor

        editor = DocxXMLEditor(comments_extended_path, rsid="00000000", author="Test", initials="T")
        for elem in editor.dom.getElementsByTagName("w15:commentEx"):
            elem.parentNode.removeChild(elem)
        editor.save()

        # Clear the editor cache
        doc._comment_manager._editors.clear()

        # The comment has a paraId and no row: resolve_comment writes the row.
        result = doc.resolve_comment(comment_id)
        assert result is True
        manager = doc._comment_manager
        row = manager._extended_row(manager.existing_comments[comment_id]["para_id"])
        assert row is not None
        assert row.getAttribute("w15:done") == "1"

        doc.close()


class TestNextCommentIdWithInvalidIds:
    """Tests for _get_next_comment_id with edge cases."""

    def test_next_comment_id_with_non_numeric_id(self, clean_workspace):
        """Test _get_next_comment_id handles non-numeric IDs gracefully.

        This tests the try/except in _get_next_comment_id (lines 365-369).
        We directly manipulate the DOM after the document is opened to add
        a malformed comment that won't trigger issues in _load_existing_comments.
        """
        doc = Document.open(clean_workspace)

        try:
            doc.add_comment("fox", "First comment")
        except TextNotFoundError:
            doc.close()
            pytest.skip("Anchor text not found in document")

        # Save to ensure comments.xml is persisted
        doc.save()

        # Get the comments editor and add a malformed comment directly to the DOM
        # This way we bypass _load_existing_comments which already ran
        comments_path = doc._workspace.workspace_path / "word" / "comments.xml"
        editor = doc._comment_manager._get_editor(comments_path)
        root = editor.get_node(tag="w:comments")

        # Add comment with non-numeric w:id directly to the DOM
        malformed_xml = """<w:comment w:id="abc">
          <w:p>
            <w:r><w:t>Comment with non-numeric id</w:t></w:r>
          </w:p>
        </w:comment>"""
        editor.append_to(root, malformed_xml)

        # Now call _get_next_comment_id which should handle the non-numeric id
        # by catching ValueError and continuing
        next_id = doc._comment_manager._get_next_comment_id()

        # The existing comment has id=0, so next should be 1
        # The non-numeric "abc" should be ignored due to the try/except
        assert next_id == 1

        doc.close()

    def test_next_comment_id_with_empty_id(self, clean_workspace):
        """Test _get_next_comment_id handles empty IDs gracefully."""
        doc = Document.open(clean_workspace)

        try:
            doc.add_comment("fox", "First comment")
        except TextNotFoundError:
            doc.close()
            pytest.skip("Anchor text not found in document")

        # Save to ensure comments.xml is persisted
        doc.save()

        # Add a comment with empty id
        comments_path = doc._workspace.workspace_path / "word" / "comments.xml"
        from docx_editor.xml_editor import DocxXMLEditor

        editor = DocxXMLEditor(comments_path, rsid="00000000", author="Test", initials="T")
        root = editor.get_node(tag="w:comments")

        # Add comment with empty w:id
        malformed_xml = """<w:comment w:id="">
          <w:p w14:paraId="BBBBBBBB" w14:textId="77777777">
            <w:r><w:t>Comment with empty id</w:t></w:r>
          </w:p>
        </w:comment>"""
        editor.append_to(root, malformed_xml)
        editor.save()

        # Reload to re-parse (keep workspace)
        doc.close(cleanup=False)
        doc = Document.open(clean_workspace)

        # Should still work - existing comment id=0, so next should be 1
        assert doc._comment_manager.next_comment_id >= 1

        doc.close()


class TestParseCommentEdgeCases:
    """Tests for _parse_comment edge cases."""

    def test_parse_comment_without_id(self, clean_workspace):
        """Test _parse_comment returns None when comment has no id attribute."""
        doc = Document.open(clean_workspace)

        try:
            doc.add_comment("fox", "Test comment")
        except TextNotFoundError:
            doc.close()
            pytest.skip("Anchor text not found in document")

        doc.save()

        # Get the comments editor and add a comment without id
        comments_path = doc._workspace.workspace_path / "word" / "comments.xml"
        editor = doc._comment_manager._get_editor(comments_path)
        root = editor.get_node(tag="w:comments")

        # Add comment without w:id attribute
        malformed_xml = """<w:comment w:author="Test">
          <w:p w14:paraId="CCCCCCCC" w14:textId="77777777">
            <w:r><w:t>Comment without id</w:t></w:r>
          </w:p>
        </w:comment>"""
        editor.append_to(root, malformed_xml)

        # Get the last comment element (the one we just added)
        comment_elems = editor.dom.getElementsByTagName("w:comment")
        last_comment = comment_elems[-1]

        # _parse_comment should return None for comment without id
        result = doc._comment_manager._parse_comment(last_comment)
        assert result is None

        doc.close()

    def test_get_comment_para_id_no_para_id_found(self, clean_workspace):
        """Test _get_comment_para_id returns None when no para_id attribute exists."""
        doc = Document.open(clean_workspace)

        try:
            doc.add_comment("fox", "Test comment")
        except TextNotFoundError:
            doc.close()
            pytest.skip("Anchor text not found in document")

        doc.save()

        # Get the comments editor - use the base XMLEditor to avoid auto-injection
        comments_path = doc._workspace.workspace_path / "word" / "comments.xml"
        from docx_editor.xml_editor import XMLEditor

        editor = XMLEditor(comments_path)
        root = editor.get_node(tag="w:comments")

        # Add comment with w:p that has no w14:paraId
        malformed_xml = """<w:comment w:id="998">
          <w:p w14:textId="77777777">
            <w:r><w:t>Comment without paraId</w:t></w:r>
          </w:p>
        </w:comment>"""
        editor.append_to(root, malformed_xml)
        editor.save()

        # Clear comment manager's cached editor
        doc._comment_manager._editors.clear()

        # Get fresh editor to load the modified comments.xml
        fresh_editor = doc._comment_manager._get_editor(comments_path)

        # Get the last comment element (the one we just added)
        comment_elems = fresh_editor.dom.getElementsByTagName("w:comment")
        last_comment = None
        for c in comment_elems:
            if c.getAttribute("w:id") == "998":
                last_comment = c
                break

        assert last_comment is not None

        # _get_comment_para_id should return None since no w14:paraId
        result = doc._comment_manager._get_comment_para_id(last_comment)
        assert result is None

        doc.close()

    def test_load_existing_comments_registers_comment_without_para_id(self, clean_workspace):
        """A comment with no paraId is registered, with ``para_id`` None."""
        doc = Document.open(clean_workspace)

        try:
            doc.add_comment("fox", "Test comment")
        except TextNotFoundError:
            doc.close()
            pytest.skip("Anchor text not found in document")

        doc.save()

        # Use base XMLEditor to avoid auto-injection of w14:paraId
        comments_path = doc._workspace.workspace_path / "word" / "comments.xml"
        from docx_editor.xml_editor import XMLEditor

        editor = XMLEditor(comments_path)
        root = editor.get_node(tag="w:comments")

        # Add comment with w:id but w:p without w14:paraId
        malformed_xml = """<w:comment w:id="997">
          <w:p w14:textId="77777777">
            <w:r><w:t>No paraId here</w:t></w:r>
          </w:p>
        </w:comment>"""
        editor.append_to(root, malformed_xml)
        editor.save()

        # Clear editor cache and call _load_existing_comments again
        doc._comment_manager._editors.clear()
        existing = doc._comment_manager._load_existing_comments()

        # Comment 997 is addressable by its w:id alone
        assert existing[997] == {"para_id": None}
        # And comment 0 keeps the paraId it was written with
        assert existing[0]["para_id"]

        doc.close()


class TestDeleteCommentRefNotInRun:
    """Test delete_comment when commentReference parent is not w:r."""

    def test_delete_comment_ref_parent_not_run(self, clean_workspace):
        """Test delete_comment when commentReference's parent is not a w:r element."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "Comment to manipulate")
        except TextNotFoundError:
            doc.close()
            pytest.skip("Anchor text not found in document")

        doc.save()

        # Manually modify the document to have commentReference outside of w:r
        document_path = doc._workspace.workspace_path / "word" / "document.xml"
        from docx_editor.xml_editor import DocxXMLEditor

        editor = DocxXMLEditor(document_path, rsid="00000000", author="Test", initials="T")

        # Find the commentReference and move it outside of its w:r parent
        for ref_elem in editor.dom.getElementsByTagName("w:commentReference"):
            if ref_elem.getAttribute("w:id") == str(comment_id):
                parent = ref_elem.parentNode
                grandparent = parent.parentNode
                if parent.nodeName == "w:r":
                    # Remove from run and append directly to paragraph
                    parent.removeChild(ref_elem)
                    grandparent.appendChild(ref_elem)
                break
        editor.save()

        # Clear editor cache and reload document editor
        doc._comment_manager._editors.clear()
        doc._comment_manager.document_editor = DocxXMLEditor(
            document_path, rsid="00000000", author="Test", initials="T"
        )

        # Now delete should use the else branch (line 312)
        result = doc.delete_comment(comment_id)
        assert result is True

        doc.close()


class TestAddCommentParentTraversal:
    """Tests for parent traversal in add_comment."""

    def test_add_comment_traverses_to_find_run(self, clean_workspace):
        """Test add_comment when w:t parent is not immediately w:r."""
        doc = Document.open(clean_workspace)

        try:
            comment_id = doc.add_comment("fox", "Test comment")
            assert isinstance(comment_id, int)
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        doc.close()


class TestSplitTextNodeComment:
    """Issue #9: comments with multi-TEXT_NODE w:t must round-trip full text."""

    def test_list_comments_full_text_with_split_w_t(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            doc.add_comment("fox", "Hello World")
        except TextNotFoundError:
            pytest.skip("Anchor text not found in document")

        editor = doc._comment_manager._get_editor(doc._comment_manager.comments_path)
        wts = editor.dom.getElementsByTagName("w:t")
        target = None
        for wt in wts:
            full = "".join(c.data for c in wt.childNodes if c.nodeType == c.TEXT_NODE)
            if "Hello World" in full:
                target = wt
                break
        assert target is not None, "comment w:t not found"

        while target.firstChild:
            target.removeChild(target.firstChild)
        owner = target.ownerDocument
        target.appendChild(owner.createTextNode("Hello "))
        target.appendChild(owner.createTextNode("World"))

        comments = doc.list_comments()
        texts = [c.text for c in comments]
        assert "Hello World" in texts
        doc.close()


def _split_first_run_text(paragraph, split_at: int) -> None:
    """Split the first ``<w:r>``'s single ``<w:t>`` text into two runs.

    Useful for synthesizing a cross-run anchor: callers locate a paragraph
    whose run contains the anchor entirely in one ``w:t``, then call this to
    rebuild it as ``<w:r><w:t>head</w:t></w:r><w:r><w:t>tail</w:t></w:r>``.
    """
    run = paragraph.getElementsByTagName("w:r")[0]
    wt = run.getElementsByTagName("w:t")[0]
    full = "".join(c.data for c in wt.childNodes if c.nodeType == c.TEXT_NODE)
    head, tail = full[:split_at], full[split_at:]
    owner = paragraph.ownerDocument

    new_run_2 = owner.createElement("w:r")
    new_t_2 = owner.createElement("w:t")
    new_t_2.setAttribute("xml:space", "preserve")
    new_t_2.appendChild(owner.createTextNode(tail))
    new_run_2.appendChild(new_t_2)

    while wt.firstChild:
        wt.removeChild(wt.firstChild)
    wt.appendChild(owner.createTextNode(head))
    wt.setAttribute("xml:space", "preserve")

    parent = run.parentNode
    next_sib = run.nextSibling
    if next_sib:
        parent.insertBefore(new_run_2, next_sib)
    else:
        parent.appendChild(new_run_2)


def _find_paragraph_with_text(doc, text: str):
    """Return the ``<w:p>`` element whose visible text contains ``text``."""
    from docx_editor.xml_editor import build_text_map

    for p in doc._document_editor.dom.getElementsByTagName("w:p"):
        if text in build_text_map(p).text:
            return p
    raise AssertionError(f"No paragraph contains {text!r}")


def _paragraph_ref(doc, text: str) -> str:
    """Find ``P{i}#{hash}`` for the paragraph containing ``text``."""
    for entry in doc.list_paragraphs():
        if text in entry:
            return entry.split("|", 1)[0]
    raise AssertionError(f"No paragraph contains {text!r}")


class TestCommentCrossBoundaryAnchor:
    """Issue #5: ``add_comment`` must locate anchors that span run boundaries.

    Mirrors the asymmetry fix between ``count_matches`` (text-map search) and
    the old ``get_node(contains=...)`` lookup, which only saw text inside a
    single ``<w:t>``.
    """

    def test_anchor_split_across_two_runs(self, clean_workspace):
        """Regression: anchor spanning two ``<w:t>`` is found and commented."""
        doc = Document.open(clean_workspace)
        try:
            para = _find_paragraph_with_text(doc, "brown fox")
            full = "The quick brown fox jumps over the lazy dog."
            _split_first_run_text(para, full.index("brown") + 2)  # split inside "brown"

            assert doc.count_matches("brown fox") >= 1
            comment_id = doc.add_comment("brown fox", "spans runs")
            assert isinstance(comment_id, int)

            # Marker pair must be present in document.xml
            starts = doc._document_editor.dom.getElementsByTagName("w:commentRangeStart")
            ends = doc._document_editor.dom.getElementsByTagName("w:commentRangeEnd")
            assert any(s.getAttribute("w:id") == str(comment_id) for s in starts)
            assert any(e.getAttribute("w:id") == str(comment_id) for e in ends)
        finally:
            doc.close()

    def test_anchor_across_formatting_boundary(self, clean_workspace):
        """Anchor splits across two runs with different ``rPr`` (bold + plain)."""
        doc = Document.open(clean_workspace)
        try:
            para = _find_paragraph_with_text(doc, "quick brown")
            # Build a paragraph with a bold "quick" run + plain " brown" run by
            # replacing the existing single run.
            run = para.getElementsByTagName("w:r")[0]
            new_xml = (
                '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">quick</w:t></w:r>'
                '<w:r><w:t xml:space="preserve"> brown fox jumps over the lazy dog.</w:t></w:r>'
            )
            doc._document_editor.replace_node(run, new_xml)

            assert doc.count_matches("quick brown") == 1
            comment_id = doc.add_comment("quick brown", "fmt boundary")
            assert isinstance(comment_id, int)
        finally:
            doc.close()

    def test_paragraph_scoped_lookup_finds_anchor(self, clean_workspace):
        """``paragraph=`` scope restricts the search and accepts the anchor."""
        doc = Document.open(clean_workspace)
        try:
            ref = _paragraph_ref(doc, "brown fox")
            comment_id = doc.add_comment("fox", "scoped", paragraph=ref)
            assert isinstance(comment_id, int)
        finally:
            doc.close()

    def test_paragraph_scoped_lookup_anchor_not_in_paragraph(self, clean_workspace):
        """Scoped search raises ``TextNotFoundError`` when anchor lives elsewhere."""
        doc = Document.open(clean_workspace)
        try:
            other_ref = _paragraph_ref(doc, "sample document")
            # "fox" exists in a different paragraph
            with pytest.raises(TextNotFoundError) as exc_info:
                doc.add_comment("fox", "wrong paragraph", paragraph=other_ref)
            assert exc_info.value.paragraph_ref == other_ref
        finally:
            doc.close()

    def test_stale_paragraph_hash_raises(self, clean_workspace):
        """Passing a hash that no longer matches raises ``HashMismatchError``."""
        doc = Document.open(clean_workspace)
        try:
            ref = _paragraph_ref(doc, "brown fox")
            # Construct a stale ref: keep the index, mangle the hash
            idx_part = ref.split("#")[0]
            stale = f"{idx_part}#0000"
            with pytest.raises(HashMismatchError):
                doc.add_comment("fox", "stale", paragraph=stale)
        finally:
            doc.close()

    def test_occurrence_index_selects_match(self, clean_workspace):
        """``occurrence=N`` picks the Nth match within the scoped paragraph."""
        doc = Document.open(clean_workspace)
        try:
            # Synthesize a paragraph with two copies of "abc"
            para = _find_paragraph_with_text(doc, "brown fox")
            run = para.getElementsByTagName("w:r")[0]
            doc._document_editor.replace_node(
                run,
                '<w:r><w:t xml:space="preserve">abc middle abc tail</w:t></w:r>',
            )
            # Recompute ref since hash changed
            ref = _paragraph_ref(doc, "abc middle")

            id0 = doc.add_comment("abc", "first", paragraph=ref, occurrence=0)
            # After adding a comment, the paragraph hash changes — recompute
            ref2 = _paragraph_ref(doc, "abc middle")
            id1 = doc.add_comment("abc", "second", paragraph=ref2, occurrence=1)
            assert id0 != id1

            # Third occurrence does not exist
            ref3 = _paragraph_ref(doc, "abc middle")
            with pytest.raises(TextNotFoundError) as exc_info:
                doc.add_comment("abc", "missing", paragraph=ref3, occurrence=2)
            assert exc_info.value.occurrence == 2
            assert exc_info.value.total_occurrences == 2
        finally:
            doc.close()

    def test_marker_placement_is_character_precise(self, clean_workspace):
        """Markers bracket exactly the anchor text (no oversized range)."""
        doc = Document.open(clean_workspace)
        try:
            comment_id = doc.add_comment("brown", "tight range")

            para = _find_paragraph_with_text(doc, "quick")
            # Collect children of w:p in document order; between
            # commentRangeStart and commentRangeEnd, visible text must equal
            # the anchor.
            in_range = False
            collected: list[str] = []
            for child in para.childNodes:
                if child.nodeType != child.ELEMENT_NODE:
                    continue
                if child.tagName == "w:commentRangeStart" and child.getAttribute("w:id") == str(comment_id):
                    in_range = True
                    continue
                if child.tagName == "w:commentRangeEnd" and child.getAttribute("w:id") == str(comment_id):
                    in_range = False
                    continue
                if in_range and child.tagName == "w:r":
                    for wt in child.getElementsByTagName("w:t"):
                        collected.append("".join(c.data for c in wt.childNodes if c.nodeType == c.TEXT_NODE))
            assert "".join(collected) == "brown"
        finally:
            doc.close()

    def test_anchor_inside_tracked_insertion(self, clean_workspace):
        """Anchor text inside a ``<w:ins>`` is found and bracketed."""
        doc = Document.open(clean_workspace)
        try:
            # Wrap the first run of a paragraph in <w:ins>
            para = _find_paragraph_with_text(doc, "sample document")
            run = para.getElementsByTagName("w:r")[0]
            new_xml = (
                '<w:ins w:id="999" w:author="Test" w:date="2026-01-01T00:00:00Z">'
                '<w:r><w:t xml:space="preserve">This is a sample document for testing the editing features.</w:t></w:r>'
                "</w:ins>"
            )
            doc._document_editor.replace_node(run, new_xml)

            comment_id = doc.add_comment("sample", "anchor inside w:ins")
            # Verify markers exist
            starts = doc._document_editor.dom.getElementsByTagName("w:commentRangeStart")
            ends = doc._document_editor.dom.getElementsByTagName("w:commentRangeEnd")
            assert any(s.getAttribute("w:id") == str(comment_id) for s in starts)
            assert any(e.getAttribute("w:id") == str(comment_id) for e in ends)
        finally:
            doc.close()

    def test_non_text_children_preserved(self, clean_workspace):
        """``<w:tab/>`` / ``<w:br/>`` in the anchor's run survive the split.

        Regression guard: an earlier implementation iterated
        ``getElementsByTagName("w:t")`` and silently dropped sibling tabs,
        breaks, and drawings when rebuilding the run.
        """
        doc = Document.open(clean_workspace)
        try:
            para = _find_paragraph_with_text(doc, "brown fox")
            run = para.getElementsByTagName("w:r")[0]
            # Run with interleaved tab + break alongside the matched text
            doc._document_editor.replace_node(
                run,
                '<w:r><w:t xml:space="preserve">The </w:t><w:tab/>'
                '<w:t xml:space="preserve">quick brown fox</w:t><w:br/>'
                '<w:t xml:space="preserve"> jumps</w:t></w:r>',
            )

            comment_id = doc.add_comment("brown", "preserves siblings")
            assert isinstance(comment_id, int)

            # The tab and break must still be in the paragraph after rebuild
            tabs = para.getElementsByTagName("w:tab")
            brs = para.getElementsByTagName("w:br")
            assert len(tabs) == 1
            assert len(brs) == 1
        finally:
            doc.close()

    def test_empty_anchor_text_raises(self, clean_workspace):
        """``add_comment("")`` raises ``CommentError`` rather than IndexError."""
        doc = Document.open(clean_workspace)
        try:
            with pytest.raises(CommentError):
                doc.add_comment("", "no anchor")
        finally:
            doc.close()

    def test_anchor_spans_many_cjk_runs(self, clean_workspace):
        """Regression: GitHub #14. CJK anchor spanning nine ``<w:t>`` runs.

        Reporter's failing case: a 29-character Chinese anchor that
        ``find_text``/``count_matches`` accepted but ``add_comment`` rejected
        with ``TextNotFoundError`` because the old lookup only saw text inside
        a single ``<w:t>`` element.
        """
        doc = Document.open(clean_workspace)
        try:
            para = _find_paragraph_with_text(doc, "brown fox")
            segments = [
                "息税前",
                "利润率",
                "，",
                "息税前利润率",
                "=",
                "息税前利润",
                "/",
                "营业收入",
                "×100%",
            ]
            runs_xml = "".join(f'<w:r><w:t xml:space="preserve">{seg}</w:t></w:r>' for seg in segments)
            run = para.getElementsByTagName("w:r")[0]
            doc._document_editor.replace_node(run, runs_xml)

            target = "".join(segments)
            assert doc.count_matches(target) == 1
            assert doc.find_text(target) is not None

            comment_id = doc.add_comment(target, "issue-14 repro")
            assert isinstance(comment_id, int)
        finally:
            doc.close()

    def test_paragraph_index_out_of_range_raises(self, clean_workspace):
        """``paragraph=`` with an out-of-range index raises ``ParagraphIndexError``."""
        doc = Document.open(clean_workspace)
        try:
            with pytest.raises(ParagraphIndexError):
                doc.add_comment("fox", "out of range", paragraph="P99#abcd")
        finally:
            doc.close()

    def test_stale_hash_long_paragraph_truncates_preview(self, clean_workspace):
        """``HashMismatchError.paragraph_preview`` truncates at 80 chars + '...'."""
        doc = Document.open(clean_workspace)
        try:
            # Synthesize a >80-char paragraph so the truncation branch fires
            para = _find_paragraph_with_text(doc, "brown fox")
            run = para.getElementsByTagName("w:r")[0]
            long_text = "a" * 100
            doc._document_editor.replace_node(run, f"<w:r><w:t>{long_text}</w:t></w:r>")

            ref = _paragraph_ref(doc, long_text[:20])
            idx_part = ref.split("#")[0]
            stale = f"{idx_part}#0000"

            with pytest.raises(HashMismatchError) as exc_info:
                doc.add_comment("aaa", "stale long", paragraph=stale)
            assert exc_info.value.paragraph_preview.endswith("...")
            # 80 chars of body + "..."
            assert len(exc_info.value.paragraph_preview) == 83
        finally:
            doc.close()

    def test_orphan_wt_raises_comment_error(self, clean_workspace):
        """``<w:t>`` without a ``<w:r>`` ancestor → ``CommentError``, not IndexError."""
        doc = Document.open(clean_workspace)
        try:
            para = _find_paragraph_with_text(doc, "brown fox")
            run = para.getElementsByTagName("w:r")[0]
            # Bare w:t as direct child of w:p — _find_ancestor_run will return None
            doc._document_editor.replace_node(run, '<w:t xml:space="preserve">orphan anchor text</w:t>')
            with pytest.raises(CommentError):
                doc.add_comment("orphan", "no run wrapper")
        finally:
            doc.close()

    def test_document_wide_occurrence_skips_earlier_matches(self, clean_workspace):
        """Unscoped search with ``occurrence>0`` walks past earlier hits."""
        doc = Document.open(clean_workspace)
        try:
            # "the" appears in multiple paragraphs of simple.docx
            assert doc.count_matches("the") >= 2
            comment_id = doc.add_comment("the", "second occurrence", occurrence=1)
            assert isinstance(comment_id, int)
        finally:
            doc.close()

    def test_document_wide_occurrence_out_of_range(self, clean_workspace):
        """Unscoped search with ``occurrence`` past last match reports ``total_occurrences``."""
        doc = Document.open(clean_workspace)
        try:
            total = doc.count_matches("fox")
            with pytest.raises(TextNotFoundError) as exc_info:
                doc.add_comment("fox", "out of range", occurrence=99)
            assert exc_info.value.occurrence == 99
            assert exc_info.value.total_occurrences == total
        finally:
            doc.close()

    def test_scoped_occurrence_when_anchor_absent(self, clean_workspace):
        """Scoped search with ``occurrence>0`` and zero matches raises plain ``TextNotFoundError``."""
        doc = Document.open(clean_workspace)
        try:
            ref = _paragraph_ref(doc, "brown fox")
            with pytest.raises(TextNotFoundError) as exc_info:
                doc.add_comment("nonexistent", "x", paragraph=ref, occurrence=1)
            # total_occurrences must NOT be set (the 'if total > 0' fall-through path)
            assert exc_info.value.total_occurrences is None
        finally:
            doc.close()


class TestAddCommentSearchResultAnchor:
    """``add_comment`` accepts a SearchResult as its anchor, which pins the
    paragraph and occurrence for it (ISSUES.md #52).
    """

    def test_comment_anchored_from_a_match(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            match = match_for(doc, "brown fox")
            comment_id = doc.add_comment(match, "from a search")

            comment = doc.list_comments()[0]
            assert comment.id == comment_id
            assert comment.text == "from a search"
        finally:
            doc.close()

    def test_matches_the_explicit_spelling(self, clean_workspace):
        """Same markers as passing text/paragraph/occurrence by hand."""
        doc = Document.open(clean_workspace)
        try:
            match = match_for(doc, "lazy dog")
            by_object = doc.add_comment(match, "note A")
            by_fields = doc.add_comment(
                match.text,
                "note B",
                paragraph=match.paragraph_ref,
                occurrence=match.paragraph_occurrence,
            )

            starts = doc._document_editor.dom.getElementsByTagName("w:commentRangeStart")
            ids = [s.getAttribute("w:id") for s in starts]
            assert str(by_object) in ids and str(by_fields) in ids
        finally:
            doc.close()

    def test_second_occurrence_is_targeted_without_bookkeeping(self, clean_workspace):
        """The 2nd "he" in P2 gets the comment, not the 1st."""
        doc = Document.open(clean_workspace)
        try:
            second = [m for m in doc.find_all("he") if m.paragraph_index == 2][1]
            assert second.paragraph_occurrence == 1
            doc.add_comment(second, "second only")

            # Everything before the marker must already carry the sentence up to
            # the SECOND "he" — anchoring to the first would cut after "T".
            para = _find_paragraph_with_text(doc, "quick brown fox")
            before_marker = para.toxml().split("<w:commentRangeStart")[0]
            assert "jumps over t" in before_marker
            assert "he lazy dog" not in before_marker
            assert len(doc.list_comments()) == 1
        finally:
            doc.close()

    def test_paragraph_or_occurrence_with_a_match_is_refused(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            match = match_for(doc, "brown fox")
            extras: list[dict[str, Any]] = [{"paragraph": match.paragraph_ref}, {"occurrence": 0}]
            for kwargs in extras:
                with pytest.raises(ValueError, match="already pins the paragraph"):
                    doc.add_comment(match, "x", **kwargs)
            assert doc.list_comments() == []
        finally:
            doc.close()

    def test_plain_text_still_searches_document_wide(self, clean_workspace):
        """paragraph=None keeps its meaning for the string form."""
        doc = Document.open(clean_workspace)
        try:
            assert isinstance(doc.add_comment("brown fox", "unscoped"), int)
        finally:
            doc.close()


def _range_start(doc, comment_id: int):
    """The live ``w:commentRangeStart`` marker of ``comment_id``."""
    for marker in doc._document_editor.dom.getElementsByTagName("w:commentRangeStart"):
        if marker.getAttribute("w:id") == str(comment_id):
            return marker
    raise AssertionError(f"No range start for comment {comment_id}")


class TestCommentParagraphRef:
    """``Comment.paragraph_ref`` names the paragraph a comment sits on (ROADMAP.md #88)."""

    def test_set_right_after_add_and_reusable_as_paragraph(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            expected = match_for(doc, "brown fox").paragraph_ref
            comment_id = doc.add_comment("brown fox", "first")

            (comment,) = doc.list_comments()
            assert comment.id == comment_id
            assert comment.paragraph_ref == expected

            doc.add_comment("lazy", "second", paragraph=comment.paragraph_ref)
            assert [c.paragraph_ref for c in doc.list_comments()] == [expected, expected]
        finally:
            doc.close()

    def test_comments_in_different_paragraphs_report_different_refs(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            fox_ref = _paragraph_ref(doc, "brown fox")
            sample_ref = _paragraph_ref(doc, "sample document")
            assert fox_ref != sample_ref
            fox_id = doc.add_comment("brown fox", "on the fox")
            sample_id = doc.add_comment("sample document", "on the sample")

            refs = {c.id: c.paragraph_ref for c in doc.list_comments()}
            assert refs == {fox_id: fox_ref, sample_id: sample_ref}
        finally:
            doc.close()

    def test_replies_inherit_the_thread_root_ref(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            expected = _paragraph_ref(doc, "brown fox")
            parent_id = doc.add_comment("brown fox", "parent")
            reply_id = doc.reply_to_comment(parent_id, "reply")
            doc.reply_to_comment(reply_id, "reply to the reply")

            def thread_refs(comment):
                return [comment.paragraph_ref] + [r for reply in comment.replies for r in thread_refs(reply)]

            (parent,) = doc.list_comments()
            assert thread_refs(parent) == [expected] * 3
        finally:
            doc.close()

    def test_reply_reports_parent_ref_over_its_own_marker(self, clean_workspace):
        """A foreign document may seat a reply's range in another paragraph."""
        doc = Document.open(clean_workspace)
        try:
            expected = _paragraph_ref(doc, "brown fox")
            parent_id = doc.add_comment("brown fox", "parent")
            reply_id = doc.reply_to_comment(parent_id, "reply")

            marker = _range_start(doc, reply_id)
            marker.parentNode.removeChild(marker)
            other = _find_paragraph_with_text(doc, "sample document")
            other.appendChild(marker)

            (parent,) = doc.list_comments()
            (reply,) = parent.replies
            assert reply.paragraph_ref == parent.paragraph_ref == expected
        finally:
            doc.close()

    def test_author_filter_does_not_change_a_reply_ref(self, clean_workspace):
        """A reply promoted to a root by ``author=`` still reports its thread's paragraph."""
        doc = Document.open(clean_workspace, author="Parent Author")
        root_id = doc.add_comment("brown fox", "root")
        doc.save()
        doc.close()

        doc = Document.open(clean_workspace, author="Reply Author")
        middle_id = doc.reply_to_comment(root_id, "middle")
        doc.save()
        doc.close()

        doc = Document.open(clean_workspace, author="Parent Author")
        try:
            leaf_id = doc.reply_to_comment(middle_id, "leaf")
            expected = _paragraph_ref(doc, "brown fox")
            other = _find_paragraph_with_text(doc, "sample document")
            for reply_id in (middle_id, leaf_id):
                marker = _range_start(doc, reply_id)
                marker.parentNode.removeChild(marker)
                other.appendChild(marker)

            # Parent filtered out: the reply is listed as a root.
            (middle,) = doc.list_comments(author="Reply Author")
            assert (middle.id, middle.paragraph_ref) == (middle_id, expected)

            # Mid-thread gap: the leaf's parent is filtered out, its root is not.
            listed = {c.id: c.paragraph_ref for c in doc.list_comments(author="Parent Author")}
            assert listed == {root_id: expected, leaf_id: expected}
        finally:
            doc.close()

    def test_cyclic_reply_chain_terminates(self, clean_workspace):
        """A foreign ``paraIdParent`` cycle must not hang the listing."""
        doc = Document.open(clean_workspace)
        try:
            first_id = doc.add_comment("brown fox", "first")
            second_id = doc.reply_to_comment(first_id, "second")
            manager = doc._comment_manager
            para_ids = {cid: manager.existing_comments[cid]["para_id"] for cid in (first_id, second_id)}
            ext_dom = manager._get_editor(manager.comments_extended_path).dom
            for ex_elem in ext_dom.getElementsByTagName("w15:commentEx"):
                if ex_elem.getAttribute("w15:paraId") == para_ids[first_id]:
                    ex_elem.setAttribute("w15:paraIdParent", para_ids[second_id])

            # Returning at all is the assertion; which comments a cycle
            # leaves listed is not this test's business. If the cycle guard
            # regresses, this test hangs instead of failing.
            assert isinstance(doc.list_comments(), list)
        finally:
            doc.close()

    def test_survives_save_and_reopen(self, clean_workspace):
        doc = Document.open(clean_workspace)
        doc.add_comment("brown fox", "persisted")
        (before,) = doc.list_comments()
        doc.save()
        doc.close()

        doc = Document.open(clean_workspace)
        try:
            (after,) = doc.list_comments()
            assert before.paragraph_ref is not None
            assert after.paragraph_ref == before.paragraph_ref
        finally:
            doc.close()

    def test_reply_to_a_comment_without_a_range_start_is_none(self, clean_workspace):
        """The thread root decides, even when the reply's own marker is intact."""
        doc = Document.open(clean_workspace)
        try:
            parent_id = doc.add_comment("brown fox", "parent")
            reply_id = doc.reply_to_comment(parent_id, "reply")
            marker = _range_start(doc, parent_id)
            marker.parentNode.removeChild(marker)
            assert _range_start(doc, reply_id) is not None

            (parent,) = doc.list_comments()
            (reply,) = parent.replies
            assert parent.paragraph_ref is None
            assert reply.paragraph_ref is None
        finally:
            doc.close()

    def test_range_start_with_a_non_numeric_id_is_skipped(self, clean_workspace):
        """A nonconforming marker id names no comment and must not break the listing."""
        doc = Document.open(clean_workspace)
        try:
            kept_id = doc.add_comment("sample document", "kept")
            broken_id = doc.add_comment("brown fox", "broken marker")
            _range_start(doc, broken_id).setAttribute("w:id", "not-a-number")

            refs = {c.id: c.paragraph_ref for c in doc.list_comments()}
            assert refs == {kept_id: _paragraph_ref(doc, "sample document"), broken_id: None}
        finally:
            doc.close()

    def test_missing_range_start_is_none(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            kept_id = doc.add_comment("sample document", "kept")
            orphan_id = doc.add_comment("brown fox", "orphaned")
            marker = _range_start(doc, orphan_id)
            marker.parentNode.removeChild(marker)

            refs = {c.id: c.paragraph_ref for c in doc.list_comments()}
            assert refs == {kept_id: _paragraph_ref(doc, "sample document"), orphan_id: None}
        finally:
            doc.close()

    def test_range_start_outside_a_body_paragraph_is_none(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            comment_id = doc.add_comment("brown fox", "between blocks")
            marker = _range_start(doc, comment_id)
            marker.parentNode.removeChild(marker)
            body = doc._document_editor.dom.getElementsByTagName("w:body")[0]
            body.insertBefore(marker, body.firstChild)

            (comment,) = doc.list_comments()
            assert comment.paragraph_ref is None
        finally:
            doc.close()

    @pytest.mark.parametrize("in_box_paragraph", [True, False], ids=["box-paragraph", "box-content"])
    def test_range_start_inside_a_text_box_is_none(self, clean_workspace, in_box_paragraph):
        """Neither the box's own paragraph nor its host paragraph is reported."""
        doc = Document.open(clean_workspace)
        try:
            comment_id = doc.add_comment("brown fox", "boxed")
            marker = _range_start(doc, comment_id)
            host = marker.parentNode
            while host.tagName != "w:p":
                host = host.parentNode
            marker.parentNode.removeChild(marker)

            dom = doc._document_editor.dom
            parent = host
            for tag in ("w:r", "w:pict", "v:textbox", "w:txbxContent"):
                parent = parent.appendChild(dom.createElement(tag))
            if in_box_paragraph:
                parent = parent.appendChild(dom.createElement("w:p"))
            parent.appendChild(marker)

            (comment,) = doc.list_comments()
            assert comment.paragraph_ref is None
        finally:
            doc.close()

    def test_document_walks_do_not_grow_with_comment_count(self, clean_workspace, monkeypatch):
        doc = Document.open(clean_workspace)
        try:
            for i in range(6):
                doc.add_comment("brown fox", f"comment {i}")

            walks = count_dom_walks(monkeypatch)
            comments = doc.list_comments()

            assert len(comments) == 6
            assert all(c.paragraph_ref for c in comments)
            assert walks.count("w:p") == 1
            assert walks.count("w:commentRangeStart") == 1
        finally:
            doc.close()

    def test_repr_shows_the_ref_only_when_set(self):
        from docx_editor.comments import Comment

        located = Comment(id=1, text="Located", author="A", date=None, paragraph_ref="P3#b2c4")
        assert repr(located).endswith(" @P3#b2c4)")
        assert "@" not in repr(Comment(id=1, text="Unlocated", author="A", date=None))


FOX = "The quick brown fox jumps over the lazy dog."


def _commented_text(doc, comment_id: int) -> str:
    """Accepted-view text between ``comment_id``'s range start and range end."""
    out: list[str] = []
    inside = False

    def walk(node) -> None:
        nonlocal inside
        for child in node.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            name = child.tagName
            if name in ("w:commentRangeStart", "w:commentRangeEnd"):
                if child.getAttribute("w:id") == str(comment_id):
                    inside = name == "w:commentRangeStart"
            elif name == "w:t":
                if inside:
                    out.append("".join(c.data for c in child.childNodes if c.nodeType == c.TEXT_NODE))
            elif name == "w:tab":
                if inside:
                    out.append("\t")
            elif name != "w:del":
                walk(child)

    walk(doc._document_editor.dom.documentElement)
    return "".join(out)


def _marker_count(doc) -> int:
    dom = doc._document_editor.dom
    return sum(
        len(dom.getElementsByTagName(tag)) for tag in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference")
    )


class TestAddCommentWholeParagraph:
    """``add_comment(comment=..., paragraph=ref)`` comments on a whole paragraph (ROADMAP.md #89)."""

    def test_keyword_form_brackets_the_whole_paragraph(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            ref = find_ref(doc, "brown fox")
            comment_id = doc.add_comment(comment="Needs a citation", paragraph=ref)

            assert _commented_text(doc, comment_id) == FOX
            assert _marker_count(doc) == 3
            # Markers are not text: the paragraph still answers to the same ref.
            assert find_ref(doc, "brown fox") == ref
            (comment,) = doc.list_comments()
            assert (comment.id, comment.text, comment.paragraph_ref) == (comment_id, "Needs a citation", ref)
        finally:
            doc.close()

    def test_explicit_none_anchor_is_the_same_form(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            comment_id = doc.add_comment(None, "Needs a citation", paragraph=find_ref(doc, "brown fox"))
            assert _commented_text(doc, comment_id) == FOX
        finally:
            doc.close()

    def test_same_markup_as_anchoring_on_the_full_text(self, clean_workspace, tmp_path):
        """The whole-paragraph form is the anchored form with the full text as the anchor."""
        import re
        import shutil

        def markup(doc) -> str:
            # w:rsidR is random per session.
            return re.sub(r'w:rsidR="[0-9A-F]+"', "", _find_paragraph_with_text(doc, "brown fox").toxml())

        other = tmp_path / "other.docx"
        shutil.copy(clean_workspace, other)
        whole = Document.open(clean_workspace)
        anchored = Document.open(other)
        try:
            ref = find_ref(whole, "brown fox")
            whole.add_comment(comment="c", paragraph=ref)
            anchored.add_comment(FOX, "c", paragraph=ref)
            assert markup(whole) == markup(anchored)
        finally:
            whole.close()
            anchored.close()

    def test_paragraph_split_across_runs(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            para = _find_paragraph_with_text(doc, "brown fox")
            _split_first_run_text(para, FOX.index("brown") + 2)
            ref = find_ref(doc, "brown fox")

            comment_id = doc.add_comment(comment="spans runs", paragraph=ref)

            assert _commented_text(doc, comment_id) == FOX
            assert len(para.getElementsByTagName("w:commentRangeStart")) == 1
            assert len(para.getElementsByTagName("w:commentRangeEnd")) == 1
        finally:
            doc.close()

    @staticmethod
    def _with_edge_insertions(path) -> tuple[Document, str]:
        """Open ``path`` with tracked insertions at the very start and end of the fox paragraph."""
        doc = Document.open(path)
        doc.insert_before("The quick", "START ", paragraph=find_ref(doc, "brown fox"))
        doc.insert_after("lazy dog.", " END", paragraph=find_ref(doc, "brown fox"))
        return doc, find_ref(doc, "brown fox")

    @staticmethod
    def _resolve_edge(doc, action: str) -> None:
        start_ins, end_ins = doc.list_revisions()
        if action in ("accept_all", "reject_all"):
            getattr(doc, action)()
        else:
            verb, edge = action.split("_")
            getattr(doc, f"{verb}_revision")((start_ins if edge == "start" else end_ins).id)

    @staticmethod
    def _marker_parents(doc) -> list[str | None]:
        """Parent tag of the range start, the range end and the reference run (None = removed)."""
        dom = doc._document_editor.dom
        out: list[str | None] = []
        for tag in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"):
            found = dom.getElementsByTagName(tag)
            node = found[0] if found else None
            if node is not None and tag == "w:commentReference":
                node = node.parentNode
            out.append(node.parentNode.tagName if node is not None else None)
        return out

    def test_covers_tracked_insertions_at_both_ends(self, clean_workspace):
        doc, ref = self._with_edge_insertions(clean_workspace)
        try:
            comment_id = doc.add_comment(comment="all of it", paragraph=ref)
            assert _commented_text(doc, comment_id) == f"START {FOX} END"
            # The edge markers sit inside the w:ins wrappers, as they do for
            # any anchor that starts or ends in inserted text.
            assert self._marker_parents(doc) == ["w:ins", "w:ins", "w:ins"]
        finally:
            doc.close()

    # What resolving the edge insertions does to the markers: (start parent,
    # end parent, reference-run parent, paragraph_ref still set). A marker
    # inside a w:ins is unwrapped with it on accept and hoisted out of it (to
    # where the insertion was) on reject, so no edge loses its marker.
    EDGE_OUTCOMES = {
        "accept_all": ("w:p", "w:p", "w:p", True),
        "accept_start": ("w:p", "w:ins", "w:ins", True),
        "accept_end": ("w:ins", "w:p", "w:p", True),
        "reject_all": ("w:p", "w:p", "w:p", True),
        "reject_start": ("w:p", "w:ins", "w:ins", True),
        "reject_end": ("w:ins", "w:p", "w:p", True),
    }

    @pytest.mark.parametrize("action", list(EDGE_OUTCOMES))
    def test_resolving_edge_insertions_matches_the_anchored_form(self, clean_workspace, tmp_path, action):
        """Accepting or rejecting the edge insertions treats a whole-paragraph
        comment exactly as it treats one anchored on the same full text.

        The markers live inside the edge ``w:ins`` wrappers: accept unwraps
        them into the paragraph, reject hoists them out before the insertion
        goes. The comment stays listed with its ``paragraph_ref`` either way.
        """
        import shutil

        other = tmp_path / "other.docx"
        shutil.copy(clean_workspace, other)
        whole, ref = self._with_edge_insertions(clean_workspace)
        anchored, anchored_ref = self._with_edge_insertions(other)
        try:
            whole.add_comment(comment="c", paragraph=ref)
            anchored.add_comment(f"START {FOX} END", "c", paragraph=anchored_ref)

            self._resolve_edge(whole, action)
            self._resolve_edge(anchored, action)

            *parents, has_ref = self.EDGE_OUTCOMES[action]
            assert self._marker_parents(whole) == parents
            assert self._marker_parents(anchored) == parents
            (comment,) = whole.list_comments()
            (anchored_comment,) = anchored.list_comments()
            expected_ref = find_ref(whole, "brown fox") if has_ref else None
            assert comment.paragraph_ref == expected_ref
            assert anchored_comment.paragraph_ref == expected_ref
            whole.save()
        finally:
            whole.close()
            anchored.close()

        doc = Document.open(clean_workspace)
        try:
            (comment,) = doc.list_comments()
            assert comment.text == "c"
            assert self._marker_parents(doc) == parents
            assert comment.paragraph_ref == (find_ref(doc, "brown fox") if has_ref else None)
        finally:
            doc.close()

    def test_paragraph_starting_and_ending_with_a_tab(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            para = _find_paragraph_with_text(doc, "brown fox")
            doc._document_editor.replace_node(
                para.getElementsByTagName("w:r")[0],
                '<w:r><w:tab/><w:t xml:space="preserve">between tabs</w:t><w:tab/></w:r>',
            )
            ref = find_ref(doc, "between tabs")

            comment_id = doc.add_comment(comment="tabs too", paragraph=ref)

            assert _commented_text(doc, comment_id) == "\tbetween tabs\t"
            assert len(para.getElementsByTagName("w:tab")) == 2
        finally:
            doc.close()

    def test_excludes_tracked_deletions(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            doc.delete("quick ", paragraph=find_ref(doc, "brown fox"))
            ref = find_ref(doc, "brown fox")
            comment_id = doc.add_comment(comment="rest", paragraph=ref)
            assert _commented_text(doc, comment_id) == FOX.replace("quick ", "")
        finally:
            doc.close()

    def _assert_refused_untouched(self, doc, ref):
        before = doc._document_editor.dom.toxml()
        next_id = doc._comment_manager.next_comment_id
        with pytest.raises(CommentError, match="has no text to anchor a comment on") as exc:
            doc.add_comment(comment="nothing here", paragraph=ref)
        assert ref in str(exc.value)
        assert doc._document_editor.dom.toxml() == before
        assert doc._comment_manager.next_comment_id == next_id
        assert doc.list_comments() == []

    def test_empty_paragraph_is_refused_before_any_mutation(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            para = _find_paragraph_with_text(doc, "brown fox")
            doc._document_editor.insert_after(para, "<w:p/>")
            empty = doc.list_paragraphs_structured(limit=None)[2]
            assert empty.text == ""
            self._assert_refused_untouched(doc, empty.ref)
        finally:
            doc.close()

    def test_paragraph_with_only_deleted_text_is_refused(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            doc.delete(FOX, paragraph=find_ref(doc, "brown fox"))
            gone = doc.list_paragraphs_structured(limit=None)[1]
            assert gone.text == ""
            self._assert_refused_untouched(doc, gone.ref)
        finally:
            doc.close()

    def test_whitespace_only_paragraph_is_commentable(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            para = _find_paragraph_with_text(doc, "brown fox")
            doc._document_editor.insert_after(para, '<w:p><w:r><w:t xml:space="preserve">  </w:t></w:r></w:p>')
            blank = doc.list_paragraphs_structured(limit=None)[2]
            comment_id = doc.add_comment(comment="blank", paragraph=blank.ref)
            assert _commented_text(doc, comment_id) == "  "
        finally:
            doc.close()

    def test_none_anchor_without_paragraph_raises(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            with pytest.raises(CommentError, match="pass paragraph="):
                doc.add_comment(comment="where?")
            with pytest.raises(CommentError, match="pass paragraph="):
                doc.add_comment(None, "where?")
            assert _marker_count(doc) == 0
        finally:
            doc.close()

    def test_occurrence_with_none_anchor_raises(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            with pytest.raises(ValueError, match="occurrence="):
                doc.add_comment(comment="c", paragraph=find_ref(doc, "brown fox"), occurrence=0)
            assert _marker_count(doc) == 0
        finally:
            doc.close()

    def test_missing_comment_raises_naming_the_parameter(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            ref = find_ref(doc, "brown fox")
            for call in (
                lambda: doc.add_comment(paragraph=ref),
                lambda: doc.add_comment("brown fox"),
                lambda: doc.add_comment("brown fox", None),
                # A single positional string is the anchor, never the comment.
                lambda: doc.add_comment("Needs a citation", paragraph=ref),
            ):
                with pytest.raises(CommentError, match="comment is required"):
                    call()
            assert _marker_count(doc) == 0
            assert doc.list_comments() == []
        finally:
            doc.close()

    def test_manager_rejects_a_missing_comment(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            ref = find_ref(doc, "brown fox")
            with pytest.raises(CommentError, match="comment_text must be a string"):
                doc._comment_manager.add_comment(None, None, paragraph=ref)
            with pytest.raises(CommentError, match="comment_text must be a string"):
                doc._comment_manager.add_comment("brown fox", None)
            assert _marker_count(doc) == 0
        finally:
            doc.close()

    def test_bad_comment_body_is_refused_before_any_mutation(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            with pytest.raises(CommentError):
                doc.add_comment(comment="line\nbreak", paragraph=find_ref(doc, "brown fox"))
            assert _marker_count(doc) == 0
        finally:
            doc.close()

    def test_stale_hash_and_out_of_range_index(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            with pytest.raises(HashMismatchError):
                doc.add_comment(comment="c", paragraph="P2#0000")
            with pytest.raises(ParagraphIndexError):
                doc.add_comment(comment="c", paragraph="P999#0000")
            assert _marker_count(doc) == 0
        finally:
            doc.close()

    def test_empty_string_anchor_is_not_whole_paragraph(self, clean_workspace):
        doc = Document.open(clean_workspace)
        try:
            with pytest.raises(CommentError, match="anchor_text must be a non-empty string"):
                doc.add_comment("", "c", paragraph=find_ref(doc, "brown fox"))
        finally:
            doc.close()

    def test_survives_save_and_reopen(self, clean_workspace):
        doc = Document.open(clean_workspace)
        ref = find_ref(doc, "brown fox")
        comment_id = doc.add_comment(comment="persisted", paragraph=ref)
        doc.save()
        doc.close()

        doc = Document.open(clean_workspace)
        try:
            (comment,) = doc.list_comments()
            assert (comment.id, comment.text, comment.paragraph_ref) == (comment_id, "persisted", ref)
            assert _commented_text(doc, comment_id) == FOX
        finally:
            doc.close()


_COMMENT_TAGS = ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference")
_W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_ANN = 'w:author="Ann" w:date="2026-01-29T16:55:00Z"'


def _markers_by_id(doc) -> dict[str, list[str]]:
    """Comment id -> the marker tags it has in the document, in document order."""
    out: dict[str, list[str]] = {}
    for node in doc._document_editor.dom.getElementsByTagName("*"):
        if node.tagName in _COMMENT_TAGS:
            out.setdefault(node.getAttribute("w:id"), []).append(node.tagName)
    return out


def _marker_nodes(doc) -> list:
    """The range marks and reference runs, in document order."""
    nodes = []
    for node in doc._document_editor.dom.getElementsByTagName("*"):
        if node.tagName in ("w:commentRangeStart", "w:commentRangeEnd"):
            nodes.append(node)
        elif node.tagName == "w:commentReference":
            nodes.append(node.parentNode)
    return nodes


def _open_body(clean_workspace, tmp_path, paragraph: str) -> Document:
    """Open a copy of ``clean_workspace`` whose body is the single ``paragraph``."""
    from conftest import replace_document_xml

    dest = tmp_path / "hand.docx"
    replace_document_xml(
        clean_workspace,
        dest,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_W_NS}><w:body>{paragraph}</w:body></w:document>",
    )
    return Document.open(dest)


class TestResolutionKeepsCommentMarkers:
    """Resolving a revision never drops the comment markers inside it
    (ROADMAP.md #102): rejecting an insertion, or accepting a deletion, moves
    them to where the removed content was before it goes."""

    INSERTED = " EXTRA words"

    def _with_insertion(self, path):
        doc = Document.open(path)
        result = doc.insert_after("lazy dog.", self.INSERTED, paragraph=find_ref(doc, "brown fox"))
        return doc, result

    @staticmethod
    def _assert_point_comment(doc, comment_id: int) -> None:
        """One start, end and reference, adjacent children of the paragraph."""
        assert _markers_by_id(doc) == {str(comment_id): list(_COMMENT_TAGS)}
        start, end, ref_run = _marker_nodes(doc)
        assert start.parentNode.tagName == "w:p"
        assert end.parentNode is start.parentNode and ref_run.parentNode is start.parentNode
        siblings = [n for n in start.parentNode.childNodes if n.nodeType == n.ELEMENT_NODE]
        i = siblings.index(start)
        assert siblings[i : i + 3] == [start, end, ref_run]

    @pytest.mark.parametrize("verb", ["reject_revision", "reject_all", "reject_group", "reject_changeset"])
    def test_comment_wholly_inside_a_rejected_insertion_becomes_a_point_comment(self, clean_workspace, verb):
        doc, result = self._with_insertion(clean_workspace)
        try:
            comment_id = doc.add_comment("EXTRA words", "keep me", paragraph=find_ref(doc, "brown fox"))
            if verb == "reject_revision":
                (rev,) = doc.list_revisions()
                assert doc.reject_revision(rev.id)
            elif verb == "reject_all":
                doc.reject_all()
            elif verb == "reject_group":
                assert doc.reject_group(result.group_id) == 1
            else:
                assert doc.reject_changeset(result.changeset_id) == 1

            assert doc.list_revisions() == []
            assert "EXTRA" not in doc._document_editor.dom.toxml()
            self._assert_point_comment(doc, comment_id)
            ref = find_ref(doc, "brown fox")
            (comment,) = doc.list_comments()
            assert (comment.id, comment.text, comment.paragraph_ref) == (comment_id, "keep me", ref)
            doc.save()
        finally:
            doc.close()

        doc = Document.open(clean_workspace)
        try:
            self._assert_point_comment(doc, comment_id)
            (comment,) = doc.list_comments()
            assert comment.paragraph_ref == find_ref(doc, "brown fox")
        finally:
            doc.close()

    def test_straddling_anchor_is_clipped_to_the_surviving_text(self, clean_workspace):
        doc, _ = self._with_insertion(clean_workspace)
        try:
            comment_id = doc.add_comment("dog. EXTRA", "straddle", paragraph=find_ref(doc, "brown fox"))
            doc.reject_all()
            assert _markers_by_id(doc) == {str(comment_id): list(_COMMENT_TAGS)}
            assert _commented_text(doc, comment_id) == "dog."
        finally:
            doc.close()

    def test_reply_keeps_its_markers_with_its_parent(self, clean_workspace):
        doc, _ = self._with_insertion(clean_workspace)
        try:
            parent_id = doc.add_comment("EXTRA words", "parent", paragraph=find_ref(doc, "brown fox"))
            reply_id = doc.reply_to_comment(parent_id, "reply")
            doc.reject_all()
            markers = _markers_by_id(doc)
            assert sorted(markers) == sorted([str(parent_id), str(reply_id)])
            for tags in markers.values():
                assert sorted(tags) == sorted(_COMMENT_TAGS)
            (parent,) = doc.list_comments()
            assert parent.paragraph_ref == find_ref(doc, "brown fox")
            assert [r.text for r in parent.replies] == ["reply"]
        finally:
            doc.close()

    def test_accepting_keeps_the_markers_where_they_were(self, clean_workspace):
        doc, _ = self._with_insertion(clean_workspace)
        try:
            comment_id = doc.add_comment("EXTRA words", "accepted", paragraph=find_ref(doc, "brown fox"))
            doc.accept_all()
            assert _markers_by_id(doc) == {str(comment_id): list(_COMMENT_TAGS)}
            assert _commented_text(doc, comment_id) == "EXTRA words"
        finally:
            doc.close()

    def test_rejected_foreign_move_to_keeps_its_markers(self, clean_workspace, tmp_path):
        doc = _open_body(
            clean_workspace,
            tmp_path,
            '<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r>'
            f'<w:moveTo w:id="14" {_ANN}><w:commentRangeStart w:id="0"/><w:r><w:t>moved</w:t></w:r>'
            '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:moveTo></w:p>',
        )
        try:
            assert doc.reject_revision(14)
            xml = doc._document_editor.dom.toxml()
            assert "moved" not in xml and "w:moveTo" not in xml
            assert _markers_by_id(doc) == {"0": list(_COMMENT_TAGS)}
            assert {n.parentNode.tagName for n in _marker_nodes(doc)} == {"w:p"}
        finally:
            doc.close()

    def test_accepted_deletion_keeps_its_markers(self, clean_workspace, tmp_path):
        doc = _open_body(
            clean_workspace,
            tmp_path,
            '<w:p><w:r><w:t xml:space="preserve">kept </w:t></w:r>'
            f'<w:del w:id="5" {_ANN}><w:commentRangeStart w:id="0"/><w:r><w:delText>gone</w:delText></w:r>'
            '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:del></w:p>',
        )
        try:
            assert doc.accept_revision(5)
            xml = doc._document_editor.dom.toxml()
            assert "gone" not in xml and "w:del " not in xml
            assert _markers_by_id(doc) == {"0": list(_COMMENT_TAGS)}
            assert {n.parentNode.tagName for n in _marker_nodes(doc)} == {"w:p"}
        finally:
            doc.close()

    def test_reference_in_a_run_with_text_does_not_resurrect_the_text(self, clean_workspace, tmp_path):
        doc = _open_body(
            clean_workspace,
            tmp_path,
            '<w:p><w:r><w:t xml:space="preserve">kept </w:t></w:r>'
            f'<w:ins w:id="5" {_ANN}><w:commentRangeStart w:id="0"/>'
            '<w:r><w:t>new</w:t><w:commentRangeEnd w:id="0"/><w:commentReference w:id="0"/></w:r></w:ins></w:p>',
        )
        try:
            assert doc.reject_revision(5)
            xml = doc._document_editor.dom.toxml()
            assert "new" not in xml and "w:ins" not in xml
            assert _markers_by_id(doc) == {"0": list(_COMMENT_TAGS)}
            (ref,) = doc._document_editor.dom.getElementsByTagName("w:commentReference")
            run = ref.parentNode
            assert run.tagName == "w:r" and run.parentNode.tagName == "w:p"
            assert [n.tagName for n in run.childNodes if n.nodeType == n.ELEMENT_NODE] == ["w:commentReference"]
        finally:
            doc.close()

    @staticmethod
    def _assert_nothing_at_block_level(doc) -> None:
        """No run and no comment mark is a direct child of w:body or w:tc."""
        for node in doc._document_editor.dom.getElementsByTagName("*"):
            if node.tagName in ("w:r", "w:commentRangeStart", "w:commentRangeEnd"):
                assert node.parentNode.tagName not in ("w:body", "w:tc"), node.toxml()

    def _reject_block_insertion(self, clean_workspace, tmp_path, body: str) -> Document:
        """Open ``body`` (holding a block-level w:ins id 7 around a paragraph
        "inserted para"), comment that paragraph, and reject the insertion."""
        doc = _open_body(clean_workspace, tmp_path, body)
        doc.add_comment("inserted para", "on the inserted paragraph")
        assert doc.reject_revision(7)
        assert "inserted para" not in doc._document_editor.dom.toxml()
        return doc

    _BLOCK_INS = f'<w:ins w:id="7" {_ANN}><w:p><w:r><w:t>inserted para</w:t></w:r></w:p></w:ins>'

    def test_block_level_insertion_rehomes_markers_to_the_previous_paragraph_end(self, clean_workspace, tmp_path):
        doc = self._reject_block_insertion(
            clean_workspace, tmp_path, "<w:p><w:r><w:t>before</w:t></w:r></w:p>" + self._BLOCK_INS
        )
        try:
            self._assert_nothing_at_block_level(doc)
            (paragraph,) = doc._document_editor.dom.getElementsByTagName("w:p")
            children = [n for n in paragraph.childNodes if n.nodeType == n.ELEMENT_NODE]
            assert children[-3:] == _marker_nodes(doc)  # point comment at the end
            assert [c.id for c in doc.list_comments()] == [0]
        finally:
            doc.close()

    def test_block_level_insertion_in_a_cell_rehomes_markers_to_the_next_paragraph_start(
        self, clean_workspace, tmp_path
    ):
        doc = self._reject_block_insertion(
            clean_workspace,
            tmp_path,
            "<w:tbl><w:tr><w:tc>"
            + self._BLOCK_INS
            + '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>after</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
        )
        try:
            self._assert_nothing_at_block_level(doc)
            (paragraph,) = doc._document_editor.dom.getElementsByTagName("w:p")
            children = [n for n in paragraph.childNodes if n.nodeType == n.ELEMENT_NODE]
            assert [children[0].tagName] + children[1:4] == ["w:pPr", *_marker_nodes(doc)]
        finally:
            doc.close()

    def test_block_level_insertion_with_no_surviving_paragraph_drops_the_whole_marker_set(
        self, clean_workspace, tmp_path
    ):
        # The range starts outside the cell; with no paragraph left beside the
        # insertion, the start goes too rather than survive as a half-pair.
        doc = _open_body(
            clean_workspace,
            tmp_path,
            '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>outside</w:t></w:r></w:p>'
            f'<w:tbl><w:tr><w:tc><w:ins w:id="7" {_ANN}><w:p><w:r><w:t>inside</w:t></w:r>'
            '<w:commentRangeEnd w:id="0"/><w:r><w:rPr/><w:commentReference w:id="0"/></w:r></w:p>'
            "</w:ins></w:tc></w:tr></w:tbl>",
        )
        try:
            assert doc.reject_revision(7)
            assert _markers_by_id(doc) == {}
            assert "outside" in doc._document_editor.dom.toxml()
        finally:
            doc.close()

    _START = '<w:commentRangeStart w:id="0"/>'
    _END = '<w:commentRangeEnd w:id="0"/><w:r><w:rPr/><w:commentReference w:id="0"/></w:r>'
    _SPANNING = {
        # A range starting in block insertion 1 and ending in block insertion 2.
        "two_insertions": "<w:p><w:r><w:t>P0</w:t></w:r></w:p>"
        f'<w:ins w:id="1" {_ANN}><w:p>{_START}<w:r><w:t>a</w:t></w:r></w:p></w:ins>'
        f'<w:ins w:id="2" {_ANN}><w:p><w:r><w:t>b</w:t></w:r>{_END}</w:p></w:ins>'
        "<w:p><w:r><w:t>P3</w:t></w:r></w:p>",
        "two_insertions_no_previous": f'<w:ins w:id="1" {_ANN}><w:p>{_START}<w:r><w:t>a</w:t></w:r></w:p></w:ins>'
        f'<w:ins w:id="2" {_ANN}><w:p><w:r><w:t>b</w:t></w:r>{_END}</w:p></w:ins>'
        "<w:p><w:r><w:t>P3</w:t></w:r></w:p>",
        # The range starts in a table cell between P0 and the insertion.
        "start_in_a_table": "<w:p><w:r><w:t>P0</w:t></w:r></w:p>"
        f"<w:tbl><w:tr><w:tc><w:p>{_START}<w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        f'<w:ins w:id="2" {_ANN}><w:p><w:r><w:t>b</w:t></w:r>{_END}</w:p></w:ins>',
        # The insertion is alone in a block content control: the paragraph
        # outside it still receives the markers.
        "inside_a_block_sdt": "<w:p><w:r><w:t>P0</w:t></w:r></w:p>"
        f'<w:sdt><w:sdtContent><w:ins w:id="2" {_ANN}><w:p>{_START}<w:r><w:t>b</w:t></w:r>{_END}</w:p></w:ins>'
        "</w:sdtContent></w:sdt>",
    }

    @pytest.mark.parametrize("body", sorted(_SPANNING))
    @pytest.mark.parametrize("order", ["reject_all", "ascending", "descending"])
    def test_rehomed_range_keeps_its_start_before_its_end(self, clean_workspace, tmp_path, body, order):
        doc = _open_body(clean_workspace, tmp_path, self._SPANNING[body])
        try:
            ids = sorted(rev.id for rev in doc.list_revisions())
            if order == "reject_all":
                doc.reject_all()
            else:
                for rev_id in ids if order == "ascending" else reversed(ids):
                    assert doc.reject_revision(rev_id)
            assert doc.list_revisions() == []
            assert _markers_by_id(doc) == {"0": list(_COMMENT_TAGS)}
            self._assert_nothing_at_block_level(doc)
        finally:
            doc.close()

    def test_text_box_fallback_copy_is_not_hoisted_twice(self, clean_workspace, tmp_path):
        box = f"<w:txbxContent><w:p>{self._START}<w:r><w:t>box</w:t></w:r>{self._END}</w:p></w:txbxContent>"
        doc = _open_body(
            clean_workspace,
            tmp_path,
            "<w:p><w:r><w:t>keep</w:t></w:r>"
            f'<w:ins w:id="1" {_ANN}><w:r>'
            '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            f'<mc:Choice Requires="wps"><w:drawing>{box}</w:drawing></mc:Choice>'
            f"<mc:Fallback><w:pict>{box}</w:pict></mc:Fallback>"
            "</mc:AlternateContent></w:r></w:ins></w:p>",
        )
        try:
            assert doc.reject_revision(1)
            assert _markers_by_id(doc) == {"0": list(_COMMENT_TAGS)}
        finally:
            doc.close()

    @pytest.mark.skipif(__import__("shutil").which("soffice") is None, reason="LibreOffice not installed")
    def test_libreoffice_opens_a_rejected_block_level_insertion(self, clean_workspace, tmp_path, monkeypatch):
        import shutil

        from test_corpus_harness import harness

        doc = self._reject_block_insertion(
            clean_workspace, tmp_path, "<w:p><w:r><w:t>before</w:t></w:r></w:p>" + self._BLOCK_INS
        )
        try:
            out = tmp_path / "out.docx"
            doc.save(out)
        finally:
            doc.close()

        monkeypatch.setattr(harness, "LO_DIR", tmp_path / "lo")
        (tmp_path / "lo").mkdir()
        monkeypatch.setattr(harness, "LO_PROFILE", tmp_path / "profile")
        record = harness.run_lo_roundtrip(out, False, shutil.which("soffice"), tmp_path / "work")
        assert record["status"] == "pass", record

    @pytest.mark.skipif(__import__("shutil").which("soffice") is None, reason="LibreOffice not installed")
    def test_libreoffice_keeps_the_comment_after_rejecting_the_trailing_edge(
        self, clean_workspace, tmp_path, monkeypatch
    ):
        import shutil
        import zipfile

        from test_corpus_harness import harness

        doc = Document.open(clean_workspace)
        try:
            doc.insert_before("The quick", "START ", paragraph=find_ref(doc, "brown fox"))
            doc.insert_after("lazy dog.", " END", paragraph=find_ref(doc, "brown fox"))
            doc.add_comment(comment="whole paragraph", paragraph=find_ref(doc, "brown fox"))
            _, end_ins = doc.list_revisions()
            doc.reject_revision(end_ins.id)
            out = tmp_path / "out.docx"
            doc.save(out)
        finally:
            doc.close()

        lo_dir = tmp_path / "lo"
        lo_dir.mkdir()
        monkeypatch.setattr(harness, "LO_DIR", lo_dir)
        monkeypatch.setattr(harness, "LO_PROFILE", tmp_path / "profile")
        record = harness.run_lo_roundtrip(out, False, shutil.which("soffice"), tmp_path / "work")
        assert record["status"] == "pass", record

        (resaved,) = lo_dir.glob("*.docx")
        with zipfile.ZipFile(resaved) as z:
            assert "whole paragraph" in z.read("word/comments.xml").decode()
            body = z.read("word/document.xml").decode()
        assert "commentRangeStart" in body and "commentReference" in body
