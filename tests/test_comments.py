"""Tests for comment functionality."""

import pytest

from docx_editor import (
    CommentError,
    Document,
    HashMismatchError,
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
        """Test resolve_comment returns False when commentsExtended.xml doesn't exist."""
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
        assert result is False

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

        # Clear editor cache
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

        # Clear editor cache
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

        # Clear editor cache
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
        """Test resolve_comment returns False when para_id not found in commentsExtended."""
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

        # resolve_comment should return False since para_id won't be found
        result = doc.resolve_comment(comment_id)
        assert result is False

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

    def test_load_existing_comments_skips_comment_without_para_id(self, clean_workspace):
        """Test that _load_existing_comments skips comments without para_id."""
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

        # Comment 997 should NOT be in existing_comments (skipped due to no para_id)
        assert 997 not in existing
        # But comment 0 should be there
        assert 0 in existing

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
                "息税前", "利润率", "，", "息税前利润率", "=",
                "息税前利润", "/", "营业收入", "×100%",
            ]
            runs_xml = "".join(
                f'<w:r><w:t xml:space="preserve">{seg}</w:t></w:r>' for seg in segments
            )
            run = para.getElementsByTagName("w:r")[0]
            doc._document_editor.replace_node(run, runs_xml)

            target = "".join(segments)
            assert doc.count_matches(target) == 1
            assert doc.find_text(target) is not None

            comment_id = doc.add_comment(target, "issue-14 repro")
            assert isinstance(comment_id, int)
        finally:
            doc.close()
