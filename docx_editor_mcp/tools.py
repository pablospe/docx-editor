"""MCP Tools for docx_editor operations."""

import os
from typing import Any, cast

from docx_editor import Document
from docx_editor.exceptions import TextNotFoundError

from .cache import CachedDocument, normalize_path
from .server import DocxMCPServer


def _get_cached_or_error(server: DocxMCPServer, path: str) -> CachedDocument | dict[str, Any]:
    """Get cached document or return error dict.

    Returns:
        CachedDocument if found, or error dict if not open.
    """
    normalized = normalize_path(path)
    cached = server.cache.get(normalized)
    if not cached:
        error: dict[str, Any] = {"success": False, "error": f"Document not open: {path}"}
        return error
    return cached


def _check_external_changes(cached: CachedDocument) -> dict | None:
    """Check for external changes, return error dict if detected."""
    if cached.has_external_changes():
        return {
            "success": False,
            "error": f"File was modified externally: {cached.path}. Use reload_document or force_save.",
        }
    return None


# =============================================================================
# Document Lifecycle Tools (Task 3.1)
# =============================================================================


def open_document(
    server: DocxMCPServer,
    path: str,
    author: str | None = None,
) -> dict[str, Any]:
    """Open a document and add to cache.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        author: Author name for track changes.

    Returns:
        Result dict with success, path, author, and optional hint.
    """
    normalized = normalize_path(path)

    # Check if file exists
    if not os.path.exists(normalized):
        return {"success": False, "error": f"File not found: {path}"}

    # Check if already cached
    cached = server.cache.get(normalized)
    if cached:
        return {
            "success": True,
            "path": normalized,
            "author": cached.author,
            "cached": True,
        }

    # Get author with session memory
    resolved_author, is_default = server.cache.get_author(author)

    # Open document
    try:
        doc = Document.open(normalized, author=resolved_author)
    except Exception as e:
        return {"success": False, "error": f"Failed to open document: {e}"}

    # Cache it
    cached = CachedDocument(path=normalized, document=doc, author=resolved_author)
    server.cache.put(cached)

    result = {
        "success": True,
        "path": normalized,
        "author": resolved_author,
    }

    # Add hint if using system default
    if is_default:
        result["hint"] = f"Author set to '{resolved_author}' (system default). Use author parameter to change."

    return result


def save_document(server: DocxMCPServer, path: str) -> dict[str, Any]:
    """Save a document to disk.

    Args:
        server: The MCP server instance.
        path: Path to the document.

    Returns:
        Result dict with success status.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    # Check for external changes
    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        cached.document.save()
        cached.clear_dirty()
        cached.update_mtime()
        return {"success": True, "path": cached.path}
    except Exception as e:
        return {"success": False, "error": f"Failed to save: {e}"}


def close_document(server: DocxMCPServer, path: str) -> dict[str, Any]:
    """Close a document and remove from cache.

    Args:
        server: The MCP server instance.
        path: Path to the document.

    Returns:
        Result dict with success status and optional warning.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    result = {"success": True, "path": cached.path}

    if cached.dirty:
        result["warning"] = "Document had unsaved changes that were discarded."

    # Close the document
    try:
        cached.document.close()
    except Exception:
        pass  # Best effort

    server.cache.remove(path)
    return result


def reload_document(server: DocxMCPServer, path: str) -> dict[str, Any]:
    """Reload a document from disk, discarding cached changes.

    Args:
        server: The MCP server instance.
        path: Path to the document.

    Returns:
        Result dict with success status and optional warning.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    result = {"success": True, "path": cached.path}
    author = cached.author
    was_dirty = cached.dirty

    # Close old document
    try:
        cached.document.close()
    except Exception:
        pass

    # Remove from cache
    server.cache.remove(path)

    # Reopen
    try:
        doc = Document.open(cached.path, author=author)
    except Exception as e:
        return {"success": False, "error": f"Failed to reload: {e}"}

    # Add back to cache
    new_cached = CachedDocument(path=cached.path, document=doc, author=author)
    server.cache.put(new_cached)

    if was_dirty:
        result["warning"] = "Unsaved changes were discarded."

    return result


def force_save(server: DocxMCPServer, path: str) -> dict[str, Any]:
    """Force save a document, ignoring external changes.

    Args:
        server: The MCP server instance.
        path: Path to the document.

    Returns:
        Result dict with success status.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    try:
        cached.document.save()
        cached.clear_dirty()
        cached.update_mtime()
        return {"success": True, "path": cached.path}
    except Exception as e:
        return {"success": False, "error": f"Failed to save: {e}"}


# =============================================================================
# Track Changes Tools (Task 3.2)
# =============================================================================


def replace_text(
    server: DocxMCPServer,
    path: str,
    old_text: str,
    new_text: str,
    occurrence: int = 0,
) -> dict[str, Any]:
    """Replace text with tracked changes.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        old_text: Text to find and replace.
        new_text: Replacement text.
        occurrence: Which occurrence to replace (0 = first).

    Returns:
        Result dict with success status and change_id.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        change_id = cached.document.replace(old_text, new_text, occurrence=occurrence)
        cached.mark_dirty()
        return {"success": True, "change_id": change_id}
    except TextNotFoundError:
        return {"success": False, "error": f"Text not found: '{old_text}'"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def delete_text(
    server: DocxMCPServer,
    path: str,
    text: str,
    occurrence: int = 0,
) -> dict[str, Any]:
    """Delete text with tracked changes.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        text: Text to delete.
        occurrence: Which occurrence to delete (0 = first).

    Returns:
        Result dict with success status and change_id.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        change_id = cached.document.delete(text, occurrence=occurrence)
        cached.mark_dirty()
        return {"success": True, "change_id": change_id}
    except TextNotFoundError:
        return {"success": False, "error": f"Text not found: '{text}'"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def insert_after(
    server: DocxMCPServer,
    path: str,
    anchor: str,
    text: str,
    occurrence: int = 0,
) -> dict[str, Any]:
    """Insert text after anchor with tracked changes.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        anchor: Text to find as insertion point.
        text: Text to insert.
        occurrence: Which occurrence of anchor to use (0 = first).

    Returns:
        Result dict with success status and change_id.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        change_id = cached.document.insert_after(anchor, text, occurrence=occurrence)
        cached.mark_dirty()
        return {"success": True, "change_id": change_id}
    except TextNotFoundError:
        return {"success": False, "error": f"Anchor not found: '{anchor}'"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def insert_before(
    server: DocxMCPServer,
    path: str,
    anchor: str,
    text: str,
    occurrence: int = 0,
) -> dict[str, Any]:
    """Insert text before anchor with tracked changes.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        anchor: Text to find as insertion point.
        text: Text to insert.
        occurrence: Which occurrence of anchor to use (0 = first).

    Returns:
        Result dict with success status and change_id.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        change_id = cached.document.insert_before(anchor, text, occurrence=occurrence)
        cached.mark_dirty()
        return {"success": True, "change_id": change_id}
    except TextNotFoundError:
        return {"success": False, "error": f"Anchor not found: '{anchor}'"}
    except Exception as e:
        return {"success": False, "error": str(e)}


# =============================================================================
# Comment Tools (Task 3.3)
# =============================================================================


def add_comment(
    server: DocxMCPServer,
    path: str,
    anchor_text: str,
    comment_text: str,
) -> dict[str, Any]:
    """Add a comment anchored to text.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        anchor_text: Text to attach comment to.
        comment_text: The comment content.

    Returns:
        Result dict with success status and comment_id.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        comment_id = cached.document.add_comment(anchor_text, comment_text)
        cached.mark_dirty()
        return {"success": True, "comment_id": comment_id}
    except TextNotFoundError:
        return {"success": False, "error": f"Anchor text not found: '{anchor_text}'"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def list_comments(
    server: DocxMCPServer,
    path: str,
    author: str | None = None,
) -> dict[str, Any]:
    """List all comments in the document.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        author: Optional filter by author.

    Returns:
        Result dict with success status and comments list.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    try:
        comments = cached.document.list_comments(author=author)
        return {
            "success": True,
            "comments": [
                {
                    "id": c.id,
                    "author": c.author,
                    "text": c.text,
                    "resolved": c.resolved,
                    "replies": [{"id": r.id, "author": r.author, "text": r.text} for r in c.replies],
                }
                for c in comments
            ],
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def reply_to_comment(
    server: DocxMCPServer,
    path: str,
    comment_id: int,
    reply_text: str,
) -> dict[str, Any]:
    """Reply to an existing comment.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        comment_id: ID of the comment to reply to.
        reply_text: The reply content.

    Returns:
        Result dict with success status and new comment_id.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        new_id = cached.document.reply_to_comment(comment_id, reply_text)
        cached.mark_dirty()
        return {"success": True, "comment_id": new_id}
    except Exception as e:
        return {"success": False, "error": str(e)}


def resolve_comment(
    server: DocxMCPServer,
    path: str,
    comment_id: int,
) -> dict[str, Any]:
    """Mark a comment as resolved.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        comment_id: ID of the comment to resolve.

    Returns:
        Result dict with success status.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        result = cached.document.resolve_comment(comment_id)
        if result:
            cached.mark_dirty()
            return {"success": True}
        return {"success": False, "error": f"Comment not found: {comment_id}"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def delete_comment(
    server: DocxMCPServer,
    path: str,
    comment_id: int,
) -> dict[str, Any]:
    """Delete a comment.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        comment_id: ID of the comment to delete.

    Returns:
        Result dict with success status.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        result = cached.document.delete_comment(comment_id)
        if result:
            cached.mark_dirty()
            return {"success": True}
        return {"success": False, "error": f"Comment not found: {comment_id}"}
    except Exception as e:
        return {"success": False, "error": str(e)}


# =============================================================================
# Revision Tools (Task 3.4)
# =============================================================================


def list_revisions(
    server: DocxMCPServer,
    path: str,
    author: str | None = None,
) -> dict[str, Any]:
    """List all tracked revisions.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        author: Optional filter by author.

    Returns:
        Result dict with success status and revisions list.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    try:
        revisions = cached.document.list_revisions(author=author)
        return {
            "success": True,
            "revisions": [
                {
                    "id": r.id,
                    "type": r.type,
                    "author": r.author,
                    "text": r.text,
                }
                for r in revisions
            ],
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def accept_revision(
    server: DocxMCPServer,
    path: str,
    revision_id: int,
) -> dict[str, Any]:
    """Accept a specific revision.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        revision_id: ID of the revision to accept.

    Returns:
        Result dict with success status.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        result = cached.document.accept_revision(revision_id)
        if result:
            cached.mark_dirty()
            return {"success": True}
        return {"success": False, "error": f"Revision not found: {revision_id}"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def reject_revision(
    server: DocxMCPServer,
    path: str,
    revision_id: int,
) -> dict[str, Any]:
    """Reject a specific revision.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        revision_id: ID of the revision to reject.

    Returns:
        Result dict with success status.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        result = cached.document.reject_revision(revision_id)
        if result:
            cached.mark_dirty()
            return {"success": True}
        return {"success": False, "error": f"Revision not found: {revision_id}"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def accept_all(
    server: DocxMCPServer,
    path: str,
    author: str | None = None,
) -> dict[str, Any]:
    """Accept all revisions.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        author: Optional filter by author.

    Returns:
        Result dict with success status and count.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        count = cached.document.accept_all(author=author)
        if count > 0:
            cached.mark_dirty()
        return {"success": True, "count": count}
    except Exception as e:
        return {"success": False, "error": str(e)}


def reject_all(
    server: DocxMCPServer,
    path: str,
    author: str | None = None,
) -> dict[str, Any]:
    """Reject all revisions.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        author: Optional filter by author.

    Returns:
        Result dict with success status and count.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        count = cached.document.reject_all(author=author)
        if count > 0:
            cached.mark_dirty()
        return {"success": True, "count": count}
    except Exception as e:
        return {"success": False, "error": str(e)}


# =============================================================================
# Read Tools (Task 3.5)
# =============================================================================


def find_text(
    server: DocxMCPServer,
    path: str,
    text: str,
) -> dict[str, Any]:
    """Check if text exists in document.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        text: Text to search for.

    Returns:
        Result dict with success status and found flag.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    try:
        match = cached.document.find_text(text)
        return {"success": True, "found": match is not None}
    except Exception as e:
        return {"success": False, "error": str(e)}


def count_matches(
    server: DocxMCPServer,
    path: str,
    text: str,
) -> dict[str, Any]:
    """Count occurrences of text in document.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        text: Text to search for.

    Returns:
        Result dict with success status and count.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    try:
        count = cached.document.count_matches(text)
        return {"success": True, "count": count}
    except Exception as e:
        return {"success": False, "error": str(e)}


def get_visible_text(
    server: DocxMCPServer,
    path: str,
) -> dict[str, Any]:
    """Get the visible text of the document.

    Args:
        server: The MCP server instance.
        path: Path to the document.

    Returns:
        Result dict with success status and text.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    try:
        text = cached.document.get_visible_text()
        return {"success": True, "text": text}
    except Exception as e:
        return {"success": False, "error": str(e)}
