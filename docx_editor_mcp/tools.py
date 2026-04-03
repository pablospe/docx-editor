"""MCP Tools for docx_editor operations."""

import logging
import os
from typing import Any, cast

from docx_editor import Document
from docx_editor.exceptions import TextNotFoundError

from .cache import CachedDocument, normalize_path
from .server import DocxMCPServer

logger = logging.getLogger(__name__)


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
        logger.debug("Failed to close document: %s", path)

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
        logger.debug("Failed to close document during reload: %s", path)

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


def list_paragraphs(
    server: DocxMCPServer,
    path: str,
    max_chars: int = 80,
    start: int = 0,
    limit: int = 0,
) -> dict[str, Any]:
    """List all paragraphs with hash-anchored references.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        max_chars: Maximum characters for preview text.
        start: Starting paragraph index, 0-based.
        limit: Maximum paragraphs to return, 0 for all.

    Returns:
        Result dict with success status and paragraphs list.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    try:
        total_count = len(cached.document._document_editor.dom.getElementsByTagName("w:p"))
        paragraphs = cached.document.list_paragraphs(max_chars=max_chars, start=start, limit=limit)
        return {"success": True, "paragraphs": paragraphs, "total": total_count}
    except Exception as e:
        return {"success": False, "error": str(e)}


def replace_text(
    server: DocxMCPServer,
    path: str,
    old_text: str,
    new_text: str,
    paragraph: str,
    occurrence: int = 0,
) -> dict[str, Any]:
    """Replace text with tracked changes.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        old_text: Text to find and replace.
        new_text: Replacement text.
        paragraph: Paragraph reference from list_paragraphs (e.g., "P2#f3c1").
        occurrence: Which occurrence within the paragraph to replace (0 = first).

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
        new_ref = cached.document.replace(old_text, new_text, paragraph=paragraph, occurrence=occurrence)
        cached.mark_dirty()
        return {"success": True, "new_ref": new_ref}
    except TextNotFoundError:
        return {"success": False, "error": f"Text not found: '{old_text}'"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def delete_text(
    server: DocxMCPServer,
    path: str,
    text: str,
    paragraph: str,
    occurrence: int = 0,
) -> dict[str, Any]:
    """Delete text with tracked changes.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        text: Text to delete.
        paragraph: Paragraph reference from list_paragraphs (e.g., "P2#f3c1").
        occurrence: Which occurrence within the paragraph to delete (0 = first).

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
        new_ref = cached.document.delete(text, paragraph=paragraph, occurrence=occurrence)
        cached.mark_dirty()
        return {"success": True, "new_ref": new_ref}
    except TextNotFoundError:
        return {"success": False, "error": f"Text not found: '{text}'"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def insert_after(
    server: DocxMCPServer,
    path: str,
    anchor: str,
    text: str,
    paragraph: str,
    occurrence: int = 0,
) -> dict[str, Any]:
    """Insert text after anchor with tracked changes.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        anchor: Text to find as insertion point.
        text: Text to insert.
        paragraph: Paragraph reference from list_paragraphs (e.g., "P2#f3c1").
        occurrence: Which occurrence of anchor within the paragraph to use (0 = first).

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
        new_ref = cached.document.insert_after(anchor, text, paragraph=paragraph, occurrence=occurrence)
        cached.mark_dirty()
        return {"success": True, "new_ref": new_ref}
    except TextNotFoundError:
        return {"success": False, "error": f"Anchor not found: '{anchor}'"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def insert_before(
    server: DocxMCPServer,
    path: str,
    anchor: str,
    text: str,
    paragraph: str,
    occurrence: int = 0,
) -> dict[str, Any]:
    """Insert text before anchor with tracked changes.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        anchor: Text to find as insertion point.
        text: Text to insert.
        paragraph: Paragraph reference from list_paragraphs (e.g., "P2#f3c1").
        occurrence: Which occurrence of anchor within the paragraph to use (0 = first).

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
        new_ref = cached.document.insert_before(anchor, text, paragraph=paragraph, occurrence=occurrence)
        cached.mark_dirty()
        return {"success": True, "new_ref": new_ref}
    except TextNotFoundError:
        return {"success": False, "error": f"Anchor not found: '{anchor}'"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def batch_edit(
    server: DocxMCPServer,
    path: str,
    operations: list[dict[str, Any]],
) -> dict[str, Any]:
    """Apply multiple edits atomically with upfront hash validation.

    All paragraph hashes are validated before any edits are applied.
    If any hash is stale, the entire batch is rejected.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        operations: List of operation dicts, each with keys:
            - action: "replace", "delete", "insert_after", or "insert_before"
            - paragraph: Hash-anchored paragraph reference (e.g., "P2#f3c1")
            - find/replace_with: For "replace" action
            - text: For "delete" (text to delete) or insert (text to insert)
            - anchor: For "insert_after"/"insert_before"
            - occurrence: Optional, defaults to 0

    Returns:
        Result dict with success status and list of change_ids.
    """
    from docx_editor import EditOperation

    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        ops = [EditOperation(**op) for op in operations]
    except (TypeError, ValueError) as e:
        return {"success": False, "error": f"Invalid operation: {e}"}

    try:
        new_refs = cached.document.batch_edit(ops)
        cached.mark_dirty()
        return {"success": True, "new_refs": new_refs}
    except TextNotFoundError as e:
        return {"success": False, "error": str(e)}
    except Exception as e:
        return {"success": False, "error": str(e)}


def rewrite_paragraph(
    server: DocxMCPServer,
    path: str,
    paragraph: str,
    new_text: str,
) -> dict[str, Any]:
    """Rewrite a paragraph with automatic word-level tracked changes.

    Diffs old vs new text and generates fine-grained tracked insertions/deletions.
    Use only when the edit cannot be decomposed into independent find/replace pairs
    (e.g., sentence restructuring, reordering).

    Args:
        server: The MCP server instance.
        path: Path to the document.
        paragraph: Paragraph reference from list_paragraphs (e.g., "P2#f3c1").
        new_text: Desired new text for the paragraph.

    Returns:
        Result dict with success status and new_ref.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        new_ref = cached.document.rewrite_paragraph(paragraph, new_text)
        cached.mark_dirty()
        return {"success": True, "new_ref": new_ref}
    except Exception as e:
        return {"success": False, "error": str(e)}


def batch_rewrite(
    server: DocxMCPServer,
    path: str,
    rewrites: list[list[str]],
) -> dict[str, Any]:
    """Rewrite multiple paragraphs with upfront hash validation.

    All paragraph hashes are validated before any rewrites are applied.
    If any hash is stale, the entire batch is rejected.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        rewrites: List of [paragraph_ref, new_text] pairs.

    Returns:
        Result dict with success status and list of new_refs.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    ext_error = _check_external_changes(cached)
    if ext_error:
        return ext_error

    try:
        tuples = [(ref, text) for ref, text in rewrites]
        new_refs = cached.document.batch_rewrite(tuples)
        cached.mark_dirty()
        return {"success": True, "new_refs": new_refs}
    except Exception as e:
        return {"success": False, "error": str(e)}


# =============================================================================
# Exploration Tools
# =============================================================================


def search_text(
    server: DocxMCPServer,
    path: str,
    query: str,
    context_chars: int = 100,
) -> dict[str, Any]:
    """Search for text in the document with surrounding context.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        query: Text to search for.
        context_chars: Characters of context before and after each match.

    Returns:
        Result dict with success status, matches, and count.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    try:
        matches = cached.document.search_text(query, context_chars=context_chars)
        return {"success": True, "matches": matches, "count": len(matches)}
    except Exception as e:
        return {"success": False, "error": str(e)}


def get_paragraph_text(
    server: DocxMCPServer,
    path: str,
    paragraphs: list[str],
) -> dict[str, Any]:
    """Read specific paragraphs in full by their hash-anchored refs.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        paragraphs: List of paragraph references (e.g., ["P1#a7b2", "P3#cc33"]).

    Returns:
        Result dict with success status and paragraphs.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    try:
        results = cached.document.get_paragraph_text(paragraphs)
        return {"success": True, "paragraphs": results}
    except Exception as e:
        return {"success": False, "error": str(e)}


def get_document_info(
    server: DocxMCPServer,
    path: str,
) -> dict[str, Any]:
    """Get document overview: paragraph count, word count, and heading outline.

    Args:
        server: The MCP server instance.
        path: Path to the document.

    Returns:
        Result dict with success status and document info.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    try:
        info = cached.document.get_document_info()
        return {"success": True, **info}
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
    max_chars: int = 10000,
) -> dict[str, Any]:
    """Get the visible text of the document.

    Args:
        server: The MCP server instance.
        path: Path to the document.
        max_chars: Maximum characters to return. 0 for no limit.

    Returns:
        Result dict with success status and text.
    """
    result = _get_cached_or_error(server, path)
    if isinstance(result, dict):
        return cast(dict[str, Any], result)
    cached = result

    try:
        text = cached.document.get_visible_text()
        if max_chars and len(text) > max_chars:
            text = text[:max_chars]
            return {
                "success": True,
                "text": text,
                "truncated": True,
                "hint": (
                    "Document too large to return in full. Use search_text to find specific "
                    "content, get_paragraph_text to read specific paragraphs, or "
                    "get_document_info for an overview."
                ),
            }
        return {"success": True, "text": text, "truncated": False}
    except Exception as e:
        return {"success": False, "error": str(e)}
