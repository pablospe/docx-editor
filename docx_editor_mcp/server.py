"""MCP Server for docx_editor with persistent DOM caching."""

import json
import logging
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass

from .cache import DocumentCache

logger = logging.getLogger(__name__)

SERVER_INSTRUCTIONS = (
    "Document editing server for .docx files with tracked changes, comments, and revisions. "
    "Use open_document to load a file, then edit with replace_text/delete_text/insert_after/insert_before. "
    "Add comments with add_comment, manage revisions with accept/reject. "
    "Save changes with save_document. Documents stay cached between operations for fast repeated edits."
)


@dataclass
class DocxMCPServer:
    """MCP Server wrapper with document cache."""

    name: str
    cache: DocumentCache

    def shutdown(self) -> None:
        """Gracefully shutdown server, saving dirty documents.

        Best-effort: continues even if individual saves fail.
        """
        for cached_doc in list(self.cache.all()):
            if cached_doc.dirty:
                try:
                    cached_doc.document.save()
                    logger.info(f"Saved dirty document: {cached_doc.path}")
                except Exception:
                    logger.exception("Failed to save %s", cached_doc.path)

        # Clear cache
        for cached_doc in list(self.cache.all()):
            self.cache.remove(cached_doc.path)


def create_server(max_documents: int = 10) -> DocxMCPServer:
    """Create a new MCP server instance.

    Args:
        max_documents: Maximum documents to keep in cache.

    Returns:
        Configured DocxMCPServer instance.
    """
    cache = DocumentCache(max_documents=max_documents)
    return DocxMCPServer(name="docx-editor", cache=cache)


def _get_server(ctx) -> DocxMCPServer:
    """Extract DocxMCPServer from FastMCP context."""
    return ctx.request_context.lifespan_context


@asynccontextmanager
async def server_lifespan(mcp) -> AsyncIterator[DocxMCPServer]:
    """Manage DocxMCPServer lifecycle."""
    server = create_server()
    try:
        yield server
    finally:
        server.shutdown()


def _create_mcp_app():
    """Create and configure the FastMCP application with all tools registered."""
    from mcp.server.fastmcp import FastMCP as _FastMCP

    mcp = _FastMCP(
        "docx-editor",
        instructions=SERVER_INSTRUCTIONS,
        lifespan=server_lifespan,
    )
    _register_tools(mcp)
    return mcp


def _register_tools(mcp) -> None:
    """Register all docx_editor tools on the FastMCP instance."""
    from mcp.server.fastmcp import Context

    from . import tools

    # -- Document Lifecycle --

    @mcp.tool()
    def open_document(path: str, ctx: Context, author: str | None = None) -> str:
        """Open a .docx document for editing. Must be called before any other operations on the file.

        Args:
            path: Absolute path to the .docx file.
            author: Author name for tracked changes. Uses system username if omitted.
        """
        return json.dumps(tools.open_document(_get_server(ctx), path, author))

    @mcp.tool()
    def save_document(path: str, ctx: Context) -> str:
        """Save changes to a document. Fails if the file was modified externally.

        Args:
            path: Path to the document (as used in open_document).
        """
        return json.dumps(tools.save_document(_get_server(ctx), path))

    @mcp.tool()
    def close_document(path: str, ctx: Context) -> str:
        """Close a document and remove it from the cache. Warns if there are unsaved changes.

        Args:
            path: Path to the document.
        """
        return json.dumps(tools.close_document(_get_server(ctx), path))

    @mcp.tool()
    def reload_document(path: str, ctx: Context) -> str:
        """Reload a document from disk, discarding all cached changes.

        Args:
            path: Path to the document.
        """
        return json.dumps(tools.reload_document(_get_server(ctx), path))

    @mcp.tool()
    def force_save(path: str, ctx: Context) -> str:
        """Force save a document, ignoring external modifications. Use with care.

        Args:
            path: Path to the document.
        """
        return json.dumps(tools.force_save(_get_server(ctx), path))

    # -- Paragraphs & Track Changes --

    @mcp.tool()
    def list_paragraphs(path: str, ctx: Context, max_chars: int = 80) -> str:
        """List all paragraphs with hash-anchored references. Call this before editing to get paragraph refs.

        Returns references like "P1#a7b2| Introduction to the..." that you must pass to edit tools.

        Args:
            path: Path to the document.
            max_chars: Maximum characters for the preview text (default 80).
        """
        return json.dumps(tools.list_paragraphs(_get_server(ctx), path, max_chars))

    @mcp.tool()
    def replace_text(path: str, old_text: str, new_text: str, paragraph: str, ctx: Context, occurrence: int = 0) -> str:
        """Replace text in a document with tracked changes (redlining).

        Args:
            path: Path to the document.
            old_text: The text to find and replace.
            new_text: The replacement text.
            paragraph: Paragraph reference from list_paragraphs (e.g., "P2#f3c1").
            occurrence: Which occurrence within the paragraph to replace (0 = first).
        """
        return json.dumps(tools.replace_text(_get_server(ctx), path, old_text, new_text, paragraph, occurrence))

    @mcp.tool()
    def delete_text(path: str, text: str, paragraph: str, ctx: Context, occurrence: int = 0) -> str:
        """Delete text from a document with tracked changes.

        Args:
            path: Path to the document.
            text: The text to delete.
            paragraph: Paragraph reference from list_paragraphs (e.g., "P2#f3c1").
            occurrence: Which occurrence within the paragraph to delete (0 = first).
        """
        return json.dumps(tools.delete_text(_get_server(ctx), path, text, paragraph, occurrence))

    @mcp.tool()
    def insert_after(path: str, anchor: str, text: str, paragraph: str, ctx: Context, occurrence: int = 0) -> str:
        """Insert text after an anchor string with tracked changes.

        Args:
            path: Path to the document.
            anchor: The text to find as the insertion point.
            text: The text to insert after the anchor.
            paragraph: Paragraph reference from list_paragraphs (e.g., "P2#f3c1").
            occurrence: Which occurrence of anchor within the paragraph to use (0 = first).
        """
        return json.dumps(tools.insert_after(_get_server(ctx), path, anchor, text, paragraph, occurrence))

    @mcp.tool()
    def insert_before(path: str, anchor: str, text: str, paragraph: str, ctx: Context, occurrence: int = 0) -> str:
        """Insert text before an anchor string with tracked changes.

        Args:
            path: Path to the document.
            anchor: The text to find as the insertion point.
            text: The text to insert before the anchor.
            paragraph: Paragraph reference from list_paragraphs (e.g., "P2#f3c1").
            occurrence: Which occurrence of anchor within the paragraph to use (0 = first).
        """
        return json.dumps(tools.insert_before(_get_server(ctx), path, anchor, text, paragraph, occurrence))

    @mcp.tool()
    def batch_edit(path: str, operations: list[dict], ctx: Context) -> str:
        """Apply multiple edits atomically. All paragraph hashes are validated before any edit is applied.

        Args:
            path: Path to the document.
            operations: List of edit operations, each a dict with keys:
                action ("replace"/"delete"/"insert_after"/"insert_before"),
                paragraph (ref from list_paragraphs),
                and action-specific fields (find/replace_with, text, anchor, occurrence).
        """
        return json.dumps(tools.batch_edit(_get_server(ctx), path, operations))

    # -- Comments --

    @mcp.tool()
    def add_comment(path: str, anchor_text: str, comment_text: str, ctx: Context) -> str:
        """Add a comment anchored to specific text in the document.

        Args:
            path: Path to the document.
            anchor_text: The text to attach the comment to.
            comment_text: The comment content.
        """
        return json.dumps(tools.add_comment(_get_server(ctx), path, anchor_text, comment_text))

    @mcp.tool()
    def list_comments(path: str, ctx: Context, author: str | None = None) -> str:
        """List all comments in the document, optionally filtered by author.

        Args:
            path: Path to the document.
            author: Filter comments by this author name.
        """
        return json.dumps(tools.list_comments(_get_server(ctx), path, author))

    @mcp.tool()
    def reply_to_comment(path: str, comment_id: int, reply_text: str, ctx: Context) -> str:
        """Reply to an existing comment.

        Args:
            path: Path to the document.
            comment_id: ID of the comment to reply to.
            reply_text: The reply content.
        """
        return json.dumps(tools.reply_to_comment(_get_server(ctx), path, comment_id, reply_text))

    @mcp.tool()
    def resolve_comment(path: str, comment_id: int, ctx: Context) -> str:
        """Mark a comment as resolved.

        Args:
            path: Path to the document.
            comment_id: ID of the comment to resolve.
        """
        return json.dumps(tools.resolve_comment(_get_server(ctx), path, comment_id))

    @mcp.tool()
    def delete_comment(path: str, comment_id: int, ctx: Context) -> str:
        """Delete a comment from the document.

        Args:
            path: Path to the document.
            comment_id: ID of the comment to delete.
        """
        return json.dumps(tools.delete_comment(_get_server(ctx), path, comment_id))

    # -- Revisions --

    @mcp.tool()
    def list_revisions(path: str, ctx: Context, author: str | None = None) -> str:
        """List all tracked revisions (insertions/deletions), optionally filtered by author.

        Args:
            path: Path to the document.
            author: Filter revisions by this author name.
        """
        return json.dumps(tools.list_revisions(_get_server(ctx), path, author))

    @mcp.tool()
    def accept_revision(path: str, revision_id: int, ctx: Context) -> str:
        """Accept a specific tracked revision, making it permanent.

        Args:
            path: Path to the document.
            revision_id: ID of the revision to accept.
        """
        return json.dumps(tools.accept_revision(_get_server(ctx), path, revision_id))

    @mcp.tool()
    def reject_revision(path: str, revision_id: int, ctx: Context) -> str:
        """Reject a specific tracked revision, reverting the change.

        Args:
            path: Path to the document.
            revision_id: ID of the revision to reject.
        """
        return json.dumps(tools.reject_revision(_get_server(ctx), path, revision_id))

    @mcp.tool()
    def accept_all(path: str, ctx: Context, author: str | None = None) -> str:
        """Accept all tracked revisions, optionally only those by a specific author.

        Args:
            path: Path to the document.
            author: Only accept revisions by this author.
        """
        return json.dumps(tools.accept_all(_get_server(ctx), path, author))

    @mcp.tool()
    def reject_all(path: str, ctx: Context, author: str | None = None) -> str:
        """Reject all tracked revisions, optionally only those by a specific author.

        Args:
            path: Path to the document.
            author: Only reject revisions by this author.
        """
        return json.dumps(tools.reject_all(_get_server(ctx), path, author))

    # -- Read --

    @mcp.tool()
    def find_text(path: str, text: str, ctx: Context) -> str:
        """Check if text exists in the document.

        Args:
            path: Path to the document.
            text: The text to search for.
        """
        return json.dumps(tools.find_text(_get_server(ctx), path, text))

    @mcp.tool()
    def count_matches(path: str, text: str, ctx: Context) -> str:
        """Count how many times text appears in the document. Use before replace/delete to verify uniqueness.

        Args:
            path: Path to the document.
            text: The text to count occurrences of.
        """
        return json.dumps(tools.count_matches(_get_server(ctx), path, text))

    @mcp.tool()
    def get_visible_text(path: str, ctx: Context) -> str:
        """Get the full visible text of the document (insertions included, deletions excluded).

        Args:
            path: Path to the document.
        """
        return json.dumps(tools.get_visible_text(_get_server(ctx), path))


def main() -> None:
    """Main entry point for running the MCP server.

    This is called when running `mcp-server-docx` or
    `python -m docx_editor_mcp`.
    """
    app = _create_mcp_app()
    app.run(transport="stdio")
