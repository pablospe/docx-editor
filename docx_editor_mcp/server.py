"""MCP Server for docx_editor with persistent DOM caching."""

import logging
from dataclasses import dataclass

from .cache import DocumentCache

logger = logging.getLogger(__name__)


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
                except Exception as e:
                    logger.error(f"Failed to save {cached_doc.path}: {e}")
                    # Continue with other documents (best-effort)

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


def main() -> None:
    """Main entry point for running the MCP server.

    This will be called when running `mcp-server-docx` or
    `python -m docx_editor_mcp`.
    """
    # TODO: Implement MCP protocol handling in Task 3
    # For now, just create the server
    server = create_server()
    logger.info(f"Started {server.name} MCP server")
