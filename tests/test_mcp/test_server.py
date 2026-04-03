"""Tests for MCP Server following TDD."""

import asyncio
from unittest.mock import MagicMock


class TestMCPServer:
    """Test MCP server initialization and lifecycle."""

    def test_server_has_name(self):
        """Server has correct name."""
        from docx_editor_mcp.server import create_server

        server = create_server()
        assert server.name == "docx-editor"

    def test_server_has_cache(self):
        """Server initializes with DocumentCache."""
        from docx_editor_mcp.server import create_server

        server = create_server()
        assert server.cache is not None
        assert server.cache.max_documents == 10

    def test_server_cache_configurable(self):
        """Server cache size is configurable."""
        from docx_editor_mcp.server import create_server

        server = create_server(max_documents=5)
        assert server.cache.max_documents == 5


class TestGracefulShutdown:
    """Test graceful shutdown with dirty document saving (Task 2.2)."""

    def test_shutdown_saves_dirty_documents(self):
        """Shutdown saves all dirty documents."""
        from docx_editor_mcp.cache import CachedDocument
        from docx_editor_mcp.server import create_server

        server = create_server()

        # Add dirty documents to cache
        mock_doc1 = MagicMock()
        mock_doc2 = MagicMock()

        cached1 = CachedDocument(path="/doc1.docx", document=mock_doc1, author="Tester")
        cached1.mark_dirty()
        cached2 = CachedDocument(path="/doc2.docx", document=mock_doc2, author="Tester")
        cached2.mark_dirty()

        server.cache.put(cached1)
        server.cache.put(cached2)

        # Trigger shutdown
        server.shutdown()

        # Both documents should be saved
        mock_doc1.save.assert_called_once()
        mock_doc2.save.assert_called_once()

    def test_shutdown_skips_clean_documents(self):
        """Shutdown doesn't save clean documents."""
        from docx_editor_mcp.cache import CachedDocument
        from docx_editor_mcp.server import create_server

        server = create_server()

        # Add clean document
        mock_doc = MagicMock()
        cached = CachedDocument(path="/doc.docx", document=mock_doc, author="Tester")
        # Not marked dirty
        server.cache.put(cached)

        server.shutdown()

        mock_doc.save.assert_not_called()

    def test_shutdown_continues_on_save_error(self):
        """Shutdown continues even if a save fails (best-effort)."""
        from docx_editor_mcp.cache import CachedDocument
        from docx_editor_mcp.server import create_server

        server = create_server()

        # Add two dirty documents, first one fails to save
        mock_doc1 = MagicMock()
        mock_doc1.save.side_effect = Exception("Disk full")
        mock_doc2 = MagicMock()

        cached1 = CachedDocument(path="/doc1.docx", document=mock_doc1, author="Tester")
        cached1.mark_dirty()
        cached2 = CachedDocument(path="/doc2.docx", document=mock_doc2, author="Tester")
        cached2.mark_dirty()

        server.cache.put(cached1)
        server.cache.put(cached2)

        # Should not raise, should continue to save doc2
        server.shutdown()

        mock_doc1.save.assert_called_once()
        mock_doc2.save.assert_called_once()

    def test_shutdown_clears_cache(self):
        """Shutdown clears the cache after saving."""
        from docx_editor_mcp.cache import CachedDocument
        from docx_editor_mcp.server import create_server

        server = create_server()

        mock_doc = MagicMock()
        cached = CachedDocument(path="/doc.docx", document=mock_doc, author="Tester")
        server.cache.put(cached)

        assert server.cache.size == 1

        server.shutdown()

        assert server.cache.size == 0


class TestFastMCPApp:
    """Test FastMCP application wiring."""

    def test_create_mcp_app(self):
        """_create_mcp_app returns a FastMCP instance."""
        from docx_editor_mcp.server import _create_mcp_app

        app = _create_mcp_app()
        assert app is not None

    def test_app_has_name(self):
        """App has correct server name."""
        from docx_editor_mcp.server import _create_mcp_app

        app = _create_mcp_app()
        assert app.name == "docx-editor"

    def test_app_has_instructions(self):
        """App has server instructions for Tool Search discoverability."""
        from docx_editor_mcp.server import SERVER_INSTRUCTIONS, _create_mcp_app

        app = _create_mcp_app()
        assert app.instructions == SERVER_INSTRUCTIONS

    def test_all_tools_registered(self):
        """All 24 tools are registered on the app."""
        from docx_editor_mcp.server import _create_mcp_app

        app = _create_mcp_app()
        tool_names = {t.name for t in app._tool_manager.list_tools()}
        expected = {
            # Lifecycle
            "open_document", "save_document", "close_document",
            "reload_document", "force_save",
            # Paragraphs & track changes
            "list_paragraphs", "replace_text", "delete_text",
            "insert_after", "insert_before", "batch_edit",
            "rewrite_paragraph", "batch_rewrite",
            # Comments
            "add_comment", "list_comments", "reply_to_comment",
            "resolve_comment", "delete_comment",
            # Revisions
            "list_revisions", "accept_revision", "reject_revision",
            "accept_all", "reject_all",
            # Read
            "find_text", "count_matches", "get_visible_text",
        }
        assert tool_names == expected

    def test_main_function_exists(self):
        """main() function exists and is callable."""
        from docx_editor_mcp.server import main

        assert callable(main)


class TestLifespan:
    """Test server lifespan context manager."""

    def test_lifespan_creates_and_shuts_down_server(self):
        """Lifespan yields a DocxMCPServer and shuts it down on exit."""
        from docx_editor_mcp.server import DocxMCPServer, server_lifespan

        async def _run():
            mock_mcp = MagicMock()
            async with server_lifespan(mock_mcp) as server:
                assert isinstance(server, DocxMCPServer)
                assert server.cache is not None
                assert server.cache.size == 0
            # After exit, shutdown was called (cache cleared)
            assert server.cache.size == 0

        asyncio.run(_run())
