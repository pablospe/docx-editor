"""Tests for ooxml pack and unpack functions."""

import pytest

from docx_editor.exceptions import DocumentNotFoundError, InvalidDocumentError
from docx_editor.ooxml.pack import pack_document
from docx_editor.ooxml.unpack import unpack_document


class TestUnpack:
    """Tests for unpack_document function."""

    def test_unpack_nonexistent_file(self, temp_dir):
        """Test error when unpacking nonexistent file."""
        with pytest.raises(DocumentNotFoundError, match="Document not found"):
            unpack_document(temp_dir / "nonexistent.docx", temp_dir / "output")

    def test_unpack_invalid_zip(self, temp_dir):
        """Test error when unpacking invalid zip file."""
        invalid_file = temp_dir / "invalid.docx"
        invalid_file.write_text("not a zip file")

        with pytest.raises(InvalidDocumentError, match="Not a valid .docx file"):
            unpack_document(invalid_file, temp_dir / "output")

    def test_unpack_returns_rsid(self, simple_docx, temp_dir):
        """Test that unpack returns an 8-character RSID."""
        rsid = unpack_document(simple_docx, temp_dir / "output")

        assert isinstance(rsid, str)
        assert len(rsid) == 8
        assert all(c in "0123456789ABCDEF" for c in rsid)


class TestPack:
    """Tests for pack_document function."""

    def test_pack_nonexistent_directory(self, temp_dir):
        """Test error when packing nonexistent directory."""
        with pytest.raises(ValueError, match="is not a directory"):
            pack_document(temp_dir / "nonexistent", temp_dir / "output.docx")

    def test_pack_wrong_extension(self, simple_docx, temp_dir):
        """Test error when output has wrong extension."""
        # First unpack to create a valid directory
        unpack_document(simple_docx, temp_dir / "unpacked")

        with pytest.raises(ValueError, match="must be a .docx, .pptx, or .xlsx"):
            pack_document(temp_dir / "unpacked", temp_dir / "output.txt")

    def test_pack_creates_parent_directory(self, simple_docx, temp_dir):
        """Test that pack creates parent directories for output."""
        unpack_document(simple_docx, temp_dir / "unpacked")

        output_path = temp_dir / "nested" / "dir" / "output.docx"
        result = pack_document(temp_dir / "unpacked", output_path)

        assert result is True
        assert output_path.exists()

    def test_pack_roundtrip(self, simple_docx, temp_dir):
        """Test unpacking and repacking a document."""
        # Unpack
        unpack_document(simple_docx, temp_dir / "unpacked")

        # Pack
        output_path = temp_dir / "repacked.docx"
        result = pack_document(temp_dir / "unpacked", output_path)

        assert result is True
        assert output_path.exists()
        assert output_path.stat().st_size > 0
