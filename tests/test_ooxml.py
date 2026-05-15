"""Tests for ooxml pack and unpack functions."""

import zipfile

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

    def test_pack_excludes_root_meta_json(self, simple_docx, temp_dir):
        """Root-level meta.json must not be packed (Word flags it as corrupt).

        Scope is workspace-root only: a same-named file under a subpath stays.
        """
        unpacked = temp_dir / "unpacked"
        unpack_document(simple_docx, unpacked)
        (unpacked / "meta.json").write_text('{"author": "test"}')
        (unpacked / "word" / "meta.json").write_text('{"unrelated": true}')

        output_path = temp_dir / "output.docx"
        pack_document(unpacked, output_path)

        with zipfile.ZipFile(output_path, "r") as zf:
            names = zf.namelist()
        assert "meta.json" not in names
        assert "word/meta.json" in names


def _read_doc_xml(docx_path):
    """Open a packed .docx and return its word/document.xml content as text."""
    import zipfile

    with zipfile.ZipFile(docx_path) as zf:
        return zf.read("word/document.xml").decode("utf-8")


class TestCondenseXmlPreservesTextContent:
    """Issue #9: condense_xml must not strip whitespace fragments from text-bearing
    OOXML elements (w:t, w:delText, w:instrText). Pre-fix only :t was protected."""

    def _write_xml(self, base_dir, body_xml):
        """Replace word/document.xml's body with body_xml, returning the dir."""
        doc_xml = base_dir / "word" / "document.xml"
        original = doc_xml.read_text(encoding="utf-8")
        head, _, tail = original.partition("<w:body>")
        _, _, end = tail.partition("</w:body>")
        new = f"{head}<w:body>{body_xml}</w:body>{end}"
        doc_xml.write_text(new, encoding="utf-8")
        return base_dir

    def test_condense_preserves_whitespace_in_delText(self, simple_docx, temp_dir):
        unpack_document(simple_docx, temp_dir / "unpacked")
        body = (
            '<w:p><w:del w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z">'
            '<w:r><w:delText xml:space="preserve">foo bar</w:delText></w:r>'
            "</w:del></w:p>"
        )
        self._write_xml(temp_dir / "unpacked", body)

        # Inject a multi-TEXT_NODE state into the w:delText, including a lone space.
        import defusedxml.minidom

        doc_xml = temp_dir / "unpacked" / "word" / "document.xml"
        dom = defusedxml.minidom.parseString(doc_xml.read_text(encoding="utf-8"))
        delText = dom.getElementsByTagName("w:delText")[0]
        while delText.firstChild:
            delText.removeChild(delText.firstChild)
        owner = delText.ownerDocument
        delText.appendChild(owner.createTextNode("foo"))
        delText.appendChild(owner.createTextNode(" "))
        delText.appendChild(owner.createTextNode("bar"))
        doc_xml.write_bytes(dom.toxml(encoding="UTF-8"))

        out = temp_dir / "out.docx"
        pack_document(temp_dir / "unpacked", out)
        repacked = _read_doc_xml(out)
        # The space must survive condensation
        assert "foo bar" in repacked

    def test_condense_preserves_whitespace_in_instrText(self, simple_docx, temp_dir):
        unpack_document(simple_docx, temp_dir / "unpacked")
        body = '<w:p><w:r><w:instrText xml:space="preserve"> HYPERLINK </w:instrText></w:r></w:p>'
        self._write_xml(temp_dir / "unpacked", body)

        out = temp_dir / "out.docx"
        pack_document(temp_dir / "unpacked", out)
        repacked = _read_doc_xml(out)
        assert " HYPERLINK " in repacked
