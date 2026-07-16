"""Tests for ooxml pack and unpack functions."""

import stat
import sys
import zipfile
from xml.parsers.expat import ExpatError

import pytest
from conftest import ENTITY_DTD_XML, NS, replace_document_xml
from defusedxml.common import EntitiesForbidden

from docx_editor.exceptions import DocumentNotFoundError, InvalidDocumentError
from docx_editor.ooxml.pack import pack_document
from docx_editor.ooxml.unpack import _is_symlink_entry, unpack_document
from docx_editor.xml_editor import DocxXMLEditor


def _build_zip(path, entries):
    """Build a ZIP at path from a list of (name_or_zipinfo, data_bytes) tuples."""
    with zipfile.ZipFile(path, "w") as zf:
        for name_or_info, data in entries:
            zf.writestr(name_or_info, data)


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

    def test_unpack_rejects_path_traversal(self, temp_dir):
        """Reject ZIP entries containing '..' segments before any extraction."""
        bad_zip = temp_dir / "bad.docx"
        _build_zip(
            bad_zip,
            [
                ("safe.txt", b"safe"),
                ("../evil.txt", b"pwned"),
            ],
        )
        output = temp_dir / "output"
        with pytest.raises(InvalidDocumentError, match="Unsafe ZIP entry"):
            unpack_document(bad_zip, output)

        # Nothing should have leaked outside output_dir.
        assert not (temp_dir / "evil.txt").exists()

    def test_unpack_rejects_absolute_path(self, temp_dir):
        """Reject ZIP entries with absolute paths."""
        bad_zip = temp_dir / "bad.docx"
        _build_zip(
            bad_zip,
            [
                ("safe.txt", b"safe"),
                ("/tmp/evil.txt", b"pwned"),
            ],
        )
        output = temp_dir / "output"
        with pytest.raises(InvalidDocumentError, match="Unsafe ZIP entry"):
            unpack_document(bad_zip, output)

    def test_unpack_rejects_symlink_entry(self, temp_dir):
        """Reject Unix symlink ZIP entries (create_system==3, mode S_IFLNK)."""
        bad_zip = temp_dir / "bad.docx"
        link_info = zipfile.ZipInfo("evil_link")
        link_info.create_system = 3  # Unix
        link_info.external_attr = (stat.S_IFLNK | 0o777) << 16
        _build_zip(
            bad_zip,
            [
                ("safe.txt", b"safe"),
                (link_info, b"/etc/passwd"),
            ],
        )
        output = temp_dir / "output"
        with pytest.raises(InvalidDocumentError, match="Symlink ZIP entry"):
            unpack_document(bad_zip, output)

        # Validation must happen before extraction — no symlink left behind.
        assert not (output / "evil_link").exists()
        assert not (output / "evil_link").is_symlink()

    def test_unpack_rejects_dotdot_with_trailing_space(self, temp_dir):
        """Reject '.. ' (Windows strips trailing space → '..' parent traversal)."""
        bad_zip = temp_dir / "bad.docx"
        _build_zip(bad_zip, [("safe.txt", b"safe"), (".. /evil.txt", b"pwned")])
        with pytest.raises(InvalidDocumentError, match="Unsafe ZIP entry"):
            unpack_document(bad_zip, temp_dir / "output")

    def test_unpack_rejects_drive_letter(self, temp_dir):
        """Reject Windows drive-letter ZIP entry names like 'C:evil.txt'."""
        bad_zip = temp_dir / "bad.docx"
        _build_zip(bad_zip, [("safe.txt", b"safe"), ("C:evil.txt", b"pwned")])
        with pytest.raises(InvalidDocumentError, match="Unsafe ZIP entry"):
            unpack_document(bad_zip, temp_dir / "output")

    def test_unpack_rejects_colon_ads(self, temp_dir):
        """Reject mid-path colons (NTFS alternate data stream, e.g. 'word/a.xml:s')."""
        bad_zip = temp_dir / "bad.docx"
        _build_zip(bad_zip, [("safe.txt", b"safe"), ("word/document.xml:stream", b"pwned")])
        with pytest.raises(InvalidDocumentError, match="Unsafe ZIP entry"):
            unpack_document(bad_zip, temp_dir / "output")

    def test_unpack_accepts_existing_empty_output_dir(self, temp_dir, simple_docx):
        """Pre-existing empty output_dir is fine (no symlinks to reject)."""
        output = temp_dir / "output"
        output.mkdir()
        rsid = unpack_document(simple_docx, output)
        assert len(rsid) == 8

    def test_unpack_accepts_existing_output_dir_without_symlinks(self, temp_dir, simple_docx):
        """Pre-existing output_dir containing only regular files is accepted."""
        output = temp_dir / "output"
        output.mkdir()
        (output / "leftover.txt").write_text("ok")
        rsid = unpack_document(simple_docx, output)
        assert len(rsid) == 8

    def test_unpack_rejects_output_path_that_is_a_file(self, temp_dir, simple_docx):
        """Refuse to unpack when output_dir already exists as a regular file."""
        output = temp_dir / "output"
        output.write_text("i am a file")
        with pytest.raises(InvalidDocumentError, match="Output path is not a directory"):
            unpack_document(simple_docx, output)

    def test_is_symlink_entry_skips_non_unix_creator(self):
        """Non-Unix-created entries are not classified as symlinks even with S_IFLNK bits."""
        info = zipfile.ZipInfo("foo")
        info.create_system = 0  # MS-DOS / FAT
        info.external_attr = (stat.S_IFLNK | 0o777) << 16
        assert _is_symlink_entry(info) is False

    @pytest.mark.skipif(sys.platform == "win32", reason="symlinks require elevation on Windows")
    def test_unpack_rejects_symlinked_output_dir(self, temp_dir, simple_docx):
        """Refuse to unpack into an output_dir that is itself a symlink."""
        real = temp_dir / "real"
        real.mkdir()
        linked = temp_dir / "linked"
        linked.symlink_to(real, target_is_directory=True)

        with pytest.raises(InvalidDocumentError, match="Output directory is a symlink"):
            unpack_document(simple_docx, linked)

    @pytest.mark.skipif(sys.platform == "win32", reason="symlinks require elevation on Windows")
    def test_unpack_rejects_preexisting_symlink_inside_output_dir(self, temp_dir, simple_docx):
        """Refuse to extract when output_dir already contains a symlink."""
        output = temp_dir / "output"
        output.mkdir()
        external = temp_dir / "external"
        external.mkdir()
        (output / "word").symlink_to(external, target_is_directory=True)

        with pytest.raises(InvalidDocumentError, match="Symlink inside output directory"):
            unpack_document(simple_docx, output)


class TestUnpackParseErrors:
    """Parse-stage failures must raise InvalidDocumentError, not the raw parser exception (ISSUES.md #35).

    Each fixture contains exactly one bad XML part — rglob order is
    nondeterministic, so a second bad part would make the "names the
    offending part" assertion flaky.
    """

    def test_unpack_wraps_entity_dtd_in_invalid_document_error(self, temp_dir):
        """Entity-bearing DTD (defusedxml refusal) surfaces as InvalidDocumentError."""
        bad_zip = temp_dir / "bad.docx"
        _build_zip(bad_zip, [("word/document.xml", ENTITY_DTD_XML.encode("utf-8"))])
        output = temp_dir / "output"

        with pytest.raises(InvalidDocumentError, match=r"Invalid XML in word/document\.xml") as excinfo:
            unpack_document(bad_zip, output)

        assert isinstance(excinfo.value.__cause__, EntitiesForbidden)
        # The partially-extracted dir this call created must be removed.
        assert not output.exists()

    def test_unpack_wraps_truncated_xml_in_invalid_document_error(self, temp_dir):
        """Malformed/truncated XML (ExpatError) surfaces as InvalidDocumentError."""
        bad_zip = temp_dir / "bad.docx"
        _build_zip(bad_zip, [("word/document.xml", b"<w:document")])
        output = temp_dir / "output"

        with pytest.raises(InvalidDocumentError, match=r"Invalid XML in word/document\.xml") as excinfo:
            unpack_document(bad_zip, output)

        assert isinstance(excinfo.value.__cause__, ExpatError)
        assert not output.exists()

    def test_unpack_wraps_non_utf8_garbage_in_invalid_document_error(self, temp_dir):
        """Byte-level garbage must not escape as a raw UnicodeDecodeError."""
        bad_zip = temp_dir / "bad.docx"
        _build_zip(bad_zip, [("word/document.xml", b"<r>\xff\xfe</r>")])
        output = temp_dir / "output"

        with pytest.raises(InvalidDocumentError, match=r"Invalid XML in word/document\.xml") as excinfo:
            unpack_document(bad_zip, output)

        assert isinstance(excinfo.value.__cause__, ExpatError)
        assert not output.exists()

    def test_unpack_honors_part_encoding_declaration(self, temp_dir):
        """A part declaring a non-UTF-8 encoding is valid XML and must unpack."""
        u16_zip = temp_dir / "u16.docx"
        part = '<?xml version="1.0" encoding="utf-16"?><r>héllo</r>'.encode("utf-16")
        _build_zip(u16_zip, [("word/document.xml", part)])
        output = temp_dir / "output"

        unpack_document(u16_zip, output)

        # toprettyxml(encoding="utf-8") re-serializes every part as UTF-8.
        pretty = (output / "word" / "document.xml").read_text(encoding="utf-8")
        assert "héllo" in pretty

    def test_unpack_parse_failure_preserves_preexisting_output_dir(self, temp_dir):
        """Cleanup must never delete a caller's pre-existing output directory."""
        bad_zip = temp_dir / "bad.docx"
        _build_zip(bad_zip, [("word/document.xml", ENTITY_DTD_XML.encode("utf-8"))])
        output = temp_dir / "output"
        output.mkdir()
        sentinel = output / "keep.txt"
        sentinel.write_text("user data")

        with pytest.raises(InvalidDocumentError, match=r"Invalid XML in word/document\.xml"):
            unpack_document(bad_zip, output)

        assert output.exists()
        assert sentinel.read_text() == "user data"


class TestUnpackMissingParts:
    """Structurally deficient input must raise InvalidDocumentError before extraction (ISSUES.md #41)."""

    def test_unpack_missing_document_xml_raises_invalid_document_error(self, temp_dir):
        """A valid zip without word/document.xml is rejected pre-extraction."""
        junk_zip = temp_dir / "junk.docx"
        _build_zip(junk_zip, [("[Content_Types].xml", b"<Types/>")])
        output = temp_dir / "output"

        with pytest.raises(InvalidDocumentError, match=r"Not a valid \.docx: missing word/document\.xml") as excinfo:
            unpack_document(junk_zip, output)

        # Message names the document the caller opened, not a cache path.
        assert str(junk_zip) in str(excinfo.value)
        # Rejection happens before mkdir/extractall — filesystem untouched.
        assert not output.exists()

    def test_unpack_missing_part_preserves_preexisting_output_dir(self, temp_dir):
        """Rejection must never delete a caller's pre-existing output directory."""
        junk_zip = temp_dir / "junk.docx"
        _build_zip(junk_zip, [("[Content_Types].xml", b"<Types/>")])
        output = temp_dir / "output"
        output.mkdir()
        sentinel = output / "keep.txt"
        sentinel.write_text("user data")

        with pytest.raises(InvalidDocumentError, match=r"missing word/document\.xml"):
            unpack_document(junk_zip, output)

        assert output.exists()
        assert sentinel.read_text() == "user data"

    def test_unpack_directory_input_raises_invalid_document_error(self, temp_dir):
        """A directory input raises the typed error, not raw IsADirectoryError."""
        dir_input = temp_dir / "iamadir.docx"
        dir_input.mkdir()

        with pytest.raises(InvalidDocumentError, match="Is a directory") as excinfo:
            unpack_document(dir_input, temp_dir / "output")

        # Rejected by the explicit check, not by IsADirectoryError leaking
        # out of zipfile and getting wrapped downstream.
        assert excinfo.value.__cause__ is None


SMART_TEXT = "“He said ‘hello’ — it’s fine”"


def _smart_quote_document() -> str:
    """One paragraph whose single <w:t> carries smart quotes and an em-dash."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {NS}>"
        "<w:body>"
        f"<w:p><w:r><w:t>{SMART_TEXT}</w:t></w:r></w:p>"
        "</w:body>"
        "</w:document>"
    )


class TestUnpackUtf8:
    """Issue #9: unpack must write UTF-8 workspace XML so non-ASCII text stays literal.

    With ascii output, toprettyxml escapes non-ASCII characters as numeric
    character references, which the line-tracking parser fragments into
    multiple TEXT_NODEs — the root cause of the issue #9 bug class.
    """

    def test_unpack_preserves_non_ascii_text(self, simple_docx, temp_dir):
        """Unpacked document.xml keeps smart quotes/em-dash literal, no charrefs."""
        docx = temp_dir / "smart.docx"
        replace_document_xml(simple_docx, docx, _smart_quote_document())

        unpack_document(docx, temp_dir / "output")

        content = (temp_dir / "output" / "word" / "document.xml").read_text(encoding="utf-8")
        assert SMART_TEXT in content
        assert "&#" not in content

    def test_unpack_yields_single_text_node_per_wt(self, simple_docx, temp_dir):
        """Each non-empty w:t parses to exactly one TEXT_NODE via the editor parser."""
        docx = temp_dir / "smart.docx"
        replace_document_xml(simple_docx, docx, _smart_quote_document())

        unpack_document(docx, temp_dir / "output")

        editor = DocxXMLEditor(temp_dir / "output" / "word" / "document.xml", rsid="00AABBCC", author="Test")
        wts = editor.dom.getElementsByTagName("w:t")
        assert wts
        for wt in wts:
            text_children = [c for c in wt.childNodes if c.nodeType == c.TEXT_NODE]
            if text_children:
                assert len(text_children) == 1
        assert any(wt.firstChild is not None and wt.firstChild.data == SMART_TEXT for wt in wts)


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

    def test_pack_excludes_meta_json_tmp(self, simple_docx, temp_dir):
        """A meta.json.tmp orphaned by a crash mid-_save_meta must not be packed."""
        unpacked = temp_dir / "unpacked"
        unpack_document(simple_docx, unpacked)
        (unpacked / "meta.json.tmp").write_text('{"author": "te')

        output_path = temp_dir / "output.docx"
        pack_document(unpacked, output_path)

        with zipfile.ZipFile(output_path, "r") as zf:
            names = zf.namelist()
        assert "meta.json.tmp" not in names

    def test_pack_deterministic_output(self, simple_docx, temp_dir):
        """Packing the same directory twice must produce byte-identical ZIPs."""
        unpack_document(simple_docx, temp_dir / "unpacked")

        out1 = temp_dir / "first.docx"
        out2 = temp_dir / "second.docx"
        pack_document(temp_dir / "unpacked", out1)
        pack_document(temp_dir / "unpacked", out2)

        assert out1.read_bytes() == out2.read_bytes()

    def test_pack_zip_entries_sorted(self, simple_docx, temp_dir):
        """ZIP entries must be in sorted order for deterministic output."""
        unpack_document(simple_docx, temp_dir / "unpacked")

        output_path = temp_dir / "output.docx"
        pack_document(temp_dir / "unpacked", output_path)

        with zipfile.ZipFile(output_path, "r") as zf:
            names = zf.namelist()
        assert names == sorted(names)

    def test_pack_fixed_timestamps(self, simple_docx, temp_dir):
        """All ZIP entries must have the fixed epoch timestamp."""
        unpack_document(simple_docx, temp_dir / "unpacked")

        output_path = temp_dir / "output.docx"
        pack_document(temp_dir / "unpacked", output_path)

        with zipfile.ZipFile(output_path, "r") as zf:
            for info in zf.infolist():
                assert info.date_time == (1980, 1, 1, 0, 0, 0)

    @pytest.mark.skipif(sys.platform == "win32", reason="symlinks require elevation on Windows")
    def test_pack_skips_symlink_files(self, simple_docx, temp_dir):
        """File symlinks in the workspace must not be packed (would leak host content)."""
        unpacked = temp_dir / "unpacked"
        unpack_document(simple_docx, unpacked)

        # Create an external secret and replace word/document.xml with a symlink to it.
        # Packing the workspace must NOT include the symlink (we'd rather break the doc
        # than leak external content).
        secret = temp_dir / "secret.txt"
        secret.write_text("OUTSIDE")

        doc_xml = unpacked / "word" / "document.xml"
        doc_xml.unlink()
        doc_xml.symlink_to(secret)

        output_path = temp_dir / "output.docx"
        pack_document(unpacked, output_path)

        with zipfile.ZipFile(output_path, "r") as zf:
            names = zf.namelist()
        assert "word/document.xml" not in names

    @pytest.mark.skipif(sys.platform == "win32", reason="symlinks require elevation on Windows")
    def test_pack_skips_symlink_directories(self, simple_docx, temp_dir):
        """Directory symlinks must not be followed during pack."""
        unpacked = temp_dir / "unpacked"
        unpack_document(simple_docx, unpacked)

        # Create an external directory with a recognizable file and symlink it in.
        external = temp_dir / "external"
        external.mkdir()
        (external / "leaked.txt").write_text("LEAKED")
        (unpacked / "leak").symlink_to(external, target_is_directory=True)

        output_path = temp_dir / "output.docx"
        pack_document(unpacked, output_path)

        with zipfile.ZipFile(output_path, "r") as zf:
            names = zf.namelist()
        assert not any(n.startswith("leak/") for n in names)

    @pytest.mark.skipif(sys.platform == "win32", reason="symlinks require elevation on Windows")
    def test_pack_rejects_symlinked_input_dir(self, simple_docx, temp_dir):
        """Refuse to pack when input_dir itself is a symlink to an external dir."""
        unpacked = temp_dir / "unpacked"
        unpack_document(simple_docx, unpacked)

        link = temp_dir / "link"
        link.symlink_to(unpacked, target_is_directory=True)

        with pytest.raises(ValueError, match="symlink"):
            pack_document(link, temp_dir / "out.docx")


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
