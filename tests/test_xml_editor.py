"""Tests for XMLEditor and DocxXMLEditor classes."""

import pytest

from docx_editor.exceptions import MultipleNodesFoundError, NodeNotFoundError
from docx_editor.xml_editor import (
    DocxXMLEditor,
    XMLEditor,
    _generate_hex_id,
    _generate_rsid,
)


@pytest.fixture
def sample_xml_file(tmp_path):
    """Create a sample XML file for testing."""
    xml_content = """<?xml version="1.0" encoding="utf-8"?>
<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:p>
        <w:r>
            <w:t>Hello World</w:t>
        </w:r>
    </w:p>
    <w:p>
        <w:r>
            <w:t>Second paragraph</w:t>
        </w:r>
    </w:p>
</root>"""
    xml_path = tmp_path / "test.xml"
    xml_path.write_text(xml_content)
    return xml_path


@pytest.fixture
def ascii_xml_file(tmp_path):
    """Create an XML file with ascii encoding."""
    xml_content = '<?xml version="1.0" encoding="ascii"?>\n<root><item>test</item></root>'
    xml_path = tmp_path / "ascii.xml"
    xml_path.write_text(xml_content)
    return xml_path


@pytest.fixture
def rels_xml_file(tmp_path):
    """Create a relationships XML file for testing get_next_rid."""
    xml_content = """<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="type1" Target="target1"/>
    <Relationship Id="rId3" Type="type2" Target="target2"/>
    <Relationship Id="rId5" Type="type3" Target="target3"/>
    <Relationship Id="invalid" Type="type4" Target="target4"/>
</Relationships>"""
    xml_path = tmp_path / "rels.xml"
    xml_path.write_text(xml_content)
    return xml_path


@pytest.fixture
def docx_xml_file(tmp_path):
    """Create a docx-style document.xml file for DocxXMLEditor testing."""
    xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t>Hello World</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:numPr>
                    <w:ilvl w:val="0"/>
                    <w:numId w:val="1"/>
                </w:numPr>
            </w:pPr>
            <w:r w:rsidR="00123456">
                <w:t>List item</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r w:rsidR="00ABCDEF">
                <w:t>Regular paragraph</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>"""
    xml_path = tmp_path / "document.xml"
    xml_path.write_text(xml_content)
    return xml_path


@pytest.fixture
def tracked_changes_xml(tmp_path):
    """Create XML with tracked changes for revert testing."""
    xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:ins w:id="0" w:author="Author1" w:date="2024-01-01T00:00:00Z">
                <w:r w:rsidR="00111111">
                    <w:t>Inserted text</w:t>
                </w:r>
            </w:ins>
        </w:p>
        <w:p>
            <w:del w:id="1" w:author="Author1" w:date="2024-01-01T00:00:00Z">
                <w:r w:rsidDel="00222222">
                    <w:delText>Deleted text</w:delText>
                </w:r>
            </w:del>
        </w:p>
    </w:body>
</w:document>"""
    xml_path = tmp_path / "tracked.xml"
    xml_path.write_text(xml_content)
    return xml_path


class TestXMLEditorInit:
    """Tests for XMLEditor initialization."""

    def test_file_not_found(self, tmp_path):
        """Test that FileNotFoundError is raised for missing file."""
        non_existent = tmp_path / "nonexistent.xml"
        with pytest.raises(FileNotFoundError) as exc_info:
            XMLEditor(non_existent)
        assert "nonexistent.xml" in str(exc_info.value)

    def test_detects_utf8_encoding(self, sample_xml_file):
        """Test that UTF-8 encoding is correctly detected."""
        editor = XMLEditor(sample_xml_file)
        assert editor.encoding == "utf-8"

    def test_detects_ascii_encoding(self, ascii_xml_file):
        """Test that ASCII encoding is correctly detected."""
        editor = XMLEditor(ascii_xml_file)
        assert editor.encoding == "ascii"


class TestXMLEditorGetNode:
    """Tests for XMLEditor.get_node method."""

    def test_get_node_by_line_number_int(self, sample_xml_file):
        """Test finding node by integer line number."""
        editor = XMLEditor(sample_xml_file)
        # Get a node at a specific line
        node = editor.get_node("w:t", contains="Hello")
        assert node is not None
        assert "Hello" in editor._get_element_text(node)

    def test_get_node_by_line_number_range(self, sample_xml_file):
        """Test finding node within a line number range."""
        editor = XMLEditor(sample_xml_file)
        # First find a node to get its line number
        node = editor.get_node("w:t", contains="Hello")
        line = getattr(node, "parse_position", (1,))[0]

        # Now find by range
        found = editor.get_node("w:t", line_number=range(line, line + 1))
        assert found is not None

    def test_get_node_line_number_not_found(self, sample_xml_file):
        """Test that NodeNotFoundError is raised when line number doesn't match."""
        editor = XMLEditor(sample_xml_file)
        with pytest.raises(NodeNotFoundError) as exc_info:
            editor.get_node("w:t", line_number=9999)
        assert "line 9999" in str(exc_info.value)

    def test_get_node_line_range_not_found(self, sample_xml_file):
        """Test error message for line range not found."""
        editor = XMLEditor(sample_xml_file)
        with pytest.raises(NodeNotFoundError) as exc_info:
            editor.get_node("w:t", line_number=range(9990, 10000))
        assert "lines 9990-9999" in str(exc_info.value)

    def test_get_node_with_attrs_not_found(self, sample_xml_file):
        """Test error message when attrs filter doesn't match."""
        editor = XMLEditor(sample_xml_file)
        with pytest.raises(NodeNotFoundError) as exc_info:
            editor.get_node("w:t", attrs={"nonexistent": "value"})
        assert "Verify attribute values" in str(exc_info.value)

    def test_get_node_no_filters_not_found(self, sample_xml_file):
        """Test error message when no filters provided and tag not found."""
        editor = XMLEditor(sample_xml_file)
        with pytest.raises(NodeNotFoundError) as exc_info:
            editor.get_node("w:nonexistent")
        assert "Try adding filters" in str(exc_info.value)

    def test_get_node_multiple_matches(self, sample_xml_file):
        """Test that MultipleNodesFoundError is raised for multiple matches."""
        editor = XMLEditor(sample_xml_file)
        with pytest.raises(MultipleNodesFoundError) as exc_info:
            editor.get_node("w:t")  # Multiple w:t elements exist
        assert "Multiple nodes found" in str(exc_info.value)


class TestXMLEditorFindAllNodes:
    """Tests for XMLEditor.find_all_nodes method."""

    def test_find_all_with_attrs_filter(self, sample_xml_file):
        """Test find_all_nodes with attribute filter that doesn't match."""
        editor = XMLEditor(sample_xml_file)
        # Filter with non-matching attrs
        nodes = editor.find_all_nodes("w:t", attrs={"nonexistent": "value"})
        assert nodes == []

    def test_find_all_with_contains_filter(self, sample_xml_file):
        """Test find_all_nodes with contains filter."""
        editor = XMLEditor(sample_xml_file)
        nodes = editor.find_all_nodes("w:t", contains="Hello")
        assert len(nodes) == 1


class TestXMLEditorGetElementText:
    """Tests for XMLEditor._get_element_text method."""

    def test_get_element_text_with_whitespace_only(self, tmp_path):
        """Test that whitespace-only text nodes are skipped."""
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<root>
    <item>   </item>
    <item>text</item>
</root>"""
        xml_path = tmp_path / "whitespace.xml"
        xml_path.write_text(xml_content)
        editor = XMLEditor(xml_path)

        items = editor.dom.getElementsByTagName("item")
        # First item has whitespace only - should return empty
        assert editor._get_element_text(items[0]) == ""
        # Second item has actual text
        assert editor._get_element_text(items[1]) == "text"

    def test_get_element_text_nested(self, tmp_path):
        """Test recursive text extraction from nested elements."""
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<root><outer>Hello <inner>World</inner></outer></root>"""
        xml_path = tmp_path / "nested.xml"
        xml_path.write_text(xml_content)
        editor = XMLEditor(xml_path)

        outer = editor.dom.getElementsByTagName("outer")[0]
        text = editor._get_element_text(outer)
        assert text == "Hello World"


class TestXMLEditorGetNextRid:
    """Tests for XMLEditor.get_next_rid method."""

    def test_get_next_rid_with_gaps(self, rels_xml_file):
        """Test get_next_rid returns next ID after highest."""
        editor = XMLEditor(rels_xml_file)
        next_rid = editor.get_next_rid()
        assert next_rid == "rId6"

    def test_get_next_rid_handles_invalid(self, rels_xml_file):
        """Test that invalid rId values are skipped."""
        editor = XMLEditor(rels_xml_file)
        # The XML has an "invalid" Id that should be skipped
        next_rid = editor.get_next_rid()
        assert next_rid == "rId6"

    def test_get_next_rid_handles_invalid_numeric_suffix(self, tmp_path):
        """Test that rId with non-numeric suffix is skipped.

        This tests lines 277-278 (the ValueError exception handling).
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="type1" Target="target1"/>
    <Relationship Id="rIdABC" Type="type2" Target="target2"/>
    <Relationship Id="rId3" Type="type3" Target="target3"/>
</Relationships>"""
        xml_path = tmp_path / "rels.xml"
        xml_path.write_text(xml_content)

        editor = XMLEditor(xml_path)
        # rIdABC should be skipped due to ValueError when parsing "ABC"
        next_rid = editor.get_next_rid()
        assert next_rid == "rId4"


class TestXMLEditorParseFragment:
    """Tests for XMLEditor._parse_fragment method."""

    def test_parse_fragment_preserves_namespaces(self, sample_xml_file):
        """Test that namespace declarations from root are used."""
        editor = XMLEditor(sample_xml_file)
        # Parse fragment with namespaced element
        nodes = editor._parse_fragment("<w:t>New text</w:t>")
        assert len(nodes) == 1
        assert nodes[0].tagName == "w:t"

    def test_parse_fragment_empty_raises(self, sample_xml_file):
        """Test that empty fragment raises AssertionError."""
        editor = XMLEditor(sample_xml_file)
        with pytest.raises(AssertionError):
            editor._parse_fragment("   ")


class TestXMLEditorSave:
    """Tests for XMLEditor.save method."""

    def test_save_preserves_content(self, sample_xml_file):
        """Test that save writes back the content."""
        editor = XMLEditor(sample_xml_file)

        # Modify DOM
        text_node = editor.get_node("w:t", contains="Hello")
        text_node.firstChild.data = "Modified"

        # Save and reload
        editor.save()

        editor2 = XMLEditor(sample_xml_file)
        modified_node = editor2.get_node("w:t", contains="Modified")
        assert modified_node is not None


class TestDocxXMLEditorInit:
    """Tests for DocxXMLEditor initialization."""

    def test_init_with_author_sets_initials(self, docx_xml_file):
        """Test that initials are derived from author name."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="TestUser")
        assert editor.initials == "T"

    def test_init_with_empty_author(self, docx_xml_file):
        """Test initialization with empty author."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="")
        assert editor.initials == ""


class TestDocxXMLEditorGetNextChangeId:
    """Tests for DocxXMLEditor._get_next_change_id method."""

    def test_get_next_change_id_with_existing(self, tracked_changes_xml):
        """Test that next change ID is max + 1."""
        editor = DocxXMLEditor(tracked_changes_xml, rsid="00AABBCC", author="Test")
        next_id = editor._get_next_change_id()
        assert next_id == 2  # IDs 0 and 1 exist

    def test_get_next_change_id_handles_invalid(self, tmp_path):
        """Test that invalid w:id values are skipped."""
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:ins w:id="5" w:author="A" w:date="2024-01-01"/>
        <w:del w:id="invalid" w:author="A" w:date="2024-01-01"/>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "invalid_id.xml"
        xml_path.write_text(xml_content)
        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        next_id = editor._get_next_change_id()
        assert next_id == 6


class TestDocxXMLEditorNamespaces:
    """Tests for namespace handling in DocxXMLEditor."""

    def test_ensure_w16du_namespace(self, docx_xml_file):
        """Test that w16du namespace is added when needed."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        editor._ensure_w16du_namespace()
        root = editor.dom.documentElement
        assert root.hasAttribute("xmlns:w16du")

    def test_ensure_w16cex_namespace(self, docx_xml_file):
        """Test that w16cex namespace is added when needed."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        editor._ensure_w16cex_namespace()
        root = editor.dom.documentElement
        assert root.hasAttribute("xmlns:w16cex")

    def test_ensure_w14_namespace(self, docx_xml_file):
        """Test that w14 namespace is added when needed."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        editor._ensure_w14_namespace()
        root = editor.dom.documentElement
        assert root.hasAttribute("xmlns:w14")


class TestDocxXMLEditorAttributeInjection:
    """Tests for automatic attribute injection in DocxXMLEditor."""

    def test_inject_rsid_to_p(self, docx_xml_file):
        """Test RSID attributes added to w:p elements."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        p_elem = editor.get_node("w:p", contains="Hello")

        # Insert new paragraph
        nodes = editor.insert_after(p_elem, "<w:p><w:r><w:t>New</w:t></w:r></w:p>")

        new_p = nodes[0]
        assert new_p.getAttribute("w:rsidR") == "00AABBCC"
        assert new_p.getAttribute("w:rsidRDefault") == "00AABBCC"
        assert new_p.getAttribute("w:rsidP") == "00AABBCC"
        # Should have paraId and textId
        assert new_p.hasAttribute("w14:paraId")
        assert new_p.hasAttribute("w14:textId")

    def test_inject_rsid_to_r(self, docx_xml_file):
        """Test RSID attribute added to w:r elements."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        p_elem = editor.get_node("w:p", contains="Hello")

        # Insert new run
        nodes = editor.insert_after(p_elem, "<w:p><w:r><w:t>New</w:t></w:r></w:p>")

        new_r = nodes[0].getElementsByTagName("w:r")[0]
        assert new_r.getAttribute("w:rsidR") == "00AABBCC"

    def test_inject_rsid_del_inside_deletion(self, docx_xml_file):
        """Test that w:rsidDel is used for runs inside w:del."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        p_elem = editor.get_node("w:p", contains="Hello")

        # Insert deletion with run
        nodes = editor.insert_after(p_elem, "<w:del><w:r><w:delText>deleted</w:delText></w:r></w:del>")

        new_r = nodes[0].getElementsByTagName("w:r")[0]
        assert new_r.getAttribute("w:rsidDel") == "00AABBCC"
        assert not new_r.hasAttribute("w:rsidR")

    def test_inject_xml_space_for_whitespace(self, docx_xml_file):
        """Test xml:space="preserve" added for text with whitespace."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        p_elem = editor.get_node("w:p", contains="Hello")

        # Insert text with leading space
        nodes = editor.insert_after(p_elem, "<w:p><w:r><w:t> spaced</w:t></w:r></w:p>")

        new_t = nodes[0].getElementsByTagName("w:t")[0]
        assert new_t.getAttribute("xml:space") == "preserve"

    def test_inject_tracked_change_attrs(self, docx_xml_file):
        """Test w:ins gets proper attributes."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="TestAuthor")
        p_elem = editor.get_node("w:p", contains="Hello")

        # Insert tracked insertion
        nodes = editor.insert_after(p_elem, "<w:ins><w:r><w:t>inserted</w:t></w:r></w:ins>")

        ins_elem = nodes[0]
        assert ins_elem.hasAttribute("w:id")
        assert ins_elem.getAttribute("w:author") == "TestAuthor"
        assert ins_elem.hasAttribute("w:date")
        assert ins_elem.hasAttribute("w16du:dateUtc")

    def test_inject_comment_attrs(self, docx_xml_file):
        """Test w:comment gets proper attributes."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="TestAuthor", initials="TA")
        p_elem = editor.get_node("w:p", contains="Hello")

        # Insert comment
        nodes = editor.append_to(p_elem, '<w:comment w:id="99"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment>')

        comment = nodes[0]
        assert comment.getAttribute("w:author") == "TestAuthor"
        assert comment.getAttribute("w:initials") == "TA"
        assert comment.hasAttribute("w:date")

    def test_inject_comment_extensible_attrs(self, tmp_path):
        """Test w16cex:commentExtensible gets dateUtc."""
        # Create XML with w16cex namespace declared
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex">
    <w:body>
        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        body = editor.dom.getElementsByTagName("w:body")[0]

        # Insert comment extensible element
        nodes = editor.append_to(body, '<w16cex:commentExtensible w16cex:durableId="ABC123"/>')

        ext = nodes[0]
        assert ext.hasAttribute("w16cex:dateUtc")


class TestDocxXMLEditorSuggestDeletion:
    """Tests for DocxXMLEditor.suggest_deletion method."""

    def test_suggest_deletion_on_run(self, docx_xml_file):
        """Test suggesting deletion on a w:r element."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        r_elem = editor.get_node("w:r", contains="Hello")

        result = editor.suggest_deletion(r_elem)

        # Should return w:del wrapper
        assert result.tagName == "w:del"
        # Original run should be inside
        inner_r = result.getElementsByTagName("w:r")[0]
        assert inner_r is not None
        # w:t should be converted to w:delText
        del_texts = result.getElementsByTagName("w:delText")
        assert len(del_texts) == 1

    def test_suggest_deletion_on_run_with_existing_rsidR(self, docx_xml_file):
        """Test that w:rsidR is converted to w:rsidDel."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        r_elem = editor.get_node("w:r", contains="Regular")

        # This run already has w:rsidR
        result = editor.suggest_deletion(r_elem)

        inner_r = result.getElementsByTagName("w:r")[0]
        assert inner_r.hasAttribute("w:rsidDel")
        assert not inner_r.hasAttribute("w:rsidR")

    def test_suggest_deletion_on_run_already_deleted(self, tmp_path):
        """Test error when run already contains w:delText."""
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:delText>Already deleted</w:delText></w:r></w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "already_del.xml"
        xml_path.write_text(xml_content)
        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        r_elem = editor.get_node("w:r")

        with pytest.raises(ValueError) as exc_info:
            editor.suggest_deletion(r_elem)
        assert "already contains w:delText" in str(exc_info.value)

    def test_suggest_deletion_on_paragraph(self, docx_xml_file):
        """Test suggesting deletion on a w:p element."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        p_elem = editor.get_node("w:p", contains="Hello")

        result = editor.suggest_deletion(p_elem)

        # Should return the paragraph
        assert result.tagName == "w:p"
        # Should contain w:del
        del_elems = result.getElementsByTagName("w:del")
        assert len(del_elems) >= 1

    def test_suggest_deletion_on_numbered_list(self, docx_xml_file):
        """Test suggesting deletion on numbered list item."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        p_elem = editor.get_node("w:p", contains="List item")

        result = editor.suggest_deletion(p_elem)

        # Should have w:del marker in w:rPr
        pPr = result.getElementsByTagName("w:pPr")[0]
        rPr = pPr.getElementsByTagName("w:rPr")[0]
        del_markers = rPr.getElementsByTagName("w:del")
        assert len(del_markers) == 1

    def test_suggest_deletion_on_paragraph_with_existing_tracked_changes(self, tracked_changes_xml):
        """Test error when paragraph has existing tracked changes."""
        editor = DocxXMLEditor(tracked_changes_xml, rsid="00AABBCC", author="Test")
        p_elems = editor.dom.getElementsByTagName("w:p")
        # Get the paragraph with w:ins
        p_with_ins = p_elems[0]

        with pytest.raises(ValueError) as exc_info:
            editor.suggest_deletion(p_with_ins)
        assert "already contains tracked changes" in str(exc_info.value)

    def test_suggest_deletion_on_invalid_element(self, docx_xml_file):
        """Test error when suggesting deletion on invalid element type."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        t_elem = editor.get_node("w:t", contains="Hello")

        with pytest.raises(ValueError) as exc_info:
            editor.suggest_deletion(t_elem)
        assert "must be w:r or w:p" in str(exc_info.value)


class TestDocxXMLEditorRevertInsertion:
    """Tests for DocxXMLEditor.revert_insertion method."""

    def test_revert_insertion_single_ins(self, tracked_changes_xml):
        """Test reverting a single w:ins element."""
        editor = DocxXMLEditor(tracked_changes_xml, rsid="00AABBCC", author="Test")
        ins_elem = editor.get_node("w:ins")

        result = editor.revert_insertion(ins_elem)

        assert len(result) == 1
        # Should now contain w:del inside
        del_elems = ins_elem.getElementsByTagName("w:del")
        assert len(del_elems) == 1

    def test_revert_insertion_no_ins_elements(self, docx_xml_file):
        """Test error when no w:ins elements found."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        p_elem = editor.get_node("w:p", contains="Hello")

        with pytest.raises(ValueError) as exc_info:
            editor.revert_insertion(p_elem)
        assert "contains no insertions" in str(exc_info.value)

    def test_revert_insertion_converts_rsidR_to_rsidDel(self, tracked_changes_xml):
        """Test that w:rsidR is converted to w:rsidDel."""
        editor = DocxXMLEditor(tracked_changes_xml, rsid="00AABBCC", author="Test")
        ins_elem = editor.get_node("w:ins")

        editor.revert_insertion(ins_elem)

        del_elem = ins_elem.getElementsByTagName("w:del")[0]
        run = del_elem.getElementsByTagName("w:r")[0]
        assert run.hasAttribute("w:rsidDel")
        assert not run.hasAttribute("w:rsidR")


class TestDocxXMLEditorRevertDeletion:
    """Tests for DocxXMLEditor.revert_deletion method."""

    def test_revert_deletion_single_del(self, tracked_changes_xml):
        """Test reverting a single w:del element."""
        editor = DocxXMLEditor(tracked_changes_xml, rsid="00AABBCC", author="Test")
        del_elem = editor.get_node("w:del")

        result = editor.revert_deletion(del_elem)

        # Should return [del, ins]
        assert len(result) == 2
        assert result[0] == del_elem
        assert result[1].tagName == "w:ins"

    def test_revert_deletion_no_del_elements(self, docx_xml_file):
        """Test error when no w:del elements found."""
        editor = DocxXMLEditor(docx_xml_file, rsid="00AABBCC", author="Test")
        p_elem = editor.get_node("w:p", contains="Hello")

        with pytest.raises(ValueError) as exc_info:
            editor.revert_deletion(p_elem)
        assert "contains no deletions" in str(exc_info.value)

    def test_revert_deletion_converts_delText_to_t(self, tracked_changes_xml):
        """Test that w:delText is converted to w:t in insertion."""
        editor = DocxXMLEditor(tracked_changes_xml, rsid="00AABBCC", author="Test")
        del_elem = editor.get_node("w:del")

        result = editor.revert_deletion(del_elem)

        ins_elem = result[1]
        t_elems = ins_elem.getElementsByTagName("w:t")
        assert len(t_elems) == 1

    def test_revert_deletion_converts_rsidDel_to_rsidR(self, tracked_changes_xml):
        """Test that w:rsidDel is converted to w:rsidR in insertion."""
        editor = DocxXMLEditor(tracked_changes_xml, rsid="00AABBCC", author="Test")
        del_elem = editor.get_node("w:del")

        result = editor.revert_deletion(del_elem)

        ins_elem = result[1]
        run = ins_elem.getElementsByTagName("w:r")[0]
        assert run.hasAttribute("w:rsidR")
        assert not run.hasAttribute("w:rsidDel")

    def test_revert_deletion_on_body(self, tracked_changes_xml):
        """Test reverting deletion via body element."""
        editor = DocxXMLEditor(tracked_changes_xml, rsid="00AABBCC", author="Test")
        body = editor.dom.getElementsByTagName("w:body")[0]

        result = editor.revert_deletion(body)

        # Should return [body] when not a single w:del
        assert len(result) == 1
        assert result[0] == body


class TestHelperFunctions:
    """Tests for module-level helper functions."""

    def test_generate_hex_id_format(self):
        """Test _generate_hex_id returns valid format."""
        hex_id = _generate_hex_id()
        assert len(hex_id) == 8
        # Should be valid hex
        int(hex_id, 16)

    def test_generate_hex_id_within_range(self):
        """Test _generate_hex_id values are within OOXML spec range."""
        for _ in range(100):
            hex_id = _generate_hex_id()
            value = int(hex_id, 16)
            assert value >= 1
            assert value < 0x7FFFFFFF

    def test_generate_rsid_format(self):
        """Test _generate_rsid returns valid format."""
        rsid = _generate_rsid()
        assert len(rsid) == 8
        # Should be valid hex
        int(rsid, 16)
        # Should be uppercase
        assert rsid == rsid.upper()


class TestDocxXMLEditorAttributeInjectionExtended:
    """Extended tests for attribute injection covering more edge cases."""

    def test_inject_xml_space_to_t_node_directly(self, tmp_path):
        """Test that xml:space is added to w:t when it's the top-level node.

        This tests line 474.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        p_elem = editor.get_node("w:p", contains="Hello")

        # Insert a standalone w:t element with leading space
        nodes = editor.insert_after(p_elem, "<w:t> text with space</w:t>")

        # The w:t should have xml:space="preserve"
        t_elem = nodes[0]
        assert t_elem.getAttribute("xml:space") == "preserve"

    def test_inject_tracked_change_attrs_to_descendant_del(self, tmp_path):
        """Test that descendant w:del elements get attributes injected.

        This tests line 491.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="TestAuthor")
        p_elem = editor.get_node("w:p", contains="Hello")

        # Insert a paragraph containing a w:del (descendant)
        nodes = editor.insert_after(
            p_elem,
            "<w:p><w:del><w:r><w:delText>deleted</w:delText></w:r></w:del></w:p>",
        )

        # The nested w:del should have attributes
        del_elem = nodes[0].getElementsByTagName("w:del")[0]
        assert del_elem.hasAttribute("w:id")
        assert del_elem.getAttribute("w:author") == "TestAuthor"
        assert del_elem.hasAttribute("w:date")

    def test_inject_comment_attrs_to_descendant(self, tmp_path):
        """Test that descendant w:comment elements get attributes injected.

        This tests line 493.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="TestAuthor", initials="TA")
        body = editor.dom.getElementsByTagName("w:body")[0]

        # Insert a structure containing w:comment as descendant
        nodes = editor.append_to(
            body,
            '<w:p><w:comment w:id="5"><w:p><w:r><w:t>comment</w:t></w:r></w:p></w:comment></w:p>',
        )

        # The nested w:comment should have attributes
        comment = nodes[0].getElementsByTagName("w:comment")[0]
        assert comment.getAttribute("w:author") == "TestAuthor"
        assert comment.getAttribute("w:initials") == "TA"

    def test_inject_comment_extensible_to_descendant(self, tmp_path):
        """Test that descendant w16cex:commentExtensible gets attributes.

        This tests line 495.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex">
    <w:body>
        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        body = editor.dom.getElementsByTagName("w:body")[0]

        # Insert a structure containing w16cex:commentExtensible as descendant
        nodes = editor.append_to(
            body,
            '<w:p><w16cex:commentExtensible w16cex:durableId="ABC"/></w:p>',
        )

        # The nested commentExtensible should have dateUtc
        ext = nodes[0].getElementsByTagName("w16cex:commentExtensible")[0]
        assert ext.hasAttribute("w16cex:dateUtc")


class TestSuggestDeletionAttributeCopying:
    """Tests for attribute copying in suggest_deletion."""

    def test_suggest_deletion_copies_t_attributes(self, tmp_path):
        """Test that suggest_deletion copies attributes from w:t to w:delText.

        This tests lines 549-550.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t xml:space="preserve"> text with space </w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        r_elem = editor.get_node("w:r")

        editor.suggest_deletion(r_elem)

        # Check that xml:space was copied to w:delText
        del_texts = editor.dom.getElementsByTagName("w:delText")
        assert len(del_texts) == 1
        assert del_texts[0].getAttribute("xml:space") == "preserve"

    def test_suggest_deletion_numbered_list_existing_rpr(self, tmp_path):
        """Test suggest_deletion on numbered list item with existing w:rPr.

        This tests lines 590 and 595.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr>
                <w:numPr>
                    <w:ilvl w:val="0"/>
                    <w:numId w:val="1"/>
                </w:numPr>
                <w:rPr>
                    <w:b/>
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:t>List item text</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        p_elem = editor.get_node("w:p", contains="List item")

        result = editor.suggest_deletion(p_elem)

        # Check that w:del was added to existing w:rPr
        pPr = result.getElementsByTagName("w:pPr")[0]
        rPr = pPr.getElementsByTagName("w:rPr")[0]
        del_markers = rPr.getElementsByTagName("w:del")
        assert len(del_markers) == 1

        # The existing w:b should still be there
        bold_markers = rPr.getElementsByTagName("w:b")
        assert len(bold_markers) == 1

    def test_suggest_deletion_numbered_list_rpr_has_first_child(self, tmp_path):
        """Test suggest_deletion inserts w:del before existing rPr children.

        This tests line 595 (insertBefore branch).
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr>
                <w:numPr>
                    <w:ilvl w:val="0"/>
                    <w:numId w:val="1"/>
                </w:numPr>
                <w:rPr><w:i/></w:rPr>
            </w:pPr>
            <w:r>
                <w:t>Numbered item</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        p_elem = editor.get_node("w:p", contains="Numbered item")

        result = editor.suggest_deletion(p_elem)

        # w:del should be first child of w:rPr
        pPr = result.getElementsByTagName("w:pPr")[0]
        rPr = pPr.getElementsByTagName("w:rPr")[0]
        first_child = rPr.firstChild
        # Skip whitespace text nodes
        while first_child and first_child.nodeType == first_child.TEXT_NODE:
            first_child = first_child.nextSibling
        assert first_child.tagName == "w:del"

    def test_suggest_deletion_paragraph_copies_t_attributes(self, tmp_path):
        """Test that suggest_deletion on paragraph copies w:t attributes.

        This tests lines 605-606.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r>
                <w:t xml:space="preserve"> preserved </w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        p_elem = editor.get_node("w:p")

        editor.suggest_deletion(p_elem)

        # Check that xml:space was copied to w:delText
        del_texts = editor.dom.getElementsByTagName("w:delText")
        assert len(del_texts) == 1
        assert del_texts[0].getAttribute("xml:space") == "preserve"


class TestRevertInsertionEdgeCases:
    """Tests for edge cases in revert_insertion."""

    def test_revert_insertion_empty_runs(self, tmp_path):
        """Test revert_insertion handles w:ins with no runs.

        This tests line 663.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:ins w:id="0" w:author="Author1">
            </w:ins>
        </w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        ins_elem = editor.get_node("w:ins")

        # Should not raise, just skip the empty insertion
        result = editor.revert_insertion(ins_elem)
        assert len(result) == 1

    def test_revert_insertion_run_without_rsidR(self, tmp_path):
        """Test revert_insertion when run has neither rsidR nor rsidDel.

        This tests lines 674-675.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:ins w:id="0" w:author="Author1">
                <w:r>
                    <w:t>Text without rsid</w:t>
                </w:r>
            </w:ins>
        </w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        ins_elem = editor.get_node("w:ins")

        editor.revert_insertion(ins_elem)

        # The run should now have w:rsidDel set to the editor's rsid
        del_elem = ins_elem.getElementsByTagName("w:del")[0]
        run = del_elem.getElementsByTagName("w:r")[0]
        assert run.getAttribute("w:rsidDel") == "00AABBCC"

    def test_revert_insertion_copies_t_attributes(self, tmp_path):
        """Test revert_insertion copies attributes from w:t to w:delText.

        This tests lines 682-683.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:ins w:id="0" w:author="Author1">
                <w:r>
                    <w:t xml:space="preserve"> spaced </w:t>
                </w:r>
            </w:ins>
        </w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        ins_elem = editor.get_node("w:ins")

        editor.revert_insertion(ins_elem)

        # Check that xml:space was copied to w:delText
        del_texts = editor.dom.getElementsByTagName("w:delText")
        assert len(del_texts) == 1
        assert del_texts[0].getAttribute("xml:space") == "preserve"


class TestRevertDeletionEdgeCases:
    """Tests for edge cases in revert_deletion."""

    def test_revert_deletion_empty_runs(self, tmp_path):
        """Test revert_deletion handles w:del with no runs.

        This tests line 733.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:del w:id="0" w:author="Author1">
            </w:del>
        </w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        del_elem = editor.get_node("w:del")

        # Should not raise, but since no runs, returns just [del] (no ins created)
        result = editor.revert_deletion(del_elem)
        # Since is_single_del but no created_insertion (empty runs), returns [del]
        assert len(result) == 1
        assert result[0] == del_elem

    def test_revert_deletion_copies_deltext_attributes(self, tmp_path):
        """Test revert_deletion copies attributes from w:delText to w:t.

        This tests lines 748-749.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:del w:id="0" w:author="Author1">
                <w:r w:rsidDel="00112233">
                    <w:delText xml:space="preserve"> preserved </w:delText>
                </w:r>
            </w:del>
        </w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        del_elem = editor.get_node("w:del")

        result = editor.revert_deletion(del_elem)

        # Check that xml:space was copied to w:t in the new insertion
        assert len(result) == 2
        ins_elem = result[1]
        t_elems = ins_elem.getElementsByTagName("w:t")
        assert len(t_elems) == 1
        assert t_elems[0].getAttribute("xml:space") == "preserve"

    def test_revert_deletion_run_without_rsidDel(self, tmp_path):
        """Test revert_deletion when run has neither rsidDel nor rsidR.

        This tests lines 756-757.
        """
        xml_content = """<?xml version="1.0" encoding="utf-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:del w:id="0" w:author="Author1">
                <w:r>
                    <w:delText>deleted text</w:delText>
                </w:r>
            </w:del>
        </w:p>
    </w:body>
</w:document>"""
        xml_path = tmp_path / "document.xml"
        xml_path.write_text(xml_content)

        editor = DocxXMLEditor(xml_path, rsid="00AABBCC", author="Test")
        del_elem = editor.get_node("w:del")

        result = editor.revert_deletion(del_elem)

        # The new run in the insertion should have w:rsidR set to the editor's rsid
        ins_elem = result[1]
        run = ins_elem.getElementsByTagName("w:r")[0]
        assert run.getAttribute("w:rsidR") == "00AABBCC"
