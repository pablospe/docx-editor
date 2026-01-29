"""XML editing utilities for OOXML documents.

This module provides XMLEditor, a tool for manipulating XML files with support for
line-number-based node finding and DOM manipulation.
"""

import html
import random
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

import defusedxml.minidom
import defusedxml.sax

from .exceptions import MultipleNodesFoundError, NodeNotFoundError


@dataclass
class TextPosition:
    """A single character's position in the source XML."""

    node: object  # The <w:t> element
    offset_in_node: int  # Character offset within the node's text
    is_inside_ins: bool  # Inside <w:ins>?
    is_inside_del: bool  # Inside <w:del>?


@dataclass
class TextMap:
    """Maps visible text to source XML nodes."""

    text: str  # Concatenated visible text
    positions: list[TextPosition]  # One per character in `text`

    def find(self, search: str, start: int = 0) -> int:
        """Find text in the visible text string. Returns index or -1."""
        return self.text.find(search, start)

    def get_nodes_for_range(self, start: int, end: int) -> list[TextPosition]:
        """Get TextPosition entries for a character range."""
        return self.positions[start:end]


@dataclass
class TextMapMatch:
    """A match found in the text map."""

    start: int  # Start index in visible text
    end: int  # End index in visible text
    text: str  # The matched text
    positions: list[TextPosition]  # TextPosition entries for the match
    spans_boundary: bool  # True if match spans different revision contexts


def find_in_text_map(text_map: TextMap, search: str, occurrence: int = 0) -> TextMapMatch | None:
    """Find the nth occurrence of text in a text map.

    Returns TextMapMatch or None if not found.
    """
    start = 0
    for i in range(occurrence + 1):
        idx = text_map.find(search, start)
        if idx == -1:
            return None
        if i < occurrence:
            start = idx + 1

    end = idx + len(search)
    positions = text_map.get_nodes_for_range(idx, end)

    # Check if match spans different revision contexts
    if positions:
        first_ins = positions[0].is_inside_ins
        spans = any(p.is_inside_ins != first_ins for p in positions)
    else:
        spans = False

    return TextMapMatch(
        start=idx,
        end=end,
        text=search,
        positions=positions,
        spans_boundary=spans,
    )


def _is_inside_element(node, tag_name: str) -> bool:
    """Check if a node is inside an element with the given tag."""
    parent = node.parentNode
    while parent:
        if parent.nodeType == parent.ELEMENT_NODE and parent.tagName == tag_name:
            return True
        parent = parent.parentNode
    return False


def build_text_map(paragraph) -> TextMap:
    """Build a text map for a paragraph element.

    Walks all <w:t> elements in the paragraph, concatenating visible text
    (excluding <w:delText>) and recording the source node and offset for
    each character position.

    Text inside <w:ins> is included (it's visible) with is_inside_ins=True.
    Text inside <w:del>/<w:delText> is excluded from visible text.
    """
    text_chars: list[str] = []
    positions: list[TextPosition] = []

    for node in paragraph.getElementsByTagName("w:t"):
        # Skip w:t inside w:del (deleted text uses w:delText, but be safe)
        if _is_inside_element(node, "w:del"):
            continue

        inside_ins = _is_inside_element(node, "w:ins")
        # Get text content from the text node
        node_text = ""
        for child in node.childNodes:
            if child.nodeType == child.TEXT_NODE:
                node_text += child.data

        for i, char in enumerate(node_text):
            text_chars.append(char)
            positions.append(
                TextPosition(
                    node=node,
                    offset_in_node=i,
                    is_inside_ins=inside_ins,
                    is_inside_del=False,
                )
            )

    return TextMap(text="".join(text_chars), positions=positions)


class XMLEditor:
    """Editor for manipulating OOXML XML files with line-number-based node finding.

    This class parses XML files and tracks the original line and column position
    of each element. This enables finding nodes by their line number in the original
    file.

    Attributes:
        xml_path: Path to the XML file being edited
        encoding: Detected encoding of the XML file ('ascii' or 'utf-8')
        dom: Parsed DOM tree with parse_position attributes on elements
    """

    def __init__(self, xml_path: str | Path):
        """Initialize with path to XML file and parse with line number tracking.

        Args:
            xml_path: Path to XML file to edit (str or Path)

        Raises:
            FileNotFoundError: If the XML file does not exist
        """
        self.xml_path = Path(xml_path)
        if not self.xml_path.exists():
            raise FileNotFoundError(f"XML file not found: {xml_path}")

        with open(self.xml_path, "rb") as f:
            header = f.read(200).decode("utf-8", errors="ignore")
        self.encoding = "ascii" if 'encoding="ascii"' in header else "utf-8"

        parser = _create_line_tracking_parser()
        self.dom = defusedxml.minidom.parse(str(self.xml_path), parser)

    def get_node(
        self,
        tag: str,
        attrs: dict[str, str] | None = None,
        line_number: int | range | None = None,
        contains: str | None = None,
    ):
        """Get a DOM element by tag and identifier.

        Finds an element by either its line number in the original file or by
        matching attribute values. Exactly one match must be found.

        Args:
            tag: The XML tag name (e.g., "w:del", "w:ins", "w:r")
            attrs: Dictionary of attribute name-value pairs to match
            line_number: Line number (int) or line range (range) in original XML file
            contains: Text string that must appear in any text node within the element

        Returns:
            The matching DOM element

        Raises:
            NodeNotFoundError: If node not found
            MultipleNodesFoundError: If multiple matches found
        """
        matches = []
        for elem in self.dom.getElementsByTagName(tag):
            # Check line_number filter
            if line_number is not None:
                parse_pos = getattr(elem, "parse_position", (None,))
                elem_line = parse_pos[0]

                # Handle both single line number and range
                if isinstance(line_number, range):
                    if elem_line not in line_number:
                        continue
                else:
                    if elem_line != line_number:
                        continue

            # Check attrs filter
            if attrs is not None:
                if not all(elem.getAttribute(attr_name) == attr_value for attr_name, attr_value in attrs.items()):
                    continue

            # Check contains filter
            if contains is not None:
                elem_text = self._get_element_text(elem)
                # Normalize: convert HTML entities to Unicode characters
                normalized_contains = html.unescape(contains)
                if normalized_contains not in elem_text:
                    continue

            # If all applicable filters passed, this is a match
            matches.append(elem)

        if not matches:
            # Build descriptive error message
            filters = []
            if line_number is not None:
                line_str = (
                    f"lines {line_number.start}-{line_number.stop - 1}"
                    if isinstance(line_number, range)
                    else f"line {line_number}"
                )
                filters.append(f"at {line_str}")
            if attrs is not None:
                filters.append(f"with attributes {attrs}")
            if contains is not None:
                filters.append(f"containing '{contains}'")

            filter_desc = " ".join(filters) if filters else ""
            base_msg = f"Node not found: <{tag}> {filter_desc}".strip()

            # Add helpful hint based on filters used
            if contains:
                hint = "Text may be split across elements or use different wording."
            elif line_number:
                hint = "Line numbers may have changed if document was modified."
            elif attrs:
                hint = "Verify attribute values are correct."
            else:
                hint = "Try adding filters (attrs, line_number, or contains)."

            raise NodeNotFoundError(f"{base_msg}. {hint}")

        if len(matches) > 1:
            raise MultipleNodesFoundError(
                f"Multiple nodes found: <{tag}>. "
                f"Add more filters (attrs, line_number, or contains) to narrow the search."
            )
        return matches[0]

    def find_all_nodes(
        self,
        tag: str,
        attrs: dict[str, str] | None = None,
        contains: str | None = None,
    ) -> list:
        """Find all DOM elements matching the given criteria.

        Unlike get_node(), this returns all matches instead of requiring exactly one.

        Args:
            tag: The XML tag name (e.g., "w:t", "w:p")
            attrs: Dictionary of attribute name-value pairs to match
            contains: Text string that must appear in any text node within the element

        Returns:
            List of matching DOM elements (may be empty)
        """
        matches = []
        for elem in self.dom.getElementsByTagName(tag):
            # Check attrs filter
            if attrs is not None:
                if not all(elem.getAttribute(attr_name) == attr_value for attr_name, attr_value in attrs.items()):
                    continue

            # Check contains filter
            if contains is not None:
                elem_text = self._get_element_text(elem)
                normalized_contains = html.unescape(contains)
                if normalized_contains not in elem_text:
                    continue

            matches.append(elem)

        return matches

    def _get_element_text(self, elem) -> str:
        """Recursively extract all text content from an element.

        Skips text nodes that contain only whitespace (spaces, tabs, newlines).

        Args:
            elem: DOM element to extract text from

        Returns:
            Concatenated text from all non-whitespace text nodes
        """
        text_parts = []
        for node in elem.childNodes:
            if node.nodeType == node.TEXT_NODE:
                # Skip whitespace-only text nodes (XML formatting)
                if node.data.strip():
                    text_parts.append(node.data)
            elif node.nodeType == node.ELEMENT_NODE:
                text_parts.append(self._get_element_text(node))
        return "".join(text_parts)

    def replace_node(self, elem, new_content: str):
        """Replace a DOM element with new XML content.

        Args:
            elem: DOM element to replace
            new_content: String containing XML to replace the node with

        Returns:
            List of all inserted nodes
        """
        parent = elem.parentNode
        nodes = self._parse_fragment(new_content)
        for node in nodes:
            parent.insertBefore(node, elem)
        parent.removeChild(elem)
        return nodes

    def insert_after(self, elem, xml_content: str):
        """Insert XML content after a DOM element.

        Args:
            elem: DOM element to insert after
            xml_content: String containing XML to insert

        Returns:
            List of all inserted nodes
        """
        parent = elem.parentNode
        next_sibling = elem.nextSibling
        nodes = self._parse_fragment(xml_content)
        for node in nodes:
            if next_sibling:
                parent.insertBefore(node, next_sibling)
            else:
                parent.appendChild(node)
        return nodes

    def insert_before(self, elem, xml_content: str):
        """Insert XML content before a DOM element.

        Args:
            elem: DOM element to insert before
            xml_content: String containing XML to insert

        Returns:
            List of all inserted nodes
        """
        parent = elem.parentNode
        nodes = self._parse_fragment(xml_content)
        for node in nodes:
            parent.insertBefore(node, elem)
        return nodes

    def append_to(self, elem, xml_content: str):
        """Append XML content as a child of a DOM element.

        Args:
            elem: DOM element to append to
            xml_content: String containing XML to append

        Returns:
            List of all inserted nodes
        """
        nodes = self._parse_fragment(xml_content)
        for node in nodes:
            elem.appendChild(node)
        return nodes

    def get_next_rid(self) -> str:
        """Get the next available rId for relationships files."""
        max_id = 0
        for rel_elem in self.dom.getElementsByTagName("Relationship"):
            rel_id = rel_elem.getAttribute("Id")
            if rel_id.startswith("rId"):
                try:
                    max_id = max(max_id, int(rel_id[3:]))
                except ValueError:
                    pass
        return f"rId{max_id + 1}"

    def save(self) -> None:
        """Save the edited XML back to the file.

        Serializes the DOM tree and writes it back to the original file path,
        preserving the original encoding (ascii or utf-8).
        """
        content = self.dom.toxml(encoding=self.encoding)
        self.xml_path.write_bytes(content)

    def _parse_fragment(self, xml_content: str):
        """Parse XML fragment and return list of imported nodes.

        Args:
            xml_content: String containing XML fragment

        Returns:
            List of DOM nodes imported into this document

        Raises:
            AssertionError: If fragment contains no element nodes
        """
        # Extract namespace declarations from the root document element
        root_elem = self.dom.documentElement
        namespaces = []
        if root_elem and root_elem.attributes:
            for i in range(root_elem.attributes.length):
                attr = root_elem.attributes.item(i)
                if attr.name.startswith("xmlns"):
                    namespaces.append(f'{attr.name}="{attr.value}"')

        ns_decl = " ".join(namespaces)
        wrapper = f"<root {ns_decl}>{xml_content}</root>"
        fragment_doc = defusedxml.minidom.parseString(wrapper)
        nodes = [self.dom.importNode(child, deep=True) for child in fragment_doc.documentElement.childNodes]
        elements = [n for n in nodes if n.nodeType == n.ELEMENT_NODE]
        assert elements, "Fragment must contain at least one element"
        return nodes


class DocxXMLEditor(XMLEditor):
    """XMLEditor that automatically applies RSID, author, and date to new elements.

    Automatically adds attributes to elements that support them when inserting:
    - w:rsidR, w:rsidRDefault, w:rsidP (for w:p and w:r elements)
    - w:author and w:date (for w:ins, w:del, w:comment elements)
    - w:id (for w:ins and w:del elements)
    """

    def __init__(self, xml_path: str | Path, rsid: str, author: str, initials: str = ""):
        """Initialize with required RSID and optional author.

        Args:
            xml_path: Path to XML file to edit
            rsid: RSID to automatically apply to new elements
            author: Author name for tracked changes and comments
            initials: Author initials for comments
        """
        super().__init__(xml_path)
        self.rsid = rsid
        self.author = author
        self.initials = initials or author[0].upper() if author else ""

    def _get_next_change_id(self) -> int:
        """Get the next available change ID by checking all tracked change elements."""
        max_id = -1
        for tag in ("w:ins", "w:del"):
            elements = self.dom.getElementsByTagName(tag)
            for elem in elements:
                change_id = elem.getAttribute("w:id")
                if change_id:
                    try:
                        max_id = max(max_id, int(change_id))
                    except ValueError:
                        pass
        return max_id + 1

    def _ensure_w16du_namespace(self) -> None:
        """Ensure w16du namespace is declared on the root element."""
        root = self.dom.documentElement
        if not root.hasAttribute("xmlns:w16du"):
            root.setAttribute(
                "xmlns:w16du",
                "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
            )

    def _ensure_w16cex_namespace(self) -> None:
        """Ensure w16cex namespace is declared on the root element."""
        root = self.dom.documentElement
        if not root.hasAttribute("xmlns:w16cex"):
            root.setAttribute(
                "xmlns:w16cex",
                "http://schemas.microsoft.com/office/word/2018/wordml/cex",
            )

    def _ensure_w14_namespace(self) -> None:
        """Ensure w14 namespace is declared on the root element."""
        root = self.dom.documentElement
        if not root.hasAttribute("xmlns:w14"):
            root.setAttribute(
                "xmlns:w14",
                "http://schemas.microsoft.com/office/word/2010/wordml",
            )

    def _inject_attributes_to_nodes(self, nodes) -> None:
        """Inject RSID, author, and date attributes into DOM nodes where applicable.

        Adds attributes to elements that support them:
        - w:r: gets w:rsidR (or w:rsidDel if inside w:del)
        - w:p: gets w:rsidR, w:rsidRDefault, w:rsidP, w14:paraId, w14:textId
        - w:t: gets xml:space="preserve" if text has leading/trailing whitespace
        - w:ins, w:del: get w:id, w:author, w:date, w16du:dateUtc
        - w:comment: gets w:author, w:date, w:initials
        - w16cex:commentExtensible: gets w16cex:dateUtc
        """
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        def is_inside_deletion(elem) -> bool:
            """Check if element is inside a w:del element."""
            parent = elem.parentNode
            while parent:
                if parent.nodeType == parent.ELEMENT_NODE and parent.tagName == "w:del":
                    return True
                parent = parent.parentNode
            return False

        def add_rsid_to_p(elem) -> None:
            if not elem.hasAttribute("w:rsidR"):
                elem.setAttribute("w:rsidR", self.rsid)
            if not elem.hasAttribute("w:rsidRDefault"):
                elem.setAttribute("w:rsidRDefault", self.rsid)
            if not elem.hasAttribute("w:rsidP"):
                elem.setAttribute("w:rsidP", self.rsid)
            # Add w14:paraId and w14:textId if not present
            if not elem.hasAttribute("w14:paraId"):
                self._ensure_w14_namespace()
                elem.setAttribute("w14:paraId", _generate_hex_id())
            if not elem.hasAttribute("w14:textId"):
                self._ensure_w14_namespace()
                elem.setAttribute("w14:textId", _generate_hex_id())

        def add_rsid_to_r(elem) -> None:
            # Use w:rsidDel for <w:r> inside <w:del>, otherwise w:rsidR
            if is_inside_deletion(elem):
                if not elem.hasAttribute("w:rsidDel"):
                    elem.setAttribute("w:rsidDel", self.rsid)
            else:
                if not elem.hasAttribute("w:rsidR"):
                    elem.setAttribute("w:rsidR", self.rsid)

        def add_tracked_change_attrs(elem) -> None:
            # Auto-assign w:id if not present
            if not elem.hasAttribute("w:id"):
                elem.setAttribute("w:id", str(self._get_next_change_id()))
            if not elem.hasAttribute("w:author"):
                elem.setAttribute("w:author", self.author)
            if not elem.hasAttribute("w:date"):
                elem.setAttribute("w:date", timestamp)
            # Add w16du:dateUtc for tracked changes
            if elem.tagName in ("w:ins", "w:del") and not elem.hasAttribute("w16du:dateUtc"):
                self._ensure_w16du_namespace()
                elem.setAttribute("w16du:dateUtc", timestamp)

        def add_comment_attrs(elem) -> None:
            if not elem.hasAttribute("w:author"):
                elem.setAttribute("w:author", self.author)
            if not elem.hasAttribute("w:date"):
                elem.setAttribute("w:date", timestamp)
            if not elem.hasAttribute("w:initials"):
                elem.setAttribute("w:initials", self.initials)

        def add_comment_extensible_date(elem) -> None:
            if not elem.hasAttribute("w16cex:dateUtc"):
                self._ensure_w16cex_namespace()
                elem.setAttribute("w16cex:dateUtc", timestamp)

        def add_xml_space_to_t(elem) -> None:
            # Add xml:space="preserve" to w:t if text has leading/trailing whitespace
            if elem.firstChild and elem.firstChild.nodeType == elem.firstChild.TEXT_NODE:
                text = elem.firstChild.data
                if text and (text[0].isspace() or text[-1].isspace()):
                    if not elem.hasAttribute("xml:space"):
                        elem.setAttribute("xml:space", "preserve")

        for node in nodes:
            if node.nodeType != node.ELEMENT_NODE:
                continue

            # Handle the node itself
            if node.tagName == "w:p":
                add_rsid_to_p(node)
            elif node.tagName == "w:r":
                add_rsid_to_r(node)
            elif node.tagName == "w:t":
                add_xml_space_to_t(node)
            elif node.tagName in ("w:ins", "w:del"):
                add_tracked_change_attrs(node)
            elif node.tagName == "w:comment":
                add_comment_attrs(node)
            elif node.tagName == "w16cex:commentExtensible":
                add_comment_extensible_date(node)

            # Process descendants
            for elem in node.getElementsByTagName("w:p"):
                add_rsid_to_p(elem)
            for elem in node.getElementsByTagName("w:r"):
                add_rsid_to_r(elem)
            for elem in node.getElementsByTagName("w:t"):
                add_xml_space_to_t(elem)
            for tag in ("w:ins", "w:del"):
                for elem in node.getElementsByTagName(tag):
                    add_tracked_change_attrs(elem)
            for elem in node.getElementsByTagName("w:comment"):
                add_comment_attrs(elem)
            for elem in node.getElementsByTagName("w16cex:commentExtensible"):
                add_comment_extensible_date(elem)

    def replace_node(self, elem, new_content: str):
        """Replace node with automatic attribute injection."""
        nodes = super().replace_node(elem, new_content)
        self._inject_attributes_to_nodes(nodes)
        return nodes

    def insert_after(self, elem, xml_content: str):
        """Insert after with automatic attribute injection."""
        nodes = super().insert_after(elem, xml_content)
        self._inject_attributes_to_nodes(nodes)
        return nodes

    def insert_before(self, elem, xml_content: str):
        """Insert before with automatic attribute injection."""
        nodes = super().insert_before(elem, xml_content)
        self._inject_attributes_to_nodes(nodes)
        return nodes

    def append_to(self, elem, xml_content: str):
        """Append to with automatic attribute injection."""
        nodes = super().append_to(elem, xml_content)
        self._inject_attributes_to_nodes(nodes)
        return nodes

    def suggest_deletion(self, elem):
        """Mark a w:r or w:p element as deleted with tracked changes.

        For w:r: wraps in <w:del>, converts <w:t> to <w:delText>
        For w:p: wraps content in <w:del>, converts <w:t> to <w:delText>

        Args:
            elem: A w:r or w:p DOM element without existing tracked changes

        Returns:
            The modified element

        Raises:
            ValueError: If element has existing tracked changes or invalid structure
        """
        if elem.nodeName == "w:r":
            # Check for existing w:delText
            if elem.getElementsByTagName("w:delText"):
                raise ValueError("w:r element already contains w:delText")

            # Convert w:t to w:delText
            for t_elem in list(elem.getElementsByTagName("w:t")):
                del_text = self.dom.createElement("w:delText")
                # Copy ALL child nodes to handle entities
                while t_elem.firstChild:
                    del_text.appendChild(t_elem.firstChild)
                # Preserve attributes like xml:space
                for i in range(t_elem.attributes.length):
                    attr = t_elem.attributes.item(i)
                    del_text.setAttribute(attr.name, attr.value)
                t_elem.parentNode.replaceChild(del_text, t_elem)

            # Update run attributes: w:rsidR to w:rsidDel
            if elem.hasAttribute("w:rsidR"):
                elem.setAttribute("w:rsidDel", elem.getAttribute("w:rsidR"))
                elem.removeAttribute("w:rsidR")
            elif not elem.hasAttribute("w:rsidDel"):
                elem.setAttribute("w:rsidDel", self.rsid)

            # Wrap in w:del
            del_wrapper = self.dom.createElement("w:del")
            parent = elem.parentNode
            parent.insertBefore(del_wrapper, elem)
            parent.removeChild(elem)
            del_wrapper.appendChild(elem)

            # Inject attributes to the deletion wrapper
            self._inject_attributes_to_nodes([del_wrapper])

            return del_wrapper

        elif elem.nodeName == "w:p":
            # Check for existing tracked changes
            if elem.getElementsByTagName("w:ins") or elem.getElementsByTagName("w:del"):
                raise ValueError("w:p element already contains tracked changes")

            # Check if it's a numbered list item
            pPr_list = elem.getElementsByTagName("w:pPr")
            is_numbered = pPr_list and pPr_list[0].getElementsByTagName("w:numPr")

            if is_numbered:
                # Add <w:del/> to w:rPr in w:pPr
                pPr = pPr_list[0]
                rPr_list = pPr.getElementsByTagName("w:rPr")

                if not rPr_list:
                    rPr = self.dom.createElement("w:rPr")
                    pPr.appendChild(rPr)
                else:
                    rPr = rPr_list[0]

                # Add <w:del/> marker
                del_marker = self.dom.createElement("w:del")
                if rPr.firstChild:
                    rPr.insertBefore(del_marker, rPr.firstChild)
                else:
                    rPr.appendChild(del_marker)

            # Convert w:t to w:delText in all runs
            for t_elem in list(elem.getElementsByTagName("w:t")):
                del_text = self.dom.createElement("w:delText")
                while t_elem.firstChild:
                    del_text.appendChild(t_elem.firstChild)
                for i in range(t_elem.attributes.length):
                    attr = t_elem.attributes.item(i)
                    del_text.setAttribute(attr.name, attr.value)
                t_elem.parentNode.replaceChild(del_text, t_elem)

            # Update run attributes: w:rsidR to w:rsidDel
            for run in elem.getElementsByTagName("w:r"):
                if run.hasAttribute("w:rsidR"):
                    run.setAttribute("w:rsidDel", run.getAttribute("w:rsidR"))
                    run.removeAttribute("w:rsidR")
                elif not run.hasAttribute("w:rsidDel"):
                    run.setAttribute("w:rsidDel", self.rsid)

            # Wrap all non-pPr children in <w:del>
            del_wrapper = self.dom.createElement("w:del")
            for child in [c for c in elem.childNodes if c.nodeName != "w:pPr"]:
                elem.removeChild(child)
                del_wrapper.appendChild(child)
            elem.appendChild(del_wrapper)

            # Inject attributes to the deletion wrapper
            self._inject_attributes_to_nodes([del_wrapper])

            return elem

        else:
            raise ValueError(f"Element must be w:r or w:p, got {elem.nodeName}")

    def revert_insertion(self, elem):
        """Reject an insertion by wrapping its content in a deletion.

        Wraps all runs inside w:ins in w:del, converting w:t to w:delText.

        Args:
            elem: Element to process (w:ins, w:p, w:body, etc.)

        Returns:
            List containing the processed element(s)

        Raises:
            ValueError: If the element contains no w:ins elements
        """
        # Collect insertions
        ins_elements = []
        if elem.tagName == "w:ins":
            ins_elements.append(elem)
        else:
            ins_elements.extend(elem.getElementsByTagName("w:ins"))

        if not ins_elements:
            raise ValueError(
                f"revert_insertion requires w:ins elements. "
                f"The provided element <{elem.tagName}> contains no insertions."
            )

        # Process all insertions - wrap all children in w:del
        for ins_elem in ins_elements:
            runs = list(ins_elem.getElementsByTagName("w:r"))
            if not runs:
                continue

            # Create deletion wrapper
            del_wrapper = self.dom.createElement("w:del")

            # Process each run
            for run in runs:
                # Convert w:t to w:delText and w:rsidR to w:rsidDel
                if run.hasAttribute("w:rsidR"):
                    run.setAttribute("w:rsidDel", run.getAttribute("w:rsidR"))
                    run.removeAttribute("w:rsidR")
                elif not run.hasAttribute("w:rsidDel"):
                    run.setAttribute("w:rsidDel", self.rsid)

                for t_elem in list(run.getElementsByTagName("w:t")):
                    del_text = self.dom.createElement("w:delText")
                    while t_elem.firstChild:
                        del_text.appendChild(t_elem.firstChild)
                    for i in range(t_elem.attributes.length):
                        attr = t_elem.attributes.item(i)
                        del_text.setAttribute(attr.name, attr.value)
                    t_elem.parentNode.replaceChild(del_text, t_elem)

            # Move all children from ins to del wrapper
            while ins_elem.firstChild:
                del_wrapper.appendChild(ins_elem.firstChild)

            # Add del wrapper back to ins
            ins_elem.appendChild(del_wrapper)

            # Inject attributes to the deletion wrapper
            self._inject_attributes_to_nodes([del_wrapper])

        return [elem]

    def revert_deletion(self, elem):
        """Reject a deletion by re-inserting the deleted content.

        Creates w:ins elements after each w:del, copying deleted content.

        Args:
            elem: Element to process (w:del, w:p, w:body, etc.)

        Returns:
            List: If elem is w:del, returns [elem, new_ins]. Otherwise returns [elem].

        Raises:
            ValueError: If the element contains no w:del elements
        """
        # Collect deletions FIRST - before we modify the DOM
        del_elements = []
        is_single_del = elem.tagName == "w:del"

        if is_single_del:
            del_elements.append(elem)
        else:
            del_elements.extend(elem.getElementsByTagName("w:del"))

        if not del_elements:
            raise ValueError(
                f"revert_deletion requires w:del elements. The provided element <{elem.tagName}> contains no deletions."
            )

        # Track created insertion (only relevant if elem is a single w:del)
        created_insertion = None

        # Process all deletions - create insertions that copy the deleted content
        for del_elem in del_elements:
            runs = list(del_elem.getElementsByTagName("w:r"))
            if not runs:
                continue

            # Create insertion wrapper
            ins_elem = self.dom.createElement("w:ins")

            for run in runs:
                # Clone the run
                new_run = run.cloneNode(True)

                # Convert w:delText to w:t
                for del_text in list(new_run.getElementsByTagName("w:delText")):
                    t_elem = self.dom.createElement("w:t")
                    while del_text.firstChild:
                        t_elem.appendChild(del_text.firstChild)
                    for i in range(del_text.attributes.length):
                        attr = del_text.attributes.item(i)
                        t_elem.setAttribute(attr.name, attr.value)
                    del_text.parentNode.replaceChild(t_elem, del_text)

                # Update run attributes: w:rsidDel to w:rsidR
                if new_run.hasAttribute("w:rsidDel"):
                    new_run.setAttribute("w:rsidR", new_run.getAttribute("w:rsidDel"))
                    new_run.removeAttribute("w:rsidDel")
                elif not new_run.hasAttribute("w:rsidR"):
                    new_run.setAttribute("w:rsidR", self.rsid)

                ins_elem.appendChild(new_run)

            # Insert the new insertion after the deletion
            nodes = self.insert_after(del_elem, ins_elem.toxml())

            # If processing a single w:del, track the created insertion
            if is_single_del and nodes:
                created_insertion = nodes[0]

        # Return based on input type
        if is_single_del and created_insertion:
            return [elem, created_insertion]
        else:
            return [elem]


def _generate_hex_id() -> str:
    """Generate random 8-character hex ID for para/durable IDs.

    Values are constrained to be less than 0x7FFFFFFF per OOXML spec.
    """
    return f"{random.randint(1, 0x7FFFFFFE):08X}"


def _generate_rsid() -> str:
    """Generate random 8-character hex RSID."""
    return "".join(random.choices("0123456789ABCDEF", k=8))


def _create_line_tracking_parser():
    """Create a SAX parser that tracks line and column numbers for each element.

    Returns:
        Configured SAX parser
    """

    def set_content_handler(dom_handler):
        def startElementNS(name, tagName, attrs):
            orig_start_cb(name, tagName, attrs)
            cur_elem = dom_handler.elementStack[-1]
            cur_elem.parse_position = (
                parser._parser.CurrentLineNumber,
                parser._parser.CurrentColumnNumber,
            )

        orig_start_cb = dom_handler.startElementNS
        dom_handler.startElementNS = startElementNS
        orig_set_content_handler(dom_handler)

    parser = defusedxml.sax.make_parser()
    orig_set_content_handler = parser.setContentHandler
    parser.setContentHandler = set_content_handler
    return parser
