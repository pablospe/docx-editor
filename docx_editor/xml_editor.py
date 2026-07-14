"""XML editing utilities for OOXML documents.

This module provides XMLEditor, a tool for manipulating XML files with support for
line-number-based node finding and DOM manipulation.
"""

import html
import io
import random
import re
import zlib
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Literal
from xml.dom.minidom import Element

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


_PARAGRAPH_REF_RE = re.compile(r"^P(\d+)#([0-9a-f]{4})$")


@dataclass
class ParagraphRef:
    """A reference to a specific paragraph by index and content hash."""

    index: int  # 1-based paragraph index
    hash: str  # 4-char lowercase hex hash

    @classmethod
    def parse(cls, ref: str) -> "ParagraphRef":
        """Parse a paragraph reference string like 'P3#a7b2'.

        Raises:
            ValueError: If the reference format is invalid
        """
        m = _PARAGRAPH_REF_RE.match(ref)
        if not m:
            raise ValueError(
                f"Invalid paragraph reference '{ref}'. Expected format: P{{index}}#{{hash}} (e.g., P3#a7b2)"
            )
        return cls(index=int(m.group(1)), hash=m.group(2))


@dataclass(frozen=True)
class ParagraphInfo:
    """Structured record for a paragraph: 1-based index, hash-anchored ref, and full visible text.

    ``str(info)`` uses the same ``"P{i}#{hash}| {text}"`` delimiter format as
    :meth:`Document.list_paragraphs`, always with the full, untruncated text.
    """

    index: int  # 1-based paragraph index
    ref: str  # "P{index}#{hash}"
    text: str  # full visible text, untruncated

    def __str__(self) -> str:
        return f"{self.ref}| {self.text}"


def compute_paragraph_hash(paragraph) -> str:
    """Compute a 4-char hex content hash for a paragraph element.

    Uses CRC32 of the paragraph's visible text (from build_text_map).
    """
    text = build_text_map(paragraph).text
    return f"{zlib.crc32(text.encode('utf-8')) & 0xFFFF:04x}"


@dataclass(frozen=True)
class TableCell:
    """Position of a paragraph's enclosing table cell.

    Coordinates are 1-based. ``col`` is the logical-grid column accounting
    for ``w:gridSpan`` of preceding cells in the same row, so a cell that
    visually sits in column 4 reports ``col=4`` even when earlier cells
    in the row are merged.
    """

    index: int  # 1-based, doc-wide, depth-first order of <w:tbl>
    row: int  # 1-based
    col: int  # 1-based logical grid (accounts for w:gridSpan)
    depth: int  # 1 = outermost, >1 = nested table


@dataclass(frozen=True)
class ListItem:
    """Raw numbering reference of a list paragraph.

    Values come straight from the paragraph's direct ``<w:pPr>/<w:numPr>``:
    ``num_id`` keys into ``word/numbering.xml``; ``ilvl`` is the 0-based
    indentation level (0 when ``<w:ilvl>`` is absent, per spec default).

    Limitations of this raw extraction: numbering inherited via a paragraph
    style (``w:pStyle``) is NOT resolved — a paragraph numbered only through
    its style reports ``list=None``. Rendered display numbers ("7.2(a)")
    are not computed (that requires resolving ``word/numbering.xml``).
    """

    num_id: int  # w:numPr/w:numId/@w:val — key into word/numbering.xml
    ilvl: int  # w:numPr/w:ilvl/@w:val — 0-based level, 0 when absent


@dataclass(frozen=True)
class ParagraphLocation:
    """Structural location of a paragraph within the document body.

    Reports table membership, list membership, and heading context. The
    shape is intentionally extensible: future releases may add other
    container kinds (header, footer, footnote), section index, etc., as
    plain optional field additions.

    ``list`` is the paragraph's raw numbering reference (see
    :class:`ListItem`), or ``None`` for non-list paragraphs.

    ``style`` is the raw ``w:pPr/w:pStyle/@w:val`` style id (a key into
    ``word/styles.xml``, e.g. ``"Heading1"``), or ``None`` when the
    paragraph carries no explicit style. No name resolution is performed.

    ``outline_level`` is the paragraph's 0-based outline level (0 ==
    Heading 1). A direct ``w:pPr/w:outlineLvl`` wins; absent that, the
    level defined by the paragraph's style (``w:basedOn`` chains resolved)
    applies. ``None`` means body text — including the spec's explicit
    "no outline" marker ``w:val="9"``.

    ``heading_path`` is the chain of nearest preceding headings that
    contains this paragraph, outermost first (e.g. ``("Chapter one",
    "Termination")``), using each heading's current visible text. A
    heading's own path lists only its ancestors, never itself.
    """

    table: TableCell | None
    list: ListItem | None = None
    style: str | None = None
    outline_level: int | None = None
    heading_path: tuple[str, ...] = ()

    @property
    def in_table(self) -> bool:
        """True iff the paragraph lives inside a ``<w:tc>`` cell."""
        return self.table is not None


def _innermost_ancestor(node, tag_name: str) -> Element | None:
    """Return the closest ancestor element with ``tag_name``, or None."""
    if node is None:  # pragma: no cover - defensive; callers guard against None
        return None
    parent = node.parentNode
    while parent is not None:
        if parent.nodeType == parent.ELEMENT_NODE and parent.tagName == tag_name:
            return parent
        parent = parent.parentNode
    return None


def _direct_grid_span(tc) -> int:
    """Return the ``w:gridSpan`` value of ``tc`` (1 if absent or malformed).

    Only inspects ``tc``'s direct ``w:tcPr`` child, so nested tables inside
    the cell never leak their own gridSpan values.
    """
    for child in tc.childNodes:
        if child.nodeType != child.ELEMENT_NODE or child.tagName != "w:tcPr":
            continue
        for gs in child.childNodes:
            if gs.nodeType == gs.ELEMENT_NODE and gs.tagName == "w:gridSpan":
                val = gs.getAttribute("w:val")
                if val:
                    try:
                        return max(1, int(val))
                    except ValueError:
                        return 1
                return 1
        return 1
    return 1


def _row_index_in_table(tbl, target_tr) -> int:
    """1-based index of ``target_tr`` among ``tbl``'s rows.

    Walks descendant ``w:tr`` elements but filters to those whose innermost
    enclosing ``w:tbl`` is ``tbl`` — so rows nested inside child tables are
    skipped, and ``<w:sdt><w:sdtContent>`` wrappers are transparent.
    """
    n = 0
    for tr in tbl.getElementsByTagName("w:tr"):
        if _innermost_ancestor(tr, "w:tbl") is not tbl:
            continue
        n += 1
        if tr is target_tr:
            return n
    raise ValueError("target_tr not found in tbl")  # pragma: no cover


def _initial_grid_offset(tr) -> int:
    """``<w:trPr>/<w:gridBefore w:val="N"/>`` — grid columns skipped at row start.

    Returns 0 when absent. A row that opens with ``gridBefore=2`` makes its
    first ``w:tc`` land at logical column 3.
    """
    for child in tr.childNodes:
        if child.nodeType != child.ELEMENT_NODE or child.tagName != "w:trPr":
            continue
        for gb in child.childNodes:
            if gb.nodeType == gb.ELEMENT_NODE and gb.tagName == "w:gridBefore":
                val = gb.getAttribute("w:val")
                if val:
                    try:
                        return max(0, int(val))
                    except ValueError:
                        return 0
                return 0
        return 0
    return 0


def _logical_col_in_row(tr, target_tc) -> int:
    """1-based logical column (gridSpan- and gridBefore-aware) of ``target_tc``.

    Walks descendant ``w:tc`` elements filtered to those whose innermost
    enclosing ``w:tr`` is ``tr`` — so cells nested inside child tables are
    skipped, and ``<w:sdt><w:sdtContent>`` cell wrappers are transparent.
    """
    col = 1 + _initial_grid_offset(tr)
    for tc in tr.getElementsByTagName("w:tc"):
        if _innermost_ancestor(tc, "w:tr") is not tr:
            continue
        if tc is target_tc:
            return col
        col += _direct_grid_span(tc)
    raise ValueError("target_tc not found in tr")  # pragma: no cover


def _doc_wide_table_index(dom, target_tbl) -> int:
    """1-based depth-first index of ``target_tbl`` among all ``<w:tbl>``."""
    for i, tbl in enumerate(dom.getElementsByTagName("w:tbl"), start=1):
        if tbl is target_tbl:
            return i
    raise ValueError("target_tbl not found in document")  # pragma: no cover


def _build_table_index(dom) -> dict:
    """Map every ``<w:tbl>`` element to its 1-based depth-first index.

    One ``getElementsByTagName("w:tbl")`` walk produces the same indices
    ``_doc_wide_table_index`` would return per call, so callers processing
    many paragraphs avoid rescanning the whole document for each one.

    minidom Elements are identity-hashable, so the node itself is the key.
    """
    return {tbl: i for i, tbl in enumerate(dom.getElementsByTagName("w:tbl"), start=1)}


def _table_depth(tbl) -> int:
    """1 = outermost table; +1 for each enclosing ``<w:tbl>`` ancestor."""
    depth = 1
    parent = tbl.parentNode
    while parent is not None:
        if parent.nodeType == parent.ELEMENT_NODE and parent.tagName == "w:tbl":
            depth += 1
        parent = parent.parentNode
    return depth


def _direct_child(elem, tag_name: str) -> Element | None:
    """Return the first direct child element of ``elem`` named ``tag_name``."""
    for child in elem.childNodes:
        if child.nodeType == child.ELEMENT_NODE and child.tagName == tag_name:
            return child
    return None


def _extract_list_item(paragraph) -> ListItem | None:
    """Raw ``<w:pPr>/<w:numPr>`` of ``paragraph``, as a :class:`ListItem`.

    Walks direct children only, so a stale ``w:numPr`` inside a
    ``<w:pPrChange>`` revision record is never picked up. Returns ``None``
    when there is no ``w:numPr``, when ``w:numId`` is absent or malformed,
    or when ``w:numId`` is 0 (the spec's "numbering disabled" marker).
    ``w:ilvl`` absent/malformed → level 0 (spec default).
    """
    ppr = _direct_child(paragraph, "w:pPr")
    if ppr is None:
        return None
    numpr = _direct_child(ppr, "w:numPr")
    if numpr is None:
        return None
    numid_elem = _direct_child(numpr, "w:numId")
    if numid_elem is None:
        return None
    try:
        num_id = int(numid_elem.getAttribute("w:val"))
    except ValueError:
        return None
    if num_id < 1:
        return None
    ilvl = 0
    ilvl_elem = _direct_child(numpr, "w:ilvl")
    if ilvl_elem is not None:
        try:
            ilvl = max(0, int(ilvl_elem.getAttribute("w:val")))
        except ValueError:
            ilvl = 0
    return ListItem(num_id=num_id, ilvl=ilvl)


def _extract_style(paragraph) -> str | None:
    """Raw ``<w:pPr>/<w:pStyle>/@w:val`` of ``paragraph``, or ``None``.

    Walks direct children only, so a stale ``w:pStyle`` inside a
    ``<w:pPrChange>`` revision record is never picked up. Returns ``None``
    when there is no ``w:pStyle`` or its ``w:val`` is absent/empty.
    """
    ppr = _direct_child(paragraph, "w:pPr")
    if ppr is None:
        return None
    pstyle = _direct_child(ppr, "w:pStyle")
    if pstyle is None:
        return None
    return pstyle.getAttribute("w:val") or None


def _parse_outline_level(outline_elem) -> int | None:
    """Parse a ``<w:outlineLvl>`` element's ``w:val`` as an outline level.

    Valid levels are ``0..8`` (0 == Heading 1). ``9`` is the spec's
    explicit "no outline level / body text" marker; it, out-of-range, and
    malformed values all return ``None``.
    """
    try:
        level = int(outline_elem.getAttribute("w:val"))
    except ValueError:
        return None
    return level if 0 <= level <= 8 else None


def _extract_outline_level(paragraph, style: str | None, style_outlines: dict[str, int]) -> int | None:
    """Effective outline level of ``paragraph`` (0-based, ``None`` = body text).

    A direct ``<w:pPr>/<w:outlineLvl>`` always wins: when the element is
    present, its value decides (``9``, out-of-range, or malformed →
    ``None`` — no style fallback, matching OOXML direct-formatting
    override semantics). When absent, the level defined by ``style`` in
    ``style_outlines`` (see :func:`_build_style_outline_map`) applies.
    Direct-child walk only, so ``<w:pPrChange>`` decoys are ignored.
    """
    ppr = _direct_child(paragraph, "w:pPr")
    outline_elem = _direct_child(ppr, "w:outlineLvl") if ppr is not None else None
    if outline_elem is not None:
        return _parse_outline_level(outline_elem)
    if style is not None:
        return style_outlines.get(style)
    return None


def _build_style_outline_map(styles_dom) -> dict[str, int]:
    """Map paragraph-style ids to their effective 0-based outline level.

    One pass over ``<w:style w:type="paragraph">`` elements collects each
    style's own ``<w:pPr>/<w:outlineLvl>`` and its ``w:basedOn`` parent; a
    second pass resolves ``basedOn`` chains (visited-set cycle guard) so a
    custom style based on e.g. ``Heading1`` without restating the level
    inherits it. A *present* ``w:outlineLvl`` terminates the chain even
    when it yields no level (``w:val="9"`` explicitly resets to body text
    — Word's ``TOCHeading`` based on ``Heading1`` relies on this). Styles
    that end up with no outline level are omitted from the map.
    """
    raw: dict[str, tuple[int | None, bool, str | None]] = {}
    for style in styles_dom.getElementsByTagName("w:style"):
        if style.getAttribute("w:type") != "paragraph":
            continue
        style_id = style.getAttribute("w:styleId")
        if not style_id:
            continue
        ppr = _direct_child(style, "w:pPr")
        outline_elem = _direct_child(ppr, "w:outlineLvl") if ppr is not None else None
        level = _parse_outline_level(outline_elem) if outline_elem is not None else None
        based_on_elem = _direct_child(style, "w:basedOn")
        based_on = based_on_elem.getAttribute("w:val") if based_on_elem is not None else None
        raw[style_id] = (level, outline_elem is not None, based_on or None)

    resolved: dict[str, int] = {}
    for style_id in raw:
        current: str | None = style_id
        visited: set[str] = set()
        while current is not None and current in raw and current not in visited:
            visited.add(current)
            level, has_own_outline, based_on = raw[current]
            if has_own_outline:
                if level is not None:
                    resolved[style_id] = level
                break
            current = based_on
    return resolved


def _compute_heading_paths(paragraphs, style_outlines: dict[str, int]) -> list[tuple[str, ...]]:
    """Heading-ancestor path for each of ``paragraphs`` (document order).

    Single forward pass maintaining a stack of open headings. A heading at
    outline level L first pops all open headings at level >= L (so its own
    recorded path lists only its ancestors, never itself or a sibling),
    then pushes its current visible text (insertions included, deletions
    excluded, per :func:`build_text_map`). Non-heading paragraphs record
    the stack as-is.
    """
    paths: list[tuple[str, ...]] = []
    stack: list[tuple[int, str]] = []  # (outline_level, heading text)
    for p in paragraphs:
        level = _extract_outline_level(p, _extract_style(p), style_outlines)
        if level is not None:
            while stack and stack[-1][0] >= level:
                stack.pop()
        paths.append(tuple(text for _, text in stack))
        if level is not None:
            stack.append((level, build_text_map(p).text))
    return paths


def _compute_paragraph_location(
    paragraph,
    table_index: dict | None = None,
    style_outlines: dict[str, int] | None = None,
    heading_path: tuple[str, ...] = (),
) -> ParagraphLocation:
    """Compute the structural location of a ``<w:p>`` element.

    Reports the innermost enclosing ``<w:tc>`` (and its table) plus the
    paragraph's raw list membership (see :func:`_extract_list_item`),
    style id, and outline level; a plain body paragraph gets ``table=None``.

    ``table_index`` is an optional ``{tbl_node: 1-based-index}`` map (see
    :func:`_build_table_index`). When supplied, the enclosing table's index
    is looked up there instead of via a whole-document rescan — the batch
    fast path. The result is identical to the ``None`` (per-call) path; a
    table missing from the map falls back to the rescan defensively.

    ``style_outlines`` is an optional precomputed ``{style_id:
    outline_level}`` map (see :func:`_build_style_outline_map`) used to
    resolve style-defined outline levels; ``None`` behaves like ``{}``.
    ``heading_path`` is threaded through verbatim — computing it needs
    whole-document context (see :func:`_compute_heading_paths`).
    """
    list_item = _extract_list_item(paragraph)
    style = _extract_style(paragraph)
    outline_level = _extract_outline_level(paragraph, style, style_outlines or {})
    tc = _innermost_ancestor(paragraph, "w:tc")
    if tc is None:
        return ParagraphLocation(
            table=None, list=list_item, style=style, outline_level=outline_level, heading_path=heading_path
        )
    tr = _innermost_ancestor(tc, "w:tr")
    tbl = _innermost_ancestor(tr, "w:tbl") if tr is not None else None
    if tr is None or tbl is None:
        # Malformed table structure — tolerate by treating as body content.
        return ParagraphLocation(
            table=None, list=list_item, style=style, outline_level=outline_level, heading_path=heading_path
        )
    if table_index is not None and tbl in table_index:
        index = table_index[tbl]
    else:
        index = _doc_wide_table_index(paragraph.ownerDocument, tbl)
    return ParagraphLocation(
        table=TableCell(
            index=index,
            row=_row_index_in_table(tbl, tr),
            col=_logical_col_in_row(tr, tc),
            depth=_table_depth(tbl),
        ),
        list=list_item,
        style=style,
        outline_level=outline_level,
        heading_path=heading_path,
    )


def _is_inside_element(node, tag_name: str) -> bool:
    """Check if a node is inside an element with the given tag."""
    parent = node.parentNode
    while parent:
        if parent.nodeType == parent.ELEMENT_NODE and parent.tagName == tag_name:
            return True
        parent = parent.parentNode
    return False


def get_text_node_data(elem) -> str:
    """Concatenate all direct TEXT_NODE children of ``elem`` into one string.

    minidom can split a single ``<w:t>``'s text across multiple TEXT_NODE
    children — reproducibly when smart quotes (U+2018/U+2019) are present
    (issue #9). Anywhere code naively read ``elem.firstChild.data`` would
    only see the first fragment. This helper returns the complete string.

    Does NOT recurse into element children. For recursive extraction across
    a subtree see ``XMLEditor._get_element_text``.
    """
    return "".join(c.data for c in elem.childNodes if c.nodeType == c.TEXT_NODE)


def build_text_map(paragraph, view: Literal["accepted", "original"] = "accepted") -> TextMap:
    """Build a text map for a paragraph element.

    Two views are supported:

    - ``"accepted"`` (default): the visible text — text inside <w:ins> is
      included (with is_inside_ins=True), text inside <w:del>/<w:delText>
      is excluded. Paragraph hashes and all editing operations use this view.
    - ``"original"``: the pre-revision text — <w:ins> and <w:moveTo> subtrees
      are excluded, <w:delText> is included, and text inside <w:del> or
      <w:moveFrom> is flagged with is_inside_del=True.

    Records the source node and offset for each character position.
    """
    if view not in ("accepted", "original"):
        raise ValueError(f"Unknown view: {view!r} (expected 'accepted' or 'original')")

    text_chars: list[str] = []
    positions: list[TextPosition] = []

    if view == "accepted":
        for node in paragraph.getElementsByTagName("w:t"):
            # Skip w:t inside w:del (deleted text uses w:delText, but be safe)
            if _is_inside_element(node, "w:del"):
                continue

            inside_ins = _is_inside_element(node, "w:ins")
            node_text = get_text_node_data(node)

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

    def collect_original(node, inside_del: bool) -> None:
        for child in node.childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            if child.tagName in ("w:ins", "w:moveTo"):
                continue
            if child.tagName in ("w:t", "w:delText"):
                node_text = get_text_node_data(child)
                for i, char in enumerate(node_text):
                    text_chars.append(char)
                    positions.append(
                        TextPosition(
                            node=child,
                            offset_in_node=i,
                            is_inside_ins=False,
                            is_inside_del=inside_del,
                        )
                    )
            else:
                collect_original(child, inside_del or child.tagName in ("w:del", "w:moveFrom"))

    collect_original(paragraph, False)
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

    def __init__(self, xml_path: str | Path, *, on_save: Callable[[], None] | None = None):
        """Initialize with path to XML file and parse with line number tracking.

        Args:
            xml_path: Path to XML file to edit (str or Path)
            on_save: Optional callback save() invokes immediately before
                writing the file — a write-ahead hook (e.g.
                Workspace.mark_dirty) so bookkeeping reaches disk even if the
                write itself crashes.

        Raises:
            FileNotFoundError: If the XML file does not exist
        """
        self.xml_path = Path(xml_path)
        self._on_save = on_save
        if not self.xml_path.exists():
            raise FileNotFoundError(f"XML file not found: {xml_path}")

        with open(self.xml_path, "rb") as f:
            header = f.read(200).decode("utf-8", errors="ignore")
        self.encoding = "ascii" if 'encoding="ascii"' in header else "utf-8"

        parser = _create_line_tracking_parser()
        self.dom = defusedxml.minidom.parse(str(self.xml_path), parser)

    def _reload_dom_from_bytes(self, xml_bytes: bytes) -> None:
        """Replace self.dom by re-parsing xml_bytes with line tracking.

        Used to restore a snapshot (e.g., rollback after a failed batch
        edit) while preserving the parse_position invariant that get_node
        with line_number relies on. Parses via BytesIO because
        defusedxml.minidom.parseString does not accept bytes when a custom
        parser is supplied.

        Args:
            xml_bytes: Raw XML bytes to parse as the new DOM.
        """
        parser = _create_line_tracking_parser()
        self.dom = defusedxml.minidom.parse(io.BytesIO(xml_bytes), parser)

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

        A whitespace-only TEXT_NODE child is treated as content (rather than
        pretty-print indentation) when any of:
          * ``elem`` carries ``xml:space="preserve"`` — the document explicitly
            marks its whitespace significant.
          * Another TEXT_NODE sibling has non-whitespace content — the lone
            whitespace is a fragment of a split text node (issue #9).
        Otherwise the whitespace TEXT_NODE is discarded as inter-element
        formatting.

        Args:
            elem: DOM element to extract text from

        Returns:
            Concatenated text content
        """
        preserve_whitespace = elem.getAttribute("xml:space") == "preserve" or any(
            n.nodeType == n.TEXT_NODE and n.data.strip() for n in elem.childNodes
        )
        text_parts = []
        for node in elem.childNodes:
            if node.nodeType == node.TEXT_NODE:
                if not node.data.strip() and not preserve_whitespace:
                    continue
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
        preserving the original encoding (ascii or utf-8). The on_save hook
        (if any) fires first, before the file is touched, so its write-ahead
        bookkeeping is on disk even if the write crashes.
        """
        if self._on_save is not None:
            self._on_save()
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

    def __init__(
        self,
        xml_path: str | Path,
        rsid: str,
        author: str,
        initials: str = "",
        *,
        on_save: Callable[[], None] | None = None,
    ):
        """Initialize with required RSID and optional author.

        Args:
            xml_path: Path to XML file to edit
            rsid: RSID to automatically apply to new elements
            author: Author name for tracked changes and comments
            initials: Author initials for comments
            on_save: Forwarded to XMLEditor — called by save() just before
                the file is written
        """
        super().__init__(xml_path, on_save=on_save)
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
            # Add xml:space="preserve" to w:t if text has leading/trailing whitespace.
            text = get_text_node_data(elem)
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
