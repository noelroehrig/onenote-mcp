"""Write-direction builders: convert slim model instances to OneNote XML strings.

The public API is two functions:
  build_page_xml(title, outlines)  → full <one:Page> XML string (for replace_page)
  build_append_xml(outline)        → <one:Page> without ID holding the outline (for append_page)

Read-direction parsers (xml → models/dicts):
  parse_notebook_skeleton(xml) → [{"id", "name", "sections": [...]}]
  parse_section_pages(xml)     → [{"id", "name", "level"}]
  parse_page(xml)              → PageContent
"""
from __future__ import annotations

import base64
import html as _html_module
import re
import struct
import xml.etree.ElementTree as ET
import zlib
from collections.abc import Iterable, Iterator
from html.parser import HTMLParser

from onenote_mcp.models import (
    FONT_FAMILY_MAX_LENGTH, FONT_FAMILY_PATTERN, HEX_COLOR_PATTERN,
    ContentItem, FloatingImage, ImagePlaceholder, InlineImage, List, ListChild, ListItem,
    Outline, PageContent, PageOutline, Paragraph, Position, TextRun, UnsupportedItem,
    UnsupportedPageObject,
)

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
ET.register_namespace("one", _ONE_NS)

# Heading styles, written as <one:QuickStyleDef> elements so OneNote treats the
# paragraphs as real headings.  The index is local to the written XML: OneNote
# maps each definition into the page's own style table, reusing an identical one.
_HEADING_STYLE_DEFS: dict[str, dict[str, str]] = {
    "h1": {"fontSize": "16.0", "fontColor": "#1E4E79", "bold": "true"},
    "h2": {"fontSize": "14.0", "fontColor": "#1E4E79", "bold": "true"},
    "h3": {"fontSize": "13.0", "fontColor": "#2E74B5", "bold": "true"},
    "h4": {"fontSize": "12.0", "fontColor": "#2E74B5", "bold": "true"},
    "h5": {"fontSize": "11.0", "fontColor": "#595959", "bold": "true"},
    "h6": {"fontSize": "11.0", "fontColor": "#595959", "italic": "true"},
}
_HEADING_QUICK_STYLE_INDEX: dict[str, str] = {
    name: str(index) for index, name in enumerate(_HEADING_STYLE_DEFS)
}


def _quick_style_def_el(name: str) -> ET.Element:
    """Return the <one:QuickStyleDef> element for heading style *name*."""
    return ET.Element(f"{{{_ONE_NS}}}QuickStyleDef", {
        "index": _HEADING_QUICK_STYLE_INDEX[name],
        "name": name,
        "highlightColor": "automatic",
        "font": "Calibri",
        "spaceBefore": "0.0",
        "spaceAfter": "0.0",
        **_HEADING_STYLE_DEFS[name],
    })


def _prepend_quick_style_defs(page_el: ET.Element) -> None:
    """Insert the QuickStyleDefs that the headings in *page_el* refer to.

    The schema requires them before <one:Title> and all page content.
    """
    used = {oe.get("quickStyleIndex") for oe in page_el.iter(f"{{{_ONE_NS}}}OE")}
    names = [name for name, index in _HEADING_QUICK_STYLE_INDEX.items() if index in used]
    for position, name in enumerate(names):
        page_el.insert(position, _quick_style_def_el(name))


# ---- CDATA handling ---------------------------------------------------------

def _cdata_placeholder(index: int) -> str:
    return f"__CDATA_{index:06d}__"


def _apply_cdata(xml_str: str, cdata_map: dict[int, str]) -> str:
    for index, html in cdata_map.items():
        if "]]>" in html:
            raise ValueError(f"Markup must not contain ']]>', it would end the CDATA section: {html!r}")
        xml_str = xml_str.replace(_cdata_placeholder(index), f"<![CDATA[{html}]]>")
    return xml_str


def _set_t_html(
    t_el: ET.Element,
    html: str,
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> None:
    """Store an HTML fragment as the content of a <one:T>.

    All text is escaped before it becomes HTML, so a '<' can only come from our
    own span markup: such fragments go into a CDATA section, others stay plain.
    """
    if "<" not in html:
        t_el.text = html
        return
    idx = cdata_counter[0]
    cdata_counter[0] += 1
    cdata_map[idx] = html
    t_el.text = _cdata_placeholder(idx)


# ---- Formatting → HTML ------------------------------------------------------

def _escape_text(text: str) -> str:
    """Escape text for <one:T>, whose content OneNote interprets as HTML.

    Escaping '>' also guarantees that text can never produce ']]>'.
    """
    return _html_module.escape(text, quote=False)


def _format_declarations(src: "TextRun | Paragraph") -> list[str]:
    """Return CSS declarations for any object carrying the shared formatting
    fields (TextRun and Paragraph both expose bold/italic/underline/
    strikethrough/color/highlight/font_size/font_family)."""
    styles: list[str] = []
    if src.bold:
        styles.append("font-weight:bold")
    if src.italic:
        styles.append("font-style:italic")
    if src.underline and src.strikethrough:
        styles.append("text-decoration:underline line-through")
    elif src.underline:
        styles.append("text-decoration:underline")
    elif src.strikethrough:
        styles.append("text-decoration:line-through")
    if src.color is not None:
        styles.append(f"color:{src.color}")
    if src.highlight is not None:
        styles.append(f"background:{src.highlight}")
    if src.font_size is not None:
        styles.append(f"font-size:{src.font_size}pt")
    if src.font_family is not None:
        styles.append(f"font-family:{src.font_family}")
    return styles


def _wrap_span(text: str, styles: list[str]) -> str:
    """Return *text* escaped and wrapped in a styled <span>, or just escaped if no styles."""
    escaped = _escape_text(text)
    if not styles:
        return escaped
    return f'<span style="{";".join(styles)}">{escaped}</span>'


def _run_to_html(run: TextRun) -> str:
    """Return an HTML string for one TextRun."""
    return _wrap_span(run.text, _format_declarations(run))


def _paragraph_html(text: str, p: Paragraph) -> str:
    """Build the HTML string for a paragraph using the text shorthand path."""
    return _wrap_span(text, _format_declarations(p))


# ---- Paragraph → OE element -------------------------------------------------

def _paragraph_to_oe(
    p: Paragraph,
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> ET.Element:
    oe = ET.Element(f"{{{_ONE_NS}}}OE")
    if p.style != "normal":
        # Refers to the QuickStyleDef that build_page_xml/build_append_xml add.
        oe.set("quickStyleIndex", _HEADING_QUICK_STYLE_INDEX[p.style])

    t_el = ET.SubElement(oe, f"{{{_ONE_NS}}}T")

    if p.text is not None:
        _set_t_html(t_el, _paragraph_html(p.text, p), cdata_map, cdata_counter)
    elif p.segments is not None:
        html = "".join(_run_to_html(run) for run in p.segments)
        _set_t_html(t_el, html, cdata_map, cdata_counter)

    _append_children(oe, _items_to_oes(p.children, cdata_map, cdata_counter))
    return oe


def _append_children(oe: ET.Element, child_oes: list[ET.Element]) -> None:
    """Indent *child_oes* one level under *oe*, in an <one:OEChildren>.

    Nothing is added when there are none: the schema requires at least one OE.
    """
    if child_oes:
        ET.SubElement(oe, f"{{{_ONE_NS}}}OEChildren").extend(child_oes)


# ---- Image size ---------------------------------------------------------------

def _add_image_size(image_el: ET.Element, width: float | None, height: float | None) -> None:
    """Add a user-set <one:Size> with whichever dimensions are given.

    With only one of them, com.py computes the other from the image's aspect
    ratio before writing, since OneNote requires both.
    """
    if width is None and height is None:
        return
    size_el = ET.SubElement(image_el, f"{{{_ONE_NS}}}Size")
    if width is not None:
        size_el.set("width", str(width))
    if height is not None:
        size_el.set("height", str(height))
    size_el.set("isSetByUser", "true")


# ---- FloatingImage → top-level Image element --------------------------------

def _floating_image_to_el(img: FloatingImage) -> ET.Element:
    """Return a top-level <one:Image> element for a FloatingImage."""
    image_el = ET.Element(f"{{{_ONE_NS}}}Image")

    if img.position is not None:
        pos_el = ET.SubElement(image_el, f"{{{_ONE_NS}}}Position")
        pos_el.set("x", str(img.position.x))
        pos_el.set("y", str(img.position.y))
        pos_el.set("z", str(img.position.z))

    _add_image_size(image_el, img.width, img.height)

    data_el = ET.SubElement(image_el, f"{{{_ONE_NS}}}Data")
    data_el.text = img.handle

    return image_el


# ---- InlineImage → OE element -----------------------------------------------

def _inline_image_to_oe(
    img: InlineImage,
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> ET.Element:
    """Return a <one:OE> element wrapping a <one:Image> for an inline image.

    Unlike floating images, inline images have no <one:Position> element —
    they are positioned by the OEChildren flow, not the page canvas.
    """
    oe = ET.Element(f"{{{_ONE_NS}}}OE")
    image_el = ET.SubElement(oe, f"{{{_ONE_NS}}}Image")

    _add_image_size(image_el, img.width, img.height)

    data_el = ET.SubElement(image_el, f"{{{_ONE_NS}}}Data")
    data_el.text = img.handle

    _append_children(oe, _items_to_oes(img.children, cdata_map, cdata_counter))
    return oe


# ---- ImagePlaceholder → OE elements -----------------------------------------
# A placeholder is a label OE, followed by a box OE when it reserves space.
# Each carries a <one:Meta> with _PLACEHOLDER_META_NAME and its role as
# content, so get_page recognizes it whatever its text says.  OneNote keeps
# Meta on an OE but drops it on an Image, so the box image is marked by its alt
# text instead, which tells it apart from an image put into the box's paragraph.

_PLACEHOLDER_META_NAME = "onenote-mcp.image-placeholder"
_PLACEHOLDER_BOX_ALT = "Image placeholder"
_PLACEHOLDER_LABEL_STYLES = ["font-weight:bold", "background:#FFEB3B"]
# The box image has one pixel per point, scaled down to at most this many
# pixels on its longer side; its user-set Size gives the exact box size.
_PLACEHOLDER_BOX_MAX_PIXELS = 1000
_PLACEHOLDER_BOX_BORDER_PIXELS = 2
_PLACEHOLDER_BOX_FILL = bytes((0xFF, 0xF9, 0xC4))
_PLACEHOLDER_BOX_BORDER = bytes((0xB0, 0x75, 0x00))


def _placeholder_meta(role: str) -> ET.Element:
    """Return the <one:Meta> marking an OE as the placeholder's *role* ("label" or "box")."""
    return ET.Element(f"{{{_ONE_NS}}}Meta", {"name": _PLACEHOLDER_META_NAME, "content": role})


def _png_chunk(kind: bytes, data: bytes) -> bytes:
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))


def _box_png(width_px: int, height_px: int) -> bytes:
    """Return an RGB PNG of *width_px* x *height_px* pixels: a light fill inside a darker border."""
    border = _PLACEHOLDER_BOX_BORDER_PIXELS
    left = min(border, width_px)
    inner = max(width_px - 2 * border, 0)
    inner_row = (
        b"\x00" + _PLACEHOLDER_BOX_BORDER * left + _PLACEHOLDER_BOX_FILL * inner
        + _PLACEHOLDER_BOX_BORDER * (width_px - left - inner)
    )
    edge_row = b"\x00" + _PLACEHOLDER_BOX_BORDER * width_px
    rows = b"".join(
        edge_row if y < border or y >= height_px - border else inner_row for y in range(height_px)
    )
    return (
        b"\x89PNG\r\n\x1a\n"
        + _png_chunk(b"IHDR", struct.pack(">IIBBBBB", width_px, height_px, 8, 2, 0, 0, 0))
        + _png_chunk(b"IDAT", zlib.compress(rows))
        + _png_chunk(b"IEND", b"")
    )


def _box_pixel_size(width: float, height: float) -> tuple[int, int]:
    """Return the pixel size of the image for a box of *width* x *height* points."""
    scale = min(1.0, _PLACEHOLDER_BOX_MAX_PIXELS / max(width, height))
    return max(1, round(width * scale)), max(1, round(height * scale))


def _placeholder_label_oe(
    description: str,
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> ET.Element:
    """Return the OE showing *description* verbatim, bold on a yellow highlight."""
    oe = ET.Element(f"{{{_ONE_NS}}}OE")
    oe.append(_placeholder_meta("label"))
    t_el = ET.SubElement(oe, f"{{{_ONE_NS}}}T")
    _set_t_html(t_el, _wrap_span(description, _PLACEHOLDER_LABEL_STYLES), cdata_map, cdata_counter)
    return oe


def _placeholder_box_oe(width: float, height: float) -> ET.Element:
    """Return the OE holding a light box image of exactly *width* x *height* points."""
    oe = ET.Element(f"{{{_ONE_NS}}}OE")
    oe.append(_placeholder_meta("box"))
    image_el = ET.SubElement(oe, f"{{{_ONE_NS}}}Image", {"alt": _PLACEHOLDER_BOX_ALT})
    _add_image_size(image_el, width, height)
    data_el = ET.SubElement(image_el, f"{{{_ONE_NS}}}Data")
    data_el.text = base64.b64encode(_box_png(*_box_pixel_size(width, height))).decode()
    return oe


def _placeholder_to_oes(
    ph: ImagePlaceholder,
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> list[ET.Element]:
    """Return the label OE, plus the box OE when *ph* reserves space."""
    oes = [_placeholder_label_oe(ph.description, cdata_map, cdata_counter)]
    if ph.width is not None and ph.height is not None:
        oes.append(_placeholder_box_oe(ph.width, ph.height))
    return oes


# ---- List items → list of OE elements (recursive) ---------------------------

def _make_list_marker(list_style: str) -> ET.Element:
    """Return a <one:Bullet> or <one:Number> element with required attributes.

    OneNote's COM API rejects <one:Bullet /> and <one:Number /> with no
    attributes (hresult -0x7ffbdfff).  The ``bullet`` attribute on Bullet and
    ``numberSequence``/``numberFormat`` attributes on Number are required.
    """
    if list_style == "bullet":
        el = ET.Element(f"{{{_ONE_NS}}}Bullet")
        el.set("bullet", "2")
        return el
    else:
        el = ET.Element(f"{{{_ONE_NS}}}Number")
        el.set("numberSequence", "0")
        el.set("numberFormat", "##.")
        return el


def _list_items_to_oes(
    items: list[ListItem],
    list_style: str,
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> list[ET.Element]:
    result: list[ET.Element] = []

    for item in items:
        # Drop empty items (no text, no segments, no children).  An empty list
        # entry carries no content and OneNote has nothing to render; silently
        # skipping it lets an otherwise-valid page write succeed.
        if item.text is None and item.segments is None and not item.children:
            continue

        oe = ET.Element(f"{{{_ONE_NS}}}OE")

        list_el = ET.SubElement(oe, f"{{{_ONE_NS}}}List")
        list_el.append(_make_list_marker(list_style))

        t_el = ET.SubElement(oe, f"{{{_ONE_NS}}}T")

        if item.text is not None:
            t_el.text = _escape_text(item.text)
        elif item.segments is not None:
            html = "".join(_run_to_html(run) for run in item.segments)
            _set_t_html(t_el, html, cdata_map, cdata_counter)

        _append_children(oe, _list_children_to_oes(item.children, list_style, cdata_map, cdata_counter))
        result.append(oe)

    return result


def _list_children_to_oes(
    children: list[ListChild],
    list_style: str,
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> list[ET.Element]:
    """Return the OEs indented under a list item: sub-items continue the list's
    style, other content items are written as they are anywhere else."""
    result: list[ET.Element] = []
    for child in children:
        if isinstance(child, ListItem):
            result.extend(_list_items_to_oes([child], list_style, cdata_map, cdata_counter))
        else:
            result.extend(_item_to_oes(child, cdata_map, cdata_counter))
    return result


# ---- List → list of OE elements ---------------------------------------------

def _list_to_oes(
    lst: List,
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> list[ET.Element]:
    return _list_items_to_oes(lst.items, lst.style, cdata_map, cdata_counter)


# ---- Items → OE elements ----------------------------------------------------

_UNSUPPORTED_ITEM_ERROR = (
    "Cannot write an item of type 'unsupported' (kind {kind!r}): it marks content the "
    "structured tools cannot write, so rewriting this page would delete that content. "
    "If deleting it is intended, remove the item from the payload deliberately."
)


def _item_to_oes(
    item: ContentItem,
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> list[ET.Element]:
    """Return the OEs for one content item (a list yields one OE per list item).

    An UnsupportedItem is read-only and raises ValueError.
    """
    if isinstance(item, Paragraph):
        return [_paragraph_to_oe(item, cdata_map, cdata_counter)]
    if isinstance(item, ImagePlaceholder):
        return _placeholder_to_oes(item, cdata_map, cdata_counter)
    if isinstance(item, InlineImage):
        return [_inline_image_to_oe(item, cdata_map, cdata_counter)]
    if isinstance(item, List):
        return _list_to_oes(item, cdata_map, cdata_counter)
    raise ValueError(_UNSUPPORTED_ITEM_ERROR.format(kind=item.kind))


def _items_to_oes(
    items: list[ContentItem],
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> list[ET.Element]:
    return [oe for item in items for oe in _item_to_oes(item, cdata_map, cdata_counter)]


def _items_to_oechildren(
    items: list[ContentItem],
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> ET.Element:
    oe_children = ET.Element(f"{{{_ONE_NS}}}OEChildren")
    oe_children.extend(_items_to_oes(items, cdata_map, cdata_counter))
    return oe_children


# ---- Outline → Element ------------------------------------------------------

def _outline_to_el(
    outline: Outline,
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> ET.Element:
    el = ET.Element(f"{{{_ONE_NS}}}Outline")

    if outline.position is not None:
        pos_el = ET.SubElement(el, f"{{{_ONE_NS}}}Position")
        pos_el.set("x", str(outline.position.x))
        pos_el.set("y", str(outline.position.y))
        pos_el.set("z", str(outline.position.z))

    # OneNote requires <one:Size> to carry BOTH width and height — emitting
    # width-only (or height-only) is rejected with hresult 0x80042001.  When
    # the caller supplies only width, pair it with height="0.0" which OneNote
    # interprets as auto-height, preserving the column-width intent without
    # forcing the caller to pick a concrete height.
    width = outline.width
    height: float | None = None

    if width is None:
        for item in outline.items:
            if isinstance(item, (ImagePlaceholder, InlineImage)) and item.width is not None:
                width = item.width
                height = item.height
                break
    else:
        for item in outline.items:
            if isinstance(item, InlineImage) and item.height is not None:
                height = item.height
                break

    if width is not None:
        size_el = ET.SubElement(el, f"{{{_ONE_NS}}}Size")
        size_el.set("width", str(width))
        size_el.set("height", str(height if height is not None else 0.0))
        size_el.set("isSetByUser", "true")

    el.append(_items_to_oechildren(outline.items, cdata_map, cdata_counter))

    return el


# ---- Public write functions --------------------------------------------------

def build_page_xml(title: str, outlines: list[Outline], images: list[FloatingImage] | None = None) -> str:
    """Return a complete <one:Page> XML string for replace_page."""
    cdata_map: dict[int, str] = {}
    cdata_counter: list[int] = [0]

    root = ET.Element(f"{{{_ONE_NS}}}Page")

    title_el = ET.SubElement(root, f"{{{_ONE_NS}}}Title")
    title_oe = ET.SubElement(title_el, f"{{{_ONE_NS}}}OE")
    title_t = ET.SubElement(title_oe, f"{{{_ONE_NS}}}T")
    title_t.text = _escape_text(title)

    for outline in outlines:
        root.append(_outline_to_el(outline, cdata_map, cdata_counter))

    for img in (images or []):
        root.append(_floating_image_to_el(img))

    _prepend_quick_style_defs(root)
    xml_str = ET.tostring(root, encoding="unicode", xml_declaration=False)
    return _apply_cdata(xml_str, cdata_map)


def build_append_xml(outline: Outline) -> str:
    """Return a <one:Page> without ID holding *outline* and the QuickStyleDefs
    its headings refer to, for append_page."""
    cdata_map: dict[int, str] = {}
    cdata_counter: list[int] = [0]

    page_el = ET.Element(f"{{{_ONE_NS}}}Page")
    page_el.append(_outline_to_el(outline, cdata_map, cdata_counter))
    _prepend_quick_style_defs(page_el)

    xml_str = ET.tostring(page_el, encoding="unicode", xml_declaration=False)
    return _apply_cdata(xml_str, cdata_map)


# ---- Read-direction helpers -------------------------------------------------

# Numeric fallback for pages with no <one:QuickStyleDef> table: map the
# quickStyleIndex value straight to a heading level.
_NUMERIC_QUICK_STYLE_TO_HEADING: dict[str, str] = {
    "1": "h1", "2": "h2", "3": "h3", "4": "h4", "5": "h5", "6": "h6",
}


def _clean_text(raw: str) -> str:
    """Strip HTML tags and unescape HTML entities from a raw text string."""
    plain = re.sub(r"<[^>]+>", "", raw)
    return _html_module.unescape(plain)


# ---- Inline formatting parser (read direction) ------------------------------
# OneNote stores per-run formatting as inline <span style="..."> markup inside
# <one:T> (carried through CDATA).  These helpers parse that markup back into
# the slim model's formatting fields so a get_page → replace_page round-trip
# preserves bold/italic/colour/size instead of flattening to plain text.

# Matches a real start/end tag (e.g. "<span", "</span>") so that plain text
# containing a stray "<" (e.g. "a < b") is NOT mistaken for markup.
_TAG_RE = re.compile(r"</?\w+[\s/>]")


def _parse_pt(value: str) -> float | None:
    """Return the leading numeric part of a CSS length like '16pt' → 16.0."""
    m = re.match(r"\s*([0-9]+(?:\.[0-9]+)?)", value)
    return float(m.group(1)) if m else None


# Basic CSS colour keywords (the names Office uses for highlights), mapped to
# hex so colours read from OneNote validate as HexColor.
_CSS_BASIC_COLORS: dict[str, str] = {
    "black": "#000000", "silver": "#C0C0C0", "gray": "#808080", "white": "#FFFFFF",
    "maroon": "#800000", "red": "#FF0000", "purple": "#800080", "fuchsia": "#FF00FF",
    "green": "#008000", "lime": "#00FF00", "olive": "#808000", "yellow": "#FFFF00",
    "navy": "#000080", "blue": "#0000FF", "teal": "#008080", "aqua": "#00FFFF",
}


def _slim_color(value: str) -> str | None:
    """Return a CSS colour as '#RGB'/'#RRGGBB', or None if it has no such form."""
    if re.fullmatch(HEX_COLOR_PATTERN, value):
        return value
    return _CSS_BASIC_COLORS.get(value.lower())


def _slim_font_family(value: str) -> str | None:
    """Return a CSS font-family list without quotes, or None if it does not fit FontFamily."""
    family = value.replace('"', "").replace("'", "").strip()
    if len(family) <= FONT_FAMILY_MAX_LENGTH and re.fullmatch(FONT_FAMILY_PATTERN, family):
        return family
    return None


def _is_bold(weight: str) -> bool:
    """Whether a CSS font-weight value ('bold', 'normal', '700', ...) is bold."""
    weight = weight.lower()
    return "bold" in weight or (weight.isdigit() and int(weight) >= 600)


def _fmt_from_style(style_str: str) -> dict:
    """Translate a CSS ``style`` declaration string into slim formatting fields.

    Returns only the keys that are explicitly set, ready to splat into a
    TextRun/Paragraph (e.g. {"bold": True, "font_size": 16.0, "color": "#1E4E79"}).
    An explicit 'normal' or 'none' yields False, so a span can switch off
    formatting its paragraph sets.  Colour and font values the slim model
    cannot express are left out, like any other unsupported CSS property.
    """
    fmt: dict = {}
    if not style_str:
        return fmt
    decoration: str | None = None
    for decl in style_str.split(";"):
        prop, sep, val = decl.partition(":")
        if not sep:
            continue
        prop = prop.strip().lower()
        val = val.strip()
        if not val:
            continue
        if prop == "font-weight":
            fmt["bold"] = _is_bold(val)
        elif prop == "font-style":
            fmt["italic"] = "italic" in val.lower()
        elif prop == "text-decoration":
            decoration = (decoration or "") + " " + val.lower()
        elif prop == "color":
            color = _slim_color(val)
            if color is not None:
                fmt["color"] = color
        elif prop in ("background", "background-color"):
            color = _slim_color(val)
            if color is not None:
                fmt["highlight"] = color
        elif prop == "font-size":
            pt = _parse_pt(val)
            if pt is not None:
                fmt["font_size"] = pt
        elif prop == "font-family":
            family = _slim_font_family(val)
            if family is not None:
                fmt["font_family"] = family
    if decoration is not None:
        fmt["underline"] = "underline" in decoration
        fmt["strikethrough"] = "line-through" in decoration
    return fmt


class _RunCollector(HTMLParser):
    """Collect (text, formatting) runs from inline span markup, honouring nesting."""

    def __init__(self, base: dict) -> None:
        super().__init__(convert_charrefs=True)
        self.runs: list[tuple[str, dict]] = []
        self._stack: list[dict] = [dict(base)]

    def handle_starttag(self, tag, attrs):
        style = dict(self._stack[-1])
        if tag == "span":
            attr_map = {k: (v or "") for k, v in attrs}
            style.update(_fmt_from_style(attr_map.get("style", "")))
        self._stack.append(style)

    def handle_startendtag(self, tag, attrs):
        if tag == "br":
            self.runs.append(("\n", dict(self._stack[-1])))

    def handle_endtag(self, tag):
        if len(self._stack) > 1:
            self._stack.pop()

    def handle_data(self, data):
        if data:
            self.runs.append((data, dict(self._stack[-1])))


def _extract_runs(raw: str, base: dict) -> list[tuple[str, dict]]:
    """Parse a raw <one:T> string into a list of (text, formatting) runs.

    *base* is the formatting the text inherits from the style attributes of
    its OE and <one:T>; inline spans override it per run.  Plain text (no
    markup) yields a single run with the base formatting.
    """
    if not raw:
        return []
    if not _TAG_RE.search(raw):
        return [(_html_module.unescape(raw), dict(base))]
    collector = _RunCollector(base)
    try:
        collector.feed(raw)
        collector.close()
    except Exception:
        # Malformed markup — degrade gracefully to stripped plain text.
        return [(_clean_text(raw), dict(base))]
    return collector.runs


def _same_value(a: object, b: object) -> bool:
    """Compare formatting values, ignoring the case of colours and font names."""
    if isinstance(a, str) and isinstance(b, str):
        return a.lower() == b.lower()
    return a == b


def _merge_runs(runs: list[tuple[str, dict]], implied: dict) -> list[tuple[str, dict]]:
    """Drop formatting that is switched off or that *implied* (the quick style)
    already supplies, then merge adjacent runs with identical formatting so
    uniformly-styled text collapses to one run."""
    merged: list[tuple[str, dict]] = []
    for text, fmt in runs:
        fmt = {
            key: value for key, value in fmt.items()
            if value is not False and not _same_value(implied.get(key), value)
        }
        if merged and merged[-1][1] == fmt:
            merged[-1] = (merged[-1][0] + text, fmt)
        else:
            merged.append((text, fmt))
    return merged


def _oe_runs(
    oe_el: ET.Element,
    quick_style_map: dict[str, ET.Element] | None,
    reproduced_styles: frozenset[str],
) -> list[tuple[str, dict]]:
    """Return the formatted text runs of every <one:T> in *oe_el*, in order.

    OneNote moves paragraph-wide formatting into the OE's ``style`` attribute
    (a <one:T> can carry one too), so it is the base the inline spans of each
    <one:T> override.  The schema allows several <one:T> per OE; their runs
    are joined.  Formatting implied by a quick style in *reproduced_styles*
    is left out.
    """
    oe_fmt = _fmt_from_style(oe_el.get("style", ""))
    runs: list[tuple[str, dict]] = []
    for t_el in oe_el.findall(f"{{{_ONE_NS}}}T"):
        base = {**oe_fmt, **_fmt_from_style(t_el.get("style", ""))}
        runs.extend(_extract_runs(t_el.text or "", base))
    return _merge_runs(runs, _quick_style_format(oe_el, quick_style_map, reproduced_styles))


def _parse_list_item(
    oe_el: ET.Element,
    quick_style_map: dict[str, ET.Element] | None = None,
) -> ListItem:
    """Parse a single list OE element into a ListItem, recursing into children.

    Inline formatting is preserved: a uniformly-formatted item keeps its text
    plus a single styled segment; a mixed item becomes multiple segments.  (The
    ListItem model has no whole-item format fields, so any formatting is carried
    via ``segments`` rather than ``text``.)
    """
    runs = _oe_runs(oe_el, quick_style_map, _LIST_ITEM_QUICK_STYLES)
    children = _parse_list_children(oe_el, quick_style_map)

    if not runs:
        return ListItem(text="", children=children)
    if len(runs) == 1 and not runs[0][1]:
        return ListItem(text=runs[0][0], children=children)
    segments = [TextRun(text=text, **fmt) for text, fmt in runs]
    return ListItem(segments=segments, children=children)


def _parse_list_children(
    oe_el: ET.Element,
    quick_style_map: dict[str, ET.Element] | None,
) -> list[ListChild]:
    """Parse the OEs indented under a list item: list OEs become sub-items,
    any other OE the content item it holds."""
    children: list[ListChild] = []
    for child_oe, box in _with_placeholder_boxes(_child_oes(oe_el)):
        if _list_style(child_oe) is not None:
            children.append(_parse_list_item(child_oe, quick_style_map))
            continue
        item = _parse_oe(child_oe, quick_style_map, box)
        if item is not None:
            children.append(item)
    return children


_HEADING_NAMES = frozenset({"h1", "h2", "h3", "h4", "h5", "h6"})

# Quick styles whose formatting a rewrite reproduces, so it is reported once,
# as the style, rather than per run: "p" is the style OneNote gives text
# written without one, and headings come back through Paragraph.style.
_PARAGRAPH_QUICK_STYLES = _HEADING_NAMES | {"p"}
_LIST_ITEM_QUICK_STYLES = frozenset({"p"})

# Character formatting a <one:QuickStyleDef> can carry.  OneNote repeats it in
# the OE's style attribute and as inline spans in <one:T>.
_QUICK_STYLE_FLAGS = ("bold", "italic", "underline", "strikethrough")


def _resolve_heading_style(
    el: ET.Element,
    quick_style_map: dict[str, ET.Element] | None,
) -> str:
    """Resolve an OE's heading style ('normal' or 'h1'–'h6').

    A page's <one:QuickStyleDef> table is AUTHORITATIVE when present: a
    paragraph is a heading only if its quickStyleIndex maps to a heading-named
    style.  Body text can sit at any index (e.g. index 1, name "p"), so a
    non-heading name means 'normal' — never a numeric guess.  Only pages
    without any QSD table use the numeric fallback mapping.
    """
    index = el.get("quickStyleIndex")
    if index is None:
        return "normal"
    if quick_style_map:
        quick_style = quick_style_map.get(index)
        name = quick_style.get("name") if quick_style is not None else None
        return name if name in _HEADING_NAMES else "normal"
    return _NUMERIC_QUICK_STYLE_TO_HEADING.get(index, "normal")


def _quick_style_format(
    el: ET.Element,
    quick_style_map: dict[str, ET.Element] | None,
    reproduced_styles: frozenset[str],
) -> dict:
    """Return the formatting *el*'s quick style supplies as slim fields, e.g.
    {"bold": True, "font_size": 16.0}, or {} unless it is in *reproduced_styles*."""
    quick_style = quick_style_map.get(el.get("quickStyleIndex")) if quick_style_map else None
    if quick_style is None or quick_style.get("name") not in reproduced_styles:
        return {}
    fmt: dict = {flag: True for flag in _QUICK_STYLE_FLAGS if quick_style.get(flag) == "true"}
    values = {
        "font_family": _slim_font_family(quick_style.get("font", "")),
        "font_size": _parse_pt(quick_style.get("fontSize", "")),
        "color": _slim_color(quick_style.get("fontColor", "")),
        "highlight": _slim_color(quick_style.get("highlightColor", "")),
    }
    fmt.update({field: value for field, value in values.items() if value is not None})
    return fmt


def _parse_paragraph_oe(
    el: ET.Element,
    quick_style_map: dict[str, ET.Element] | None,
    children: list[ContentItem],
) -> Paragraph:
    """Parse an OE element containing text into a Paragraph.

    Heading style is resolved via _resolve_heading_style (QSD table
    authoritative, numeric fallback only when no QSD table exists).
    Formatting from the OE and <one:T> style attributes and inline
    <span style="…"> markup is preserved, except what the quick style already
    supplies: uniformly-formatted text collapses to ``text`` plus
    whole-paragraph format fields, while mixed formatting becomes ``segments``.
    """
    style = _resolve_heading_style(el, quick_style_map)
    runs = _oe_runs(el, quick_style_map, _PARAGRAPH_QUICK_STYLES)

    if not runs:
        return Paragraph(type="paragraph", text="", style=style, children=children)
    if len(runs) == 1:
        text, fmt = runs[0]
        return Paragraph(type="paragraph", text=text, style=style, children=children, **fmt)
    segments = [TextRun(text=text, **fmt) for text, fmt in runs]
    return Paragraph(type="paragraph", segments=segments, style=style, children=children)


def _parse_size(el: ET.Element) -> tuple[float | None, float | None]:
    """Return the width and height of *el*'s <one:Size>, each None when absent."""
    size_el = el.find(f"{{{_ONE_NS}}}Size")
    if size_el is None:
        return None, None
    width, height = size_el.get("width"), size_el.get("height")
    return (
        float(width) if width is not None else None,
        float(height) if height is not None else None,
    )


def _parse_position(el: ET.Element) -> Position | None:
    """Return *el*'s <one:Position> as a Position, or None when absent."""
    pos_el = el.find(f"{{{_ONE_NS}}}Position")
    if pos_el is None:
        return None
    return Position(
        x=float(pos_el.get("x", 0)),
        y=float(pos_el.get("y", 0)),
        z=int(pos_el.get("z", 0)),
    )


def _parse_inline_image(img_el: ET.Element, children: list[ContentItem]) -> InlineImage | None:
    """Parse an outline <one:Image> into an InlineImage, or None without a handle."""
    data_el = img_el.find(f"{{{_ONE_NS}}}Data")
    if data_el is None or not (data_el.text or "").strip():
        return None  # missing/empty handle: treat as absent, don't raise
    width, height = _parse_size(img_el)
    return InlineImage(
        type="inline_image", handle=data_el.text.strip(), width=width, height=height,
        children=children,
    )


# ---- Content the slim model cannot represent (read direction) ---------------

# Elements holding such content, by the kind they are reported as.  Per the
# OneNote 2013 schema all of them can be the content of an OE (InkWord mixed
# with <one:T>); those in UNSUPPORTED_PAGE_OBJECT_TAGS can also sit directly
# on the page canvas.
_UNSUPPORTED_KINDS: dict[str, str] = {
    f"{{{_ONE_NS}}}{name}": kind
    for name, kind in (
        ("Table", "table"),
        ("InkDrawing", "ink"),
        ("InkParagraph", "ink"),
        ("InkWord", "ink"),
        ("InsertedFile", "file"),
        ("MediaFile", "media"),
        ("FutureObject", "unknown"),
    )
}

# The unsupported elements the schema allows as direct <one:Page> children.
# get_page reports them in PageContent.unsupported; replace_page keeps them.
UNSUPPORTED_PAGE_OBJECT_TAGS: frozenset[str] = frozenset(
    f"{{{_ONE_NS}}}{name}" for name in ("InkDrawing", "InsertedFile", "MediaFile", "FutureObject")
)


def _unsupported_content(oe_el: ET.Element) -> ET.Element | None:
    """Return the first child of *oe_el* the slim model cannot represent, if any."""
    return next((child for child in oe_el if child.tag in _UNSUPPORTED_KINDS), None)


def _plain_text(elements: Iterable[ET.Element]) -> str:
    """Join the text of the <one:T> and the recognized text of the <one:InkWord>
    elements among *elements*, in order, with single spaces."""
    pieces: list[str] = []
    for el in elements:
        if el.tag == f"{{{_ONE_NS}}}T":
            pieces.append(_clean_text(el.text or "").strip())
        elif el.tag == f"{{{_ONE_NS}}}InkWord":
            pieces.append(el.get("recognizedText", "").strip())
    return " ".join(piece for piece in pieces if piece)


def _table_text(table_el: ET.Element) -> str:
    """Return a table's cell text: one line per row, cells separated by ' | '."""
    return "\n".join(
        " | ".join(_plain_text(cell.iter()) for cell in row.findall(f"{{{_ONE_NS}}}Cell"))
        for row in table_el.findall(f"{{{_ONE_NS}}}Row")
    )


def _readable_text(el: ET.Element) -> str:
    """Return the plain text an agent can read for an unsupported element."""
    if el.tag == f"{{{_ONE_NS}}}Table":
        return _table_text(el)
    if el.tag in (f"{{{_ONE_NS}}}InsertedFile", f"{{{_ONE_NS}}}MediaFile"):
        return el.get("preferredName", "")
    return _plain_text(el.iter())


def _parse_unsupported_oe(oe_el: ET.Element, children: list[ContentItem]) -> UnsupportedItem:
    """Parse an OE holding unsupported content into an UnsupportedItem.

    Its text covers the whole OE content, so typed text that shares the OE
    with ink words is kept with them.
    """
    texts = (_readable_text(child) for child in oe_el if child.tag != f"{{{_ONE_NS}}}OEChildren")
    return UnsupportedItem(
        type="unsupported",
        kind=_UNSUPPORTED_KINDS[_unsupported_content(oe_el).tag],
        text=" ".join(text for text in texts if text),
        children=children,
    )


def _parse_unsupported_page_object(el: ET.Element) -> UnsupportedPageObject:
    """Parse an unsupported direct child of <one:Page> into an UnsupportedPageObject."""
    width, height = _parse_size(el)
    return UnsupportedPageObject(
        kind=_UNSUPPORTED_KINDS[el.tag],
        text=_readable_text(el),
        position=_parse_position(el),
        width=width,
        height=height,
    )


# ---- Image placeholders (read direction) ------------------------------------

def _placeholder_role(oe_el: ET.Element) -> str | None:
    """Return the placeholder role ("label" or "box") the Meta of *oe_el* marks, or None."""
    return next(
        (meta.get("content") for meta in oe_el.findall(f"{{{_ONE_NS}}}Meta")
         if meta.get("name") == _PLACEHOLDER_META_NAME),
        None,
    )


def _is_standalone(oe_el: ET.Element) -> bool:
    """Whether *oe_el* has neither a list marker nor indented children.

    A placeholder OE that gained either is read as ordinary content, which
    the image_placeholder item could not hold.
    """
    return oe_el.find(f"{{{_ONE_NS}}}List") is None and oe_el.find(f"{{{_ONE_NS}}}OEChildren") is None


def _is_placeholder_label(oe_el: ET.Element) -> bool:
    return _placeholder_role(oe_el) == "label" and _is_standalone(oe_el)


def _is_placeholder_box(oe_el: ET.Element) -> bool:
    """Whether *oe_el* is a placeholder box still holding the generated image."""
    image_el = oe_el.find(f"{{{_ONE_NS}}}Image")
    return (
        _placeholder_role(oe_el) == "box" and _is_standalone(oe_el)
        and image_el is not None and image_el.get("alt") == _PLACEHOLDER_BOX_ALT
    )


def _with_placeholder_boxes(
    oe_elements: list[ET.Element],
) -> Iterator[tuple[ET.Element, ET.Element | None]]:
    """Yield each OE with None, except a placeholder label directly followed by
    a placeholder box: it is yielded with that box, which is not yielded alone.

    A box without its label is yielded alone and reads as an inline image.
    """
    index = 0
    while index < len(oe_elements):
        oe = oe_elements[index]
        following = oe_elements[index + 1] if index + 1 < len(oe_elements) else None
        if following is not None and _is_placeholder_label(oe) and _is_placeholder_box(following):
            yield oe, following
            index += 2
        else:
            yield oe, None
            index += 1


def _oe_text(oe_el: ET.Element) -> str:
    """Return the text of every <one:T> in *oe_el*, without its formatting."""
    return "".join(
        text
        for t_el in oe_el.findall(f"{{{_ONE_NS}}}T")
        for text, _ in _extract_runs(t_el.text or "", {})
    )


def _parse_placeholder(label_oe: ET.Element, box_oe: ET.Element | None) -> ImagePlaceholder:
    """Parse a placeholder label, and the box after it if any, into an ImagePlaceholder."""
    width, height = _parse_size(box_oe.find(f"{{{_ONE_NS}}}Image")) if box_oe is not None else (None, None)
    return ImagePlaceholder(
        type="image_placeholder", description=_oe_text(label_oe), width=width, height=height,
    )


# ---- OE classification (read direction) -------------------------------------

def _child_oes(el: ET.Element) -> list[ET.Element]:
    """Return the OEs of every <one:OEChildren> directly under *el*, in order."""
    return [
        oe
        for oe_children_el in el.findall(f"{{{_ONE_NS}}}OEChildren")
        for oe in oe_children_el.findall(f"{{{_ONE_NS}}}OE")
    ]


def _list_style(oe_el: ET.Element) -> str | None:
    """Return 'bullet' or 'numbered' for a list OE holding text, else None.

    A bulleted image, table or ink OE is read as that content, which keeps
    the content but not its bullet.
    """
    list_el = oe_el.find(f"{{{_ONE_NS}}}List")
    if list_el is None or oe_el.find(f"{{{_ONE_NS}}}Image") is not None:
        return None
    if _unsupported_content(oe_el) is not None:
        return None
    return "bullet" if list_el.find(f"{{{_ONE_NS}}}Bullet") is not None else "numbered"


def _parse_oe(
    el: ET.Element,
    quick_style_map: dict[str, ET.Element] | None = None,
    box: ET.Element | None = None,
) -> "ContentItem | None":
    """Classify a non-list OE element and return the matching ContentItem.

    *box* is the placeholder box paired with a placeholder label by
    _with_placeholder_boxes.  Returns None for an image without a handle.  The
    OE's indented children are parsed into the item's ``children``.
    """
    if _is_placeholder_label(el):
        return _parse_placeholder(el, box)
    children = _parse_outline_items(_child_oes(el), quick_style_map)
    if _unsupported_content(el) is not None:
        return _parse_unsupported_oe(el, children)
    img_el = el.find(f"{{{_ONE_NS}}}Image")
    if img_el is not None:
        return _parse_inline_image(img_el, children)
    return _parse_paragraph_oe(el, quick_style_map, children)


def _parse_outline_items(
    oe_elements: list[ET.Element],
    quick_style_map: dict[str, ET.Element] | None = None,
) -> list["ContentItem"]:
    """Walk OE elements and group consecutive list OEs into List objects."""
    result: list[ContentItem] = []
    current_list_style: str | None = None
    current_list_items: list[ListItem] = []

    def _flush_list() -> None:
        nonlocal current_list_style, current_list_items
        if current_list_style is not None and current_list_items:
            result.append(List(type="list", style=current_list_style, items=current_list_items))
        current_list_style = None
        current_list_items = []

    for oe, box in _with_placeholder_boxes(oe_elements):
        style = _list_style(oe)
        if style is not None:
            if style != current_list_style:
                _flush_list()
                current_list_style = style

            current_list_items.append(_parse_list_item(oe, quick_style_map))
        else:
            _flush_list()
            item = _parse_oe(oe, quick_style_map, box)
            if item is not None:
                result.append(item)

    _flush_list()
    return result


# ---- Public read functions --------------------------------------------------

def parse_notebook_skeleton(xml: str) -> list[dict]:
    """Parse hierarchy XML into notebooks, each with its sections (id+name, no pages).

    Shape: [{"id", "name", "sections": [{"id", "name"}]}].  This is the lean
    navigation skeleton returned by get_notebooks — small enough to fetch the
    whole notebook/section map without pulling any pages.
    """
    root = ET.fromstring(xml)
    notebooks: list[dict] = []
    for nb_el in root.findall(f"{{{_ONE_NS}}}Notebook"):
        sections = [
            {"id": sec_el.get("ID", ""), "name": sec_el.get("name", "")}
            for sec_el in nb_el.findall(f"{{{_ONE_NS}}}Section")
        ]
        notebooks.append({
            "id": nb_el.get("ID", ""),
            "name": nb_el.get("name", ""),
            "sections": sections,
        })
    return notebooks


def parse_section_pages(xml: str) -> list[dict]:
    """Parse a section-scoped hierarchy XML into its pages as [{"id", "name", "level"}].

    The input root is a <one:Section> (section-anchored GetHierarchy); pages are
    its direct <one:Page> children, in section order.  ``level`` is the page's
    pageLevel, 1 for a top-level page (also when OneNote omits the attribute).
    Returns an empty list if the section has no pages.
    """
    root = ET.fromstring(xml)
    return [
        {
            "id": page_el.get("ID", ""),
            "name": page_el.get("name", ""),
            "level": int(page_el.get("pageLevel", "1")),
        }
        for page_el in root.iter(f"{{{_ONE_NS}}}Page")
    ]


def _parse_top_level_image(img_el: ET.Element) -> FloatingImage:
    """Parse a floating <one:Image> direct child of <one:Page> into a FloatingImage."""
    data_el = img_el.find(f"{{{_ONE_NS}}}Data")
    handle = data_el.text or "" if data_el is not None else ""

    width: float | None = None
    height: float | None = None
    size_el = img_el.find(f"{{{_ONE_NS}}}Size")
    if size_el is not None:
        w_str = size_el.get("width")
        h_str = size_el.get("height")
        if w_str is not None and h_str is not None:
            width = float(w_str)
            height = float(h_str)

    return FloatingImage(handle=handle, width=width, height=height, position=_parse_position(img_el))


def parse_page(xml: str) -> PageContent:
    """Parse a full <one:Page> XML (with mcpref image handles) into a PageContent."""
    root = ET.fromstring(xml)

    # Map quickStyleIndex to the page's <one:QuickStyleDef> elements, which
    # carry a ``name`` (e.g. "h1", "p") and the style's formatting.
    # _parse_paragraph_oe resolves headings through this map; pages without
    # defs use the numeric fallback.
    quick_style_map: dict[str, ET.Element] = {
        def_el.get("index"): def_el
        for def_el in root.findall(f"{{{_ONE_NS}}}QuickStyleDef")
        if def_el.get("index") is not None and def_el.get("name") is not None
    }

    # Extract title: the text of every <one:T>, OEs joined by spaces
    title = ""
    title_el = root.find(f"{{{_ONE_NS}}}Title")
    if title_el is not None:
        title = " ".join(
            "".join(_clean_text(t_el.text or "") for t_el in oe_el.findall(f"{{{_ONE_NS}}}T"))
            for oe_el in title_el.findall(f"{{{_ONE_NS}}}OE")
        )

    # Extract outlines (the schema allows several <one:OEChildren> per outline)
    outlines: list[Outline] = []
    for outline_el in root.findall(f"{{{_ONE_NS}}}Outline"):
        width, height = _parse_size(outline_el)
        items = _parse_outline_items(_child_oes(outline_el), quick_style_map or None)
        outlines.append(PageOutline(
            position=_parse_position(outline_el), width=width, height=height, items=items,
        ))

    # Extract top-level floating images (direct <one:Image> children of <one:Page>)
    images: list[FloatingImage] = []
    for img_el in root.findall(f"{{{_ONE_NS}}}Image"):
        images.append(_parse_top_level_image(img_el))

    unsupported = [
        _parse_unsupported_page_object(el) for el in root if el.tag in UNSUPPORTED_PAGE_OBJECT_TAGS
    ]

    return PageContent(title=title, outlines=outlines, images=images, unsupported=unsupported)
