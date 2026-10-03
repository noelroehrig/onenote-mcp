"""Write-direction builders: convert slim model instances to OneNote XML strings.

The public API is two functions:
  build_page_xml(title, outlines)  → full <one:Page> XML string (for replace_page)
  build_outline_xml(outline)       → single <one:Outline> XML string with xmlns (for append_page)

Read-direction parsers (xml → models/dicts):
  parse_notebook_skeleton(xml) → [{"id", "name", "sections": [...]}]
  parse_section_pages(xml)     → [{"id", "name"}]
  parse_page(xml)              → PageContent
"""
from __future__ import annotations

import html as _html_module
import re
import xml.etree.ElementTree as ET
from html.parser import HTMLParser

from onenote_mcp.models import (
    FONT_FAMILY_MAX_LENGTH, FONT_FAMILY_PATTERN, HEX_COLOR_PATTERN,
    ContentItem, FloatingImage, ImagePlaceholder, InlineImage, List, ListItem,
    Outline, PageContent, Paragraph, Position, TextRun,
)

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
ET.register_namespace("one", _ONE_NS)

# Heading visual defaults for the write direction.
# These CSS declarations are emitted as an inline span on the paragraph when
# p.style is h1–h6.  quickStyleIndex is NOT set; instead visual properties are
# carried in HTML so they render correctly regardless of the page's QuickStyleDef
# table.
_HEADING_STYLES: dict[str, list[str]] = {
    "h1": ["font-size:16pt", "font-weight:bold", "color:#1E4E79"],
    "h2": ["font-size:14pt", "font-weight:bold", "color:#1E4E79"],
    "h3": ["font-size:13pt", "font-weight:bold", "color:#2E74B5"],
    "h4": ["font-size:12pt", "font-weight:bold", "color:#2E74B5"],
    "h5": ["font-size:11pt", "font-weight:bold", "color:#595959"],
    "h6": ["font-size:11pt", "font-style:italic", "color:#595959"],
}


def _heading_default_styles(style: str) -> list[str]:
    """Return the CSS declaration list for heading style *style*, or [] for 'normal'."""
    return list(_HEADING_STYLES.get(style, []))


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
    # quickStyleIndex is intentionally NOT set — it indexes into the page's
    # QuickStyleDef table whose layout varies per page.  Heading visuals are
    # applied via inline CSS spans instead.

    t_el = ET.SubElement(oe, f"{{{_ONE_NS}}}T")

    heading_defaults = _heading_default_styles(p.style)

    if p.text is not None:
        # Heading defaults first, then explicit paragraph-level fields so the
        # latter win (last value wins within a single style="" attribute).
        merged = heading_defaults + _format_declarations(p)
        _set_t_html(t_el, _wrap_span(p.text, merged), cdata_map, cdata_counter)

    elif p.segments is not None:
        segments_html = "".join(_run_to_html(run) for run in p.segments)
        if heading_defaults:
            # Wrap the segment HTML in an outer span carrying the heading
            # visual defaults.  Per-segment formatting is unaffected.
            style_str = ";".join(heading_defaults)
            html = f'<span style="{style_str}">{segments_html}</span>'
        else:
            html = segments_html
        _set_t_html(t_el, html, cdata_map, cdata_counter)

    return oe


# ---- FloatingImage → top-level Image element --------------------------------

def _floating_image_to_el(img: FloatingImage) -> ET.Element:
    """Return a top-level <one:Image> element for a FloatingImage."""
    image_el = ET.Element(f"{{{_ONE_NS}}}Image")

    if img.position is not None:
        pos_el = ET.SubElement(image_el, f"{{{_ONE_NS}}}Position")
        pos_el.set("x", str(img.position.x))
        pos_el.set("y", str(img.position.y))
        pos_el.set("z", str(img.position.z))

    if img.width is not None and img.height is not None:
        size_el = ET.SubElement(image_el, f"{{{_ONE_NS}}}Size")
        size_el.set("width", str(img.width))
        size_el.set("height", str(img.height))
        size_el.set("isSetByUser", "true")

    data_el = ET.SubElement(image_el, f"{{{_ONE_NS}}}Data")
    data_el.text = img.handle

    return image_el


# ---- InlineImage → OE element -----------------------------------------------

def _inline_image_to_oe(img: InlineImage) -> ET.Element:
    """Return a <one:OE> element wrapping a <one:Image> for an inline image.

    Unlike floating images, inline images have no <one:Position> element —
    they are positioned by the OEChildren flow, not the page canvas.
    """
    oe = ET.Element(f"{{{_ONE_NS}}}OE")
    image_el = ET.SubElement(oe, f"{{{_ONE_NS}}}Image")

    if img.width is not None and img.height is not None:
        size_el = ET.SubElement(image_el, f"{{{_ONE_NS}}}Size")
        size_el.set("width", str(img.width))
        size_el.set("height", str(img.height))
        size_el.set("isSetByUser", "true")

    data_el = ET.SubElement(image_el, f"{{{_ONE_NS}}}Data")
    data_el.text = img.handle

    return oe


# ---- ImagePlaceholder → OE element ------------------------------------------

def _placeholder_to_oe(
    ph: ImagePlaceholder,
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> ET.Element:
    oe = ET.Element(f"{{{_ONE_NS}}}OE")
    oe.set("alignment", "left")

    t_el = ET.SubElement(oe, f"{{{_ONE_NS}}}T")
    html = (
        f'<span style="background:#ffeb3b;color:#222;font-weight:bold;'
        f'padding:4px 8px;border:2px dashed #b07500;">'
        f'[INSERT IMAGE: {_escape_text(ph.description)}]</span>'
    )
    _set_t_html(t_el, html, cdata_map, cdata_counter)

    return oe


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

        if item.children:
            oe_children = ET.SubElement(oe, f"{{{_ONE_NS}}}OEChildren")
            for child_oe in _list_items_to_oes(item.children, list_style, cdata_map, cdata_counter):
                oe_children.append(child_oe)

        result.append(oe)

    return result


# ---- List → list of OE elements ---------------------------------------------

def _list_to_oes(
    lst: List,
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> list[ET.Element]:
    return _list_items_to_oes(lst.items, lst.style, cdata_map, cdata_counter)


# ---- Items → OEChildren element ---------------------------------------------

def _items_to_oechildren(
    items: list[ContentItem],
    cdata_map: dict[int, str],
    cdata_counter: list[int],
) -> ET.Element:
    oe_children = ET.Element(f"{{{_ONE_NS}}}OEChildren")

    for item in items:
        if isinstance(item, Paragraph):
            oe_children.append(_paragraph_to_oe(item, cdata_map, cdata_counter))
        elif isinstance(item, ImagePlaceholder):
            oe_children.append(_placeholder_to_oe(item, cdata_map, cdata_counter))
        elif isinstance(item, InlineImage):
            oe_children.append(_inline_image_to_oe(item))
        elif isinstance(item, List):
            for oe in _list_to_oes(item, cdata_map, cdata_counter):
                oe_children.append(oe)

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

    xml_str = ET.tostring(root, encoding="unicode", xml_declaration=False)
    return _apply_cdata(xml_str, cdata_map)


def build_outline_xml(outline: Outline) -> str:
    """Return a single <one:Outline> XML string with xmlns for append_page."""
    cdata_map: dict[int, str] = {}
    cdata_counter: list[int] = [0]

    el = _outline_to_el(outline, cdata_map, cdata_counter)

    xml_str = ET.tostring(el, encoding="unicode", xml_declaration=False)
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


def _fmt_from_style(style_str: str) -> dict:
    """Translate a CSS ``style`` declaration string into slim formatting fields.

    Returns only the keys that are explicitly set, ready to splat into a
    TextRun/Paragraph (e.g. {"bold": True, "font_size": 16.0, "color": "#1E4E79"}).
    Colour and font values the slim model cannot express are left out, like
    any other unsupported CSS property.
    """
    fmt: dict = {}
    if not style_str:
        return fmt
    decoration = ""
    for decl in style_str.split(";"):
        prop, sep, val = decl.partition(":")
        if not sep:
            continue
        prop = prop.strip().lower()
        val = val.strip()
        if not val:
            continue
        if prop == "font-weight" and "bold" in val.lower():
            fmt["bold"] = True
        elif prop == "font-style" and "italic" in val.lower():
            fmt["italic"] = True
        elif prop == "text-decoration":
            decoration += " " + val.lower()
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
    if "underline" in decoration:
        fmt["underline"] = True
    if "line-through" in decoration:
        fmt["strikethrough"] = True
    return fmt


class _RunCollector(HTMLParser):
    """Collect (text, formatting) runs from inline span markup, honouring nesting."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.runs: list[tuple[str, dict]] = []
        self._stack: list[dict] = [{}]

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


def _extract_runs(raw: str) -> list[tuple[str, dict]]:
    """Parse a raw <one:T> string into a list of (text, formatting) runs.

    Plain text (no markup) yields a single unformatted run.  Adjacent runs with
    identical formatting are merged so uniformly-styled text collapses to one run.
    """
    if not raw:
        return []
    if not _TAG_RE.search(raw):
        return [(_html_module.unescape(raw), {})]
    collector = _RunCollector()
    try:
        collector.feed(raw)
        collector.close()
    except Exception:
        # Malformed markup — degrade gracefully to stripped plain text.
        return [(_clean_text(raw), {})]
    merged: list[tuple[str, dict]] = []
    for text, fmt in collector.runs:
        if merged and merged[-1][1] == fmt:
            merged[-1] = (merged[-1][0] + text, fmt)
        else:
            merged.append((text, fmt))
    return merged


def _parse_list_item(oe_el: ET.Element) -> ListItem:
    """Parse a single list OE element into a ListItem, recursing into children.

    Inline formatting is preserved: a uniformly-formatted item keeps its text
    plus a single styled segment; a mixed item becomes multiple segments.  (The
    ListItem model has no whole-item format fields, so any formatting is carried
    via ``segments`` rather than ``text``.)
    """
    t_el = oe_el.find(f"{{{_ONE_NS}}}T")
    raw = t_el.text or "" if t_el is not None else ""
    runs = _extract_runs(raw)

    children: list[ListItem] = []
    oe_children_el = oe_el.find(f"{{{_ONE_NS}}}OEChildren")
    if oe_children_el is not None:
        for child_oe in oe_children_el.findall(f"{{{_ONE_NS}}}OE"):
            if child_oe.find(f"{{{_ONE_NS}}}List") is not None:
                children.append(_parse_list_item(child_oe))

    if not runs:
        return ListItem(text="", children=children)
    if len(runs) == 1 and not runs[0][1]:
        return ListItem(text=runs[0][0], children=children)
    segments = [TextRun(text=text, **fmt) for text, fmt in runs]
    return ListItem(segments=segments, children=children)


_HEADING_NAMES = frozenset({"h1", "h2", "h3", "h4", "h5", "h6"})


def _resolve_heading_style(
    el: ET.Element,
    quick_style_map: dict[str, str] | None,
) -> str:
    """Resolve an OE's heading style ('normal' or 'h1'–'h6').

    A page's <one:QuickStyleDef> table is AUTHORITATIVE when present: a
    paragraph is a heading only if its quickStyleIndex maps to a heading-named
    style.  Body text can sit at any index (e.g. index 1, name "p"), so a
    non-heading name means 'normal' — never a numeric guess.  Only pages
    without any QSD table use the numeric fallback mapping.
    """
    quick_style = el.get("quickStyleIndex")
    if quick_style is None:
        return "normal"
    if quick_style_map:
        mapped_name = quick_style_map.get(quick_style)
        return mapped_name if mapped_name in _HEADING_NAMES else "normal"
    return _NUMERIC_QUICK_STYLE_TO_HEADING.get(quick_style, "normal")


def _parse_paragraph_oe(
    el: ET.Element,
    quick_style_map: dict[str, str] | None = None,
) -> Paragraph:
    """Parse an OE element containing text into a Paragraph.

    Heading style is resolved via _resolve_heading_style (QSD table
    authoritative, numeric fallback only when no QSD table exists).
    Inline <span style="…"> formatting is preserved: uniformly-formatted text
    collapses to ``text`` plus whole-paragraph format fields, while mixed
    formatting becomes ``segments``.
    """
    style = _resolve_heading_style(el, quick_style_map)

    t_el = el.find(f"{{{_ONE_NS}}}T")
    raw = t_el.text or "" if t_el is not None else ""
    runs = _extract_runs(raw)

    if not runs:
        return Paragraph(type="paragraph", text="", style=style)
    if len(runs) == 1:
        text, fmt = runs[0]
        return Paragraph(type="paragraph", text=text, style=style, **fmt)
    segments = [TextRun(text=text, **fmt) for text, fmt in runs]
    return Paragraph(type="paragraph", segments=segments, style=style)


def _parse_oe(
    el: ET.Element,
    quick_style_map: dict[str, str] | None = None,
) -> "ContentItem | None":
    """Classify an OE element and return the appropriate ContentItem, or None for list OEs.

    *quick_style_map* is threaded through for paragraph handling; image branches
    ignore the map but accept it to keep the call signature compatible.
    """
    if el.find(f"{{{_ONE_NS}}}List") is not None:
        return None

    img_el = el.find(f"{{{_ONE_NS}}}Image")
    if img_el is not None:
        data_el = img_el.find(f"{{{_ONE_NS}}}Data")
        if data_el is None or not (data_el.text or "").strip():
            return None  # missing/empty handle — treat as absent, don't raise

        handle = data_el.text.strip()

        width: float | None = None
        height: float | None = None
        size_el = img_el.find(f"{{{_ONE_NS}}}Size")
        if size_el is not None:
            w_str = size_el.get("width")
            h_str = size_el.get("height")
            if w_str is not None:
                width = float(w_str)
            if h_str is not None:
                height = float(h_str)

        return InlineImage(type="inline_image", handle=handle, width=width, height=height)

    return _parse_paragraph_oe(el, quick_style_map)


def _parse_outline_items(
    oe_elements: list[ET.Element],
    quick_style_map: dict[str, str] | None = None,
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

    for oe in oe_elements:
        list_el = oe.find(f"{{{_ONE_NS}}}List")
        if list_el is not None:
            # Determine list style
            if list_el.find(f"{{{_ONE_NS}}}Bullet") is not None:
                style = "bullet"
            else:
                style = "numbered"

            if style != current_list_style:
                _flush_list()
                current_list_style = style

            current_list_items.append(_parse_list_item(oe))
        else:
            _flush_list()
            item = _parse_oe(oe, quick_style_map)
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
    """Parse a section-scoped hierarchy XML into its pages as [{"id", "name"}].

    The input root is a <one:Section> (section-anchored GetHierarchy); pages are
    its direct <one:Page> children.  Returns an empty list if the section has no
    pages.
    """
    root = ET.fromstring(xml)
    return [
        {"id": page_el.get("ID", ""), "name": page_el.get("name", "")}
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

    pos: Position | None = None
    pos_el = img_el.find(f"{{{_ONE_NS}}}Position")
    if pos_el is not None:
        pos = Position(
            x=float(pos_el.get("x", 0)),
            y=float(pos_el.get("y", 0)),
            z=int(pos_el.get("z", 0)),
        )

    return FloatingImage(handle=handle, width=width, height=height, position=pos)


def parse_page(xml: str) -> PageContent:
    """Parse a full <one:Page> XML (with mcpref image handles) into a PageContent."""
    root = ET.fromstring(xml)

    # Build quick-style map from <one:QuickStyleDef> direct children of the page.
    # Each def has an ``index`` attribute and a ``name`` attribute (e.g. "h1", "p").
    # _parse_paragraph_oe uses this map to resolve quickStyleIndex to a heading
    # level; pages without defs use the numeric fallback.
    quick_style_map: dict[str, str] = {
        def_el.get("index"): def_el.get("name")
        for def_el in root.findall(f"{{{_ONE_NS}}}QuickStyleDef")
        if def_el.get("index") is not None and def_el.get("name") is not None
    }

    # Extract title
    title = ""
    title_el = root.find(f"{{{_ONE_NS}}}Title")
    if title_el is not None:
        oe_el = title_el.find(f"{{{_ONE_NS}}}OE")
        if oe_el is not None:
            t_el = oe_el.find(f"{{{_ONE_NS}}}T")
            if t_el is not None:
                raw = t_el.text or ""
                title = _clean_text(raw)

    # Extract outlines
    outlines: list[Outline] = []
    for outline_el in root.findall(f"{{{_ONE_NS}}}Outline"):
        # Position
        pos: Position | None = None
        pos_el = outline_el.find(f"{{{_ONE_NS}}}Position")
        if pos_el is not None:
            pos = Position(
                x=float(pos_el.get("x", 0)),
                y=float(pos_el.get("y", 0)),
                z=int(pos_el.get("z", 0)),
            )

        # Width
        width: float | None = None
        size_el = outline_el.find(f"{{{_ONE_NS}}}Size")
        if size_el is not None:
            w_str = size_el.get("width")
            if w_str is not None:
                width = float(w_str)

        # Items
        items: list[ContentItem] = []
        oe_children_el = outline_el.find(f"{{{_ONE_NS}}}OEChildren")
        if oe_children_el is not None:
            oe_elements = oe_children_el.findall(f"{{{_ONE_NS}}}OE")
            items = _parse_outline_items(oe_elements, quick_style_map or None)

        outlines.append(Outline(position=pos, width=width, items=items))

    # Extract top-level floating images (direct <one:Image> children of <one:Page>)
    images: list[FloatingImage] = []
    for img_el in root.findall(f"{{{_ONE_NS}}}Image"):
        images.append(_parse_top_level_image(img_el))

    return PageContent(title=title, outlines=outlines, images=images)
