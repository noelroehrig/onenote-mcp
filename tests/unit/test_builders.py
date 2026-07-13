"""Unit tests for onenote_mcp.builders — no COM dependency."""
import xml.etree.ElementTree as ET
import pytest
from onenote_mcp.builders import (
    _run_to_html, _paragraph_html, _heading_default_styles, _cdata_placeholder, _apply_cdata,
    build_page_xml, build_outline_xml,
    parse_page,
    _clean_text, _floating_image_to_el,
)
from onenote_mcp.models import (
    TextRun, Paragraph, ImagePlaceholder, InlineImage, ListItem, List,
    FloatingImage, Outline, Position, PageContent,
)

_NS = {"one": "http://schemas.microsoft.com/office/onenote/2013/onenote"}


# ---------------------------------------------------------------------------
# CDATA helpers
# ---------------------------------------------------------------------------

def test_cdata_placeholder_format():
    assert _cdata_placeholder(0) == "__CDATA_000000__"
    assert _cdata_placeholder(42) == "__CDATA_000042__"


def test_apply_cdata_replaces_placeholder():
    cdata_map = {0: "<b>hello</b>"}
    result = _apply_cdata("text __CDATA_000000__ end", cdata_map)
    assert "<![CDATA[<b>hello</b>]]>" in result


def test_apply_cdata_multiple():
    cdata_map = {0: "first", 1: "second"}
    source = "__CDATA_000000__ and __CDATA_000001__"
    result = _apply_cdata(source, cdata_map)
    assert "<![CDATA[first]]>" in result
    assert "<![CDATA[second]]>" in result


# ---------------------------------------------------------------------------
# _run_to_html
# ---------------------------------------------------------------------------

def test_run_to_html_plain():
    result = _run_to_html(TextRun(text="hello"))
    assert result == "hello"


def test_run_to_html_bold():
    result = _run_to_html(TextRun(text="hello", bold=True))
    assert result == '<span style="font-weight:bold">hello</span>'


def test_run_to_html_italic():
    result = _run_to_html(TextRun(text="hello", italic=True))
    assert "font-style:italic" in result


def test_run_to_html_underline():
    result = _run_to_html(TextRun(text="hello", underline=True))
    assert "text-decoration:underline" in result


def test_run_to_html_strikethrough():
    result = _run_to_html(TextRun(text="hello", strikethrough=True))
    assert "text-decoration:line-through" in result


def test_run_to_html_underline_and_strikethrough():
    result = _run_to_html(TextRun(text="hello", underline=True, strikethrough=True))
    assert "text-decoration:underline line-through" in result


def test_run_to_html_color():
    result = _run_to_html(TextRun(text="hello", color="#ff0000"))
    assert "color:#ff0000" in result


def test_run_to_html_highlight():
    result = _run_to_html(TextRun(text="hello", highlight="#ffeb3b"))
    assert "background:#ffeb3b" in result


def test_run_to_html_font_size():
    result = _run_to_html(TextRun(text="hello", font_size=12.0))
    assert "font-size:12.0pt" in result


def test_run_to_html_font_family():
    result = _run_to_html(TextRun(text="hello", font_family="Courier New"))
    assert "font-family:Courier New" in result


def test_run_to_html_multiple_styles():
    result = _run_to_html(TextRun(text="hello", bold=True, color="#ff0000"))
    assert "font-weight:bold" in result
    assert "color:#ff0000" in result
    assert 'style="' in result


# ---------------------------------------------------------------------------
# _paragraph_html
# ---------------------------------------------------------------------------

def test_heading_default_styles_h1():
    styles = _heading_default_styles("h1")
    assert "font-size:16pt" in styles
    assert "font-weight:bold" in styles
    assert "color:#1E4E79" in styles


def test_heading_default_styles_h6():
    styles = _heading_default_styles("h6")
    assert "font-size:11pt" in styles
    assert "font-style:italic" in styles
    assert "color:#595959" in styles
    assert not any("font-weight" in s for s in styles), "h6 must not have font-weight"


def test_heading_default_styles_normal():
    assert _heading_default_styles("normal") == []


def test_paragraph_html_plain():
    p = Paragraph(type="paragraph", text="plain text")
    result = _paragraph_html("plain text", p)
    assert result == "plain text"


def test_paragraph_html_bold():
    p = Paragraph(type="paragraph", text="bold text", bold=True)
    result = _paragraph_html("bold text", p)
    assert "font-weight:bold" in result
    assert "<span" in result


# ---------------------------------------------------------------------------
# _clean_text
# ---------------------------------------------------------------------------

def test_clean_text_strips_tags():
    result = _clean_text("<span>hello</span>")
    assert result == "hello"


def test_clean_text_unescapes_entities():
    # html.unescape("&amp;lt;") -> "&lt;"
    result = _clean_text("&amp;lt;")
    assert result == "&lt;"


def test_clean_text_plain():
    result = _clean_text("plain text")
    assert result == "plain text"


# ---------------------------------------------------------------------------
# build_page_xml  (write direction)
# ---------------------------------------------------------------------------

def _parse_xml(xml_str: str) -> ET.Element:
    """Parse an XML string that may contain CDATA sections."""
    # ET.fromstring handles CDATA — the parser merges it into element text
    return ET.fromstring(xml_str)


def test_build_page_xml_has_title():
    xml = build_page_xml("My Title", [])
    root = _parse_xml(xml)
    t_el = root.find("one:Title/one:OE/one:T", _NS)
    assert t_el is not None
    assert t_el.text == "My Title"


def test_build_page_xml_plain_paragraph():
    outline = Outline(items=[Paragraph(type="paragraph", text="Hello world")])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)
    t_el = root.find("one:Outline/one:OEChildren/one:OE/one:T", _NS)
    assert t_el is not None
    assert t_el.text == "Hello world"


def test_build_page_xml_bold_paragraph():
    outline = Outline(items=[Paragraph(type="paragraph", text="Bold", bold=True)])
    xml = build_page_xml("Title", [outline])
    # The CDATA span should be in the raw XML string
    assert "<![CDATA[" in xml
    assert "font-weight:bold" in xml


def test_build_page_xml_h1_heading():
    # h1 headings must use inline CSS spans, not quickStyleIndex.
    # quickStyleIndex is page-relative and cannot be reliably hard-coded.
    outline = Outline(items=[Paragraph(type="paragraph", text="Heading", style="h1")])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)
    oe_el = root.find("one:Outline/one:OEChildren/one:OE", _NS)
    assert oe_el is not None
    assert oe_el.get("quickStyleIndex") is None, "h1 must NOT emit quickStyleIndex"
    # CSS properties must appear in the raw XML (inside CDATA)
    assert "font-size:16pt" in xml
    assert "font-weight:bold" in xml
    assert "color:#1E4E79" in xml


def test_build_page_xml_h6_heading():
    # h6: italic, no bold, color #595959, font-size 11pt
    outline = Outline(items=[Paragraph(type="paragraph", text="H6", style="h6")])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)
    oe_el = root.find("one:Outline/one:OEChildren/one:OE", _NS)
    assert oe_el is not None
    assert oe_el.get("quickStyleIndex") is None, "h6 must NOT emit quickStyleIndex"
    assert "font-size:11pt" in xml
    assert "font-style:italic" in xml
    assert "color:#595959" in xml
    assert "font-weight:bold" not in xml


def test_build_page_xml_normal_no_quick_style():
    outline = Outline(items=[Paragraph(type="paragraph", text="Normal", style="normal")])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)
    oe_el = root.find("one:Outline/one:OEChildren/one:OE", _NS)
    assert oe_el is not None
    assert oe_el.get("quickStyleIndex") is None


def test_build_page_xml_bullet_list():
    lst = List(type="list", style="bullet", items=[ListItem(text="item")])
    outline = Outline(items=[lst])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)
    bullet_el = root.find("one:Outline/one:OEChildren/one:OE/one:List/one:Bullet", _NS)
    assert bullet_el is not None


def test_build_page_xml_numbered_list():
    lst = List(type="list", style="numbered", items=[ListItem(text="item")])
    outline = Outline(items=[lst])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)
    number_el = root.find("one:Outline/one:OEChildren/one:OE/one:List/one:Number", _NS)
    assert number_el is not None


def test_build_page_xml_list_drops_empty_items():
    # An empty list item (no text/segments/children) is tolerated by the model
    # and silently dropped by the builder so the rest of the page still writes.
    empty = ListItem()
    lst = List(type="list", style="numbered", items=[ListItem(text="real"), empty])
    outline = Outline(items=[lst])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)
    oes = root.findall("one:Outline/one:OEChildren/one:OE", _NS)
    texts = [oe.find("one:T", _NS).text for oe in oes if oe.find("one:T", _NS) is not None]
    assert texts == ["real"], f"empty item should be dropped, got {texts}"


def test_build_page_xml_list_all_empty_items():
    # A list whose only item is empty produces no OE children (no crash).
    lst = List(type="list", style="bullet", items=[ListItem()])
    outline = Outline(items=[lst])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)
    assert root.find("one:Outline/one:OEChildren/one:OE/one:List", _NS) is None


def test_list_item_empty_is_allowed_but_both_set_rejected():
    # Empty is fine now; supplying both text and segments is still an error.
    ListItem()  # no raise
    with pytest.raises(ValueError, match="not both"):
        ListItem(text="x", segments=[TextRun(text="y")])


def test_build_page_xml_image_placeholder_cdata():
    ph = ImagePlaceholder(type="image_placeholder", description="Screenshot of dashboard")
    outline = Outline(items=[ph])
    xml = build_page_xml("Title", [outline])
    assert "[INSERT IMAGE:" in xml


def test_build_page_xml_outline_position():
    pos = Position(x=36.0, y=86.4, z=0)
    outline = Outline(position=pos, items=[Paragraph(type="paragraph", text="x")])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)
    pos_el = root.find("one:Outline/one:Position", _NS)
    assert pos_el is not None
    assert pos_el.get("x") == "36.0"
    assert pos_el.get("y") == "86.4"
    assert pos_el.get("z") == "0"


def test_build_page_xml_outline_width():
    outline = Outline(width=300.0, items=[Paragraph(type="paragraph", text="x")])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)
    size_el = root.find("one:Outline/one:Size", _NS)
    assert size_el is not None
    assert size_el.get("width") == "300.0"
    # OneNote rejects <one:Size> with width-only (hresult 0x80042001).
    # When the caller supplies width but no height we must pair it with a
    # placeholder height ("0.0" is OneNote's auto-height sentinel).
    assert size_el.get("height") is not None, "Size must always carry both dimensions"
    assert size_el.get("height") == "0.0"


def test_build_page_xml_outline_size_never_width_only():
    """Regression guard: every <one:Size> in the output must carry both dimensions.

    OneNote rejects width-only or height-only Size elements with hresult
    0x80042001 (hrObjectDoesNotExist).  This test sweeps a representative set
    of outline shapes and asserts that any emitted Size has both attributes.
    """
    outlines = [
        # width only
        Outline(width=700.0, items=[Paragraph(type="paragraph", text="a")]),
        # no width
        Outline(items=[Paragraph(type="paragraph", text="b")]),
        # width + placeholder with explicit dims
        Outline(width=500.0, items=[ImagePlaceholder(type="image_placeholder", description="x", width=400.0, height=200.0)]),
        # no width, inline image carries dims
        Outline(items=[InlineImage(type="inline_image", handle="mcpref:deadbeef", width=300.0, height=150.0)]),
    ]
    xml = build_page_xml("Title", outlines)
    root = _parse_xml(xml)
    for size_el in root.iter(f"{{{_NS['one']}}}Size"):
        w = size_el.get("width")
        h = size_el.get("height")
        assert (w is None) == (h is None), (
            f"<one:Size> must carry both width and height or neither — got width={w!r} height={h!r}"
        )


def test_build_page_xml_multiple_outlines():
    o1 = Outline(items=[Paragraph(type="paragraph", text="first")])
    o2 = Outline(items=[Paragraph(type="paragraph", text="second")])
    xml = build_page_xml("Title", [o1, o2])
    root = _parse_xml(xml)
    outlines = root.findall("one:Outline", _NS)
    assert len(outlines) == 2


def test_build_page_xml_segments():
    segments = [TextRun(text="bold part", bold=True), TextRun(text=" plain")]
    p = Paragraph(type="paragraph", segments=segments)
    outline = Outline(items=[p])
    xml = build_page_xml("Title", [outline])
    # Formatted segment produces CDATA with a span
    assert "<![CDATA[" in xml
    assert "font-weight:bold" in xml


# ---------------------------------------------------------------------------
# build_outline_xml
# ---------------------------------------------------------------------------

def test_build_outline_xml_is_outline_element():
    outline = Outline(items=[Paragraph(type="paragraph", text="content")])
    xml = build_outline_xml(outline)
    root = _parse_xml(xml)
    _ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
    assert root.tag == f"{{{_ONE_NS}}}Outline"


def test_build_outline_xml_paragraph():
    outline = Outline(items=[Paragraph(type="paragraph", text="content")])
    xml = build_outline_xml(outline)
    root = _parse_xml(xml)
    t_el = root.find("one:OEChildren/one:OE/one:T", _NS)
    assert t_el is not None
    assert t_el.text == "content"


# ---------------------------------------------------------------------------
# Hierarchy parsers  (read direction)
# ---------------------------------------------------------------------------

_ONE_NS_URI = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_XMLNS = f'xmlns:one="{_ONE_NS_URI}"'


def test_parse_notebook_skeleton_includes_sections_not_pages():
    xml = (
        '<one:Notebooks xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">'
        '<one:Notebook ID="nb1" name="Work">'
        '<one:Section ID="sec1" name="S1"><one:Page ID="p1" name="P1"/></one:Section>'
        '<one:Section ID="sec2" name="S2"/>'
        '</one:Notebook>'
        '<one:Notebook ID="nb2" name="Personal"/>'
        '</one:Notebooks>'
    )
    from onenote_mcp.builders import parse_notebook_skeleton
    result = parse_notebook_skeleton(xml)
    assert result == [
        {"id": "nb1", "name": "Work", "sections": [
            {"id": "sec1", "name": "S1"},
            {"id": "sec2", "name": "S2"},
        ]},
        {"id": "nb2", "name": "Personal", "sections": []},
    ]


def test_parse_notebook_skeleton_empty():
    xml = '<one:Notebooks xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"/>'
    from onenote_mcp.builders import parse_notebook_skeleton
    assert parse_notebook_skeleton(xml) == []


def test_parse_section_pages_returns_id_and_name():
    xml = (
        '<one:Section xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"'
        ' ID="sec1" name="Vorlagen">'
        '<one:Page ID="p1" name="Template A"/>'
        '<one:Page ID="p2" name="Template B"/>'
        '</one:Section>'
    )
    from onenote_mcp.builders import parse_section_pages
    assert parse_section_pages(xml) == [
        {"id": "p1", "name": "Template A"},
        {"id": "p2", "name": "Template B"},
    ]


def test_parse_section_pages_empty_section():
    xml = (
        '<one:Section xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"'
        ' ID="sec1" name="Empty"/>'
    )
    from onenote_mcp.builders import parse_section_pages
    assert parse_section_pages(xml) == []


# ---------------------------------------------------------------------------
# parse_page  (read direction)
# ---------------------------------------------------------------------------

def _page_xml(*body_parts: str) -> str:
    """Wrap body_parts inside a <one:Page> with namespace declaration."""
    body = "".join(body_parts)
    return f'<one:Page {_XMLNS}>{body}</one:Page>'


def test_parse_page_title():
    xml = _page_xml(
        '<one:Title><one:OE><one:T>My Title</one:T></one:OE></one:Title>'
    )
    page = parse_page(xml)
    assert page.title == "My Title"


def test_parse_page_empty_outlines():
    xml = _page_xml()
    page = parse_page(xml)
    assert page.outlines == []


def test_parse_page_paragraph():
    xml = _page_xml(
        '<one:Outline>'
        '  <one:OEChildren>'
        '    <one:OE><one:T>Hello paragraph</one:T></one:OE>'
        '  </one:OEChildren>'
        '</one:Outline>'
    )
    page = parse_page(xml)
    assert len(page.outlines) == 1
    items = page.outlines[0].items
    assert len(items) == 1
    assert isinstance(items[0], Paragraph)
    assert items[0].text == "Hello paragraph"


def test_parse_page_h1_heading():
    xml = _page_xml(
        '<one:Outline>'
        '  <one:OEChildren>'
        '    <one:OE quickStyleIndex="1"><one:T>Big Heading</one:T></one:OE>'
        '  </one:OEChildren>'
        '</one:Outline>'
    )
    page = parse_page(xml)
    p = page.outlines[0].items[0]
    assert isinstance(p, Paragraph)
    assert p.style == "h1"


def test_parse_page_inline_image():
    """Images inside outline OE elements are returned as InlineImage items."""
    xml = _page_xml(
        '<one:Outline>'
        '  <one:OEChildren>'
        '    <one:OE>'
        '      <one:Image>'
        '        <one:Size width="100.0" height="200.0"/>'
        '        <one:Data>mcpref:abc123def456</one:Data>'
        '      </one:Image>'
        '    </one:OE>'
        '  </one:OEChildren>'
        '</one:Outline>'
    )
    page = parse_page(xml)
    items = page.outlines[0].items
    assert len(items) == 1, f"Expected 1 InlineImage item, got {len(items)}"
    img = items[0]
    assert isinstance(img, InlineImage)
    assert img.handle == "mcpref:abc123def456"
    assert img.width == 100.0
    assert img.height == 200.0


def test_parse_page_inline_image_without_data_skipped():
    """OE with <one:Image> but no <one:Data> is silently dropped."""
    xml = _page_xml(
        '<one:Outline>'
        '  <one:OEChildren>'
        '    <one:OE><one:Image/></one:OE>'
        '  </one:OEChildren>'
        '</one:Outline>'
    )
    page = parse_page(xml)
    assert page.outlines[0].items == [], (
        "OE with Image but no Data should be silently skipped"
    )


def test_parse_page_bullet_list():
    xml = _page_xml(
        '<one:Outline>'
        '  <one:OEChildren>'
        '    <one:OE><one:List><one:Bullet/></one:List><one:T>Item A</one:T></one:OE>'
        '    <one:OE><one:List><one:Bullet/></one:List><one:T>Item B</one:T></one:OE>'
        '  </one:OEChildren>'
        '</one:Outline>'
    )
    page = parse_page(xml)
    items = page.outlines[0].items
    assert len(items) == 1
    lst = items[0]
    assert isinstance(lst, List)
    assert lst.style == "bullet"
    assert len(lst.items) == 2
    assert lst.items[0].text == "Item A"
    assert lst.items[1].text == "Item B"


def test_parse_page_numbered_list():
    xml = _page_xml(
        '<one:Outline>'
        '  <one:OEChildren>'
        '    <one:OE><one:List><one:Number/></one:List><one:T>Step 1</one:T></one:OE>'
        '  </one:OEChildren>'
        '</one:Outline>'
    )
    page = parse_page(xml)
    lst = page.outlines[0].items[0]
    assert isinstance(lst, List)
    assert lst.style == "numbered"


def test_parse_page_nested_list():
    xml = _page_xml(
        '<one:Outline>'
        '  <one:OEChildren>'
        '    <one:OE>'
        '      <one:List><one:Bullet/></one:List>'
        '      <one:T>Parent</one:T>'
        '      <one:OEChildren>'
        '        <one:OE>'
        '          <one:List><one:Bullet/></one:List>'
        '          <one:T>Child</one:T>'
        '        </one:OE>'
        '      </one:OEChildren>'
        '    </one:OE>'
        '  </one:OEChildren>'
        '</one:Outline>'
    )
    page = parse_page(xml)
    lst = page.outlines[0].items[0]
    assert isinstance(lst, List)
    parent_item = lst.items[0]
    assert parent_item.text == "Parent"
    assert len(parent_item.children) == 1
    assert parent_item.children[0].text == "Child"


def test_parse_page_position():
    xml = _page_xml(
        '<one:Outline>'
        '  <one:Position x="36.0" y="86.4" z="2"/>'
        '  <one:OEChildren>'
        '    <one:OE><one:T>content</one:T></one:OE>'
        '  </one:OEChildren>'
        '</one:Outline>'
    )
    page = parse_page(xml)
    pos = page.outlines[0].position
    assert pos is not None
    assert pos.x == 36.0
    assert pos.y == 86.4
    assert pos.z == 2


def test_parse_page_width():
    xml = _page_xml(
        '<one:Outline>'
        '  <one:Size width="250.0" isSetByUser="true"/>'
        '  <one:OEChildren>'
        '    <one:OE><one:T>content</one:T></one:OE>'
        '  </one:OEChildren>'
        '</one:Outline>'
    )
    page = parse_page(xml)
    assert page.outlines[0].width == 250.0


def test_parse_page_h1_resolved_via_qsd_table():
    """QSD-based resolution: quickStyleIndex="2" + QSD index="2" name="h1" → style=="h1".

    This proves that the new map-based resolution works for OneNote-native pages
    where the heading index does NOT match the heading number (e.g. index 2 is h1).
    """
    xml = _page_xml(
        '<one:QuickStyleDef index="0" name="PageTitle"/>',
        '<one:QuickStyleDef index="1" name="p"/>',
        '<one:QuickStyleDef index="2" name="h1"/>',
        '<one:Outline>'
        '  <one:OEChildren>'
        '    <one:OE quickStyleIndex="2"><one:T>Big Heading</one:T></one:OE>'
        '  </one:OEChildren>'
        '</one:Outline>'
    )
    page = parse_page(xml)
    p = page.outlines[0].items[0]
    assert isinstance(p, Paragraph)
    assert p.style == "h1", (
        f"Expected style='h1' via QSD table (index 2 → 'h1'), got {p.style!r}"
    )


def test_parse_page_quickstyleindex_unknown_falls_back():
    """Numeric fallback: no QSD elements, quickStyleIndex="3" → style=="h3".

    Pages without a QuickStyleDef table resolve headings via the numeric
    index mapping.
    """
    xml = _page_xml(
        '<one:Outline>'
        '  <one:OEChildren>'
        '    <one:OE quickStyleIndex="3"><one:T>Heading Three</one:T></one:OE>'
        '  </one:OEChildren>'
        '</one:Outline>'
    )
    page = parse_page(xml)
    p = page.outlines[0].items[0]
    assert isinstance(p, Paragraph)
    assert p.style == "h3", (
        f"Expected fallback style='h3' for quickStyleIndex='3', got {p.style!r}"
    )


def test_parse_page_multiple_outlines():
    xml = _page_xml(
        '<one:Outline>'
        '  <one:OEChildren><one:OE><one:T>first</one:T></one:OE></one:OEChildren>'
        '</one:Outline>'
        '<one:Outline>'
        '  <one:OEChildren><one:OE><one:T>second</one:T></one:OE></one:OEChildren>'
        '</one:Outline>'
    )
    page = parse_page(xml)
    assert len(page.outlines) == 2


# ---------------------------------------------------------------------------
# The QSD table is authoritative for heading resolution
# ---------------------------------------------------------------------------

def test_parse_page_body_at_index_1_is_normal_not_h1():
    """Body text at quickStyleIndex='1' (QSD name 'p') must read as 'normal',
    NOT 'h1'.

    A non-heading QSD name at any index must never fall back to the numeric
    mapping — that turns ordinary body paragraphs into headings.
    """
    xml = _page_xml(
        '<one:QuickStyleDef index="0" name="PageTitle"/>',
        '<one:QuickStyleDef index="1" name="p"/>',
        '<one:Outline><one:OEChildren>'
        '  <one:OE quickStyleIndex="1"><one:T>Ordinary body text</one:T></one:OE>'
        '</one:OEChildren></one:Outline>'
    )
    page = parse_page(xml)
    p = page.outlines[0].items[0]
    assert isinstance(p, Paragraph)
    assert p.style == "normal", f"body text must be 'normal', got {p.style!r}"


def test_parse_page_index_not_in_qsd_table_is_normal():
    """When a QSD table exists but does not define this index, do NOT guess a
    heading from the number — treat it as normal."""
    xml = _page_xml(
        '<one:QuickStyleDef index="0" name="p"/>',
        '<one:Outline><one:OEChildren>'
        '  <one:OE quickStyleIndex="2"><one:T>Body, index absent from QSD</one:T></one:OE>'
        '</one:OEChildren></one:Outline>'
    )
    page = parse_page(xml)
    assert page.outlines[0].items[0].style == "normal"


# ---------------------------------------------------------------------------
# Inline formatting preservation on read (#7)
# ---------------------------------------------------------------------------

def test_parse_paragraph_uniform_bold_preserved_as_text_plus_field():
    xml = _page_xml(
        '<one:Outline><one:OEChildren>'
        "  <one:OE><one:T><![CDATA[<span style='font-weight:bold'>Bold body</span>]]></one:T></one:OE>"
        '</one:OEChildren></one:Outline>'
    )
    p = parse_page(xml).outlines[0].items[0]
    assert isinstance(p, Paragraph)
    assert p.text == "Bold body"
    assert p.bold is True
    assert p.segments is None


def test_parse_paragraph_font_size_and_color_preserved():
    xml = _page_xml(
        '<one:Outline><one:OEChildren>'
        "  <one:OE><one:T><![CDATA[<span style=\"font-size:16pt;color:#1E4E79\">Big blue</span>]]></one:T></one:OE>"
        '</one:OEChildren></one:Outline>'
    )
    p = parse_page(xml).outlines[0].items[0]
    assert p.text == "Big blue"
    assert p.font_size == 16.0
    assert p.color == "#1E4E79"


def test_parse_paragraph_mixed_formatting_becomes_segments():
    xml = _page_xml(
        '<one:Outline><one:OEChildren>'
        "  <one:OE><one:T><![CDATA[<span style='font-weight:bold'>Bold</span> then plain]]></one:T></one:OE>"
        '</one:OEChildren></one:Outline>'
    )
    p = parse_page(xml).outlines[0].items[0]
    assert p.text is None
    assert p.segments is not None
    assert [s.text for s in p.segments] == ["Bold", " then plain"]
    assert p.segments[0].bold is True
    assert p.segments[1].bold is False


def test_parse_paragraph_lang_only_span_is_plain_text():
    """A span carrying only a non-style attribute (e.g. lang) is not formatting."""
    xml = _page_xml(
        '<one:Outline><one:OEChildren>'
        "  <one:OE><one:T><![CDATA[<span lang=en-US>Plain despite span</span>]]></one:T></one:OE>"
        '</one:OEChildren></one:Outline>'
    )
    p = parse_page(xml).outlines[0].items[0]
    assert p.text == "Plain despite span"
    assert p.bold is False and p.segments is None


def test_parse_list_item_formatting_preserved_as_segments():
    xml = _page_xml(
        '<one:Outline><one:OEChildren>'
        "  <one:OE><one:List><one:Bullet/></one:List>"
        "    <one:T><![CDATA[<span style='font-style:italic'>Italic item</span>]]></one:T></one:OE>"
        '</one:OEChildren></one:Outline>'
    )
    lst = parse_page(xml).outlines[0].items[0]
    assert isinstance(lst, List)
    item = lst.items[0]
    assert item.text is None
    assert item.segments is not None
    assert item.segments[0].text == "Italic item"
    assert item.segments[0].italic is True


def test_roundtrip_bold_paragraph_preserves_formatting():
    """build → parse round-trip now preserves whole-paragraph bold (was lost)."""
    outline = Outline(items=[Paragraph(type="paragraph", text="Important", bold=True)])
    xml = build_page_xml("Doc", [outline])
    p = parse_page(xml).outlines[0].items[0]
    assert p.text == "Important"
    assert p.bold is True


def test_roundtrip_segments_preserve_per_run_formatting():
    segments = [
        TextRun(text="red ", color="#ff0000"),
        TextRun(text="and italic", italic=True),
    ]
    outline = Outline(items=[Paragraph(type="paragraph", segments=segments)])
    xml = build_page_xml("Doc", [outline])
    p = parse_page(xml).outlines[0].items[0]
    assert p.segments is not None
    assert [s.text for s in p.segments] == ["red ", "and italic"]
    assert p.segments[0].color == "#ff0000"
    assert p.segments[1].italic is True


# ---------------------------------------------------------------------------
# Round-trip tests
# ---------------------------------------------------------------------------

def test_roundtrip_plain_paragraph():
    outline = Outline(items=[Paragraph(type="paragraph", text="Round-trip text")])
    xml = build_page_xml("Round-trip Title", [outline])
    page = parse_page(xml)
    assert page.title == "Round-trip Title"
    assert isinstance(page.outlines[0].items[0], Paragraph)
    assert page.outlines[0].items[0].text == "Round-trip text"


def test_roundtrip_heading():
    # Explicit trade-off: our inline-CSS heading emissions do not carry a
    # quickStyleIndex, so the read path cannot reconstruct the heading style
    # from the XML we wrote.  Text survives; style degrades to "normal".
    # Reading a OneNote-native page with a proper QuickStyleDef table still
    # resolves heading styles correctly (see
    # test_parse_page_h1_resolved_via_qsd_table).
    outline = Outline(items=[Paragraph(type="paragraph", text="Chapter", style="h2")])
    xml = build_page_xml("Doc", [outline])
    page = parse_page(xml)
    p = page.outlines[0].items[0]
    assert isinstance(p, Paragraph)
    assert p.text == "Chapter"
    # style will be "normal" — this is the accepted round-trip degradation


def test_roundtrip_bullet_list():
    lst = List(type="list", style="bullet", items=[
        ListItem(text="Alpha"),
        ListItem(text="Beta"),
    ])
    outline = Outline(items=[lst])
    xml = build_page_xml("ListPage", [outline])
    page = parse_page(xml)
    result = page.outlines[0].items[0]
    assert isinstance(result, List)
    assert result.style == "bullet"
    assert len(result.items) == 2
    assert result.items[0].text == "Alpha"
    assert result.items[1].text == "Beta"


def test_roundtrip_numbered_list():
    lst = List(type="list", style="numbered", items=[
        ListItem(text="Step 1"),
        ListItem(text="Step 2"),
    ])
    outline = Outline(items=[lst])
    xml = build_page_xml("Steps", [outline])
    page = parse_page(xml)
    result = page.outlines[0].items[0]
    assert isinstance(result, List)
    assert result.style == "numbered"
    assert result.items[1].text == "Step 2"


# ---------------------------------------------------------------------------
# FloatingImage — _floating_image_to_el and build_page_xml / parse_page
# ---------------------------------------------------------------------------

def test_floating_image_to_el_tag():
    """_floating_image_to_el returns a <one:Image> element (not wrapped in OE)."""
    img = FloatingImage(handle="mcpref:abc123def456")
    el = _floating_image_to_el(img)
    _ONE_NS_URI = "http://schemas.microsoft.com/office/onenote/2013/onenote"
    assert el.tag == f"{{{_ONE_NS_URI}}}Image"


def test_floating_image_to_el_has_data():
    """_floating_image_to_el embeds the handle in <one:Data>."""
    img = FloatingImage(handle="mcpref:abc123def456")
    el = _floating_image_to_el(img)
    data_el = el.find("one:Data", _NS)
    assert data_el is not None
    assert data_el.text == "mcpref:abc123def456"


def test_floating_image_to_el_position():
    """_floating_image_to_el adds <one:Position> when position is set."""
    pos = Position(x=36.0, y=71.0, z=1)
    img = FloatingImage(handle="mcpref:abc123def456", position=pos)
    el = _floating_image_to_el(img)
    pos_el = el.find("one:Position", _NS)
    assert pos_el is not None
    assert pos_el.get("x") == "36.0"
    assert pos_el.get("y") == "71.0"
    assert pos_el.get("z") == "1"


def test_floating_image_to_el_no_position_when_none():
    """_floating_image_to_el omits <one:Position> when position is None."""
    img = FloatingImage(handle="mcpref:abc123def456")
    el = _floating_image_to_el(img)
    pos_el = el.find("one:Position", _NS)
    assert pos_el is None


def test_floating_image_to_el_size():
    """_floating_image_to_el adds <one:Size> when width and height are set."""
    img = FloatingImage(handle="mcpref:abc123def456", width=72.0, height=72.0)
    el = _floating_image_to_el(img)
    size_el = el.find("one:Size", _NS)
    assert size_el is not None
    assert size_el.get("width") == "72.0"
    assert size_el.get("height") == "72.0"
    assert size_el.get("isSetByUser") == "true"


def test_floating_image_to_el_no_size_when_absent():
    """_floating_image_to_el omits <one:Size> when dimensions are absent."""
    img = FloatingImage(handle="mcpref:abc123def456")
    el = _floating_image_to_el(img)
    size_el = el.find("one:Size", _NS)
    assert size_el is None


def test_build_page_xml_floating_image_at_page_level():
    """build_page_xml with images puts <one:Image> as direct child of <one:Page>."""
    img = FloatingImage(handle="mcpref:abc123def456", width=72.0, height=72.0,
                        position=Position(x=36.0, y=156.0, z=1))
    xml = build_page_xml("Title", [], images=[img])
    root = _parse_xml(xml)
    # Image must be a direct child of root (not inside an Outline)
    image_el = root.find("one:Image", _NS)
    assert image_el is not None, "Expected <one:Image> as direct child of <one:Page>"
    data_el = image_el.find("one:Data", _NS)
    assert data_el is not None
    assert data_el.text == "mcpref:abc123def456"


def test_build_page_xml_floating_image_not_inside_outline():
    """build_page_xml with images does NOT nest the image inside any Outline."""
    img = FloatingImage(handle="mcpref:abc123def456")
    xml = build_page_xml("Title", [], images=[img])
    root = _parse_xml(xml)
    # Should be zero images inside outlines
    nested = root.findall("one:Outline/one:OEChildren/one:OE/one:Image", _NS)
    assert not nested, "FloatingImage must not be nested inside an Outline"


def test_parse_page_floating_image_in_images_field():
    """parse_page returns top-level <one:Image> in content.images (not in outlines)."""
    xml = _page_xml(
        '<one:Image>'
        '  <one:Position x="36.0" y="71.0" z="0"/>'
        '  <one:Size width="679.0" height="88.0" isSetByUser="true"/>'
        '  <one:Data>mcpref:abc123def456</one:Data>'
        '</one:Image>'
    )
    page = parse_page(xml)
    assert page.images, "Expected at least one FloatingImage in content.images"
    img = page.images[0]
    assert isinstance(img, FloatingImage)
    assert img.handle == "mcpref:abc123def456"
    assert img.width == 679.0
    assert img.height == 88.0
    assert img.position is not None
    assert img.position.x == 36.0
    assert img.position.y == 71.0


def test_parse_page_floating_image_not_in_outlines():
    """parse_page does NOT create synthetic Outline objects for floating images."""
    xml = _page_xml(
        '<one:Image>'
        '  <one:Data>mcpref:abc123def456</one:Data>'
        '</one:Image>'
    )
    page = parse_page(xml)
    assert page.outlines == [], "No synthetic Outline should be created for a floating image"


def test_roundtrip_floating_image():
    """FloatingImage survives a build_page_xml -> parse_page round-trip."""
    pos = Position(x=36.0, y=156.0, z=1)
    img = FloatingImage(handle="mcpref:abc123def456", width=72.0, height=72.0, position=pos)
    xml = build_page_xml("Title", [], images=[img])
    page = parse_page(xml)
    assert len(page.images) == 1
    recovered = page.images[0]
    assert isinstance(recovered, FloatingImage)
    assert recovered.handle == "mcpref:abc123def456"
    assert recovered.width == 72.0
    assert recovered.height == 72.0
    assert recovered.position is not None
    assert recovered.position.x == 36.0
    assert recovered.position.y == 156.0
    assert recovered.position.z == 1


def test_multiple_floating_images():
    """build_page_xml with multiple images preserves all at page level."""
    imgs = [
        FloatingImage(handle="mcpref:aaaaaaaaaaaa", position=Position(x=36.0, y=71.0, z=0)),
        FloatingImage(handle="mcpref:bbbbbbbbbbbb", position=Position(x=36.0, y=156.0, z=1)),
    ]
    xml = build_page_xml("Title", [], images=imgs)
    root = _parse_xml(xml)
    image_els = root.findall("one:Image", _NS)
    assert len(image_els) == 2, "Expected exactly 2 top-level <one:Image> elements"


# ---------------------------------------------------------------------------
# InlineImage — write direction
# ---------------------------------------------------------------------------

def test_build_page_xml_inline_image_with_size():
    """Inline image with size emits Outline > OEChildren > OE > Image with Size and Data."""
    img = InlineImage(type="inline_image", handle="mcpref:xxxxxxxxxxxx", width=100.0, height=200.0)
    outline = Outline(items=[img])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)

    # Verify structure: Outline > OEChildren > OE > Image
    oe_el = root.find("one:Outline/one:OEChildren/one:OE", _NS)
    assert oe_el is not None, "Expected OE inside OEChildren"
    image_el = oe_el.find("one:Image", _NS)
    assert image_el is not None, "Expected <one:Image> inside OE"

    size_el = image_el.find("one:Size", _NS)
    assert size_el is not None, "Expected <one:Size> inside Image"
    assert size_el.get("width") == "100.0"
    assert size_el.get("height") == "200.0"
    assert size_el.get("isSetByUser") == "true"

    data_el = image_el.find("one:Data", _NS)
    assert data_el is not None, "Expected <one:Data> inside Image"
    assert data_el.text == "mcpref:xxxxxxxxxxxx"

    # Must NOT have a Position element (inline images are not canvas-positioned)
    pos_el = image_el.find("one:Position", _NS)
    assert pos_el is None, "Inline images must not have a Position element"


def test_build_page_xml_inline_image_without_size():
    """Inline image without dimensions omits <one:Size> but keeps <one:Data>."""
    img = InlineImage(type="inline_image", handle="mcpref:xxxxxxxxxxxx")
    outline = Outline(items=[img])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)

    image_el = root.find("one:Outline/one:OEChildren/one:OE/one:Image", _NS)
    assert image_el is not None

    size_el = image_el.find("one:Size", _NS)
    assert size_el is None, "Size element must be omitted when width/height are absent"

    data_el = image_el.find("one:Data", _NS)
    assert data_el is not None
    assert data_el.text == "mcpref:xxxxxxxxxxxx"


def test_build_page_xml_inline_image_outline_width_inferred():
    """Outline with no explicit width infers size from InlineImage's width and height."""
    img = InlineImage(type="inline_image", handle="mcpref:xxxxxxxxxxxx", width=350.0, height=250.0)
    outline = Outline(items=[img])  # no explicit width
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)

    # Outline Size must only be emitted when both width and height are available
    # (OneNote rejects width-only Size on outlines containing images).
    size_el = root.find("one:Outline/one:Size", _NS)
    assert size_el is not None, "Outline Size must be inferred from InlineImage width+height"
    assert size_el.get("width") == "350.0"
    assert size_el.get("height") == "250.0"


# ---------------------------------------------------------------------------
# InlineImage — round-trip
# ---------------------------------------------------------------------------

def test_roundtrip_inline_image():
    """InlineImage survives a build_page_xml -> parse_page round-trip."""
    img = InlineImage(type="inline_image", handle="mcpref:abc123def456", width=100.0, height=200.0)
    outline = Outline(items=[img])
    xml = build_page_xml("RoundTrip", [outline])
    page = parse_page(xml)

    assert len(page.outlines) == 1
    items = page.outlines[0].items
    assert len(items) == 1

    recovered = items[0]
    assert isinstance(recovered, InlineImage)
    assert recovered.handle == "mcpref:abc123def456"
    assert recovered.width == 100.0
    assert recovered.height == 200.0
