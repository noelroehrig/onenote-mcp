"""Unit tests for onenote_mcp.builders — no COM dependency."""
import base64
import xml.etree.ElementTree as ET
import zlib
import pytest
from onenote_mcp.builders import (
    _run_to_html, _paragraph_html, _cdata_placeholder, _apply_cdata,
    build_page_xml, build_append_xml,
    parse_page,
    _box_pixel_size, _box_png, _clean_text, _floating_image_to_el,
)
from onenote_mcp.images import pixel_size
from onenote_mcp.models import (
    TextRun, Paragraph, ImagePlaceholder, InlineImage, ListItem, List,
    FloatingImage, Outline, Position, PageContent, UnsupportedItem,
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


def _quick_style_defs(root: ET.Element) -> dict[str, ET.Element]:
    """Return the page's QuickStyleDefs by index."""
    return {qsd.get("index"): qsd for qsd in root.findall("one:QuickStyleDef", _NS)}


def test_build_page_xml_h1_heading():
    # The heading look lives in a QuickStyleDef named "h1", so OneNote treats
    # the paragraph as a real heading; the text itself carries no inline CSS.
    outline = Outline(items=[Paragraph(type="paragraph", text="Heading", style="h1")])
    root = _parse_xml(build_page_xml("Title", [outline]))
    assert root[0].tag == f"{{{_NS['one']}}}QuickStyleDef", "QuickStyleDefs must precede <one:Title>"
    oe_el = root.find("one:Outline/one:OEChildren/one:OE", _NS)
    qsd = _quick_style_defs(root)[oe_el.get("quickStyleIndex")]
    assert qsd.get("name") == "h1"
    assert qsd.get("font") == "Calibri"
    assert qsd.get("fontSize") == "16.0"
    assert qsd.get("fontColor") == "#1E4E79"
    assert qsd.get("bold") == "true"
    assert oe_el.find("one:T", _NS).text == "Heading"


def test_build_page_xml_h6_heading():
    # h6: italic, not bold, color #595959, font size 11
    outline = Outline(items=[Paragraph(type="paragraph", text="H6", style="h6")])
    root = _parse_xml(build_page_xml("Title", [outline]))
    oe_el = root.find("one:Outline/one:OEChildren/one:OE", _NS)
    qsd = _quick_style_defs(root)[oe_el.get("quickStyleIndex")]
    assert qsd.get("name") == "h6"
    assert qsd.get("fontSize") == "11.0"
    assert qsd.get("fontColor") == "#595959"
    assert qsd.get("italic") == "true"
    assert qsd.get("bold") is None


def test_build_page_xml_normal_no_quick_style():
    outline = Outline(items=[Paragraph(type="paragraph", text="Normal", style="normal")])
    xml = build_page_xml("Title", [outline])
    root = _parse_xml(xml)
    oe_el = root.find("one:Outline/one:OEChildren/one:OE", _NS)
    assert oe_el is not None
    assert oe_el.get("quickStyleIndex") is None
    assert root.find("one:QuickStyleDef", _NS) is None


def test_build_page_xml_declares_each_used_heading_style_once():
    outline = Outline(items=[
        Paragraph(type="paragraph", text="A", style="h1"),
        Paragraph(type="paragraph", text="B", style="h3"),
        Paragraph(type="paragraph", text="C", style="h1"),
        Paragraph(type="paragraph", text="D"),
    ])
    root = _parse_xml(build_page_xml("Title", [outline]))
    defs = _quick_style_defs(root)
    assert sorted(qsd.get("name") for qsd in defs.values()) == ["h1", "h3"]
    oes = root.findall("one:Outline/one:OEChildren/one:OE", _NS)
    names = [defs[oe.get("quickStyleIndex")].get("name") if oe.get("quickStyleIndex") else None for oe in oes]
    assert names == ["h1", "h3", "h1", None]


def test_build_page_xml_heading_keeps_explicit_formatting_inline():
    outline = Outline(items=[Paragraph(type="paragraph", text="Red", style="h2", color="#C00000")])
    xml = build_page_xml("Title", [outline])
    assert '<span style="color:#C00000">Red</span>' in xml


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
    assert '<![CDATA[<span style="font-weight:bold;background:#FFEB3B">Screenshot of dashboard</span>]]>' in xml
    assert "INSERT IMAGE" not in xml


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
# build_append_xml
# ---------------------------------------------------------------------------

def test_build_append_xml_wraps_outline_in_page_without_id():
    outline = Outline(items=[Paragraph(type="paragraph", text="content")])
    root = _parse_xml(build_append_xml(outline))
    assert root.tag == f"{{{_NS['one']}}}Page"
    assert root.get("ID") is None
    assert [child.tag for child in root] == [f"{{{_NS['one']}}}Outline"]


def test_build_append_xml_paragraph():
    outline = Outline(items=[Paragraph(type="paragraph", text="content")])
    root = _parse_xml(build_append_xml(outline))
    t_el = root.find("one:Outline/one:OEChildren/one:OE/one:T", _NS)
    assert t_el is not None
    assert t_el.text == "content"


def test_build_append_xml_declares_heading_style_before_outline():
    outline = Outline(items=[Paragraph(type="paragraph", text="Appended", style="h2")])
    root = _parse_xml(build_append_xml(outline))
    qsd, outline_el = root
    assert qsd.tag == f"{{{_NS['one']}}}QuickStyleDef"
    assert qsd.get("name") == "h2"
    assert outline_el.find("one:OEChildren/one:OE", _NS).get("quickStyleIndex") == qsd.get("index")


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


def test_parse_section_pages_returns_id_name_and_level():
    xml = (
        '<one:Section xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"'
        ' ID="sec1" name="Templates">'
        '<one:Page ID="p1" name="Template A" pageLevel="1"/>'
        '<one:Page ID="p2" name="Template B"/>'
        '</one:Section>'
    )
    from onenote_mcp.builders import parse_section_pages
    assert parse_section_pages(xml) == [
        {"id": "p1", "name": "Template A", "level": 1},
        {"id": "p2", "name": "Template B", "level": 1},
    ]


def test_parse_section_pages_keeps_section_order_and_subpage_levels():
    xml = (
        '<one:Section xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="sec1">'
        '<one:Page ID="parent" name="Parent" pageLevel="1"/>'
        '<one:Page ID="child" name="Child" pageLevel="2" isSubPage="true"/>'
        '<one:Page ID="grandchild" name="Grandchild" pageLevel="3" isSubPage="true"/>'
        '<one:Page ID="next" name="Next" pageLevel="1"/>'
        '</one:Section>'
    )
    from onenote_mcp.builders import parse_section_pages
    assert [(page["id"], page["level"]) for page in parse_section_pages(xml)] == [
        ("parent", 1), ("child", 2), ("grandchild", 3), ("next", 1),
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


def test_parse_page_title_joins_several_t_elements():
    xml = _page_xml(
        '<one:Title><one:OE><one:T>My </one:T><one:T selected="all"></one:T><one:T>Title</one:T></one:OE></one:Title>'
    )
    assert parse_page(xml).title == "My Title"


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


def test_parse_page_reports_the_measured_outline_height():
    xml = _page_xml(
        '<one:Outline>'
        '  <one:Size width="93.8" height="134.27"/>'
        '  <one:OEChildren><one:OE><one:T>content</one:T></one:OE></one:OEChildren>'
        '</one:Outline>'
    )
    outline = parse_page(xml).model_dump(exclude_defaults=True)["outlines"][0]
    assert (outline["width"], outline["height"]) == (93.8, 134.27)


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


def test_parse_page_heading_drops_bold_supplied_by_its_quick_style():
    """OneNote repeats a QuickStyleDef's bold as an inline span in <one:T>;
    it is reported once, as part of the heading style."""
    xml = _page_xml(
        '<one:QuickStyleDef index="0" name="p" font="Calibri" fontSize="11.0"/>',
        '<one:QuickStyleDef index="1" name="h1" font="Calibri" fontSize="16.0" bold="true"/>',
        '<one:Outline><one:OEChildren>'
        "  <one:OE quickStyleIndex=\"1\"><one:T><![CDATA[<span style='font-weight:bold'>Integration test</span>]]></one:T></one:OE>"
        '</one:OEChildren></one:Outline>'
    )
    p = parse_page(xml).outlines[0].items[0]
    assert (p.style, p.text, p.bold, p.segments) == ("h1", "Integration test", False, None)


def test_parse_page_heading_keeps_formatting_beyond_its_quick_style():
    xml = _page_xml(
        '<one:QuickStyleDef index="1" name="h3" font="Calibri" fontSize="13.0" bold="true"/>',
        '<one:Outline><one:OEChildren>'
        '  <one:OE quickStyleIndex="1"><one:T><![CDATA['
        "<span style='font-weight:bold'>Plain </span>"
        "<span style='font-weight:bold;font-style:italic'>italic</span>"
        ']]></one:T></one:OE>'
        '</one:OEChildren></one:Outline>'
    )
    p = parse_page(xml).outlines[0].items[0]
    assert p.style == "h3"
    assert [(s.text, s.bold, s.italic) for s in p.segments] == [("Plain ", False, False), ("italic", False, True)]


def test_parse_page_bold_on_heading_without_bold_style_is_kept():
    """A OneNote-native heading style has no bold; bold on its text is real formatting."""
    xml = _page_xml(
        '<one:QuickStyleDef index="2" name="h1" font="Calibri" fontSize="14.0"/>',
        '<one:Outline><one:OEChildren>'
        "  <one:OE quickStyleIndex=\"2\"><one:T><![CDATA[<span style='font-weight:bold'>Bold heading</span>]]></one:T></one:OE>"
        '</one:OEChildren></one:Outline>'
    )
    p = parse_page(xml).outlines[0].items[0]
    assert (p.style, p.text, p.bold) == ("h1", "Bold heading", True)


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
    outline = Outline(items=[Paragraph(type="paragraph", text="Chapter", style="h2")])
    xml = build_page_xml("Doc", [outline])
    page = parse_page(xml)
    p = page.outlines[0].items[0]
    assert isinstance(p, Paragraph)
    assert p.text == "Chapter"
    assert p.style == "h2"
    assert p.bold is False


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


def test_floating_image_to_el_height_only():
    """With only height, Size carries just that; com.py adds the width before writing."""
    el = _floating_image_to_el(FloatingImage(handle="mcpref:abc123def456", height=50.0))
    assert el.find("one:Size", _NS).attrib == {"height": "50.0", "isSetByUser": "true"}


def test_build_page_xml_inline_image_width_only():
    """With only width, Size carries just that; com.py adds the height before writing."""
    outline = Outline(items=[InlineImage(type="inline_image", handle="mcpref:abc123def456", width=240.0)])
    root = _parse_xml(build_page_xml("Title", [outline]))
    size_el = root.find("one:Outline/one:OEChildren/one:OE/one:Image/one:Size", _NS)
    assert size_el.attrib == {"width": "240.0", "isSetByUser": "true"}


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


# ---------------------------------------------------------------------------
# Text escaping: OneNote reads <one:T> content as HTML, so text is escaped on
# write and must come back unchanged from parse_page.
# ---------------------------------------------------------------------------

_SPECIAL_TEXTS = [
    "a < b > c",
    "Fish & Chips, &amp; and &lt;",
    "<b>not bold</b>",
    "ends a CDATA section ]]> early",
    "\"double\" and 'single' quotes",
    "Grüße: äöü ÄÖÜ ß",
]


def _first_item(outline: Outline, title: str = "Doc"):
    return parse_page(build_page_xml(title, [outline])).outlines[0].items[0]


@pytest.mark.parametrize("text", _SPECIAL_TEXTS)
def test_special_text_round_trips_in_plain_paragraph(text):
    p = _first_item(Outline(items=[Paragraph(type="paragraph", text=text)]))
    assert p.text == text
    assert p.segments is None and p.bold is False


@pytest.mark.parametrize("text", _SPECIAL_TEXTS)
def test_special_text_round_trips_in_styled_paragraph(text):
    para = Paragraph(type="paragraph", text=text, bold=True, color="#C00000", font_family="Courier New")
    p = _first_item(Outline(items=[para]))
    assert p.text == text
    assert p.bold is True
    assert p.color == "#C00000"
    assert p.font_family == "Courier New"


@pytest.mark.parametrize("text", _SPECIAL_TEXTS)
def test_special_text_round_trips_in_heading(text):
    p = _first_item(Outline(items=[Paragraph(type="paragraph", text=text, style="h1")]))
    assert p.text == text


@pytest.mark.parametrize("style", ["normal", "h2"])
@pytest.mark.parametrize("text", _SPECIAL_TEXTS)
def test_special_text_round_trips_in_segments(text, style):
    segments = [TextRun(text=text), TextRun(text=text, italic=True)]
    p = _first_item(Outline(items=[Paragraph(type="paragraph", segments=segments, style=style)]))
    assert [s.text for s in p.segments] == [text, text]
    assert p.segments[1].italic is True


@pytest.mark.parametrize("text", _SPECIAL_TEXTS)
def test_special_text_round_trips_in_list_items(text):
    items = [ListItem(text=text), ListItem(segments=[TextRun(text=text, bold=True)])]
    lst = _first_item(Outline(items=[List(type="list", style="bullet", items=items)]))
    assert lst.items[0].text == text
    assert lst.items[1].segments[0].text == text


@pytest.mark.parametrize("text", _SPECIAL_TEXTS)
def test_special_text_round_trips_in_title(text):
    assert parse_page(build_page_xml(text, [])).title == text


@pytest.mark.parametrize("text", _SPECIAL_TEXTS)
def test_special_text_round_trips_in_image_placeholder(text):
    ph = ImagePlaceholder(type="image_placeholder", description=text, width=40.0, height=30.0)
    assert _first_item(Outline(items=[ph])) == ph


def test_t_content_is_escaped_html():
    """What OneNote receives: escaped text, wrapped in a span only when styled."""
    text = "<b>x</b> & ]]>"
    outline = Outline(items=[
        Paragraph(type="paragraph", text=text),
        Paragraph(type="paragraph", text=text, bold=True),
    ])
    root = ET.fromstring(build_page_xml("Doc", [outline]))
    plain_t, styled_t = [
        oe.find("one:T", _NS).text for oe in root.iterfind("one:Outline/one:OEChildren/one:OE", _NS)
    ]
    assert plain_t == "&lt;b&gt;x&lt;/b&gt; &amp; ]]&gt;"
    assert styled_t == '<span style="font-weight:bold">&lt;b&gt;x&lt;/b&gt; &amp; ]]&gt;</span>'


def test_cdata_terminator_only_closes_cdata_sections():
    text = "ends ]]> early"
    outline = Outline(items=[
        Paragraph(type="paragraph", text=text),
        Paragraph(type="paragraph", text=text, italic=True),
        Paragraph(type="paragraph", segments=[TextRun(text=text, bold=True)]),
        List(type="list", style="numbered", items=[ListItem(segments=[TextRun(text=text, bold=True)])]),
        ImagePlaceholder(type="image_placeholder", description=text),
    ])
    xml = build_page_xml(text, [outline])
    assert xml.count("<![CDATA[") == 4
    assert xml.count("]]>") == xml.count("<![CDATA[")


def test_apply_cdata_rejects_cdata_terminator():
    with pytest.raises(ValueError, match=r"\]\]>"):
        _apply_cdata("__CDATA_000000__", {0: "a ]]> b"})


# ---------------------------------------------------------------------------
# Style values read from OneNote are normalized to what the models accept
# ---------------------------------------------------------------------------

def _styled_paragraph(style: str, text: str = "styled") -> Paragraph:
    xml = _page_xml(
        '<one:Outline><one:OEChildren>'
        f"  <one:OE><one:T><![CDATA[<span style='{style}'>{text}</span>]]></one:T></one:OE>"
        '</one:OEChildren></one:Outline>'
    )
    return parse_page(xml).outlines[0].items[0]


def test_parse_named_highlight_maps_to_hex():
    p = _styled_paragraph("background:yellow;mso-highlight:yellow")
    assert p.text == "styled"
    assert p.highlight == "#FFFF00"


def test_parse_unsupported_color_value_is_left_out():
    p = _styled_paragraph("color:windowtext;font-weight:bold")
    assert p.text == "styled"
    assert p.color is None
    assert p.bold is True


def test_parse_quoted_font_family_is_unquoted():
    p = _styled_paragraph('font-family:"Segoe UI",Arial')
    assert p.font_family == "Segoe UI,Arial"


def test_parse_unsupported_font_family_is_left_out():
    p = _styled_paragraph("font-family:ＭＳ ゴシック")
    assert p.text == "styled"
    assert p.font_family is None


# ---------------------------------------------------------------------------
# Paragraph formatting from the OE and <one:T> style attributes: OneNote moves
# paragraph-wide formatting there, spans inside <one:T> override it per run.
# ---------------------------------------------------------------------------

_P_QSD = (
    '<one:QuickStyleDef index="0" name="p" fontColor="automatic" '
    'highlightColor="automatic" font="Calibri" fontSize="11.0"/>'
)
_H1_QSD = (
    '<one:QuickStyleDef index="1" name="h1" fontColor="#1E4E79" '
    'highlightColor="automatic" font="Calibri" fontSize="16.0" bold="true"/>'
)
_BULLET = '<one:List><one:Bullet bullet="2"/></one:List>'


def _outline_xml(*oes: str, defs: str = "") -> str:
    return _page_xml(defs, "<one:Outline><one:OEChildren>", *oes, "</one:OEChildren></one:Outline>")


def _items(*oes: str, defs: str = "") -> list[dict]:
    """Parse *oes* as one outline and return its items as get_page reports them."""
    page = parse_page(_outline_xml(*oes, defs=defs))
    return page.model_dump(exclude_defaults=True)["outlines"][0]["items"]


def _reread(page: PageContent) -> PageContent:
    """Write *page* back as replace_page would and read the result again."""
    return parse_page(build_page_xml(page.title, page.outlines, page.images))


def test_parse_paragraph_reads_oe_style():
    assert _items('<one:OE style="font-size:22.0pt"><one:T>x</one:T></one:OE>') == [
        {"type": "paragraph", "text": "x", "font_size": 22.0},
    ]


def test_parse_paragraph_reads_t_style():
    assert _items('<one:OE><one:T style="color:#C00000">x</one:T></one:OE>') == [
        {"type": "paragraph", "text": "x", "color": "#C00000"},
    ]


def test_parse_paragraph_oe_style_combines_with_uniform_span():
    items = _items(
        '<one:OE style="color:#9C0000;font-family:Courier New">'
        "<one:T><![CDATA[<span style='font-weight:bold;background:#FFFF00'>x</span>]]></one:T></one:OE>"
    )
    assert items == [{
        "type": "paragraph", "text": "x", "bold": True, "color": "#9C0000",
        "highlight": "#FFFF00", "font_family": "Courier New",
    }]


def test_parse_paragraph_span_overrides_oe_style_per_run():
    items = _items(
        '<one:OE style="color:#C00000;font-family:Courier New">'
        "<one:T><![CDATA[<span style='color:#0000FF'>blue</span> red]]></one:T></one:OE>"
    )
    assert items == [{"type": "paragraph", "segments": [
        {"text": "blue", "color": "#0000FF", "font_family": "Courier New"},
        {"text": " red", "color": "#C00000", "font_family": "Courier New"},
    ]}]


def test_parse_paragraph_span_can_switch_off_oe_formatting():
    items = _items(
        '<one:OE style="font-weight:bold;font-style:italic;text-decoration:underline"><one:T><![CDATA['
        "bold <span style='font-weight:normal;font-style:normal;text-decoration:none'>plain</span>"
        "]]></one:T></one:OE>"
    )
    assert items == [{"type": "paragraph", "segments": [
        {"text": "bold ", "bold": True, "italic": True, "underline": True},
        {"text": "plain"},
    ]}]


def test_parse_paragraph_joins_several_t_elements():
    items = _items(
        '<one:OE><one:T>Hello </one:T><one:T style="font-weight:bold">big</one:T><one:T> world</one:T></one:OE>'
    )
    assert items == [{"type": "paragraph", "segments": [
        {"text": "Hello "}, {"text": "big", "bold": True}, {"text": " world"},
    ]}]


def test_parse_paragraph_style_repeating_its_quick_style_reads_as_plain():
    items = _items(
        '<one:OE quickStyleIndex="0" style="font-family:Calibri;font-size:11.0pt"><one:T>x</one:T></one:OE>',
        defs=_P_QSD,
    )
    assert items == [{"type": "paragraph", "text": "x"}]


def test_parse_paragraph_reports_formatting_beyond_its_quick_style():
    items = _items(
        '<one:OE quickStyleIndex="0" style="font-family:Arial;font-size:11.0pt"><one:T>x</one:T></one:OE>',
        defs=_P_QSD,
    )
    assert items == [{"type": "paragraph", "text": "x", "font_family": "Arial"}]


def test_parse_heading_reports_only_the_heading_style():
    """Font, size, colour and bold all repeat the h1 QuickStyleDef (colour in
    another case), so the heading reads back without per-paragraph formatting."""
    items = _items(
        '<one:OE quickStyleIndex="1" style="font-family:Calibri;font-size:16.0pt;color:#1e4e79">'
        "<one:T><![CDATA[<span style='font-weight:bold'>Head</span>]]></one:T></one:OE>",
        defs=_P_QSD + _H1_QSD,
    )
    assert items == [{"type": "paragraph", "text": "Head", "style": "h1"}]


def test_parse_paragraph_keeps_formatting_of_a_quick_style_a_rewrite_drops():
    """Only 'p' and heading styles come back on a rewrite, so the formatting
    another quick style supplies is reported to keep the paragraph's look."""
    qsd = '<one:QuickStyleDef index="2" name="code" font="Consolas" fontSize="10.0"/>'
    items = _items(
        '<one:OE quickStyleIndex="2" style="font-family:Consolas;font-size:10.0pt"><one:T>x = 1</one:T></one:OE>',
        defs=qsd,
    )
    assert items == [{"type": "paragraph", "text": "x = 1", "font_family": "Consolas", "font_size": 10.0}]


def test_parse_list_item_reads_oe_style_as_segments():
    items = _items(f'<one:OE style="font-size:14.0pt">{_BULLET}<one:T>big item</one:T></one:OE>')
    assert items == [{"type": "list", "style": "bullet", "items": [
        {"segments": [{"text": "big item", "font_size": 14.0}]},
    ]}]


def test_parse_list_item_style_repeating_p_quick_style_reads_as_text():
    items = _items(
        f'<one:OE quickStyleIndex="0" style="font-family:Calibri;font-size:11.0pt">{_BULLET}<one:T>item</one:T></one:OE>',
        defs=_P_QSD,
    )
    assert items == [{"type": "list", "style": "bullet", "items": [{"text": "item"}]}]


def test_parse_list_item_keeps_heading_quick_style_formatting():
    """A list item has no heading style, so the bold of a bulleted heading is
    reported instead of dropped."""
    items = _items(
        f"<one:OE quickStyleIndex=\"1\">{_BULLET}<one:T><![CDATA[<span style='font-weight:bold'>item</span>]]></one:T></one:OE>",
        defs=_H1_QSD,
    )
    assert items == [{"type": "list", "style": "bullet", "items": [{"segments": [{"text": "item", "bold": True}]}]}]


def test_roundtrip_oe_style_formatting_is_stable():
    page = parse_page(_outline_xml(
        '<one:OE style="color:#C00000;font-family:Courier New">'
        "<one:T><![CDATA[<span style='font-weight:bold'>x</span> y]]></one:T></one:OE>",
        f'<one:OE style="font-size:14.0pt">{_BULLET}<one:T>item</one:T></one:OE>',
    ))
    assert _reread(page).model_dump(exclude_defaults=True) == page.model_dump(exclude_defaults=True)


# ---------------------------------------------------------------------------
# Indented content: <one:OEChildren> under paragraphs, images and list items
# ---------------------------------------------------------------------------

def test_parse_paragraph_children():
    items = _items(
        "<one:OE><one:T>parent</one:T><one:OEChildren>"
        "<one:OE><one:T>child</one:T><one:OEChildren><one:OE><one:T>grandchild</one:T></one:OE></one:OEChildren></one:OE>"
        f"<one:OE>{_BULLET}<one:T>a</one:T></one:OE>"
        f"<one:OE>{_BULLET}<one:T>b</one:T></one:OE>"
        "</one:OEChildren></one:OE>"
    )
    assert items == [{"type": "paragraph", "text": "parent", "children": [
        {"type": "paragraph", "text": "child", "children": [{"type": "paragraph", "text": "grandchild"}]},
        {"type": "list", "style": "bullet", "items": [{"text": "a"}, {"text": "b"}]},
    ]}]


def test_parse_oe_without_content_keeps_its_children():
    items = _items("<one:OE><one:OEChildren><one:OE><one:T>child</one:T></one:OE></one:OEChildren></one:OE>")
    assert items == [{"type": "paragraph", "text": "", "children": [{"type": "paragraph", "text": "child"}]}]


def test_parse_inline_image_children():
    items = _items(
        "<one:OE><one:Image><one:Data>mcpref:abc123def456</one:Data></one:Image>"
        "<one:OEChildren><one:OE><one:T>caption</one:T></one:OE></one:OEChildren></one:OE>"
    )
    assert items == [{"type": "inline_image", "handle": "mcpref:abc123def456", "children": [
        {"type": "paragraph", "text": "caption"},
    ]}]


def test_parse_page_reads_every_oechildren_of_an_outline():
    xml = _page_xml(
        "<one:Outline>"
        "<one:OEChildren><one:OE><one:T>first</one:T></one:OE></one:OEChildren>"
        "<one:OEChildren><one:OE><one:T>second</one:T></one:OE></one:OEChildren>"
        "</one:Outline>"
    )
    assert [item.text for item in parse_page(xml).outlines[0].items] == ["first", "second"]


def test_parse_list_item_keeps_non_list_children_in_order():
    items = _items(
        f"<one:OE>{_BULLET}<one:T>step</one:T><one:OEChildren>"
        f"<one:OE>{_BULLET}<one:T>sub</one:T></one:OE>"
        "<one:OE><one:Image><one:Data>mcpref:abc123def456</one:Data></one:Image></one:OE>"
        "<one:OE><one:T>note</one:T></one:OE>"
        "</one:OEChildren></one:OE>"
    )
    assert items == [{"type": "list", "style": "bullet", "items": [{"text": "step", "children": [
        {"text": "sub"},
        {"type": "inline_image", "handle": "mcpref:abc123def456"},
        {"type": "paragraph", "text": "note"},
    ]}]}]


def test_parse_bulleted_image_is_read_as_image_between_lists():
    items = _items(
        f"<one:OE>{_BULLET}<one:T>a</one:T></one:OE>",
        f"<one:OE>{_BULLET}<one:Image><one:Data>mcpref:abc123def456</one:Data></one:Image></one:OE>",
        f"<one:OE>{_BULLET}<one:T>b</one:T></one:OE>",
    )
    assert items == [
        {"type": "list", "style": "bullet", "items": [{"text": "a"}]},
        {"type": "inline_image", "handle": "mcpref:abc123def456"},
        {"type": "list", "style": "bullet", "items": [{"text": "b"}]},
    ]


def test_build_page_xml_nests_paragraph_children_after_the_text():
    outline = Outline(items=[
        Paragraph(type="paragraph", text="parent", children=[Paragraph(type="paragraph", text="child")]),
        Paragraph(type="paragraph", text="leaf"),
    ])
    root = _parse_xml(build_page_xml("Title", [outline]))
    parent_oe, leaf_oe = root.findall("one:Outline/one:OEChildren/one:OE", _NS)
    assert [el.tag for el in parent_oe] == [f"{{{_ONE_NS_URI}}}T", f"{{{_ONE_NS_URI}}}OEChildren"]
    assert parent_oe.find("one:OEChildren/one:OE/one:T", _NS).text == "child"
    assert leaf_oe.find("one:OEChildren", _NS) is None


def test_build_page_xml_nests_inline_image_children():
    image = InlineImage(type="inline_image", handle="mcpref:abc123def456",
                        children=[Paragraph(type="paragraph", text="caption")])
    root = _parse_xml(build_page_xml("Title", [Outline(items=[image])]))
    oe = root.find("one:Outline/one:OEChildren/one:OE", _NS)
    assert [el.tag for el in oe] == [f"{{{_ONE_NS_URI}}}Image", f"{{{_ONE_NS_URI}}}OEChildren"]
    assert oe.find("one:OEChildren/one:OE/one:T", _NS).text == "caption"


def test_build_page_xml_writes_non_list_children_of_list_items():
    lst = List(type="list", style="numbered", items=[ListItem(text="step", children=[
        ListItem(text="sub"), Paragraph(type="paragraph", text="note"),
    ])])
    root = _parse_xml(build_page_xml("Title", [Outline(items=[lst])]))
    child_oes = root.findall("one:Outline/one:OEChildren/one:OE/one:OEChildren/one:OE", _NS)
    assert [oe.find("one:T", _NS).text for oe in child_oes] == ["sub", "note"]
    assert [oe.find("one:List/one:Number", _NS) is not None for oe in child_oes] == [True, False]


def test_build_page_xml_omits_oechildren_when_all_sub_items_are_empty():
    lst = List(type="list", style="bullet", items=[ListItem(text="item", children=[ListItem()])])
    root = _parse_xml(build_page_xml("Title", [Outline(items=[lst])]))
    assert root.find("one:Outline/one:OEChildren/one:OE/one:OEChildren", _NS) is None


def test_roundtrip_nested_content_is_stable():
    page = parse_page(_outline_xml(
        "<one:OE><one:T>parent</one:T><one:OEChildren>"
        "<one:OE><one:T>child</one:T></one:OE>"
        f"<one:OE>{_BULLET}<one:T>bullet</one:T><one:OEChildren>"
        f"<one:OE>{_BULLET}<one:T>sub</one:T></one:OE>"
        "<one:OE><one:T>note under the bullet</one:T></one:OE>"
        "</one:OEChildren></one:OE>"
        "</one:OEChildren></one:OE>",
        "<one:OE><one:Image><one:Data>mcpref:abc123def456</one:Data></one:Image>"
        "<one:OEChildren><one:OE><one:T>caption</one:T></one:OE></one:OEChildren></one:OE>",
    ))
    first = page.model_dump(exclude_defaults=True)
    assert first["outlines"][0]["items"][0]["children"][1]["items"][0]["children"][1] == {
        "type": "paragraph", "text": "note under the bullet",
    }
    reread = _reread(page)
    assert reread.model_dump(exclude_defaults=True) == first
    assert _reread(reread).model_dump(exclude_defaults=True) == first


# ---------------------------------------------------------------------------
# Content the slim model cannot represent: read-only "unsupported" markers
# ---------------------------------------------------------------------------

def _cell(*texts: str) -> str:
    oes = "".join(f"<one:OE><one:T>{text}</one:T></one:OE>" for text in texts)
    return f"<one:Cell><one:OEChildren>{oes}</one:OEChildren></one:Cell>"


_BOLD_ADA = "<![CDATA[<span style='font-weight:bold'>Ada</span>]]>"
_TABLE = (
    '<one:Table bordersVisible="true"><one:Columns>'
    '<one:Column index="0" width="80.0"/><one:Column index="1" width="80.0"/>'
    "</one:Columns>"
    f"<one:Row>{_cell('Name')}{_cell('Age')}</one:Row>"
    f"<one:Row>{_cell(_BOLD_ADA, 'Lovelace')}{_cell('36')}</one:Row>"
    "</one:Table>"
)


def test_parse_table_becomes_unsupported_item_with_cell_text():
    assert _items(f"<one:OE>{_TABLE}</one:OE>") == [
        {"type": "unsupported", "kind": "table", "text": "Name | Age\nAda Lovelace | 36"},
    ]


def test_parse_ink_paragraph_reports_recognized_text():
    items = _items(
        "<one:OE><one:InkParagraph>"
        '<one:InkWord recognizedText="hello"><one:CallbackID callbackID="{1}"/></one:InkWord>'
        '<one:InkWord recognizedText="world"><one:CallbackID callbackID="{2}"/></one:InkWord>'
        "</one:InkParagraph></one:OE>"
    )
    assert items == [{"type": "unsupported", "kind": "ink", "text": "hello world"}]


def test_parse_text_mixed_with_ink_words_keeps_both_in_order():
    items = _items(
        "<one:OE><one:T>Meet</one:T>"
        '<one:InkWord recognizedText="Bob"><one:CallbackID callbackID="{1}"/></one:InkWord>'
        "<one:T><![CDATA[<span style='font-weight:bold'>today</span>]]></one:T>"
        '<one:InkWord><one:CallbackID callbackID="{2}"/></one:InkWord>'
        "</one:OE>"
    )
    assert items == [{"type": "unsupported", "kind": "ink", "text": "Meet Bob today"}]


@pytest.mark.parametrize("content, expected", [
    ('<one:InkDrawing><one:CallbackID callbackID="{1}"/></one:InkDrawing>',
     {"type": "unsupported", "kind": "ink"}),
    ('<one:InsertedFile pathCache="C:\\cache\\x.bin" preferredName="report.pdf"/>',
     {"type": "unsupported", "kind": "file", "text": "report.pdf"}),
    ('<one:MediaFile pathCache="C:\\cache\\y.bin" preferredName="Recording.wma"/>',
     {"type": "unsupported", "kind": "media", "text": "Recording.wma"}),
    ("<one:FutureObject/>", {"type": "unsupported", "kind": "unknown"}),
], ids=["ink-drawing", "inserted-file", "media-file", "future-object"])
def test_parse_other_oe_content_becomes_unsupported_item(content, expected):
    assert _items(f"<one:OE>{content}</one:OE>") == [expected]


def test_parse_unsupported_item_keeps_its_children():
    items = _items(
        f"<one:OE>{_TABLE}<one:OEChildren><one:OE><one:T>below the table</one:T></one:OE></one:OEChildren></one:OE>"
    )
    assert items[0]["children"] == [{"type": "paragraph", "text": "below the table"}]


def test_parse_bulleted_table_is_read_as_unsupported_item_between_lists():
    items = _items(
        f"<one:OE>{_BULLET}<one:T>a</one:T></one:OE>",
        f"<one:OE>{_BULLET}{_TABLE}</one:OE>",
        f"<one:OE>{_BULLET}<one:T>b</one:T></one:OE>",
    )
    assert [item["type"] for item in items] == ["list", "unsupported", "list"]
    assert items[1]["kind"] == "table"


def test_parse_unsupported_content_nested_under_list_item():
    items = _items(
        f"<one:OE>{_BULLET}<one:T>step</one:T><one:OEChildren><one:OE>{_TABLE}</one:OE></one:OEChildren></one:OE>"
    )
    assert items[0]["items"][0]["children"] == [
        {"type": "unsupported", "kind": "table", "text": "Name | Age\nAda Lovelace | 36"},
    ]


def test_parse_page_reports_page_level_ink_drawing():
    xml = _page_xml(
        "<one:InkDrawing>"
        '<one:Position x="100.0" y="200.0" z="4"/><one:Size width="50.0" height="25.0"/>'
        '<one:CallbackID callbackID="{1}"/>'
        "</one:InkDrawing>"
    )
    page = parse_page(xml)
    assert page.outlines == []
    assert page.model_dump(exclude_defaults=True)["unsupported"] == [
        {"kind": "ink", "position": {"x": 100.0, "y": 200.0, "z": 4}, "width": 50.0, "height": 25.0},
    ]


def test_parse_page_reports_page_level_inserted_file_by_name():
    xml = _page_xml('<one:InsertedFile pathCache="C:\\cache\\x.bin" preferredName="slides.pptx"/>')
    assert parse_page(xml).model_dump(exclude_defaults=True)["unsupported"] == [
        {"kind": "file", "text": "slides.pptx"},
    ]


def test_parse_page_without_unsupported_objects_omits_the_list():
    xml = _page_xml("<one:Outline><one:OEChildren><one:OE><one:T>x</one:T></one:OE></one:OEChildren></one:Outline>")
    assert "unsupported" not in parse_page(xml).model_dump(exclude_defaults=True)


_TABLE_ITEM = UnsupportedItem(type="unsupported", kind="table", text="a | b")


@pytest.mark.parametrize("outline", [
    Outline(items=[Paragraph(type="paragraph", text="kept"), _TABLE_ITEM]),
    Outline(items=[Paragraph(type="paragraph", text="p", children=[_TABLE_ITEM])]),
    Outline(items=[List(type="list", style="bullet", items=[ListItem(text="i", children=[_TABLE_ITEM])])]),
    Outline(items=[InlineImage(type="inline_image", handle="mcpref:abc123def456", children=[_TABLE_ITEM])]),
], ids=["outline", "paragraph-child", "list-item-child", "image-child"])
def test_build_page_xml_rejects_unsupported_items(outline):
    with pytest.raises(ValueError, match=r"'unsupported' \(kind 'table'\).*delete that content.*remove the item"):
        build_page_xml("Title", [outline])


def test_build_append_xml_rejects_unsupported_items():
    with pytest.raises(ValueError, match=r"'unsupported' \(kind 'ink'\)"):
        build_append_xml(Outline(items=[UnsupportedItem(type="unsupported", kind="ink")]))


def test_read_page_with_table_cannot_be_written_back_unchanged():
    page = parse_page(_outline_xml("<one:OE><one:T>intro</one:T></one:OE>", f"<one:OE>{_TABLE}</one:OE>"))
    with pytest.raises(ValueError, match="kind 'table'"):
        build_page_xml(page.title, page.outlines, page.images)


# ---------------------------------------------------------------------------
# Screenshot placeholders: a marked label OE, plus a box OE reserving the space
# ---------------------------------------------------------------------------

_LABEL_META = '<one:Meta name="onenote-mcp.image-placeholder" content="label"/>'
_BOX_META = '<one:Meta name="onenote-mcp.image-placeholder" content="box"/>'


def _label_oe(text: str) -> str:
    return f"<one:OE>{_LABEL_META}<one:T>{text}</one:T></one:OE>"


def _box_oe(alt: str = "Image placeholder", width: float = 300.0, height: float = 150.0) -> str:
    return (
        f'<one:OE>{_BOX_META}<one:Image alt="{alt}"><one:Size width="{width}" height="{height}" isSetByUser="true"/>'
        "<one:Data>mcpref:abc123def456</one:Data></one:Image></one:OE>"
    )


def _png_rows(png: bytes, width: int) -> list[bytes]:
    """Return the pixel rows of an RGB PNG written by _box_png (one IDAT chunk)."""
    length = int.from_bytes(png[33:37], "big")
    raw = zlib.decompress(png[41:41 + length])
    stride = 1 + 3 * width
    return [raw[start + 1:start + stride] for start in range(0, len(raw), stride)]


def _placeholder_oes(ph: ImagePlaceholder) -> list[ET.Element]:
    root = _parse_xml(build_page_xml("Doc", [Outline(items=[ph])]))
    return root.findall("one:Outline/one:OEChildren/one:OE", _NS)


def test_build_placeholder_marks_its_label_and_reserves_a_box_of_exactly_its_size():
    label, box = _placeholder_oes(
        ImagePlaceholder(type="image_placeholder", description="Screenshot des Dialogs", width=600.0, height=270.0)
    )
    assert label.find("one:Meta", _NS).attrib == {"name": "onenote-mcp.image-placeholder", "content": "label"}
    assert label.find("one:T", _NS).text == (
        '<span style="font-weight:bold;background:#FFEB3B">Screenshot des Dialogs</span>'
    )
    assert box.find("one:Meta", _NS).attrib == {"name": "onenote-mcp.image-placeholder", "content": "box"}
    image = box.find("one:Image", _NS)
    assert image.get("alt") == "Image placeholder"
    assert image.find("one:Size", _NS).attrib == {"width": "600.0", "height": "270.0", "isSetByUser": "true"}
    assert pixel_size(base64.b64decode(image.find("one:Data", _NS).text)) == (600, 270)


def test_build_placeholder_without_size_writes_the_label_only():
    (label,) = _placeholder_oes(ImagePlaceholder(type="image_placeholder", description="Diagramm"))
    assert label.find("one:Meta", _NS).get("content") == "label"
    assert label.find("one:Image", _NS) is None


def test_box_png_is_a_light_fill_inside_a_border():
    png = _box_png(10, 6)
    assert pixel_size(png) == (10, 6)
    rows = _png_rows(png, 10)
    border, fill = bytes((0xB0, 0x75, 0x00)), bytes((0xFF, 0xF9, 0xC4))
    assert len(rows) == 6
    assert rows[0] == border * 10
    assert rows[3] == border * 2 + fill * 6 + border * 2


@pytest.mark.parametrize("size, pixels", [
    ((600.0, 400.0), (600, 400)),
    ((2000.0, 500.0), (1000, 250)),
    ((300.0, 4000.0), (75, 1000)),
    ((0.2, 0.1), (1, 1)),
])
def test_box_pixel_size_keeps_the_aspect_and_caps_the_longer_side(size, pixels):
    assert _box_pixel_size(*size) == pixels


def test_parse_placeholder_label_and_box_read_back_as_one_item():
    assert _items(_label_oe("Screenshot here"), _box_oe(width=300.0, height=150.0)) == [
        {"type": "image_placeholder", "description": "Screenshot here", "width": 300.0, "height": 150.0},
    ]


@pytest.mark.parametrize("ph", [
    ImagePlaceholder(type="image_placeholder", description="Bildschirmfoto: Anmeldung", width=480.0, height=320.0),
    ImagePlaceholder(type="image_placeholder", description="label only"),
], ids=["with-box", "label-only"])
def test_placeholder_read_write_read_is_stable(ph):
    first = parse_page(build_page_xml("Doc", [Outline(items=[ph])]))
    assert first.outlines[0].items == [ph]
    assert _reread(first).outlines[0].items == [ph]


def test_parse_placeholder_is_recognized_by_its_meta_not_its_text():
    assert _items(
        _label_oe("anything at all"),
        "<one:OE><one:T>[INSERT IMAGE: an old-style callout]</one:T></one:OE>",
    ) == [
        {"type": "image_placeholder", "description": "anything at all"},
        {"type": "paragraph", "text": "[INSERT IMAGE: an old-style callout]"},
    ]


def test_parse_image_put_into_the_box_paragraph_is_not_the_placeholder():
    assert _items(_label_oe("Screenshot here"), _box_oe(alt="a pasted screenshot")) == [
        {"type": "image_placeholder", "description": "Screenshot here"},
        {"type": "inline_image", "handle": "mcpref:abc123def456", "width": 300.0, "height": 150.0},
    ]


def test_parse_box_without_its_label_reads_as_an_inline_image():
    assert _items(_box_oe()) == [
        {"type": "inline_image", "handle": "mcpref:abc123def456", "width": 300.0, "height": 150.0},
    ]


def test_parse_placeholder_label_with_indented_content_reads_as_a_paragraph():
    label = (
        f"<one:OE>{_LABEL_META}<one:T>Screenshot here</one:T>"
        "<one:OEChildren><one:OE><one:T>a note</one:T></one:OE></one:OEChildren></one:OE>"
    )
    assert [item["type"] for item in _items(label, _box_oe())] == ["paragraph", "inline_image"]


def test_placeholders_nested_under_paragraphs_and_list_items_round_trip():
    ph = ImagePlaceholder(type="image_placeholder", description="nested", width=200.0, height=100.0)
    outline = Outline(items=[
        Paragraph(type="paragraph", text="parent", children=[ph]),
        List(type="list", style="numbered", items=[ListItem(text="step", children=[ph])]),
    ])
    items = parse_page(build_page_xml("Doc", [outline])).outlines[0].items
    assert items[0].children == [ph]
    assert items[1].items[0].children == [ph]
