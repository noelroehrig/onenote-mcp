"""Unit tests for onenote_mcp.models — no COM dependency."""
import pytest
from pydantic import ValidationError
from onenote_mcp.models import (
    TextRun, Paragraph, ImagePlaceholder, InlineImage,
    ListItem, List, Position, Outline, PageContent,
)


# ---------------------------------------------------------------------------
# TextRun
# ---------------------------------------------------------------------------

def test_text_run_defaults():
    run = TextRun(text="hello")
    assert run.text == "hello"
    assert run.bold is False
    assert run.italic is False
    assert run.underline is False
    assert run.strikethrough is False
    assert run.color is None
    assert run.highlight is None
    assert run.font_size is None
    assert run.font_family is None


def test_text_run_all_fields():
    run = TextRun(
        text="styled",
        bold=True,
        italic=True,
        underline=True,
        strikethrough=True,
        color="#ff0000",
        highlight="#ffeb3b",
        font_size=14.0,
        font_family="Courier New",
    )
    assert run.text == "styled"
    assert run.bold is True
    assert run.italic is True
    assert run.underline is True
    assert run.strikethrough is True
    assert run.color == "#ff0000"
    assert run.highlight == "#ffeb3b"
    assert run.font_size == 14.0
    assert run.font_family == "Courier New"


# ---------------------------------------------------------------------------
# Paragraph
# ---------------------------------------------------------------------------

def test_paragraph_with_text():
    p = Paragraph(type="paragraph", text="hello", segments=None)
    assert p.text == "hello"
    assert p.segments is None


def test_paragraph_with_segments():
    p = Paragraph(type="paragraph", segments=[TextRun(text="hi")], text=None)
    assert p.segments is not None
    assert len(p.segments) == 1
    assert p.segments[0].text == "hi"
    assert p.text is None


def test_paragraph_rejects_both_text_and_segments():
    with pytest.raises(ValidationError, match="exactly one"):
        Paragraph(type="paragraph", text="hello", segments=[TextRun(text="hi")])


def test_paragraph_rejects_neither():
    with pytest.raises(ValidationError, match="exactly one"):
        Paragraph(type="paragraph", text=None, segments=None)


def test_paragraph_all_heading_styles():
    for style in ["normal", "h1", "h2", "h3", "h4", "h5", "h6"]:
        p = Paragraph(type="paragraph", text="heading", style=style)
        assert p.style == style


def test_paragraph_formatting_fields():
    p = Paragraph(
        type="paragraph",
        text="formatted",
        bold=True,
        italic=True,
        underline=True,
        strikethrough=True,
        color="#123456",
        highlight="#abcdef",
        font_size=12.0,
        font_family="Arial",
    )
    assert p.bold is True
    assert p.italic is True
    assert p.underline is True
    assert p.strikethrough is True
    assert p.color == "#123456"
    assert p.highlight == "#abcdef"
    assert p.font_size == 12.0
    assert p.font_family == "Arial"


def test_paragraph_type_literal():
    p = Paragraph(type="paragraph", text="x")
    assert p.type == "paragraph"


# ---------------------------------------------------------------------------
# ImagePlaceholder
# ---------------------------------------------------------------------------

def test_image_placeholder_minimal():
    ph = ImagePlaceholder(type="image_placeholder", description="Screenshot of login")
    assert ph.description == "Screenshot of login"
    assert ph.width is None
    assert ph.height is None


def test_image_placeholder_with_size():
    ph = ImagePlaceholder(type="image_placeholder", description="desc", width=300.0, height=150.0)
    assert ph.width == 300.0
    assert ph.height == 150.0


def test_image_placeholder_type_literal():
    ph = ImagePlaceholder(type="image_placeholder", description="x")
    assert ph.type == "image_placeholder"


# ---------------------------------------------------------------------------
# ListItem
# ---------------------------------------------------------------------------

def test_list_item_text():
    item = ListItem(text="item one")
    assert item.text == "item one"
    assert item.segments is None


def test_list_item_segments():
    item = ListItem(segments=[TextRun(text="bold item", bold=True)])
    assert item.segments is not None
    assert item.segments[0].bold is True
    assert item.text is None


def test_list_item_rejects_both():
    with pytest.raises(ValidationError, match="exactly one"):
        ListItem(text="x", segments=[TextRun(text="y")])


def test_list_item_allows_neither():
    # An empty list item is tolerated (LLMs sometimes emit a trailing {} entry);
    # the builder drops it rather than failing the whole page write.
    item = ListItem(text=None, segments=None)
    assert item.text is None
    assert item.segments is None


def test_list_item_children():
    child = ListItem(text="child item")
    parent = ListItem(text="parent item", children=[child])
    assert len(parent.children) == 1
    assert parent.children[0].text == "child item"


def test_list_item_deeply_nested():
    grandchild = ListItem(text="grandchild")
    child = ListItem(text="child", children=[grandchild])
    parent = ListItem(text="parent", children=[child])
    assert parent.children[0].children[0].text == "grandchild"


# ---------------------------------------------------------------------------
# List
# ---------------------------------------------------------------------------

def test_list_bullet():
    lst = List(type="list", style="bullet", items=[ListItem(text="a"), ListItem(text="b")])
    assert lst.style == "bullet"
    assert len(lst.items) == 2


def test_list_numbered():
    lst = List(type="list", style="numbered", items=[ListItem(text="one")])
    assert lst.style == "numbered"


def test_list_type_literal():
    lst = List(type="list", style="bullet", items=[ListItem(text="x")])
    assert lst.type == "list"


# ---------------------------------------------------------------------------
# Position
# ---------------------------------------------------------------------------

def test_position_valid():
    pos = Position(x=36.0, y=86.4, z=0)
    assert pos.x == 36.0
    assert pos.y == 86.4
    assert pos.z == 0


def test_position_zero():
    pos = Position(x=0, y=0)
    assert pos.x == 0.0
    assert pos.y == 0.0


def test_position_negative_x():
    with pytest.raises(ValidationError, match="'x' must be >= 0"):
        Position(x=-1.0, y=0)


def test_position_negative_y():
    with pytest.raises(ValidationError, match="'y' must be >= 0"):
        Position(x=0, y=-0.5)


def test_position_z_default():
    pos = Position(x=10.0, y=10.0)
    assert pos.z == 0


# ---------------------------------------------------------------------------
# Outline
# ---------------------------------------------------------------------------

def test_outline_no_position():
    outline = Outline(position=None, width=None, items=[])
    assert outline.position is None
    assert outline.width is None


def test_outline_with_position():
    pos = Position(x=36.0, y=86.4)
    outline = Outline(position=pos, items=[])
    assert outline.position is not None
    assert outline.position.x == 36.0


def test_outline_with_width():
    outline = Outline(position=None, width=250.0, items=[])
    assert outline.width == 250.0


# ---------------------------------------------------------------------------
# PageContent
# ---------------------------------------------------------------------------

def test_page_content():
    outline = Outline(items=[Paragraph(type="paragraph", text="body")])
    page = PageContent(title="My Page", outlines=[outline])
    assert page.title == "My Page"
    assert len(page.outlines) == 1


# ---------------------------------------------------------------------------
# InlineImage
# ---------------------------------------------------------------------------

def test_inline_image_validates_handle_prefix():
    """InlineImage rejects handles that do not start with 'mcpref:'."""
    with pytest.raises(ValidationError):
        InlineImage(type="inline_image", handle="not-a-handle")
