"""Pydantic v2 models for the OneNote slim MCP API.

These models define the JSON schema that the slim MCP tools accept and return.
FastMCP exposes these types as machine-readable JSON schema to the model.
"""

from __future__ import annotations
from typing import Annotated, Any, Literal, Union
from pydantic import (
    BaseModel, Discriminator, Field, StringConstraints, Tag, field_validator, model_validator,
)

# Style values are written into an HTML style="..." attribute, so they are
# restricted to characters that cannot end the declaration or the attribute.
HEX_COLOR_PATTERN = r"^#(?:[0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$"
FONT_FAMILY_PATTERN = r"^[A-Za-z0-9 ,.\-]+$"
FONT_FAMILY_MAX_LENGTH = 64

HexColor = Annotated[str, StringConstraints(pattern=HEX_COLOR_PATTERN)]
FontFamily = Annotated[
    str, StringConstraints(pattern=FONT_FAMILY_PATTERN, max_length=FONT_FAMILY_MAX_LENGTH)
]


class TextRun(BaseModel):
    """A single formatted text span."""

    text: str = Field(description="The text content of this span.")
    bold: bool = Field(False, description="Bold text.")
    italic: bool = Field(False, description="Italic text.")
    underline: bool = Field(False, description="Underlined text.")
    strikethrough: bool = Field(False, description="Strikethrough text.")
    color: HexColor | None = Field(None, description="Text colour as '#RGB' or '#RRGGBB', e.g. '#ff0000'. None means default.")
    highlight: HexColor | None = Field(None, description="Highlight colour as '#RGB' or '#RRGGBB', e.g. '#ffeb3b'. None means no highlight.")
    font_size: float | None = Field(None, description="Font size in points. None means default.")
    font_family: FontFamily | None = Field(None, description="Font family name, e.g. 'Courier New': letters, digits, spaces, hyphens, commas and periods, at most 64 characters. None means default.")


class Paragraph(BaseModel):
    """A text paragraph."""

    type: Literal["paragraph"]
    text: str | None = Field(None, description="Plain text for a uniformly styled paragraph. Mutually exclusive with 'segments' — provide exactly one.")
    segments: list[TextRun] | None = Field(None, description="Formatted text runs for mixed inline formatting. Mutually exclusive with 'text' — provide exactly one. Whole-paragraph formatting fields (bold, color, etc.) are ignored when segments is set.")
    style: Literal["normal", "h1", "h2", "h3", "h4", "h5", "h6"] = Field("normal", description="Heading style. 'normal' for body text, 'h1'–'h6' for headings.")
    bold: bool = Field(False, description="Bold. Applied to the whole paragraph when 'text' is used; ignored when 'segments' is used.")
    italic: bool = Field(False, description="Italic. Applied to whole paragraph when 'text' is used; ignored when 'segments' is used.")
    underline: bool = Field(False, description="Underline. Applied to whole paragraph when 'text' is used; ignored when 'segments' is used.")
    strikethrough: bool = Field(False, description="Strikethrough. Applied to whole paragraph when 'text' is used; ignored when 'segments' is used.")
    color: HexColor | None = Field(None, description="Text colour as '#RGB' or '#RRGGBB'. Applied to whole paragraph when 'text' is used; ignored when 'segments' is used.")
    highlight: HexColor | None = Field(None, description="Highlight colour as '#RGB' or '#RRGGBB'. Applied to whole paragraph when 'text' is used; ignored when 'segments' is used.")
    font_size: float | None = Field(None, description="Font size in points. Applied to whole paragraph when 'text' is used; ignored when 'segments' is used.")
    font_family: FontFamily | None = Field(None, description="Font family name (letters, digits, spaces, hyphens, commas and periods, at most 64 characters). Applied to whole paragraph when 'text' is used; ignored when 'segments' is used.")
    children: list[ContentItem] = Field(default_factory=list, description="Content indented one level under this paragraph, in order.")

    @model_validator(mode="after")
    def _check_text_or_segments(self) -> Paragraph:
        if self.text is not None and self.segments is not None:
            raise ValueError("Provide exactly one of 'text' or 'segments', not both.")
        if self.text is None and self.segments is None:
            raise ValueError("Provide exactly one of 'text' or 'segments'; neither was set.")
        return self


class ImagePlaceholder(BaseModel):
    """A screenshot placeholder: a highlighted label above a light box reserving the screenshot's space."""

    type: Literal["image_placeholder"]
    description: str = Field(description="What the user should paste here, in the user's language, e.g. 'Screenshot of the login screen'. Shown verbatim as the label.")
    width: float | None = Field(None, gt=0, description="Box width in points. Give both width and height, or neither for the label alone.")
    height: float | None = Field(None, gt=0, description="Box height in points. Give both width and height, or neither for the label alone.")

    @model_validator(mode="after")
    def _check_both_dimensions_or_neither(self) -> ImagePlaceholder:
        if (self.width is None) != (self.height is None):
            raise ValueError("Give both 'width' and 'height' to reserve the screenshot's space, or neither for the label alone.")
        return self


class InlineImage(BaseModel):
    """An image embedded inside an outline (e.g. a pasted screenshot)."""

    type: Literal["inline_image"]
    handle: str = Field(description="Image handle from a get_page response, e.g. 'mcpref:a1b2c3d4e5f6'. Must start with 'mcpref:'.")
    width: float | None = Field(None, description="Display width in points. Give only width or only height to scale proportionally; omit both for the natural size.")
    height: float | None = Field(None, description="Display height in points. Give only width or only height to scale proportionally; omit both for the natural size.")
    children: list[ContentItem] = Field(default_factory=list, description="Content indented one level under this image, in order.")

    @field_validator("handle")
    @classmethod
    def _handle_must_start_with_mcpref(cls, value: str) -> str:
        if not value.startswith("mcpref:"):
            raise ValueError("'handle' must start with 'mcpref:'.")
        return value


class ListItem(BaseModel):
    """A single list item (supports nesting via children)."""

    text: str | None = Field(None, description="Plain text for this list item. Mutually exclusive with 'segments' — provide exactly one.")
    segments: list[TextRun] | None = Field(None, description="Formatted text runs for this item. Mutually exclusive with 'text' — provide exactly one.")
    children: list[ListChild] = Field(default_factory=list, description="Content indented under this item, in order: sub-items (objects without 'type') for nested list levels, or any outline item with a 'type' (paragraph, inline_image, ...) for other indented content.")

    @model_validator(mode="after")
    def _check_text_or_segments(self) -> ListItem:
        if self.text is not None and self.segments is not None:
            raise ValueError("Provide exactly one of 'text' or 'segments', not both.")
        # An item with neither text nor segments is tolerated (LLMs sometimes emit
        # a trailing empty {} entry in a list).  The builder drops such empty items
        # rather than failing the whole page write.
        return self


class List(BaseModel):
    """A bullet or numbered list."""

    type: Literal["list"]
    style: Literal["bullet", "numbered"] = Field(description="'bullet' for an unordered list (•), 'numbered' for an ordered list (1. 2. 3.).")
    items: list[ListItem] = Field(description="List items in order, top to bottom.")


# Tables are not part of the slim model — write them with replace_page_xml /
# append_page_xml using raw OneNote XML.  On read they appear as UnsupportedItem.

UnsupportedKind = Literal["table", "ink", "file", "media", "unknown"]
_UNSUPPORTED_KIND_DESCRIPTION = (
    "What the content is: 'table'; 'ink' (handwriting or a drawing, including typed text "
    "that shares a line with handwriting); 'file' (an attached file); 'media' (an audio or "
    "video recording); 'unknown' (content from a newer OneNote version)."
)
_UNSUPPORTED_TEXT_DESCRIPTION = (
    "Readable plain text of the content when OneNote provides it: a table's cells (one line "
    "per row, cells separated by ' | '), the recognized text of handwriting, or a file's name."
)


class UnsupportedItem(BaseModel):
    """Read-only marker for outline content the structured tools cannot write.

    replace_page and append_page reject a payload that still contains one:
    rewriting the page would delete that content, so remove the item
    deliberately only if deleting it is intended.
    """

    type: Literal["unsupported"]
    kind: UnsupportedKind = Field(description=_UNSUPPORTED_KIND_DESCRIPTION)
    text: str = Field("", description=_UNSUPPORTED_TEXT_DESCRIPTION)
    children: list[ContentItem] = Field(default_factory=list, description="Content indented one level under this item, in order.")


ContentItem = Annotated[
    Union[Paragraph, ImagePlaceholder, InlineImage, List, UnsupportedItem],
    Field(discriminator="type"),
]


def _list_child_tag(value: Any) -> str:
    """Tell a list sub-item (no 'type') from any other content item."""
    has_type = "type" in value if isinstance(value, dict) else hasattr(value, "type")
    return "content" if has_type else "item"


ListChild = Annotated[
    Union[Annotated[ListItem, Tag("item")], Annotated[ContentItem, Tag("content")]],
    Discriminator(_list_child_tag),
]

Paragraph.model_rebuild()
InlineImage.model_rebuild()
ListItem.model_rebuild()
UnsupportedItem.model_rebuild()


class Position(BaseModel):
    """Canvas coordinates for positioning an outline block."""

    x: float = Field(description="Horizontal offset from the left edge of the page in points (1 pt = 1/72 in).")
    y: float = Field(description="Vertical offset from the top of the page in points (1 pt = 1/72 in).")
    z: int = Field(0, description="Stacking order. Higher values appear on top of lower values.")

    @model_validator(mode="after")
    def _check_non_negative(self) -> Position:
        if self.x < 0:
            raise ValueError("'x' must be >= 0.")
        if self.y < 0:
            raise ValueError("'y' must be >= 0.")
        return self


class FloatingImage(BaseModel):
    """A floating image positioned directly on the page canvas (not inside an outline)."""

    handle: str = Field(description="Image handle from a get_page response, e.g. 'mcpref:a1b2c3d4e5f6'. Must start with 'mcpref:'.")
    width: float | None = Field(None, description="Display width in points. Give only width or only height to scale proportionally; omit both for the natural size.")
    height: float | None = Field(None, description="Display height in points. Give only width or only height to scale proportionally; omit both for the natural size.")
    position: Position | None = Field(None, description="Canvas position. Omit to let OneNote place the image automatically.")

    @field_validator("handle")
    @classmethod
    def _handle_must_start_with_mcpref(cls, value: str) -> str:
        if not value.startswith("mcpref:"):
            raise ValueError("'handle' must start with 'mcpref:'.")
        return value


class UnsupportedPageObject(BaseModel):
    """Read-only report of an object on the page canvas the structured tools cannot write."""

    kind: UnsupportedKind = Field(description=_UNSUPPORTED_KIND_DESCRIPTION)
    text: str = Field("", description=_UNSUPPORTED_TEXT_DESCRIPTION)
    position: Position | None = Field(None, description="Canvas position, when OneNote reports one.")
    width: float | None = Field(None, description="Width in points, when OneNote reports one.")
    height: float | None = Field(None, description="Height in points, when OneNote reports one.")


class Outline(BaseModel):
    """A positioned block of content on the page canvas."""

    position: Position | None = Field(None, description="Canvas position. Omit to let OneNote place the outline automatically.")
    width: float | None = Field(None, description="Column width in points. Omit to let OneNote auto-size to content.")
    items: list[ContentItem] = Field(description="Content items stacked vertically inside this outline block.")


class PageOutline(Outline):
    """An outline as get_page reports it.

    Only the read side has ``height``, so the write schemas stay unchanged; a
    write payload that still carries it validates as an Outline without it.
    """

    height: float | None = Field(None, description="Read-only: the outline's height in points as measured by OneNote. replace_page and append_page ignore it.")


class PageContent(BaseModel):
    """Full structured page representation (returned by get_page)."""

    title: str = Field(description="The page title.")
    outlines: list[PageOutline] = Field(description="All content blocks on the page canvas.")
    images: list[FloatingImage] = Field(
        default_factory=list,
        description=(
            "Floating images positioned directly on the page canvas. "
            "These are peers of outlines — not nested inside them. "
            "Pass this list unchanged to replace_page to preserve images when rewriting a page."
        ),
    )
    unsupported: list[UnsupportedPageObject] = Field(
        default_factory=list,
        description=(
            "Read-only: objects on the page canvas the structured tools cannot write, such as "
            "handwriting and drawings. replace_page and append_page keep them in place at their "
            "position, so place new content where it does not overlap them."
        ),
    )


