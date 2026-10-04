"""Pydantic v2 models for the OneNote slim MCP API.

These models define the JSON schema that the slim MCP tools accept and return.
FastMCP exposes these types as machine-readable JSON schema to the model.
"""

from __future__ import annotations
from typing import Annotated, Literal, Union
from pydantic import BaseModel, Field, StringConstraints, field_validator, model_validator

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

    @model_validator(mode="after")
    def _check_text_or_segments(self) -> Paragraph:
        if self.text is not None and self.segments is not None:
            raise ValueError("Provide exactly one of 'text' or 'segments', not both.")
        if self.text is None and self.segments is None:
            raise ValueError("Provide exactly one of 'text' or 'segments'; neither was set.")
        return self


class ImagePlaceholder(BaseModel):
    """A highlighted callout box for a screenshot the user will paste manually."""

    type: Literal["image_placeholder"]
    description: str = Field(description="What the user should paste here, e.g. 'Screenshot of the login screen'. Shown inside the yellow highlight box.")
    width: float | None = Field(None, description="Expected screenshot width in points. Used to size the placeholder to match the final image.")
    height: float | None = Field(None, description="Expected screenshot height in points. Used to size the placeholder to match the final image.")


class InlineImage(BaseModel):
    """An image embedded inside an outline (e.g. a pasted screenshot)."""

    type: Literal["inline_image"]
    handle: str = Field(description="Image handle from a get_page response, e.g. 'mcpref:a1b2c3d4e5f6'. Must start with 'mcpref:'.")
    width: float | None = Field(None, description="Display width in points. Give only width or only height to scale proportionally; omit both for the natural size.")
    height: float | None = Field(None, description="Display height in points. Give only width or only height to scale proportionally; omit both for the natural size.")

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
    children: list[ListItem] = Field(default_factory=list, description="Nested sub-items for indented list levels.")

    @model_validator(mode="after")
    def _check_text_or_segments(self) -> ListItem:
        if self.text is not None and self.segments is not None:
            raise ValueError("Provide exactly one of 'text' or 'segments', not both.")
        # An item with neither text nor segments is tolerated (LLMs sometimes emit
        # a trailing empty {} entry in a list).  The builder drops such empty items
        # rather than failing the whole page write.
        return self


ListItem.model_rebuild()


class List(BaseModel):
    """A bullet or numbered list."""

    type: Literal["list"]
    style: Literal["bullet", "numbered"] = Field(description="'bullet' for an unordered list (•), 'numbered' for an ordered list (1. 2. 3.).")
    items: list[ListItem] = Field(description="List items in order, top to bottom.")


# Tables are not part of the slim model — write them with replace_page_xml /
# append_page_xml using raw OneNote XML.

ContentItem = Annotated[
    Union[Paragraph, ImagePlaceholder, InlineImage, List],
    Field(discriminator="type"),
]


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


class Outline(BaseModel):
    """A positioned block of content on the page canvas."""

    position: Position | None = Field(None, description="Canvas position. Omit to let OneNote place the outline automatically.")
    width: float | None = Field(None, description="Column width in points. Omit to let OneNote auto-size to content.")
    items: list[ContentItem] = Field(description="Content items stacked vertically inside this outline block.")


class PageContent(BaseModel):
    """Full structured page representation (returned by get_page)."""

    title: str = Field(description="The page title.")
    outlines: list[Outline] = Field(description="All content blocks on the page canvas.")
    images: list[FloatingImage] = Field(
        default_factory=list,
        description=(
            "Floating images positioned directly on the page canvas. "
            "These are peers of outlines — not nested inside them. "
            "Pass this list unchanged to replace_page to preserve images when rewriting a page."
        ),
    )


