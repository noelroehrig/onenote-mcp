"""FastMCP server exposing the OneNote COM wrapper as MCP tools (stdio transport).

OneNote XML namespace: xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"

Minimal page skeleton:
    <one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="...">
      <one:Title>
        <one:OE><one:T>Page Title</one:T></one:OE>
      </one:Title>
      <one:Outline>
        <one:OEChildren>
          <one:OE><one:T>Body text here</one:T></one:OE>
          <one:OE>
            <one:Image>
              <one:Data>mcpref:a1b2c3d4e5f6</one:Data>
            </one:Image>
          </one:OE>
        </one:OEChildren>
      </one:Outline>
    </one:Page>

Images returned by get_page use short opaque handles inside <one:Image><one:Data>mcpref:...</one:Data></one:Image>.
Copy those handles verbatim into write XML — the server resolves them back to bytes before calling OneNote.
"""

import asyncio
import os
import sys
from collections.abc import Callable, Mapping

from mcp.server.fastmcp import FastMCP

from onenote_mcp.com import (
    OneNoteError,
    ping as _ping_com,
    validate_handles as _validate_handles_com,
    get_image_data as _get_image_data_com,
    list_hierarchy as _list_hierarchy_com,
    list_sections as _list_sections_com,
    list_pages as _list_pages_com,
    get_page as _get_page_com,
    create_page as _create_page_com,
    replace_page as _replace_page_com,
    append_page as _append_page_com,
)
from onenote_mcp.models import FloatingImage, Outline, PageContent
from onenote_mcp.builders import (
    build_page_xml, build_append_xml, parse_notebook_skeleton, parse_section_pages, parse_page,
)

# Tool functions in definition order; create_server registers them.
_TOOLS: list[tuple[Callable, bool]] = []


def _tool(raw_xml: bool = False) -> Callable[[Callable], Callable]:
    """Collect a tool function; raw_xml marks the *_xml escape hatches."""
    def register(fn: Callable) -> Callable:
        _TOOLS.append((fn, raw_xml))
        return fn
    return register


def _fail(exc: Exception) -> ValueError:
    """Normalize an exception into a ValueError whose message starts with a
    machine-readable code (timeout / backend_error / bad_request).

    OneNoteError already prefixes its code via __str__.  Plain ValueErrors come
    from request/model validation, so they map to ``bad_request`` unless they
    already carry a known code prefix.
    """
    msg = str(exc)
    if isinstance(exc, OneNoteError):
        return ValueError(msg)
    if msg.split(":", 1)[0] in {"timeout", "backend_error", "bad_request"}:
        return ValueError(msg)
    return ValueError(f"bad_request: {msg}")


@_tool()
async def ping() -> dict:
    """Fast health check — does NOT touch heavy page content.

    Returns quickly even while another tool call is stuck on a wedged OneNote,
    so it is the right way to tell "the MCP server is alive but OneNote is busy"
    apart from "the server is down".

    Returns:
      {"server": "ok", "onenote_responsive": true|false}

    onenote_responsive is false when OneNote does not answer a lightweight call
    within the ping timeout (it is likely showing a modal dialog or syncing).
    When false, page reads/writes will also time out until OneNote recovers —
    dismiss any open OneNote dialog and retry.
    """
    responsive = await asyncio.to_thread(_ping_com)
    return {"server": "ok", "onenote_responsive": responsive}


@_tool()
async def validate_handles(handles: list[str]) -> dict:
    """Check whether image handles can still be written back, without writing.

    handles: a list of 'mcpref:...' image handles from a previous get_page.

    Returns {handle: true|false} where true means the handle resolves to image
    bytes (cached or lazily fetchable) and is safe to pass to replace_page /
    append_page.  A false result means the handle is stale — re-read the source
    page with get_page to refresh it.  Fast: no page write is attempted.
    """
    return await asyncio.to_thread(_validate_handles_com, handles)


@_tool()
async def get_image_data(handle: str) -> str:
    """Fetch the raw base64 bytes for a single image handle, on demand.

    handle: an 'mcpref:...' handle from a previous get_page response.

    get_page reads pages WITHOUT image bytes by default, so use this when you
    actually need the pixels of one specific image (e.g. to inspect or transfer
    it) rather than pulling every image inline. Returns the image as a base64
    string. Raises a 'bad_request' error if the handle is unknown or stale —
    re-read the source page with get_page to refresh it.
    """
    try:
        return await asyncio.to_thread(_get_image_data_com, handle)
    except (OneNoteError, ValueError) as exc:
        raise _fail(exc) from exc


@_tool()
async def get_notebooks() -> list[dict]:
    """Return open notebooks and their sections (no pages) — the navigation skeleton.

    This is the first step for navigating to content. It is deliberately small:
    it returns every notebook with its sections but NO pages, so you can locate
    the right section cheaply even when notebooks hold hundreds of pages.

    Returns:
    [
      {
        "id": "<notebook-id>",
        "name": "Notebook Name",
        "sections": [
          {"id": "<section-id>", "name": "Section Name"}
        ]
      }
    ]

    Workflow:
      1. get_notebooks() — find the section you want (e.g. "Vorlagen") and its id.
      2. list_pages(section_id=<id>) — list just that section's pages.
      3. get_page / replace_page / append_page with a page id from step 2.
    Use a section id with create_page; use a page id with get_page/replace_page.

    Note: if the server is configured with a notebook allowlist
    (ONENOTE_ALLOWED_NOTEBOOKS), restricted notebooks are simply absent here
    and their content is not accessible through any tool.
    """
    try:
        xml = await asyncio.to_thread(_list_sections_com)
        return parse_notebook_skeleton(xml)
    except (OneNoteError, ValueError) as exc:
        raise _fail(exc) from exc


@_tool()
async def list_pages(section_id: str) -> list[dict]:
    """Return the pages of a single section as [{"id": ..., "name": ...}].

    section_id: a section id from get_notebooks().

    Returns only that section's pages — not the whole notebook — so it stays
    small regardless of how many other sections exist. Pass a returned page id
    to get_page, replace_page, or append_page.
    """
    try:
        xml = await asyncio.to_thread(_list_pages_com, section_id)
        return parse_section_pages(xml)
    except (OneNoteError, ValueError) as exc:
        raise _fail(exc) from exc


@_tool(raw_xml=True)
async def list_hierarchy_xml() -> str:
    """Advanced / escape-hatch only. Prefer get_notebooks + list_pages for navigation.

    Return the ENTIRE OneNote notebook/section/page tree as raw XML in one call.
    This can be very large (every page in every notebook); use the structured
    get_notebooks() + list_pages() drill-down instead unless you specifically
    need the full tree as XML.

    The returned string is a complete OneNote XML document.
    xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"
    Top-level element is <one:Notebooks>.
    """
    try:
        return await asyncio.to_thread(_list_hierarchy_com)
    except OneNoteError as exc:
        raise _fail(exc) from exc


@_tool()
async def get_page(page_id: str, include_binary: bool = False) -> dict:
    """Return a OneNote page as structured JSON.

    Prefer this over get_page_xml unless you need raw XML.

    page_id: the id from list_pages.
    include_binary: when False (default), the page is read WITHOUT image bytes —
                    fast and small even for image-heavy pages. Images still
                    appear with their 'mcpref:' handles, width/height and
                    position; the actual bytes are fetched lazily only if you
                    write a handle back with replace_page/append_page. Set True
                    only when you specifically need the raw image data embedded
                    (larger payload, slower on image-heavy pages).

    Returns a PageContent-shaped dict:
    {
      "title": "Page Title",
      "outlines": [
        {
          "position": {"x": 36.0, "y": 86.4},
          "width": 500.0,
          "items": [
            {"type": "paragraph", "text": "Plain body text"},
            {"type": "paragraph", "text": "A heading", "style": "h2", "bold": true},
            {"type": "list", "style": "bullet", "items": [{"text": "Item 1"}]},
            {"type": "inline_image", "handle": "mcpref:abc...", "width": 300.0, "height": 200.0}
          ]
        }
      ],
      "images": [
        {"handle": "mcpref:abc123", "width": 300.0, "height": 200.0,
         "position": {"x": 36.0, "y": 156.0, "z": 1}}
      ]
    }

    COMPACT OUTPUT: fields at their default values are OMITTED. An absent field
    always means the default from the replace_page schema — style="normal",
    bold/italic/underline/strikethrough=false, color/highlight/font_size/
    font_family unset, position/width unset (auto-placed), no children, no
    floating images. So a plain paragraph is just {"type": "paragraph",
    "text": "..."} and any styling you DO see is really on the page.

    Floating images, when present, appear in the top-level `images` list.
    Pass content["images"] verbatim to replace_page's `images` parameter to
    preserve them when rewriting a page (omit it if absent).

    The returned structure is directly compatible with replace_page — you can read a
    page, modify its outlines, and pass outlines + images straight to replace_page.
    """
    try:
        xml = await asyncio.to_thread(_get_page_com, page_id, include_binary)
        return parse_page(xml).model_dump(exclude_defaults=True)
    except (OneNoteError, ValueError) as exc:
        raise _fail(exc) from exc


@_tool(raw_xml=True)
async def get_page_xml(page_id: str) -> str:
    """Advanced / escape-hatch only. Prefer get_page (structured JSON) instead.

    Return the full XML of a OneNote page, with images as mcpref handles.

    page_id: the id from list_pages.

    Returns a complete <one:Page> document.
    xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"

    Images appear as <one:Image><one:Data>mcpref:xxxxxxxxxxxx</one:Data></one:Image>.
    Copy mcpref handles verbatim into write XML — the server resolves them to bytes.
    """
    try:
        return await asyncio.to_thread(_get_page_com, page_id)
    except OneNoteError as exc:
        raise _fail(exc) from exc


@_tool()
async def create_page(section_id: str, title: str, parent_page_id: str | None = None) -> str:
    """Create a new page in the given section and return its page ID.

    section_id: ID attribute of a <one:Section> from get_notebooks.
    title: required display title for the new page.
    parent_page_id: optional. When given, the new page becomes the LAST SUB-PAGE
                    of that page: it is placed after the parent's existing
                    sub-pages and indented one level below the parent. The
                    parent must be a page id in the same section. Omit for a
                    normal top-level page appended at the end of the section.

    Returns the new page's ID string. After creation, call get_page with
    the returned ID to retrieve the page structure, then replace_page to
    write content.
    """
    try:
        return await asyncio.to_thread(_create_page_com, section_id, title, parent_page_id)
    except (OneNoteError, ValueError) as exc:
        raise _fail(exc) from exc


@_tool(raw_xml=True)
async def replace_page_xml(page_id: str, page_xml: str) -> str:
    """Advanced / escape-hatch only. Prefer replace_page (typed slim tool) unless you
    need direct XML control or are debugging a OneNote schema issue.

    Replace the full content of a OneNote page.

    page_id: ID attribute of the target page (from list_pages or create_page).
    page_xml: a COMPLETE <one:Page> document — not a fragment. The ID attribute
              on the root element is overwritten automatically to match page_id.

    Returns "ok" on success.

    xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"

    Minimal page skeleton:
        <one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="...">
          <one:Title><one:OE><one:T>Title text</one:T></one:OE></one:Title>
          <one:Outline><one:OEChildren>
            <one:OE><one:T>Body paragraph</one:T></one:OE>
            <one:OE><one:Image><one:Data>mcpref:a1b2c3d4e5f6</one:Data></one:Image></one:OE>
          </one:OEChildren></one:Outline>
        </one:Page>

    IMAGE HANDLING
    <one:Data> may contain either a mcpref handle (e.g. 'mcpref:a1b2c3d4e5f6')
    copied verbatim from a get_page response, or raw base64 image bytes.
    Handles are resolved to their bytes before writing to OneNote; raw base64
    is passed through unchanged. Prefer handles: they keep requests small.

    If a mcpref handle is not recognized (e.g. from a stale session), the call
    fails with an error. Fix: re-read the source page with get_page to refresh
    the handles, then retry.

    MANUAL-INSERT PLACEHOLDER
    When the user has asked for a screenshot that does not yet exist in OneNote,
    do NOT invent a mcpref handle or image bytes. Instead emit a
    highlighted callout positioned and sized to match where the screenshot will go:

        <one:Outline>
          <one:Position x="36.0" y="86.4" z="0"/>
          <one:Size width="500.0" height="400.0"/>
          <one:OEChildren>
            <one:OE alignment="left">
              <one:T><![CDATA[<span style="background:#ffeb3b;color:#222;font-weight:bold;padding:4px 8px;border:2px dashed #b07500;">[INSERT IMAGE: <short description of what goes here>]</span>]]></one:T>
            </one:OE>
          </one:OEChildren>
        </one:Outline>

    Position: read the existing page XML (get_page) to find where surrounding
    content sits, then set x/y so the placeholder lands exactly where the
    screenshot will be pasted. OneNote coordinates are in points (1 pt = 1/72 in).
    Size: set width/height to match the expected screenshot dimensions so the
    placeholder occupies the same space the final image will fill.
    Description: [INSERT IMAGE: ...] must name the expected content
    (book page, figure number, context) so the user knows which screenshot to paste.
    """
    try:
        await asyncio.to_thread(_replace_page_com, page_id, page_xml)
        return "ok"
    except OneNoteError as exc:
        raise _fail(exc) from exc


@_tool(raw_xml=True)
async def append_page_xml(page_id: str, content_xml: str) -> str:
    """Advanced / escape-hatch only. Prefer append_page (typed slim tool) unless you
    need direct XML control or are debugging a OneNote schema issue.

    Append a single content element to an existing OneNote page.

    page_id: ID attribute of the target page (from list_pages or create_page).
    content_xml: a SINGLE formally correct OneNote XML element — NOT a full
                 <one:Page> document. The xmlns:one declaration MUST appear on
                 the element itself.

    Valid top-level elements: <one:Outline>, <one:Image>, <one:Table>, etc.

    Example — add one paragraph to a page:

        <one:Outline xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
          <one:OEChildren>
            <one:OE><one:T>New paragraph text</one:T></one:OE>
          </one:OEChildren>
        </one:Outline>

    IMAGE HANDLING
    <one:Data> may contain either a mcpref handle (e.g. 'mcpref:a1b2c3d4e5f6')
    copied verbatim from a get_page response, or raw base64 image bytes.
    Handles are resolved to their bytes before writing to OneNote; raw base64
    is passed through unchanged. Prefer handles: they keep requests small.

    Returns "ok" on success.
    """
    try:
        await asyncio.to_thread(_append_page_com, page_id, content_xml)
        return "ok"
    except OneNoteError as exc:
        raise _fail(exc) from exc


@_tool()
async def replace_page(page_id: str, title: str, outlines: list[Outline], images: list[FloatingImage] | None = None) -> str:
    """Replace the full content of a OneNote page using structured typed content.

    Prefer this over replace_page_xml unless you need direct XML control.

    page_id: the page ID from list_pages or create_page.
    title:   new page title — required and replaces the existing title.
    outlines: list of Outline objects that define the page canvas layout.
              Each Outline is an independently positioned block. Items within
              an Outline stack vertically and may be:
                - {"type": "paragraph", "text": "...", "style": "normal|h1-h6", ...}
                - {"type": "list", "style": "bullet|numbered", "items": [...]}
                - {"type": "image_placeholder", "description": "what to paste here",
                   "width": ..., "height": ...}
                  → generates a yellow highlight callout; the user pastes the screenshot
                    on top of it after viewing the page.
                - {"type": "inline_image", "handle": "mcpref:...", "width": ..., "height": ...}
                  → an image embedded inside the outline (e.g. a pasted screenshot).
                    Use the handle verbatim from a get_page response.
    images: optional list of FloatingImage objects positioned directly on the page canvas.
            These are peers of outlines, NOT nested inside them. Each image must have a
            handle from get_page (starts with 'mcpref:'). Pass content.images from a
            get_page call verbatim to preserve images when rewriting a page. Omit or pass
            None if the page has no floating images.

    Returns "ok" on success.
    Note: all existing page content is deleted before the new content is written.
    """
    try:
        page_xml = build_page_xml(title, outlines, images)
        await asyncio.to_thread(_replace_page_com, page_id, page_xml)
        return "ok"
    except (OneNoteError, ValueError) as exc:
        raise _fail(exc) from exc


@_tool()
async def append_page(page_id: str, outline: Outline) -> str:
    """Append a structured content block to an existing OneNote page.

    Prefer this over append_page_xml unless you need direct XML control.

    page_id: the page ID from list_pages or create_page.
    outline: a single Outline object containing the content to append.
             Existing page content is preserved.
             Outline items may be paragraph, list, image, or image_placeholder
             — see replace_page for the full item schema.

    Returns "ok" on success.
    """
    try:
        page_xml = build_append_xml(outline)
        await asyncio.to_thread(_append_page_com, page_id, page_xml)
        return "ok"
    except (OneNoteError, ValueError) as exc:
        raise _fail(exc) from exc


def _raw_xml_disabled(environ: Mapping[str, str]) -> bool:
    """Parse ONENOTE_DISABLE_RAW_XML: '1'/'true' disable the raw-XML tools,
    unset, empty, '0' or 'false' keep them (case-insensitive)."""
    raw = environ.get("ONENOTE_DISABLE_RAW_XML", "")
    value = raw.strip().lower()
    if value in ("1", "true"):
        return True
    if value in ("", "0", "false"):
        return False
    raise ValueError(f"ONENOTE_DISABLE_RAW_XML must be 1, true, 0, false or empty, got {raw!r}")


def create_server(raw_xml_tools: bool = True) -> FastMCP:
    """Build the MCP server; with raw_xml_tools=False the four *_xml tools are not registered."""
    server = FastMCP("onenote")
    for fn, is_raw_xml in _TOOLS:
        if raw_xml_tools or not is_raw_xml:
            server.add_tool(fn)
    return server


def main() -> None:
    """Entry point for the ``onenote-mcp`` console script (stdio transport)."""
    try:
        raw_xml_disabled = _raw_xml_disabled(os.environ)
    except ValueError as exc:
        sys.exit(f"onenote-mcp: {exc}")
    create_server(raw_xml_tools=not raw_xml_disabled).run()


if __name__ == "__main__":
    main()
