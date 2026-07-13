"""E2E tests for inline image round-trip through the slim MCP path.

These tests go through server.py via the MCP stdio transport (mcp_session fixture).
They require OneNote desktop to be running; if it is not available the suite is
skipped automatically via the onenote_available fixture chain.
"""
import pytest

from tests.e2e.mcp_client import mcp_session  # noqa: F401 — re-export for fixture


# A minimal 1×1 transparent PNG encoded as base64.
# Small enough to keep the test source readable; large enough to be a valid image.
_TINY_PNG_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYA"
    "AjCB0C8AAAAASUVORK5CYII="
)

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_XMLNS = f'xmlns:one="{_ONE_NS}"'

# Image dimensions we embed in the setup XML.
_IMG_WIDTH = 100.0
_IMG_HEIGHT = 80.0


def _find_inline_image(content_dict: dict) -> dict | None:
    """Return the first inline_image item found in a get_page result, or None."""
    for outline in content_dict["outlines"]:
        for item in outline.get("items", []):
            if item.get("type") == "inline_image":
                return item
    return None


def _make_page_xml(title: str, img_data: str, width: float, height: float) -> str:
    """Return a complete <one:Page> XML string with one outline containing one inline image."""
    return (
        f'<one:Page {_XMLNS}>'
        f'  <one:Title><one:OE><one:T>{title}</one:T></one:OE></one:Title>'
        f'  <one:Outline>'
        f'    <one:OEChildren>'
        f'      <one:OE>'
        f'        <one:Image>'
        f'          <one:Size width="{width}" height="{height}" isSetByUser="true"/>'
        f'          <one:Data>{img_data}</one:Data>'
        f'        </one:Image>'
        f'      </one:OE>'
        f'    </one:OEChildren>'
        f'  </one:Outline>'
        f'</one:Page>'
    )


@pytest.mark.e2e
def test_inline_image_round_trip(mcp_session, claudespike_section_id):
    """Inline image survives get_page → replace_page round-trip through the slim path.

    Setup: create a page with a known inline image via replace_page_xml (XML escape-hatch).
    Step 1: read back via slim get_page; assert InlineImage is present with correct shape.
    Step 2: feed outlines from step 1 straight back to replace_page.
    Step 3: read back again; assert the image still survives.
    """
    # --- Setup: create a fresh page -----------------------------------------
    title = "test_inline_image_round_trip_auto"
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": title,
    })
    assert isinstance(page_id, str) and page_id, "create_page must return a non-empty page ID"

    # Write a known inline image via XML escape-hatch so we don't depend on
    # any pre-existing fixture page.
    page_xml = _make_page_xml(title, _TINY_PNG_BASE64, _IMG_WIDTH, _IMG_HEIGHT)
    result = mcp_session.call_tool("replace_page_xml", {
        "page_id": page_id,
        "page_xml": page_xml,
    })
    assert result == "ok", f"replace_page_xml setup failed: {result}"

    # --- Step 1: read back via slim get_page --------------------------------
    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    assert "outlines" in content, "get_page must return an 'outlines' key"
    assert content["outlines"], "Expected at least one outline"

    img_item = _find_inline_image(content)
    assert img_item is not None, (
        "Expected at least one inline_image item in outlines after reading the page. "
        f"Got outlines: {content['outlines']}"
    )
    assert img_item["handle"].startswith("mcpref:"), (
        f"inline_image handle must start with 'mcpref:', got: {img_item['handle']!r}"
    )
    assert isinstance(img_item.get("width"), (int, float)), (
        f"inline_image width must be numeric, got: {img_item.get('width')!r}"
    )
    assert isinstance(img_item.get("height"), (int, float)), (
        f"inline_image height must be numeric, got: {img_item.get('height')!r}"
    )
    # OneNote may quantize dimensions by a sub-pixel amount; use a 1% tolerance.
    assert abs(img_item["width"] - _IMG_WIDTH) < _IMG_WIDTH * 0.01, (
        f"Expected width≈{_IMG_WIDTH}, got {img_item['width']}"
    )
    assert abs(img_item["height"] - _IMG_HEIGHT) < _IMG_HEIGHT * 0.01, (
        f"Expected height≈{_IMG_HEIGHT}, got {img_item['height']}"
    )

    # --- Step 2: write the slim outlines back via replace_page ---------------
    write_result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": title,
        "outlines": content["outlines"],
        "images": content.get("images", []),
    })
    assert write_result == "ok", f"replace_page round-trip write failed: {write_result}"

    # --- Step 3: read again; image must still be present --------------------
    content2 = mcp_session.call_tool("get_page", {"page_id": page_id})
    img_item2 = _find_inline_image(content2)
    assert img_item2 is not None, (
        "Inline image was lost after slim-path replace_page round-trip. "
        f"Outlines after second read: {content2.get('outlines')}"
    )
    assert img_item2["handle"].startswith("mcpref:"), (
        f"Round-trip handle must start with 'mcpref:', got: {img_item2['handle']!r}"
    )
    assert isinstance(img_item2.get("width"), (int, float))
    assert isinstance(img_item2.get("height"), (int, float))


@pytest.mark.e2e
def test_inline_image_without_size_in_width_only_outline(mcp_session, claudespike_section_id):
    """Image-containing outline with explicit width but NO image height writes cleanly.

    OneNote rejects an image-containing outline whose <one:Size> carries width
    without height (hresult -0x7ffbdfff).  The builder compensates by emitting
    height="0.0" (auto-height) when the caller sets outline.width but the inline
    image has no height of its own.  This test pins down that OneNote accepts
    that shape and the image survives the write.
    """
    # --- Setup: page with a known image so we can obtain a mcpref handle ----
    title = "test_inline_image_width_only_outline_auto"
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": title,
    })
    page_xml = _make_page_xml(title, _TINY_PNG_BASE64, _IMG_WIDTH, _IMG_HEIGHT)
    result = mcp_session.call_tool("replace_page_xml", {
        "page_id": page_id,
        "page_xml": page_xml,
    })
    assert result == "ok", f"replace_page_xml setup failed: {result}"

    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    seeded = _find_inline_image(content)
    assert seeded is not None, "Setup image not found after seeding the page"
    handle = seeded["handle"]

    # --- Write: outline with explicit width; image with no width/height -----
    outlines = [{
        "width": 300.0,
        "items": [
            {"type": "paragraph", "text": "Text above the image"},
            {"type": "inline_image", "handle": handle},
        ],
    }]
    write_result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": title,
        "outlines": outlines,
    })
    assert write_result == "ok", (
        f"replace_page with width-only outline containing a size-less inline image "
        f"failed: {write_result!r} — height='0.0' auto-height shape regressed?"
    )

    # --- Read back: the image must have survived -----------------------------
    content2 = mcp_session.call_tool("get_page", {"page_id": page_id})
    img_item = _find_inline_image(content2)
    assert img_item is not None, (
        "Inline image lost after writing a width-only outline. "
        f"Outlines after write: {content2.get('outlines')}"
    )
    assert img_item["handle"].startswith("mcpref:")
