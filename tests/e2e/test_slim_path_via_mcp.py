"""E2E tests for the slim JSON-based endpoint path — exercised through the MCP server.

All calls go through server.py via the MCP stdio transport using the
mcp_session fixture.  No direct imports of com.py or builders.py
are used for test logic.  Pydantic models from onenote_mcp.models are imported
only for schema reference; test payloads are plain dicts (what the LLM actually
sends over the MCP transport).

Requires: OneNote desktop running, at least one notebook open.

Run with:  pytest -m e2e tests/e2e/test_slim_path_via_mcp.py -v
"""
import xml.etree.ElementTree as ET
import pytest

from tests.e2e.mcp_client import mcp_session  # noqa: F401 — re-export fixture

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_NS = {"one": _ONE_NS}

# A minimal 1x1 transparent PNG for seeding test pages that need an image.
# Same constant used in test_inline_images.py.
_TINY_PNG_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYA"
    "AjCB0C8AAAAASUVORK5CYII="
)

# 32x32 solid red PNG — used for the floating-image tests.
_RED_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAAJ0lEQVR42u3NsQkAAAjA"
    "sP7/tF7hIASyp6lTCQQCgUAgEAgEgi/BAjLD/C5w/SM9AAAAAElFTkSuQmCC"
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _seed_page_with_floating_image(mcp_session, page_id: str, title: str, png_b64: str) -> str:
    """Seed page with a floating image via replace_page_xml; return 'ok'."""
    ET.register_namespace("one", _ONE_NS)
    page = ET.Element(f"{{{_ONE_NS}}}Page")
    page.set("ID", page_id)
    title_el = ET.SubElement(page, f"{{{_ONE_NS}}}Title")
    ET.SubElement(ET.SubElement(title_el, f"{{{_ONE_NS}}}OE"), f"{{{_ONE_NS}}}T").text = title
    img = ET.SubElement(page, f"{{{_ONE_NS}}}Image")
    ET.SubElement(img, f"{{{_ONE_NS}}}Data").text = png_b64
    page_xml = ET.tostring(page, encoding="unicode")
    return mcp_session.call_tool("replace_page_xml", {"page_id": page_id, "page_xml": page_xml})


def _seed_page_with_inline_image(mcp_session, page_id: str, title: str, png_b64: str,
                                  width: float = 100.0, height: float = 100.0) -> str:
    """Seed page with an inline image (inside an Outline) via replace_page_xml; return 'ok'."""
    ONE = _ONE_NS
    page_xml = (
        f'<one:Page xmlns:one="{ONE}" ID="{page_id}">'
        f'  <one:Title><one:OE><one:T>{title}</one:T></one:OE></one:Title>'
        f'  <one:Outline>'
        f'    <one:OEChildren>'
        f'      <one:OE>'
        f'        <one:Image>'
        f'          <one:Size width="{width}" height="{height}" isSetByUser="true"/>'
        f'          <one:Data>{png_b64}</one:Data>'
        f'        </one:Image>'
        f'      </one:OE>'
        f'    </one:OEChildren>'
        f'  </one:Outline>'
        f'</one:Page>'
    )
    return mcp_session.call_tool("replace_page_xml", {"page_id": page_id, "page_xml": page_xml})


def _find_inline_image(content_dict):
    """Walk outlines to find the first inline_image item; return it or None."""
    for outline in content_dict.get("outlines", []):
        for item in outline.get("items", []):
            if item.get("type") == "inline_image":
                return item
    return None


# ---------------------------------------------------------------------------
# Navigation, read and write basics
# ---------------------------------------------------------------------------

@pytest.mark.e2e
def test_navigation_skeleton_and_section_pages(mcp_session, claudespike_section_id):
    """get_notebooks returns notebooks+sections (no pages); list_pages drills into one."""
    notebooks = mcp_session.call_tool("get_notebooks")
    assert isinstance(notebooks, list) and notebooks, "Expected at least one notebook"
    first_nb = notebooks[0]
    assert first_nb.get("id") and first_nb.get("name")
    assert isinstance(first_nb.get("sections"), list), "notebook must carry a sections list"
    # Skeleton must NOT contain pages anywhere.
    for nb in notebooks:
        for sec in nb["sections"]:
            assert "pages" not in sec, "get_notebooks must not return pages"

    # Drill into the ClaudeSpike section: list_pages returns a flat page list.
    pages = mcp_session.call_tool("list_pages", {"section_id": claudespike_section_id})
    assert isinstance(pages, list), f"list_pages should return a list, got {type(pages).__name__}"
    for page in pages:
        assert page.get("id") and "name" in page


@pytest.mark.e2e
def test_get_page_returns_page_content(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: get_page test",
    })
    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": "Slim via MCP: get_page test",
        "outlines": [
            {"items": [{"type": "paragraph", "text": "Some content for get_page test"}]}
        ],
    })
    assert result == "ok", f"replace_page returned {result!r}"

    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    assert isinstance(content, dict), f"get_page should return a dict, got {type(content).__name__}"
    assert content.get("title"), "Expected non-empty title in page content"
    assert isinstance(content.get("outlines"), list), "outlines should be a list"


@pytest.mark.e2e
def test_replace_page_with_paragraphs(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: paragraphs",
    })
    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": "Slim via MCP: paragraphs",
        "outlines": [
            {"items": [
                {"type": "paragraph", "text": "First paragraph"},
                {"type": "paragraph", "text": "Second paragraph"},
            ]}
        ],
    })
    assert result == "ok", f"replace_page returned {result!r}"

    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    assert content.get("title") == "Slim via MCP: paragraphs", (
        f"Expected title 'Slim via MCP: paragraphs', got: {content.get('title')!r}"
    )
    assert content.get("outlines"), "Expected at least one outline in parsed page"
    all_texts = []
    for outline in content["outlines"]:
        for item in outline.get("items", []):
            if item.get("text"):
                all_texts.append(item["text"])
    assert any(t in ("First paragraph", "Second paragraph") for t in all_texts), (
        f"Expected paragraph text in content, got texts: {all_texts}"
    )


@pytest.mark.e2e
def test_replace_page_with_heading(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: heading test",
    })
    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": "Slim via MCP: heading test",
        "outlines": [
            {"items": [{"type": "paragraph", "text": "Big Title", "style": "h1"}]}
        ],
    })
    assert result == "ok", f"replace_page returned {result!r}"

    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    items = [item for outline in content.get("outlines", []) for item in outline.get("items", [])]
    assert [(item.get("text"), item.get("style")) for item in items] == [("Big Title", "h1")], (
        f"Expected the h1 'Big Title' to read back as a heading, got: {items}"
    )


@pytest.mark.e2e
def test_replace_page_with_bullet_list(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: bullet list test",
    })
    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": "Slim via MCP: bullet list test",
        "outlines": [
            {"items": [
                {"type": "list", "style": "bullet", "items": [
                    {"text": "Alpha", "children": []},
                    {"text": "Beta", "children": []},
                ]}
            ]}
        ],
    })
    assert result == "ok", f"replace_page returned {result!r}"

    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    bullet_lists = []
    for outline in content.get("outlines", []):
        for item in outline.get("items", []):
            if item.get("type") == "list" and item.get("style") == "bullet":
                bullet_lists.append(item)
    assert bullet_lists, "Expected at least one bullet list in content"
    found_list = bullet_lists[0]
    assert len(found_list.get("items", [])) >= 2, (
        f"Expected at least 2 list items, got {len(found_list.get('items', []))}"
    )


@pytest.mark.e2e
def test_replace_page_with_formatted_text(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: formatted text test",
    })
    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": "Slim via MCP: formatted text test",
        "outlines": [
            {"items": [
                {"type": "paragraph", "segments": [
                    {"text": "Bold", "bold": True},
                    {"text": " and normal"},
                ]}
            ]}
        ],
    })
    assert result == "ok", f"replace_page returned {result!r}"

    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    assert content, "get_page returned empty result after writing formatted text"


@pytest.mark.e2e
def test_append_page_slim(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: append test",
    })
    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": "Slim via MCP: append test",
        "outlines": [
            {"items": [{"type": "paragraph", "text": "Initial content"}]}
        ],
    })
    assert result == "ok", f"replace_page returned {result!r}"

    # append_page takes a single Outline object
    result2 = mcp_session.call_tool("append_page", {
        "page_id": page_id,
        "outline": {"items": [{"type": "paragraph", "text": "Appended via slim"}]},
    })
    assert result2 == "ok", f"append_page returned {result2!r}"

    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    all_texts = []
    for outline in content.get("outlines", []):
        for item in outline.get("items", []):
            if item.get("text"):
                all_texts.append(item["text"])
    assert any("Appended via slim" in t for t in all_texts), (
        f"Expected 'Appended via slim' in paragraph texts, got: {all_texts}"
    )


@pytest.mark.e2e
def test_replace_page_with_positioned_outline(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: positioned outline test",
    })
    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": "Slim via MCP: positioned outline test",
        "outlines": [
            {
                "position": {"x": 36.0, "y": 100.0, "z": 0},
                "items": [{"type": "paragraph", "text": "Positioned"}],
            }
        ],
    })
    assert result == "ok", f"replace_page returned {result!r}"

    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    assert content, "get_page returned empty result after writing positioned outline"


@pytest.mark.e2e
def test_replace_page_slim_with_image(mcp_session, claudespike_section_id):
    """Floating image: seed via XML, read handle, write back via slim, assert image persists."""
    # Step 1: create page
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: image round-trip",
    })

    # Step 2: seed with raw base64 floating image via XML escape-hatch
    result = _seed_page_with_floating_image(mcp_session, page_id, "Slim via MCP: image round-trip", _RED_PNG_B64)
    assert result == "ok", f"replace_page_xml seed failed: {result!r}"

    # Step 3: read back via slim get_page — floating images appear in content["images"]
    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    images = content.get("images", [])
    assert images, (
        "Expected a FloatingImage in content['images'] after seeding with raw PNG"
    )
    image_item = images[0]
    assert image_item.get("handle", "").startswith("mcpref:"), (
        f"Expected image handle to start with 'mcpref:', got: {image_item.get('handle')!r}"
    )

    # Step 4: write back via slim replace_page using the dict handle
    result2 = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": "Slim via MCP: image round-trip",
        "outlines": [],
        "images": [image_item],
    })
    assert result2 == "ok", f"replace_page slim image write failed: {result2!r}"

    # Step 5: read back raw XML and assert <one:Data> element is present
    final_xml = mcp_session.call_tool("get_page_xml", {"page_id": page_id})
    final_root = ET.fromstring(final_xml)
    data_elements = final_root.findall(".//one:Data", _NS)
    assert data_elements, (
        "Expected <one:Data> element in XML after slim-path image round-trip"
    )


@pytest.mark.e2e
def test_cross_page_image_transfer_via_slim_path(mcp_session, claudespike_section_id):
    """Image handle from page A must be usable in replace_page for a different page B."""
    # Step 1: create and seed source page A
    source_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: cross-page source",
    })
    result = _seed_page_with_floating_image(mcp_session, source_id, "Slim via MCP: cross-page source", _RED_PNG_B64)
    assert result == "ok", f"replace_page_xml seed failed: {result!r}"

    # Step 2: read source page A to get mcpref handle
    source_content = mcp_session.call_tool("get_page", {"page_id": source_id})
    source_images = source_content.get("images", [])
    assert source_images, "Expected a FloatingImage in source page images after seeding"
    image_item = source_images[0]
    assert image_item.get("handle", "").startswith("mcpref:"), (
        f"Expected mcpref handle from source page, got: {image_item.get('handle')!r}"
    )

    # Step 3: create target page B (fresh, no image)
    target_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: cross-page target",
    })
    assert target_id != source_id, "Source and target must be different pages"

    # Step 4: write source handle to target page via slim replace_page
    result2 = mcp_session.call_tool("replace_page", {
        "page_id": target_id,
        "title": "Slim via MCP: cross-page target",
        "outlines": [],
        "images": [image_item],
    })
    assert result2 == "ok", f"replace_page cross-page write failed: {result2!r}"

    # Step 5: verify image landed on target page
    target_xml = mcp_session.call_tool("get_page_xml", {"page_id": target_id})
    target_root = ET.fromstring(target_xml)
    data_elements = target_root.findall(".//one:Data", _NS)
    assert data_elements, (
        "Expected <one:Data> element in target page after cross-page image transfer"
    )


# ---------------------------------------------------------------------------
# Fresh-page inline image write (verification)
# ---------------------------------------------------------------------------

@pytest.mark.e2e
def test_replace_page_with_inline_image_on_fresh_page(mcp_session, claudespike_section_id):
    """Inline image written to a fresh page (never had an image) must survive.

    Only the SOURCE page is seeded via the XML escape-hatch; the target page
    gets no seeding at all, making this a true fresh-page write through the
    slim path.

    If this test FAILS, do NOT mark it xfail here — surface the failure to the
    user.  Possible follow-ups: validation-time guard rejecting InlineImage items
    on fresh pages, or auto-promote to FloatingImage when the target page has no
    image-bearing outline.
    """
    # ------------------------------------------------------------------
    # Step 1: create source page A and seed it with an inline image via XML
    # ------------------------------------------------------------------
    source_title = "inline-fresh-source"
    source_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": source_title,
    })
    assert isinstance(source_id, str) and source_id, "create_page must return a non-empty page ID"

    seed_result = _seed_page_with_inline_image(
        mcp_session, source_id, source_title, _TINY_PNG_BASE64, width=100.0, height=100.0
    )
    assert seed_result == "ok", f"replace_page_xml seed of source page failed: {seed_result!r}"

    # ------------------------------------------------------------------
    # Step 2: read A via slim get_page; extract the inline_image handle
    # ------------------------------------------------------------------
    source_content = mcp_session.call_tool("get_page", {"page_id": source_id})
    inline_img = _find_inline_image(source_content)
    assert inline_img is not None, (
        "Expected an inline_image item in source page after seeding. "
        f"Got outlines: {source_content.get('outlines')}"
    )
    handle = inline_img.get("handle", "")
    assert handle.startswith("mcpref:"), (
        f"inline_image handle must start with 'mcpref:', got: {handle!r}"
    )

    # ------------------------------------------------------------------
    # Step 3: create a FRESH target page B (no prior image content)
    # ------------------------------------------------------------------
    target_title = "inline-fresh-target"
    target_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": target_title,
    })
    assert isinstance(target_id, str) and target_id, "create_page for target must return a non-empty ID"
    assert target_id != source_id, "Source and target must be different pages"

    # ------------------------------------------------------------------
    # Step 4: write the inline image to fresh page B via slim replace_page
    # ------------------------------------------------------------------
    write_result = mcp_session.call_tool("replace_page", {
        "page_id": target_id,
        "title": target_title,
        "outlines": [
            {
                "position": None,
                "width": None,
                "items": [
                    {
                        "type": "inline_image",
                        "handle": handle,
                        "width": 100.0,
                        "height": 100.0,
                    }
                ],
            }
        ],
        "images": [],
    })
    assert write_result == "ok", f"replace_page on fresh target page failed: {write_result!r}"

    # ------------------------------------------------------------------
    # Step 5: read B back via slim get_page; assert inline_image is present
    # ------------------------------------------------------------------
    target_content = mcp_session.call_tool("get_page", {"page_id": target_id})
    target_img = _find_inline_image(target_content)
    assert target_img is not None, (
        "Inline image was NOT found on fresh target page after slim replace_page. "
        f"target outlines: {target_content.get('outlines')}"
    )
    target_handle = target_img.get("handle", "")
    assert target_handle.startswith("mcpref:"), (
        f"Inline image on target page must have a mcpref: handle, got: {target_handle!r}"
    )

    # ------------------------------------------------------------------
    # Step 6: verify via raw XML that bytes are physically on target page
    # ------------------------------------------------------------------
    target_xml = mcp_session.call_tool("get_page_xml", {"page_id": target_id})
    assert isinstance(target_xml, str), f"get_page_xml must return str, got {type(target_xml).__name__}"
    target_root = ET.fromstring(target_xml)
    data_elements = target_root.findall(".//one:Data", _NS)
    assert data_elements, (
        "Expected <one:Data> element in target page XML — image bytes must be physically present. "
        f"Raw XML (first 500 chars): {target_xml[:500]}"
    )


# ---------------------------------------------------------------------------
# OneNote rejects <one:Size> that carries width without height (hresult
# 0x80042001); the builder pairs a caller-supplied width with height="0.0".
# Mirrors a real Claude Desktop payload: multiple positioned outlines with
# width but no height.
# ---------------------------------------------------------------------------

@pytest.mark.e2e
def test_replace_page_with_width_only_outlines(mcp_session, claudespike_section_id):
    """Multi-outline page where every outline has explicit width but no height.

    Every emitted <one:Size> must carry height="0.0" alongside the width —
    OneNote rejects width-only Size (hresult 0x80042001).  Uses a realistic
    page shape: 4 positioned outlines, headings with custom CSS, lists, an
    image placeholder.  Asserts the write succeeds and content survives the
    round-trip.
    """
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: width-only outlines",
    })
    outlines = [
        {
            "position": {"x": 36.0, "y": 70.0, "z": 0},
            "width": 700.0,
            "items": [
                {"type": "paragraph", "text": "WARM-UP", "style": "h1",
                 "bold": True, "color": "#ffffff", "font_size": 22, "highlight": "#2e7d32"},
                {"type": "paragraph", "text": "Body paragraph", "bold": True},
            ],
        },
        {
            "position": {"x": 36.0, "y": 470.0, "z": 1},
            "width": 700.0,
            "items": [
                {"type": "paragraph", "text": "AGENDA", "style": "h1"},
                {"type": "list", "style": "bullet", "items": [
                    {"text": "First bullet"},
                    {"text": "Second bullet"},
                ]},
            ],
        },
        {
            "position": {"x": 36.0, "y": 660.0, "z": 2},
            "width": 700.0,
            "items": [
                {"type": "paragraph", "text": "INTRODUCTION", "style": "h1"},
                {"type": "image_placeholder", "description": "Diagram here",
                 "width": 600.0, "height": 270.0},
            ],
        },
        {
            "position": {"x": 800.0, "y": 70.0, "z": 3},
            "width": 470.0,
            "items": [
                {"type": "paragraph", "text": "SOLUTIONS", "style": "h2"},
                {"type": "paragraph", "text": "Detail row", "font_size": 10},
            ],
        },
    ]
    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": "Slim via MCP: width-only outlines",
        "outlines": outlines,
    })
    assert result == "ok", f"replace_page returned {result!r} — width-only Size regressed?"

    # Sanity: the page now has at least the four outlines we wrote.
    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    assert len(content.get("outlines", [])) >= 4, (
        f"Expected at least 4 outlines after write, got: {content.get('outlines')}"
    )
    all_texts = []
    for outline in content["outlines"]:
        for item in outline.get("items", []):
            if item.get("text"):
                all_texts.append(item["text"])
    for needle in ("WARM-UP", "AGENDA", "INTRODUCTION", "SOLUTIONS"):
        assert any(needle in t for t in all_texts), (
            f"Expected {needle!r} in page after width-only-outline write, got: {all_texts}"
        )


# ---------------------------------------------------------------------------
# LLMs sometimes emit a trailing empty {} entry in a list.  Empty list items
# are tolerated and silently dropped so the rest of the page still writes,
# instead of failing argument validation at the MCP boundary.
# ---------------------------------------------------------------------------

@pytest.mark.e2e
def test_replace_page_drops_empty_list_item(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: empty list item",
    })
    outlines = [
        {
            "position": {"x": 36.0, "y": 70.0, "z": 0},
            "width": 700.0,
            "items": [
                {"type": "paragraph", "text": "Task 4: Example list", "bold": True},
                {"type": "list", "style": "numbered", "items": [
                    {"text": "Real entry"},
                    {},  # empty item Claude sometimes emits — must not blow up
                ]},
            ],
        },
    ]
    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": "Slim via MCP: empty list item",
        "outlines": outlines,
    })
    assert result == "ok", f"replace_page returned {result!r} — empty list item regressed?"

    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    all_texts = []
    for outline in content.get("outlines", []):
        for item in outline.get("items", []):
            if item.get("text"):
                all_texts.append(item["text"])
            for li in item.get("items", []):
                if li.get("text"):
                    all_texts.append(li["text"])
    assert any("Real entry" in t for t in all_texts), (
        f"Expected 'Real entry' to survive, got: {all_texts}"
    )
    # The empty item must not have produced a stray blank list bullet.
    assert not any(t.strip() == "" for t in all_texts), (
        f"Empty list item should have been dropped, got blank text in: {all_texts}"
    )


# ---------------------------------------------------------------------------
# UpdatePageContent is an upsert and cannot delete — only the DeletePageContent
# COM method removes objects.  This test replaces a page twice and verifies the
# first round's content is actually GONE, not just that the second round's
# content is present.
# ---------------------------------------------------------------------------

@pytest.mark.e2e
def test_replace_page_deletes_old_content(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "Slim via MCP: delete old content",
    })

    old_marker = "OLDCONTENT_MARKER_7Q2"
    new_marker = "NEWCONTENT_MARKER_9Z5"

    first = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": "Slim via MCP: delete old content",
        "outlines": [
            {
                "position": {"x": 36.0, "y": 70.0, "z": 0},
                "width": 600.0,
                "items": [
                    {"type": "paragraph", "text": old_marker},
                    {"type": "list", "style": "bullet", "items": [
                        {"text": f"{old_marker} bullet one"},
                        {"text": f"{old_marker} bullet two"},
                    ]},
                ],
            },
        ],
    })
    assert first == "ok", f"first replace_page returned {first!r}"

    def _all_text(content):
        out = []
        for outline in content.get("outlines", []):
            for item in outline.get("items", []):
                if item.get("text"):
                    out.append(item["text"])
                for li in item.get("items", []):
                    if li.get("text"):
                        out.append(li["text"])
        return out

    # Sanity: old content is really there after the first write.
    after_first = _all_text(mcp_session.call_tool("get_page", {"page_id": page_id}))
    assert any(old_marker in t for t in after_first), (
        f"Old marker not present after first write: {after_first}"
    )

    second = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": "Slim via MCP: delete old content",
        "outlines": [
            {
                "position": {"x": 36.0, "y": 70.0, "z": 0},
                "width": 600.0,
                "items": [{"type": "paragraph", "text": new_marker}],
            },
        ],
    })
    assert second == "ok", f"second replace_page returned {second!r}"

    after_second = _all_text(mcp_session.call_tool("get_page", {"page_id": page_id}))
    assert any(new_marker in t for t in after_second), (
        f"New marker missing after replace: {after_second}"
    )
    # The whole point: old content must be GONE, proving the delete-pass works.
    assert not any(old_marker in t for t in after_second), (
        f"Old content survived the replace — delete-pass is broken: {after_second}"
    )
