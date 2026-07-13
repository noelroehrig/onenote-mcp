"""E2E tests for the XML escape-hatch path — exercised through the MCP server.

All calls go through server.py via the MCP stdio transport using the
mcp_session fixture.  No direct imports of com.py or builders.py.

Image round-trip coverage lives in test_slim_path_via_mcp.py
(test_replace_page_slim_with_image, test_cross_page_image_transfer_via_slim_path)
and test_inline_images.py (test_inline_image_round_trip) — com._IMAGE_CACHE
internals are not observable through the MCP transport boundary.

Requires: OneNote desktop running, at least one notebook open.

Run with:  pytest -m e2e tests/e2e/test_xml_path_via_mcp.py -v
"""
import xml.etree.ElementTree as ET
import pytest

from tests.e2e.mcp_client import mcp_session  # noqa: F401 — re-export fixture

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
ET.register_namespace("one", _ONE_NS)
_NS = {"one": _ONE_NS}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _page_xml_with_text(page_id: str, title: str, body: str) -> str:
    """Build a minimal <one:Page> XML with title and one body paragraph."""
    page = ET.Element(f"{{{_ONE_NS}}}Page")
    page.set("ID", page_id)
    title_el = ET.SubElement(page, f"{{{_ONE_NS}}}Title")
    title_oe = ET.SubElement(title_el, f"{{{_ONE_NS}}}OE")
    ET.SubElement(title_oe, f"{{{_ONE_NS}}}T").text = title
    outline = ET.SubElement(page, f"{{{_ONE_NS}}}Outline")
    oe_children = ET.SubElement(outline, f"{{{_ONE_NS}}}OEChildren")
    ET.SubElement(ET.SubElement(oe_children, f"{{{_ONE_NS}}}OE"), f"{{{_ONE_NS}}}T").text = body
    return ET.tostring(page, encoding="unicode", xml_declaration=False)


def _clean_text(raw: str) -> str:
    """Strip HTML tags from raw text."""
    import re
    import html
    plain = re.sub(r"<[^>]+>", "", raw)
    return html.unescape(plain)


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

@pytest.mark.e2e
def test_list_hierarchy_xml_returns_xml(mcp_session):
    result = mcp_session.call_tool("list_hierarchy_xml")
    assert result, "list_hierarchy_xml returned empty string"
    assert "<" in result, "result does not look like XML"
    assert "Notebook" in result, "no Notebook tag found in hierarchy XML"


@pytest.mark.e2e
def test_create_page_returns_id(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "E2E XML via MCP: create test",
    })
    assert page_id, "create_page returned empty ID"
    assert isinstance(page_id, str), "create_page did not return a string ID"


@pytest.mark.e2e
def test_replace_page_xml_round_trip(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "E2E XML via MCP: replace round-trip",
    })
    page_xml = _page_xml_with_text(page_id, "E2E XML via MCP: replace round-trip", "Hello from replace_page_xml")
    result = mcp_session.call_tool("replace_page_xml", {"page_id": page_id, "page_xml": page_xml})
    assert result == "ok", f"replace_page_xml returned {result!r}"

    returned = mcp_session.call_tool("get_page_xml", {"page_id": page_id})
    assert isinstance(returned, str), f"get_page_xml must return a string, got {type(returned).__name__}"
    root = ET.fromstring(returned)
    t_elements = root.findall(f".//{{{_ONE_NS}}}T")
    texts = [_clean_text(el.text or "") for el in t_elements]
    assert any("Hello from replace_page_xml" in t for t in texts), (
        f"Expected 'Hello from replace_page_xml' in T elements, got: {texts}"
    )


@pytest.mark.e2e
def test_append_page_xml(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "E2E XML via MCP: append test",
    })
    init_xml = _page_xml_with_text(page_id, "E2E XML via MCP: append test", "Initial content")
    result = mcp_session.call_tool("replace_page_xml", {"page_id": page_id, "page_xml": init_xml})
    assert result == "ok", f"replace_page_xml returned {result!r}"

    outline_el = ET.Element(f"{{{_ONE_NS}}}Outline")
    ET.register_namespace("one", _ONE_NS)
    oe_children = ET.SubElement(outline_el, f"{{{_ONE_NS}}}OEChildren")
    ET.SubElement(ET.SubElement(oe_children, f"{{{_ONE_NS}}}OE"), f"{{{_ONE_NS}}}T").text = "Appended paragraph"
    outline_xml = ET.tostring(outline_el, encoding="unicode")

    result2 = mcp_session.call_tool("append_page_xml", {"page_id": page_id, "content_xml": outline_xml})
    assert result2 == "ok", f"append_page_xml returned {result2!r}"

    returned = mcp_session.call_tool("get_page_xml", {"page_id": page_id})
    assert "Appended paragraph" in returned, (
        "Expected 'Appended paragraph' in returned page XML"
    )
