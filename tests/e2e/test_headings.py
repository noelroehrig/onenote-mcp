"""E2E test: headings written via replace_page and append_page are real OneNote headings.

The written page declares a <one:QuickStyleDef> per heading level, so OneNote
stores the paragraphs with a heading style and get_page reads them back as such.

Uses the MCP stdio harness (mcp_session fixture).  Requires OneNote desktop running with
at least one notebook open.  The test page is left in place after the test (no
cleanup); the unique title avoids collisions on re-runs.
"""

from __future__ import annotations

import time
import xml.etree.ElementTree as ET

import pytest

from tests.e2e.mcp_client import mcp_session  # noqa: F401 — re-export fixture


_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_NS = {"one": _ONE_NS}


@pytest.mark.e2e
def test_headings_round_trip(mcp_session, claudespike_section_id):
    title = f"test-heading-{int(time.time())}"
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": title,
    })

    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": title,
        "outlines": [{"items": [
            {"type": "paragraph", "text": "Big Heading", "style": "h1"},
            {"type": "paragraph", "text": "Small Heading", "style": "h3"},
            {"type": "paragraph", "text": "Body text"},
        ]}],
    })
    assert result == "ok", f"replace_page returned {result!r}"
    result = mcp_session.call_tool("append_page", {
        "page_id": page_id,
        "outline": {"items": [{"type": "paragraph", "text": "Appended Heading", "style": "h2"}]},
    })
    assert result == "ok", f"append_page returned {result!r}"

    page = mcp_session.call_tool("get_page", {"page_id": page_id})
    items = [item for outline in page["outlines"] for item in outline["items"]]
    assert [(item["text"], item.get("style", "normal")) for item in items] == [
        ("Big Heading", "h1"),
        ("Small Heading", "h3"),
        ("Body text", "normal"),
        ("Appended Heading", "h2"),
    ]
    # Heading bold comes from the style, so it is not reported as direct formatting.
    assert not any(item.get("bold") for item in items), items

    # OneNote itself knows the paragraphs as headings: their quickStyleIndex
    # points at a QuickStyleDef named after the heading level.
    root = ET.fromstring(mcp_session.call_tool("get_page_xml", {"page_id": page_id}))
    names = {qsd.get("index"): qsd.get("name") for qsd in root.findall("one:QuickStyleDef", _NS)}
    referenced = [
        names.get(oe.get("quickStyleIndex"))
        for oe in root.iterfind("one:Outline/one:OEChildren/one:OE", _NS)
    ]
    assert referenced == ["h1", "h3", "p", "h2"]
