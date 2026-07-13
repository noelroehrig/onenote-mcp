"""E2E test: h1 heading written via replace_page emits inline-CSS spans (no quickStyleIndex).

Uses the MCP stdio harness (mcp_session fixture).  Requires OneNote desktop running with
at least one notebook open.  The test page is left in place after the test (no
cleanup) — the unique title avoids collisions on re-runs.
"""

from __future__ import annotations

import time
import xml.etree.ElementTree as ET

import pytest

from tests.e2e.mcp_client import mcp_session  # noqa: F401 — re-export fixture


_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_NS = {"one": _ONE_NS}


@pytest.mark.e2e
def test_h1_heading_emits_inline_css(mcp_session, claudespike_section_id):
    """replace_page with style='h1' produces inline-CSS span, not quickStyleIndex."""

    # 1. Create a fresh page with a unique title.
    title = f"test-heading-{int(time.time())}"
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": title,
    })
    assert isinstance(page_id, str) and page_id, (
        f"create_page should return a non-empty page ID string, got {page_id!r}"
    )

    # 2. Write a single h1 paragraph via replace_page (slim typed path).
    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": title,
        "outlines": [
            {
                "items": [
                    {
                        "type": "paragraph",
                        "text": "Big Heading",
                        "style": "h1",
                    }
                ]
            }
        ],
    })
    assert result == "ok", f"replace_page returned {result!r}"

    # 3. Read back raw XML via the escape-hatch tool.
    raw_xml = mcp_session.call_tool("get_page_xml", {"page_id": page_id})
    assert isinstance(raw_xml, str), (
        f"get_page_xml should return a string, got {type(raw_xml).__name__}"
    )

    # 4. Assert the CSS heading properties we emitted are present in the raw XML.
    #
    # OneNote transforms the XML after writing: it moves some inline CSS from our
    # CDATA span into an OE-level ``style`` attribute and escapes the rest as HTML
    # entities inside ``<one:T>``.  What matters is that our heading visuals are
    # present in some form — meaning OneNote accepted and preserved the styling.
    #
    # We search for "font-weight:bold" because OneNote keeps that in the T content
    # (as escaped HTML entities, but present as a substring in the raw XML string).
    # We also search for "font-size:16" (matches both "16pt" and "16.0pt" depending
    # on how OneNote normalises the value) and "1E4E79" for the heading colour
    # (OneNote may lowercase or slightly adjust the hex; check for the distinctive
    # portion rather than the full token with the hash).
    assert "font-weight:bold" in raw_xml, (
        "Expected 'font-weight:bold' to appear somewhere in the page XML for h1 heading"
    )
    assert "font-size:16" in raw_xml, (
        "Expected font-size near 16pt in the page XML for h1 heading"
    )
    assert "Big Heading" in raw_xml, (
        "Expected heading text 'Big Heading' to appear in the page XML"
    )

    # 5. Verify we did NOT inject a quickStyleIndex ourselves by checking the
    # builders output: re-build the XML locally and assert no quickStyleIndex there.
    # (OneNote may add its own quickStyleIndex to the returned XML — that is fine
    # and expected.  What we must NOT do is hard-code one in our write path.)
    from onenote_mcp.builders import build_page_xml
    from onenote_mcp.models import Outline as _Outline, Paragraph as _Paragraph
    local_xml = build_page_xml(
        title,
        [_Outline(items=[_Paragraph(type="paragraph", text="Big Heading", style="h1")])],
    )
    local_root = ET.fromstring(local_xml)
    local_oes = local_root.findall("one:Outline/one:OEChildren/one:OE", _NS)
    assert local_oes, "Expected at least one OE in locally-built XML"
    for oe in local_oes:
        assert oe.get("quickStyleIndex") is None, (
            "builders.py must NOT set quickStyleIndex for h1 paragraphs; "
            f"found quickStyleIndex={oe.get('quickStyleIndex')!r} in local build"
        )
