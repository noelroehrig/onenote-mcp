"""E2E test: text with HTML/XML special characters survives a write and read.

OneNote reads <one:T> content as HTML, so the server escapes text on write and
get_page must return it unchanged.  Requires OneNote desktop with at least one
notebook open.  The test page is left in place after the test (no cleanup).
"""

from __future__ import annotations

import time

import pytest

from tests.e2e.mcp_client import mcp_session  # noqa: F401 (re-export fixture)

SPECIAL = "a < b > c & d ]]> <b>not bold</b> &amp; \"double\" 'single' Grüße äöü ÄÖÜ ß"


@pytest.mark.e2e
def test_special_characters_round_trip(mcp_session, claudespike_section_id):
    title = f"test-escaping-{int(time.time())} {SPECIAL}"
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": title,
    })
    assert mcp_session.call_tool("get_page", {"page_id": page_id})["title"] == title

    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": title,
        "outlines": [{"items": [
            {"type": "paragraph", "text": SPECIAL},
            {"type": "paragraph", "text": SPECIAL, "bold": True, "color": "#C00000",
             "highlight": "#FFFF00", "font_family": "Courier New"},
            {"type": "paragraph", "segments": [{"text": SPECIAL}, {"text": SPECIAL, "italic": True}]},
            {"type": "list", "style": "bullet", "items": [{"text": SPECIAL}]},
        ]}],
    })
    assert result == "ok", f"replace_page returned {result!r}"

    page = mcp_session.call_tool("get_page", {"page_id": page_id})
    assert page["title"] == title
    plain, styled, mixed, bullets = page["outlines"][0]["items"]

    assert plain["text"] == SPECIAL

    assert styled["text"] == SPECIAL
    assert styled["bold"] is True
    assert styled["color"].upper() == "#C00000"
    assert styled["highlight"].upper() == "#FFFF00"
    assert styled["font_family"] == "Courier New"

    assert "".join(s["text"] for s in mixed["segments"]) == SPECIAL + SPECIAL
    assert mixed["segments"][-1]["text"] == SPECIAL
    assert mixed["segments"][-1].get("italic") is True

    assert bullets["items"][0]["text"] == SPECIAL
