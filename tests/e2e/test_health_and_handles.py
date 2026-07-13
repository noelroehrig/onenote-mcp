"""E2E tests for the health-check and handle-validation tools, plus the
include_binary toggle on get_page.

These go through server.py via the MCP stdio transport (mcp_session fixture)
and require OneNote desktop running; the suite is skipped automatically if
OneNote is unavailable.
"""
import pytest

from tests.e2e.mcp_client import mcp_session  # noqa: F401 — re-export for fixture

_TINY_PNG_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYA"
    "AjCB0C8AAAAASUVORK5CYII="
)
_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_XMLNS = f'xmlns:one="{_ONE_NS}"'


@pytest.mark.e2e
def test_ping_reports_server_and_onenote(mcp_session):
    """ping returns a fast, structured health status; OneNote is responsive here."""
    result = mcp_session.call_tool("ping")
    assert result["server"] == "ok"
    # The e2e suite only runs when OneNote is reachable, so it must be responsive.
    assert result["onenote_responsive"] is True


@pytest.mark.e2e
def test_validate_handles_true_for_fresh_false_for_bogus(mcp_session, claudespike_section_id):
    """Handles from a just-read page validate True; a fabricated handle is False."""
    title = "test_validate_handles_auto"
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": title,
    })
    page_xml = (
        f'<one:Page {_XMLNS}>'
        f'<one:Title><one:OE><one:T>{title}</one:T></one:OE></one:Title>'
        f'<one:Outline><one:OEChildren><one:OE><one:Image>'
        f'<one:Size width="100.0" height="80.0" isSetByUser="true"/>'
        f'<one:Data>{_TINY_PNG_BASE64}</one:Data>'
        f'</one:Image></one:OE></one:OEChildren></one:Outline>'
        f'</one:Page>'
    )
    assert mcp_session.call_tool("replace_page_xml", {"page_id": page_id, "page_xml": page_xml}) == "ok"

    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    handle = None
    for outline in content["outlines"]:
        for item in outline.get("items", []):
            if item.get("type") == "inline_image":
                handle = item["handle"]
    assert handle and handle.startswith("mcpref:")

    result = mcp_session.call_tool("validate_handles", {
        "handles": [handle, "mcpref:deadbeef0000"],
    })
    assert result[handle] is True
    assert result["mcpref:deadbeef0000"] is False


@pytest.mark.e2e
def test_get_image_data_returns_base64(mcp_session, claudespike_section_id):
    """get_image_data fetches the bytes for a handle read with the no-binary default."""
    title = "test_get_image_data_auto"
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": title,
    })
    page_xml = (
        f'<one:Page {_XMLNS}>'
        f'<one:Title><one:OE><one:T>{title}</one:T></one:OE></one:Title>'
        f'<one:Outline><one:OEChildren><one:OE><one:Image>'
        f'<one:Size width="100.0" height="80.0" isSetByUser="true"/>'
        f'<one:Data>{_TINY_PNG_BASE64}</one:Data>'
        f'</one:Image></one:OE></one:OEChildren></one:Outline>'
        f'</one:Page>'
    )
    assert mcp_session.call_tool("replace_page_xml", {"page_id": page_id, "page_xml": page_xml}) == "ok"

    content = mcp_session.call_tool("get_page", {"page_id": page_id})  # no binary
    handle = next(
        item["handle"]
        for outline in content["outlines"]
        for item in outline.get("items", [])
        if item.get("type") == "inline_image"
    )
    data = mcp_session.call_tool("get_image_data", {"handle": handle})
    assert isinstance(data, str) and len(data) > 0
    # PNG signature in base64 begins with "iVBOR".
    assert data.replace("\r", "").replace("\n", "").startswith("iVBOR")


@pytest.mark.e2e
def test_get_page_include_binary_round_trip(mcp_session, claudespike_section_id):
    """include_binary=True embeds bytes inline and still round-trips through replace_page."""
    title = "test_include_binary_auto"
    page_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": title,
    })
    page_xml = (
        f'<one:Page {_XMLNS}>'
        f'<one:Title><one:OE><one:T>{title}</one:T></one:OE></one:Title>'
        f'<one:Outline><one:OEChildren><one:OE><one:Image>'
        f'<one:Size width="100.0" height="80.0" isSetByUser="true"/>'
        f'<one:Data>{_TINY_PNG_BASE64}</one:Data>'
        f'</one:Image></one:OE></one:OEChildren></one:Outline>'
        f'</one:Page>'
    )
    assert mcp_session.call_tool("replace_page_xml", {"page_id": page_id, "page_xml": page_xml}) == "ok"

    content = mcp_session.call_tool("get_page", {"page_id": page_id, "include_binary": True})
    assert content["outlines"]

    def _has_image(c):
        return any(
            item.get("type") == "inline_image"
            for outline in c["outlines"]
            for item in outline.get("items", [])
        )

    assert _has_image(content)
    # Round-trip the binary-read content back unchanged.
    assert mcp_session.call_tool("replace_page", {
        "page_id": page_id,
        "title": title,
        "outlines": content["outlines"],
        "images": content.get("images", []),
    }) == "ok"
    assert _has_image(mcp_session.call_tool("get_page", {"page_id": page_id, "include_binary": True}))
