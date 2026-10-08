"""E2E test: get_page reports each outline's height as OneNote measured it.

Right after a write OneNote reports a height that grows with the content: ten
paragraphs are ten times as tall as one, and an outline holding an image is
as tall as the image.  replace_page ignores the read-only height, so the read
outlines can be written back unchanged.  Requires OneNote desktop; the test
page is left in place.
"""
import pytest

from tests.e2e.mcp_client import mcp_session  # noqa: F401 (re-export fixture)
from tests.png import png_base64

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _paragraphs(count: int) -> list[dict]:
    return [{"type": "paragraph", "text": f"paragraph {i}"} for i in range(count)]


def _image_handle(mcp_session, page_id: str, title: str) -> str:
    """Seed the page with one inline image and return its handle."""
    page_xml = (
        f'<one:Page xmlns:one="{_ONE_NS}">'
        f"<one:Title><one:OE><one:T>{title}</one:T></one:OE></one:Title>"
        "<one:Outline><one:OEChildren><one:OE><one:Image>"
        f"<one:Data>{png_base64(4, 8)}</one:Data>"
        "</one:Image></one:OE></one:OEChildren></one:Outline></one:Page>"
    )
    assert mcp_session.call_tool("replace_page_xml", {"page_id": page_id, "page_xml": page_xml}) == "ok"
    return mcp_session.call_tool("get_page", {"page_id": page_id})["outlines"][0]["items"][0]["handle"]


@pytest.mark.e2e
def test_outline_height_grows_with_its_content(mcp_session, claudespike_section_id):
    title = "test_outline_height_auto"
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": title})
    image = {"type": "inline_image", "handle": _image_handle(mcp_session, page_id, title),
             "width": 100.0, "height": 200.0}
    outlines = [
        {"position": {"x": 36.0, "y": 100.0}, "items": _paragraphs(1)},
        {"position": {"x": 300.0, "y": 100.0}, "items": _paragraphs(10)},
        {"position": {"x": 600.0, "y": 100.0}, "items": [image]},
    ]
    assert mcp_session.call_tool("replace_page", {"page_id": page_id, "title": title, "outlines": outlines}) == "ok"

    read = mcp_session.call_tool("get_page", {"page_id": page_id})["outlines"]
    one, ten, pictured = (outline["height"] for outline in read)
    assert one > 0
    assert ten == pytest.approx(10 * one, rel=0.02)
    assert pictured == pytest.approx(200.0, abs=0.5)

    # Written back as read, the heights are ignored and the layout is unchanged.
    assert mcp_session.call_tool("replace_page", {"page_id": page_id, "title": title, "outlines": read}) == "ok"
    reread = mcp_session.call_tool("get_page", {"page_id": page_id})["outlines"]
    assert [outline["height"] for outline in reread] == pytest.approx([one, ten, pictured], abs=0.5)
