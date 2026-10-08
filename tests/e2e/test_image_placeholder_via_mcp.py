"""E2E test: a screenshot placeholder reserves its space and reads back as itself.

replace_page writes an image_placeholder as a marked label OE plus a marked
box OE holding a light image of the requested size.  OneNote must keep the
OE Meta markers and the image's alt text, so get_page reports the same item,
and the box must occupy exactly width x height points.  Requires OneNote
desktop; the test page is left in place.
"""
import xml.etree.ElementTree as ET

import pytest

from tests.e2e.mcp_client import mcp_session  # noqa: F401 (re-export fixture)

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_NS = {"one": _ONE_NS}

_BOXED = {"type": "image_placeholder", "description": "Screenshot des Anmeldedialogs einfügen",
          "width": 400.0, "height": 250.0}
_LABEL_ONLY = {"type": "image_placeholder", "description": "Diagramm aus Kapitel 3"}
_BETWEEN = {"type": "paragraph", "text": "Text zwischen den Platzhaltern"}


def _approx_items(items: list[dict]) -> list[dict]:
    return [{**item, **{key: pytest.approx(item[key], abs=0.01) for key in ("width", "height") if key in item}}
            for item in items]


@pytest.mark.e2e
def test_image_placeholder_reserves_its_size_and_round_trips(mcp_session, claudespike_section_id):
    title = "test_image_placeholder_auto"
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": title})
    items = [_BOXED, _BETWEEN, _LABEL_ONLY]
    assert mcp_session.call_tool("replace_page", {
        "page_id": page_id, "title": title, "outlines": [{"items": items}],
    }) == "ok"

    (outline,) = mcp_session.call_tool("get_page", {"page_id": page_id})["outlines"]
    assert outline["items"] == _approx_items(items)
    # The box takes its full height on the page, below the label line.
    assert 250.0 < outline["height"] < 300.0

    page = ET.fromstring(mcp_session.call_tool("get_page_xml", {"page_id": page_id}))
    label_oe, box_oe = page.findall("one:Outline/one:OEChildren/one:OE", _NS)[:2]
    assert label_oe.find("one:Meta", _NS).attrib == {"name": "onenote-mcp.image-placeholder", "content": "label"}
    assert box_oe.find("one:Meta", _NS).attrib == {"name": "onenote-mcp.image-placeholder", "content": "box"}
    image = box_oe.find("one:Image", _NS)
    assert image.get("alt") == "Image placeholder"
    size = image.find("one:Size", _NS)
    assert (float(size.get("width")), float(size.get("height"))) == pytest.approx((400.0, 250.0), abs=0.01)

    # Read, write, read: the placeholders come back unchanged.
    assert mcp_session.call_tool("replace_page", {
        "page_id": page_id, "title": title, "outlines": [outline],
    }) == "ok"
    (reread,) = mcp_session.call_tool("get_page", {"page_id": page_id})["outlines"]
    assert reread["items"] == _approx_items(items)
