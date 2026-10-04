"""Unit tests for write preparation in com.py: partial image sizes and the
<one:Page> wrapper accepted by append.

A fake app stands in for OneNote.Application and records the XML it receives.
"""
import base64
import xml.etree.ElementTree as ET

import pytest

from tests.png import png_base64 as _png_base64

try:
    from onenote_mcp import com
    from onenote_mcp.com import OneNoteError, _complete_image_sizes
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not _COM_AVAILABLE,
    reason="com.py module-level imports (comtypes/winreg) unavailable in this environment",
)

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_XMLNS = f'xmlns:one="{_ONE_NS}"'


def _image_page(size_attrs: str, data: str) -> str:
    return (
        f'<one:Page {_XMLNS}><one:Outline><one:OEChildren><one:OE><one:Image>'
        f'<one:Size {size_attrs} isSetByUser="true"/><one:Data>{data}</one:Data>'
        f'</one:Image></one:OE></one:OEChildren></one:Outline></one:Page>'
    )


def _size(xml: str) -> dict[str, str]:
    return ET.fromstring(xml).find(f".//{{{_ONE_NS}}}Size").attrib


def test_width_only_gets_height_from_aspect_ratio():
    xml = _complete_image_sizes(_image_page('width="240"', _png_base64(200, 100)))
    assert _size(xml) == {"width": "240", "isSetByUser": "true", "height": "120.0"}


def test_height_only_gets_width_from_aspect_ratio():
    xml = _complete_image_sizes(_image_page('height="50"', _png_base64(200, 100)))
    assert _size(xml)["width"] == "100.0"


def test_complete_size_is_unchanged():
    xml = _complete_image_sizes(_image_page('width="240" height="10"', _png_base64(200, 100)))
    assert _size(xml) == {"width": "240", "height": "10", "isSetByUser": "true"}


def test_unreadable_image_with_partial_size_is_a_bad_request():
    with pytest.raises(OneNoteError) as exc:
        _complete_image_sizes(_image_page('width="240"', base64.b64encode(b"no image").decode()))
    assert exc.value.code == "bad_request"
    assert "Pass both width and height" in str(exc.value)


class _FakeApp:
    """Stand-in for OneNote.Application recording UpdatePageContent calls."""

    def __init__(self) -> None:
        self.updates: list[str] = []

    def UpdatePageContent(self, xml: str, date_expected_last_modified: float) -> None:
        self.updates.append(xml)


@pytest.fixture
def app(monkeypatch):
    monkeypatch.delenv("ONENOTE_ALLOWED_NOTEBOOKS", raising=False)
    fake = _FakeApp()
    monkeypatch.setattr(com, "_app", lambda: fake)
    return fake


def test_append_wraps_a_single_element_in_the_page(app):
    com._append_page_impl("{PAGE}", f'<one:Outline {_XMLNS} objectID="x"><one:OEChildren/></one:Outline>')

    (xml,) = app.updates
    page = ET.fromstring(xml)
    assert page.get("ID") == "{PAGE}"
    assert [child.tag for child in page] == [f"{{{_ONE_NS}}}Outline"]
    assert page[0].get("objectID") is None


def test_append_sends_the_children_of_a_page_wrapper(app):
    com._append_page_impl("{PAGE}", (
        f'<one:Page {_XMLNS}>'
        '<one:QuickStyleDef index="0" name="h2" font="Calibri" fontSize="14.0"/>'
        '<one:Outline><one:OEChildren><one:OE quickStyleIndex="0"><one:T>x</one:T></one:OE></one:OEChildren></one:Outline>'
        '</one:Page>'
    ))

    (xml,) = app.updates
    page = ET.fromstring(xml)
    assert page.get("ID") == "{PAGE}"
    assert [child.tag for child in page] == [f"{{{_ONE_NS}}}QuickStyleDef", f"{{{_ONE_NS}}}Outline"]


def test_append_completes_a_width_only_image(app):
    com._append_page_impl("{PAGE}", _image_page('width="240"', _png_base64(200, 100)))

    (xml,) = app.updates
    assert _size(xml)["height"] == "120.0"
