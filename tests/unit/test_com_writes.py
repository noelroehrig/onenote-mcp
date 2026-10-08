"""Unit tests for write preparation in com.py: partial image sizes, the
<one:Page> wrapper accepted by append, and the delete pass of replace.

A fake app stands in for OneNote.Application and records the XML it receives.
"""
import base64
import xml.etree.ElementTree as ET

import pytest

from tests.png import png_base64 as _png_base64

try:
    import comtypes
    from onenote_mcp import com
    from onenote_mcp.builders import UNSUPPORTED_PAGE_OBJECT_TAGS
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


# ---------------------------------------------------------------------------
# replace: delete pass, kept page objects and partial writes
# ---------------------------------------------------------------------------

_CURRENT_PAGE = (
    f'<one:Page {_XMLNS} ID="{{PAGE}}">'
    '<one:QuickStyleDef index="0" name="p"/>'
    '<one:Title objectID="title"><one:OE><one:T>Old</one:T></one:OE></one:Title>'
    '<one:Outline objectID="outline-1"><one:OEChildren><one:OE><one:T>a</one:T></one:OE></one:OEChildren></one:Outline>'
    '<one:InkDrawing objectID="ink-1"><one:CallbackID callbackID="{1}"/></one:InkDrawing>'
    '<one:Outline objectID="outline-2"><one:OEChildren><one:OE><one:T>b</one:T></one:OE></one:OEChildren></one:Outline>'
    '<one:InsertedFile objectID="file-1" preferredName="a.pdf"/>'
    '</one:Page>'
)
_NEW_PAGE = (
    f'<one:Page {_XMLNS}><one:Title><one:OE><one:T>New</one:T></one:OE></one:Title>'
    '<one:Outline><one:OEChildren><one:OE><one:T>c</one:T></one:OE></one:OEChildren></one:Outline></one:Page>'
)
_HRESULT = -2147213298  # 0x8004200E, hrPageObjectDoesNotExist


class _FakeReplaceApp:
    """Stand-in for OneNote.Application logging the replace calls in order.

    DeletePageContent raises a COMError for the ids in *undeletable*.
    """

    def __init__(self, undeletable: frozenset[str] = frozenset()) -> None:
        self.calls: list[tuple[str, str]] = []
        self._undeletable = undeletable

    def GetPageContent(self, page_id: str, page_info: int) -> str:
        return _CURRENT_PAGE

    def DeletePageContent(self, page_id: str, object_id: str, date_expected_last_modified: float) -> None:
        self.calls.append(("delete", object_id))
        if object_id in self._undeletable:
            raise comtypes.COMError(_HRESULT, "locked", None)

    def UpdatePageContent(self, xml: str, date_expected_last_modified: float) -> None:
        self.calls.append(("update", xml))


def _use(monkeypatch, fake) -> None:
    monkeypatch.delenv("ONENOTE_ALLOWED_NOTEBOOKS", raising=False)
    monkeypatch.setattr(com, "_app", lambda: fake)


def _deleted(fake: _FakeReplaceApp) -> list[str]:
    return [arg for call, arg in fake.calls if call == "delete"]


def test_replace_deletes_every_object_but_the_title_then_writes(monkeypatch):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset())

    assert _deleted(fake) == ["outline-1", "ink-1", "outline-2", "file-1"]
    assert [call for call, _ in fake.calls][-1] == "update"
    assert ET.fromstring(fake.calls[-1][1]).find(f"{{{_ONE_NS}}}Title").get("objectID") == "title"


def test_replace_keeps_objects_whose_tag_is_kept(monkeypatch):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    com._replace_page_impl("{PAGE}", _NEW_PAGE, UNSUPPORTED_PAGE_OBJECT_TAGS)

    assert _deleted(fake) == ["outline-1", "outline-2"]


def test_replace_attempts_every_delete_and_writes_before_reporting_a_partial_write(monkeypatch):
    fake = _FakeReplaceApp(undeletable=frozenset({"outline-1", "file-1"}))
    _use(monkeypatch, fake)
    with pytest.raises(OneNoteError) as exc:
        com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset())

    assert _deleted(fake) == ["outline-1", "ink-1", "outline-2", "file-1"]
    assert [call for call, _ in fake.calls][-1] == "update"
    assert exc.value.code == "partial_write"
    message = str(exc.value)
    assert message.startswith("partial_write: The new content was written, but 2 old page object(s)")
    assert "Outline outline-1 (hresult=-0x7ffbdff2)" in message
    assert "InsertedFile file-1 (hresult=-0x7ffbdff2)" in message
    assert "ink-1" not in message


def test_public_replace_page_passes_the_partial_write_through(monkeypatch):
    fake = _FakeReplaceApp(undeletable=frozenset({"outline-2"}))
    _use(monkeypatch, fake)
    with pytest.raises(OneNoteError) as exc:
        com.replace_page("{PAGE}", _NEW_PAGE)

    assert exc.value.code == "partial_write"
    assert "Outline outline-2" in str(exc.value)
