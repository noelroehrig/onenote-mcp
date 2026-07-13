"""Unit tests for com.py image-handle helpers — no COM dependency."""
import xml.etree.ElementTree as ET
import pytest

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_NS = {"one": _ONE_NS}


def _import_com_helpers():
    """Import com helpers, skipping if COM-related imports fail at module level."""
    from onenote_mcp.com import (
        _strip_data_to_mcpref, _resolve_mcpref_to_data, _IMAGE_CACHE, OneNoteError,
    )
    return _strip_data_to_mcpref, _resolve_mcpref_to_data, _IMAGE_CACHE, OneNoteError


# Try to import at module level. If comtypes/winreg aren't importable, skip all.
try:
    from onenote_mcp.com import (
        _strip_data_to_mcpref, _resolve_mcpref_to_data, _IMAGE_CACHE, OneNoteError,
    )
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not _COM_AVAILABLE,
    reason="com.py module-level imports (comtypes/winreg) unavailable in this environment",
)


@pytest.fixture(autouse=True)
def clear_cache():
    if _COM_AVAILABLE:
        _IMAGE_CACHE.clear()
    yield
    if _COM_AVAILABLE:
        _IMAGE_CACHE.clear()


def _make_page_xml(data_text: str) -> str:
    """Build a minimal OneNote page XML with a single Image/Data element."""
    ET.register_namespace("one", _ONE_NS)
    page = ET.Element(f"{{{_ONE_NS}}}Page")
    img = ET.SubElement(page, f"{{{_ONE_NS}}}Image")
    data = ET.SubElement(img, f"{{{_ONE_NS}}}Data")
    data.text = data_text
    return ET.tostring(page, encoding="unicode")


def _make_page_xml_two_images(data1: str, data2: str) -> str:
    ET.register_namespace("one", _ONE_NS)
    page = ET.Element(f"{{{_ONE_NS}}}Page")
    for data_text in (data1, data2):
        img = ET.SubElement(page, f"{{{_ONE_NS}}}Image")
        data = ET.SubElement(img, f"{{{_ONE_NS}}}Data")
        data.text = data_text
    return ET.tostring(page, encoding="unicode")


# ---------------------------------------------------------------------------
# _strip_data_to_mcpref
# ---------------------------------------------------------------------------

def test_strip_replaces_base64_with_mcpref_handle():
    xml = _make_page_xml("AAAA")
    result = _strip_data_to_mcpref(xml)
    root = ET.fromstring(result)
    data_el = root.find(f"{{{_ONE_NS}}}Image/{{{_ONE_NS}}}Data")
    assert data_el is not None
    assert data_el.text is not None
    assert data_el.text.startswith("mcpref:")
    # Cache should be populated
    handle = data_el.text[len("mcpref:"):]
    assert handle in _IMAGE_CACHE


def test_strip_leaves_existing_mcpref_handles():
    xml = _make_page_xml("mcpref:existinghandle1")
    result = _strip_data_to_mcpref(xml)
    root = ET.fromstring(result)
    data_el = root.find(f"{{{_ONE_NS}}}Image/{{{_ONE_NS}}}Data")
    assert data_el is not None
    assert data_el.text == "mcpref:existinghandle1"
    # Should not have added to cache
    assert "existinghandle1" not in _IMAGE_CACHE


def test_strip_empty_data_ignored():
    xml = _make_page_xml("")
    result = _strip_data_to_mcpref(xml)
    root = ET.fromstring(result)
    data_el = root.find(f"{{{_ONE_NS}}}Image/{{{_ONE_NS}}}Data")
    assert data_el is not None
    # Empty text should remain empty/unchanged (not replaced)
    text = data_el.text or ""
    assert not text.startswith("mcpref:")


def test_strip_multiple_images():
    xml = _make_page_xml_two_images("AAAA", "BBBB")
    result = _strip_data_to_mcpref(xml)
    root = ET.fromstring(result)
    data_els = root.findall(f"{{{_ONE_NS}}}Image/{{{_ONE_NS}}}Data")
    assert len(data_els) == 2
    handles = set()
    for data_el in data_els:
        assert data_el.text is not None
        assert data_el.text.startswith("mcpref:")
        handles.add(data_el.text[len("mcpref:"):])
    # Two different base64 strings → two different handles → two cache entries
    assert len(handles) == 2
    assert len(_IMAGE_CACHE) == 2


def test_strip_same_bytes_same_handle():
    xml = _make_page_xml_two_images("SAMEBYTES", "SAMEBYTES")
    result = _strip_data_to_mcpref(xml)
    root = ET.fromstring(result)
    data_els = root.findall(f"{{{_ONE_NS}}}Image/{{{_ONE_NS}}}Data")
    assert len(data_els) == 2
    handle1 = data_els[0].text
    handle2 = data_els[1].text
    # Same content → same sha1 handle (content-addressed)
    assert handle1 == handle2


# ---------------------------------------------------------------------------
# _resolve_mcpref_to_data
# ---------------------------------------------------------------------------

def test_resolve_known_handle():
    import hashlib
    original = "ORIGINALBASE64DATA"
    sha1_12 = hashlib.sha1(original.strip().encode()).hexdigest()[:12]
    _IMAGE_CACHE[sha1_12] = original

    xml = _make_page_xml(f"mcpref:{sha1_12}")
    result = _resolve_mcpref_to_data(xml)
    root = ET.fromstring(result)
    data_el = root.find(f"{{{_ONE_NS}}}Image/{{{_ONE_NS}}}Data")
    assert data_el is not None
    assert data_el.text == original


def test_resolve_leaves_non_mcpref_data():
    xml = _make_page_xml("plainbase64data")
    result = _resolve_mcpref_to_data(xml)
    root = ET.fromstring(result)
    data_el = root.find(f"{{{_ONE_NS}}}Image/{{{_ONE_NS}}}Data")
    assert data_el is not None
    assert data_el.text == "plainbase64data"


def test_resolve_unknown_handle_raises():
    xml = _make_page_xml("mcpref:unknownhandle1")
    with pytest.raises(OneNoteError):
        _resolve_mcpref_to_data(xml)


def test_round_trip():
    original = "MYROUNDTRIPBASE64=="
    xml = _make_page_xml(original)
    stripped = _strip_data_to_mcpref(xml)
    resolved = _resolve_mcpref_to_data(stripped)
    root = ET.fromstring(resolved)
    data_el = root.find(f"{{{_ONE_NS}}}Image/{{{_ONE_NS}}}Data")
    assert data_el is not None
    assert data_el.text == original


# ---------------------------------------------------------------------------
# _resolve_mcpref_to_data — outline-wrapped image paths (cross-page scenario)
# The helpers below use the Outline-wrapped XML structure that build_page_xml
# produces, exercising the exact code path used by the slim path.
# ---------------------------------------------------------------------------

_XMLNS = f'xmlns:one="{_ONE_NS}"'


def _data_xml_outline_wrapped(data_text: str) -> str:
    """Build page XML with an image nested inside Outline > OEChildren > OE."""
    return (
        f'<one:Page {_XMLNS}>'
        f'<one:Outline><one:OEChildren><one:OE><one:Image>'
        f'<one:Data>{data_text}</one:Data>'
        f'</one:Image></one:OE></one:OEChildren></one:Outline>'
        f'</one:Page>'
    )


def test_resolve_unknown_handle_raises_onenote_error():
    """_resolve_mcpref_to_data raises OneNoteError for unknown mcpref handles."""
    xml = _data_xml_outline_wrapped("mcpref:unknown123456")
    with pytest.raises(OneNoteError, match="Unknown image handle"):
        _resolve_mcpref_to_data(xml)


def test_resolve_known_handle_substitutes_base64():
    """_resolve_mcpref_to_data replaces a known handle with cached base64."""
    _IMAGE_CACHE["abc123def456"] = "base64imagedata=="
    xml = _data_xml_outline_wrapped("mcpref:abc123def456")
    resolved = _resolve_mcpref_to_data(xml)
    root = ET.fromstring(resolved)
    data_el = root.find(f".//{{{_ONE_NS}}}Data")
    assert data_el is not None
    assert data_el.text == "base64imagedata=="


def test_resolve_non_mcpref_content_passes_through():
    """_resolve_mcpref_to_data leaves real base64 content untouched."""
    xml = _data_xml_outline_wrapped("realbase64data==")
    resolved = _resolve_mcpref_to_data(xml)
    root = ET.fromstring(resolved)
    data_el = root.find(f".//{{{_ONE_NS}}}Data")
    assert data_el.text == "realbase64data=="


def test_resolve_multiple_handles_all_resolved():
    """_resolve_mcpref_to_data resolves all mcpref handles in one pass."""
    _IMAGE_CACHE["aaa111bbb222"] = "first_image_b64"
    _IMAGE_CACHE["ccc333ddd444"] = "second_image_b64"
    xml = (
        f'<one:Page {_XMLNS}>'
        '<one:Outline><one:OEChildren>'
        '<one:OE><one:Image><one:Data>mcpref:aaa111bbb222</one:Data></one:Image></one:OE>'
        '<one:OE><one:Image><one:Data>mcpref:ccc333ddd444</one:Data></one:Image></one:OE>'
        '</one:OEChildren></one:Outline>'
        '</one:Page>'
    )
    resolved = _resolve_mcpref_to_data(xml)
    root = ET.fromstring(resolved)
    data_els = root.findall(f".//{{{_ONE_NS}}}Data")
    texts = [el.text for el in data_els]
    assert "first_image_b64" in texts
    assert "second_image_b64" in texts


def test_resolve_unknown_handle_message_includes_handle():
    """OneNoteError message includes the bad handle and helpful guidance."""
    xml = _data_xml_outline_wrapped("mcpref:bad0handle0000")
    with pytest.raises(OneNoteError, match="mcpref:bad0handle0000"):
        _resolve_mcpref_to_data(xml)
