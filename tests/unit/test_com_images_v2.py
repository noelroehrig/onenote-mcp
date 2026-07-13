"""Unit tests for the no-binary image-handle path and structured error codes.

These exercise the callbackID→handle minting, lazy byte resolution, handle
validation, and OneNoteError code semantics — all without a live COM call.
A tiny fake ``app`` stands in for OneNote.Application where GetBinaryPageContent
is needed.
"""
import xml.etree.ElementTree as ET
import pytest

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

try:
    from onenote_mcp.com import (
        _strip_callbacks_to_mcpref,
        _resolve_handle_bytes,
        _resolve_mcpref_to_data,
        validate_handles,
        OneNoteError,
        _IMAGE_CACHE,
        _HANDLE_SOURCES,
        _MCPREF_PREFIX,
    )
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not _COM_AVAILABLE,
    reason="com.py module-level imports (comtypes/winreg) unavailable in this environment",
)


@pytest.fixture(autouse=True)
def clear_caches():
    if _COM_AVAILABLE:
        _IMAGE_CACHE.clear()
        _HANDLE_SOURCES.clear()
    yield
    if _COM_AVAILABLE:
        _IMAGE_CACHE.clear()
        _HANDLE_SOURCES.clear()


class _FakeApp:
    """Minimal stand-in for OneNote.Application.GetBinaryPageContent."""

    def __init__(self, mapping):
        self._mapping = mapping  # (page_id, callback_id) -> base64 str
        self.calls = []

    def GetBinaryPageContent(self, page_id, callback_id):
        self.calls.append((page_id, callback_id))
        return self._mapping[(page_id, callback_id)]


def _no_binary_image_xml(callback_id: str, *, with_size=True) -> str:
    """Build a no-binary page XML (image as <one:CallbackID>, no <one:Data>)."""
    size = (
        f'<one:Size width="100.0" height="80.0" isSetByUser="true"/>'
        if with_size else ""
    )
    return (
        f'<one:Page xmlns:one="{_ONE_NS}" ID="{{PAGE}}">'
        f'<one:Outline><one:OEChildren><one:OE><one:Image>'
        f'{size}<one:CallbackID callbackID="{callback_id}"/>'
        f'</one:Image></one:OE></one:OEChildren></one:Outline>'
        f'</one:Page>'
    )


# ---------------------------------------------------------------------------
# _strip_callbacks_to_mcpref
# ---------------------------------------------------------------------------

def test_callback_becomes_mcpref_handle_and_registers_source():
    page_id = "{PAGE-1}"
    cb = "{CB-AAA}{1}{B0}"
    xml = _no_binary_image_xml(cb)
    out = _strip_callbacks_to_mcpref(xml, page_id)

    root = ET.fromstring(out)
    data = root.find(f".//{{{_ONE_NS}}}Data")
    assert data is not None and data.text.startswith(_MCPREF_PREFIX)

    handle = data.text[len(_MCPREF_PREFIX):]
    assert _HANDLE_SOURCES[handle] == (page_id, cb)
    # No bytes are fetched on read — only the source is registered.
    assert handle not in _IMAGE_CACHE


def test_same_page_and_callback_yield_stable_handle():
    page_id, cb = "{PAGE-1}", "{CB-AAA}{1}{B0}"
    h1 = ET.fromstring(_strip_callbacks_to_mcpref(_no_binary_image_xml(cb), page_id))
    h2 = ET.fromstring(_strip_callbacks_to_mcpref(_no_binary_image_xml(cb), page_id))
    t1 = h1.find(f".//{{{_ONE_NS}}}Data").text
    t2 = h2.find(f".//{{{_ONE_NS}}}Data").text
    assert t1 == t2


def test_image_without_callback_is_left_alone():
    xml = (
        f'<one:Page xmlns:one="{_ONE_NS}"><one:Outline><one:OEChildren><one:OE>'
        f'<one:Image><one:Size width="10" height="10"/></one:Image>'
        f'</one:OE></one:OEChildren></one:Outline></one:Page>'
    )
    out = _strip_callbacks_to_mcpref(xml, "{PAGE-1}")
    root = ET.fromstring(out)
    assert root.find(f".//{{{_ONE_NS}}}Data") is None
    assert not _HANDLE_SOURCES


# ---------------------------------------------------------------------------
# Lazy resolution via GetBinaryPageContent
# ---------------------------------------------------------------------------

def test_resolve_handle_bytes_fetches_lazily_and_caches():
    page_id, cb = "{PAGE-1}", "{CB-AAA}{1}{B0}"
    xml = _strip_callbacks_to_mcpref(_no_binary_image_xml(cb), page_id)
    handle = ET.fromstring(xml).find(f".//{{{_ONE_NS}}}Data").text[len(_MCPREF_PREFIX):]

    app = _FakeApp({(page_id, cb): "BASE64BYTES=="})
    # First resolve fetches via COM…
    assert _resolve_handle_bytes(handle, app) == "BASE64BYTES=="
    assert app.calls == [(page_id, cb)]
    # …and caches, so a second resolve does not call COM again.
    assert _resolve_handle_bytes(handle, app) == "BASE64BYTES=="
    assert app.calls == [(page_id, cb)]


def test_resolve_handle_without_app_treats_known_source_as_unresolvable():
    page_id, cb = "{PAGE-1}", "{CB-AAA}{1}{B0}"
    xml = _strip_callbacks_to_mcpref(_no_binary_image_xml(cb), page_id)
    handle = ET.fromstring(xml).find(f".//{{{_ONE_NS}}}Data").text[len(_MCPREF_PREFIX):]
    with pytest.raises(OneNoteError) as exc:
        _resolve_handle_bytes(handle, app=None)
    assert exc.value.code == "bad_request"


def test_resolve_mcpref_to_data_uses_app_for_source_handles():
    page_id, cb = "{PAGE-1}", "{CB-AAA}{1}{B0}"
    stripped = _strip_callbacks_to_mcpref(_no_binary_image_xml(cb), page_id)
    handle_text = ET.fromstring(stripped).find(f".//{{{_ONE_NS}}}Data").text

    write_xml = (
        f'<one:Page xmlns:one="{_ONE_NS}"><one:Outline><one:OEChildren><one:OE>'
        f'<one:Image><one:Data>{handle_text}</one:Data></one:Image>'
        f'</one:OE></one:OEChildren></one:Outline></one:Page>'
    )
    app = _FakeApp({(page_id, cb): "RESOLVED=="})
    out = _resolve_mcpref_to_data(write_xml, app)
    assert ET.fromstring(out).find(f".//{{{_ONE_NS}}}Data").text == "RESOLVED=="


# ---------------------------------------------------------------------------
# validate_handles
# ---------------------------------------------------------------------------

def test_validate_handles_reports_cached_source_and_unknown():
    _IMAGE_CACHE["cached00aaaa"] = "bytes"
    _HANDLE_SOURCES["source00bbbb"] = ("{P}", "{CB}")
    result = validate_handles([
        "mcpref:cached00aaaa",
        "mcpref:source00bbbb",
        "mcpref:unknown00cccc",
        "source00bbbb",  # bare key without prefix also accepted
    ])
    assert result == {
        "mcpref:cached00aaaa": True,
        "mcpref:source00bbbb": True,
        "mcpref:unknown00cccc": False,
        "source00bbbb": True,
    }


# ---------------------------------------------------------------------------
# _IMAGE_CACHE eviction (LRU with a size cap)
# ---------------------------------------------------------------------------

def _put(handle: str, data: str, *, refetchable: bool) -> None:
    """Insert via the real cache API, optionally registering a lazy source."""
    from onenote_mcp import com
    if refetchable:
        _HANDLE_SOURCES[handle] = ("{P}", f"{{CB-{handle}}}")
    com._cache_put(handle, data)


def test_cache_evicts_refetchable_lru_first(monkeypatch):
    from onenote_mcp import com
    monkeypatch.setattr(com, "_IMAGE_CACHE_MAX_BYTES", 25)
    _put("refetch000001", "X" * 10, refetchable=True)   # oldest, re-fetchable
    _put("content000001", "Y" * 10, refetchable=False)  # content-addressed
    _put("newest0000001", "Z" * 10, refetchable=True)   # pushes size to 30 > 25
    # The oldest RE-FETCHABLE entry goes; the content-addressed one survives
    # even though it is older than the newcomer.
    assert "refetch000001" not in _IMAGE_CACHE
    assert "content000001" in _IMAGE_CACHE
    assert "newest0000001" in _IMAGE_CACHE
    # The evicted handle is still resolvable later — its source is remembered.
    assert "refetch000001" in _HANDLE_SOURCES


def test_cache_evicts_content_addressed_as_last_resort(monkeypatch):
    from onenote_mcp import com
    monkeypatch.setattr(com, "_IMAGE_CACHE_MAX_BYTES", 15)
    _put("content000001", "Y" * 10, refetchable=False)
    _put("content000002", "Z" * 10, refetchable=False)  # 20 > 15, nothing re-fetchable
    assert "content000001" not in _IMAGE_CACHE
    assert "content000002" in _IMAGE_CACHE


def test_cache_never_evicts_just_inserted_entry(monkeypatch):
    from onenote_mcp import com
    monkeypatch.setattr(com, "_IMAGE_CACHE_MAX_BYTES", 5)
    _put("oversized0001", "B" * 50, refetchable=False)  # alone exceeds the cap
    assert "oversized0001" in _IMAGE_CACHE


def test_cache_get_refreshes_lru_order(monkeypatch):
    from onenote_mcp import com
    monkeypatch.setattr(com, "_IMAGE_CACHE_MAX_BYTES", 25)
    _put("older00000001", "A" * 10, refetchable=True)
    _put("middle0000001", "B" * 10, refetchable=True)
    assert com._cache_get("older00000001") == "A" * 10  # touch → most recent
    _put("newest0000001", "C" * 10, refetchable=True)   # forces one eviction
    assert "middle0000001" not in _IMAGE_CACHE, "LRU victim should be the untouched entry"
    assert "older00000001" in _IMAGE_CACHE


def test_cache_put_is_a_replace_not_a_duplicate(monkeypatch):
    from onenote_mcp import com
    monkeypatch.setattr(com, "_IMAGE_CACHE_MAX_BYTES", 100)
    _put("replaced00001", "A" * 30, refetchable=False)
    _put("replaced00001", "B" * 40, refetchable=False)
    assert _IMAGE_CACHE["replaced00001"] == "B" * 40
    assert sum(len(v) for v in _IMAGE_CACHE.values()) == 40


# ---------------------------------------------------------------------------
# OneNoteError code semantics
# ---------------------------------------------------------------------------

def test_onenote_error_defaults_to_backend_error_and_prefixes_str():
    err = OneNoteError("boom")
    assert err.code == "backend_error"
    assert str(err) == "backend_error: boom"


def test_onenote_error_timeout_code_in_str():
    err = OneNoteError("no response", code="timeout")
    assert err.code == "timeout"
    assert str(err).startswith("timeout: ")
