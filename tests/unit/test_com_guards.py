"""Unit tests for com.py request guards — XML validation and the notebook
allowlist.  No live COM: a tiny fake ``app`` stands in for OneNote.Application
where a hierarchy read is needed.
"""
import xml.etree.ElementTree as ET
import pytest

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

try:
    from onenote_mcp import com
    from onenote_mcp.com import (
        OneNoteError,
        _allowed_notebooks,
        _ensure_allowed,
        allowlist_config_error,
        _filter_notebooks_xml,
        _parse_user_xml,
    )
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not _COM_AVAILABLE,
    reason="com.py module-level imports (comtypes/winreg) unavailable in this environment",
)


# ---------------------------------------------------------------------------
# _parse_user_xml — malformed XML must map to bad_request, not raw ParseError
# ---------------------------------------------------------------------------

def test_parse_user_xml_accepts_well_formed():
    _parse_user_xml(f'<one:Outline xmlns:one="{_ONE_NS}"/>', "content")  # no raise


def test_parse_user_xml_rejects_malformed_as_bad_request():
    with pytest.raises(OneNoteError) as exc:
        _parse_user_xml("<one:Page unclosed", "page")
    assert exc.value.code == "bad_request"
    assert "Malformed page XML" in str(exc.value)


def test_parse_user_xml_rejects_multiple_roots():
    with pytest.raises(OneNoteError) as exc:
        _parse_user_xml("<a/><b/>", "content")
    assert exc.value.code == "bad_request"


def test_replace_page_rejects_malformed_xml_without_com():
    """The public write functions fail fast on bad XML — before any COM call."""
    with pytest.raises(OneNoteError) as exc:
        com.replace_page("{PAGE-1}", "<not xml")
    assert exc.value.code == "bad_request"


def test_append_page_rejects_malformed_xml_without_com():
    with pytest.raises(OneNoteError) as exc:
        com.append_page("{PAGE-1}", "no xml at all")
    assert exc.value.code == "bad_request"


# ---------------------------------------------------------------------------
# _allowed_notebooks — env parsing
# ---------------------------------------------------------------------------

def test_allowed_notebooks_unset_means_unrestricted(monkeypatch):
    monkeypatch.delenv("ONENOTE_ALLOWED_NOTEBOOKS", raising=False)
    assert _allowed_notebooks() is None


@pytest.mark.parametrize("value", ["", "   ", " , ,"])
def test_allowed_notebooks_empty_allows_nothing(monkeypatch, value):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", value)
    assert _allowed_notebooks() == set()


def test_allowed_notebooks_parses_and_strips_names(monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", " ClaudeSpike , Project Notes,,")
    assert _allowed_notebooks() == {"ClaudeSpike", "Project Notes"}


@pytest.mark.parametrize("value", ["${user_config.allowed_notebooks}", "ClaudeSpike, ${EXTRA}"])
def test_allowed_notebooks_unexpanded_placeholder_is_config_error(monkeypatch, value):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", value)
    with pytest.raises(OneNoteError) as exc:
        _allowed_notebooks()
    assert exc.value.code == "config_error"
    assert repr(value) in str(exc.value)


@pytest.mark.parametrize("value", [None, "", "ClaudeSpike", "Budget $5"])
def test_allowlist_config_error_none_without_placeholder(monkeypatch, value):
    if value is None:
        monkeypatch.delenv("ONENOTE_ALLOWED_NOTEBOOKS", raising=False)
    else:
        monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", value)
    assert allowlist_config_error() is None


def test_allowlist_config_error_names_variable_and_value(monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "${user_config.allowed_notebooks}")
    message = allowlist_config_error()
    assert message is not None
    assert "ONENOTE_ALLOWED_NOTEBOOKS" in message
    assert "'${user_config.allowed_notebooks}'" in message


# ---------------------------------------------------------------------------
# _filter_notebooks_xml
# ---------------------------------------------------------------------------

_HIERARCHY_XML = (
    f'<one:Notebooks xmlns:one="{_ONE_NS}">'
    '<one:Notebook ID="nb-allowed" name="ClaudeSpike">'
    '  <one:Section ID="sec-allowed" name="S1">'
    '    <one:Page ID="page-allowed" name="P1"/>'
    '  </one:Section>'
    '</one:Notebook>'
    '<one:Notebook ID="nb-secret" name="Private">'
    '  <one:Section ID="sec-secret" name="Secret">'
    '    <one:Page ID="page-secret" name="Diary"/>'
    '  </one:Section>'
    '</one:Notebook>'
    '</one:Notebooks>'
)


def test_filter_notebooks_noop_when_unrestricted(monkeypatch):
    monkeypatch.delenv("ONENOTE_ALLOWED_NOTEBOOKS", raising=False)
    assert _filter_notebooks_xml(_HIERARCHY_XML) == _HIERARCHY_XML


def test_filter_notebooks_drops_disallowed(monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "ClaudeSpike")
    out = _filter_notebooks_xml(_HIERARCHY_XML)
    root = ET.fromstring(out)
    names = [nb.get("name") for nb in root.findall(f"{{{_ONE_NS}}}Notebook")]
    assert names == ["ClaudeSpike"]
    assert "Diary" not in out


def test_filter_notebooks_empty_allowlist_hides_all(monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "")
    root = ET.fromstring(_filter_notebooks_xml(_HIERARCHY_XML))
    assert root.findall(f"{{{_ONE_NS}}}Notebook") == []


def test_filter_notebooks_placeholder_raises_config_error(monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "${user_config.allowed_notebooks}")
    with pytest.raises(OneNoteError) as exc:
        _filter_notebooks_xml(_HIERARCHY_XML)
    assert exc.value.code == "config_error"


# ---------------------------------------------------------------------------
# _ensure_allowed
# ---------------------------------------------------------------------------

class _FakeApp:
    """Stand-in for OneNote.Application providing only GetHierarchy."""

    def __init__(self, xml: str) -> None:
        self._xml = xml
        self.calls = 0

    def GetHierarchy(self, anchor: str, scope: int) -> str:
        self.calls += 1
        return self._xml


def test_ensure_allowed_noop_without_allowlist(monkeypatch):
    monkeypatch.delenv("ONENOTE_ALLOWED_NOTEBOOKS", raising=False)
    app = _FakeApp(_HIERARCHY_XML)
    _ensure_allowed(app, "page-secret", "Page")  # no raise
    assert app.calls == 0, "unrestricted mode must not cost a hierarchy read"


def test_ensure_allowed_passes_for_allowed_page_and_section(monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "ClaudeSpike")
    app = _FakeApp(_HIERARCHY_XML)
    _ensure_allowed(app, "page-allowed", "Page")
    _ensure_allowed(app, "sec-allowed", "Section")


def test_ensure_allowed_rejects_disallowed_page(monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "ClaudeSpike")
    app = _FakeApp(_HIERARCHY_XML)
    with pytest.raises(OneNoteError) as exc:
        _ensure_allowed(app, "page-secret", "Page")
    assert exc.value.code == "bad_request"
    assert "ONENOTE_ALLOWED_NOTEBOOKS" in str(exc.value)


def test_ensure_allowed_rejects_disallowed_section(monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "ClaudeSpike")
    app = _FakeApp(_HIERARCHY_XML)
    with pytest.raises(OneNoteError) as exc:
        _ensure_allowed(app, "sec-secret", "Section")
    assert exc.value.code == "bad_request"


def test_ensure_allowed_rejects_unknown_id(monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "ClaudeSpike")
    app = _FakeApp(_HIERARCHY_XML)
    with pytest.raises(OneNoteError) as exc:
        _ensure_allowed(app, "page-nonexistent", "Page")
    assert exc.value.code == "bad_request"
    assert "not found" in str(exc.value)


def test_ensure_allowed_empty_allowlist_rejects_everything(monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "")
    app = _FakeApp(_HIERARCHY_XML)
    for object_id, kind in [("page-allowed", "Page"), ("sec-allowed", "Section")]:
        with pytest.raises(OneNoteError) as exc:
            _ensure_allowed(app, object_id, kind)
        assert exc.value.code == "bad_request"


def test_ensure_allowed_placeholder_raises_config_error_without_com(monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "${user_config.allowed_notebooks}")
    app = _FakeApp(_HIERARCHY_XML)
    with pytest.raises(OneNoteError) as exc:
        _ensure_allowed(app, "page-allowed", "Page")
    assert exc.value.code == "config_error"
    assert app.calls == 0
