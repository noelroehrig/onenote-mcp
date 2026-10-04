"""Unit tests for OneNote COM server discovery when OneNote is not installed.

Points the registry lookup at GUID_NULL, which no COM server ever registers,
so the "not installed" path runs against the real registry without OneNote.
"""
import pytest

try:
    from onenote_mcp import com
    from onenote_mcp.com import OneNoteError, _discover_onenote_exe
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not _COM_AVAILABLE,
    reason="com.py module-level imports (comtypes/winreg) unavailable in this environment",
)

_UNREGISTERED_KEY = r"CLSID\{00000000-0000-0000-0000-000000000000}\LocalServer32"


@pytest.fixture
def onenote_not_installed(monkeypatch):
    monkeypatch.setattr(com, "_CLSID_LOCAL_SERVER", _UNREGISTERED_KEY)
    monkeypatch.setattr(com, "_onenote_mod", None)


def test_discover_raises_backend_error_when_not_registered(onenote_not_installed):
    with pytest.raises(OneNoteError) as exc:
        _discover_onenote_exe()
    assert exc.value.code == "backend_error"
    assert "not registered for COM automation" in str(exc.value)


def test_ping_reports_unresponsive_when_not_registered(onenote_not_installed):
    """ping keeps its contract (False, never raises) on machines without OneNote."""
    assert com.ping() is False
