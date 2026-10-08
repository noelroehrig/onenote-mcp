"""Unit tests for the ping tool: OneNote reachability and configuration errors
are reported side by side.  The COM probe is stubbed; no OneNote involved.
"""
import asyncio

import pytest

try:
    from onenote_mcp import server
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not _COM_AVAILABLE,
    reason="com.py module-level imports (comtypes/winreg) unavailable in this environment",
)


@pytest.fixture
def onenote_responsive(monkeypatch):
    monkeypatch.setattr(server, "_ping_com", lambda: True)


def test_ping_without_config_error(monkeypatch, onenote_responsive):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "ClaudeSpike")
    result = asyncio.run(server.ping())
    assert result == {"server": "ok", "onenote_responsive": True, "config_error": None}


def test_ping_reaches_onenote_and_reports_placeholder(monkeypatch, onenote_responsive):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "${user_config.allowed_notebooks}")
    result = asyncio.run(server.ping())
    assert result["server"] == "ok"
    assert result["onenote_responsive"] is True
    assert "ONENOTE_ALLOWED_NOTEBOOKS" in result["config_error"]
