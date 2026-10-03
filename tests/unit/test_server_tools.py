"""Unit tests for ONENOTE_DISABLE_RAW_XML: which tools the server registers.

Builds the server in-process the way main() does and lists its tools; no
OneNote and no stdio transport involved.
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

_RAW_XML_TOOLS = {"list_hierarchy_xml", "get_page_xml", "replace_page_xml", "append_page_xml"}
_STRUCTURED_TOOLS = {
    "ping", "validate_handles", "get_image_data", "get_notebooks", "list_pages",
    "get_page", "create_page", "replace_page", "append_page",
}


def _tool_names(environ: dict[str, str]) -> set[str]:
    raw_xml_disabled = server._raw_xml_disabled(environ)
    tools = asyncio.run(server.create_server(raw_xml_tools=not raw_xml_disabled).list_tools())
    return {tool.name for tool in tools}


@pytest.mark.parametrize("environ", [
    {},
    {"ONENOTE_DISABLE_RAW_XML": ""},
    {"ONENOTE_DISABLE_RAW_XML": " "},
    {"ONENOTE_DISABLE_RAW_XML": "0"},
    {"ONENOTE_DISABLE_RAW_XML": "false"},
    {"ONENOTE_DISABLE_RAW_XML": "FALSE"},
])
def test_all_13_tools_unless_disabled(environ):
    names = _tool_names(environ)
    assert names == _STRUCTURED_TOOLS | _RAW_XML_TOOLS
    assert len(names) == 13


@pytest.mark.parametrize("value", ["1", "true", "TRUE", "True", " true "])
def test_raw_xml_tools_not_registered_when_disabled(value):
    names = _tool_names({"ONENOTE_DISABLE_RAW_XML": value})
    assert names == _STRUCTURED_TOOLS
    assert len(names) == 9


@pytest.mark.parametrize("value", ["yes", "2", "off", "disabled"])
def test_invalid_value_refuses_to_start(monkeypatch, value):
    monkeypatch.setenv("ONENOTE_DISABLE_RAW_XML", value)
    with pytest.raises(SystemExit) as exc:
        server.main()
    message = str(exc.value.code)
    assert "ONENOTE_DISABLE_RAW_XML" in message
    assert repr(value) in message
