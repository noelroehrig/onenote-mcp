"""Unit tests for the registered tools: which tools ONENOTE_DISABLE_RAW_XML
leaves, their generated schemas, and how the structured tools handle content
they cannot write.

Builds the server in-process the way main() does; no OneNote and no stdio
transport involved (COM entry points are stubbed where a tool would reach them).
"""
import asyncio

import pytest

try:
    from mcp.server.fastmcp.exceptions import ToolError
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


# ---------------------------------------------------------------------------
# Nested and unsupported content through the structured tools
# ---------------------------------------------------------------------------

def _input_schema(tool_name: str) -> dict:
    tools = asyncio.run(server.create_server().list_tools())
    return next(tool.inputSchema for tool in tools if tool.name == tool_name)


@pytest.mark.parametrize("tool_name", ["replace_page", "append_page"])
def test_write_tool_schemas_describe_nested_and_unsupported_items(tool_name):
    defs = _input_schema(tool_name)["$defs"]
    for model in ("Paragraph", "InlineImage", "ListItem", "UnsupportedItem"):
        assert "children" in defs[model]["properties"], model
    assert defs["UnsupportedItem"]["properties"]["kind"]["enum"] == ["table", "ink", "file", "media", "unknown"]


_UNSUPPORTED_OUTLINE = {"items": [
    {"type": "paragraph", "text": "kept", "children": [{"type": "unsupported", "kind": "table", "text": "a | b"}]},
]}


@pytest.mark.parametrize("tool_name, com_name, arguments", [
    ("replace_page", "_replace_page_com", {"page_id": "p", "title": "t", "outlines": [_UNSUPPORTED_OUTLINE]}),
    ("append_page", "_append_page_com", {"page_id": "p", "outline": _UNSUPPORTED_OUTLINE}),
])
def test_write_tools_reject_unsupported_items_before_touching_onenote(monkeypatch, tool_name, com_name, arguments):
    calls = []
    monkeypatch.setattr(server, com_name, lambda *args: calls.append(args))
    with pytest.raises(ToolError, match=r"bad_request: Cannot write an item of type 'unsupported' \(kind 'table'\)"):
        asyncio.run(server.create_server().call_tool(tool_name, arguments))
    assert calls == []


def test_only_the_structured_replace_keeps_unsupported_page_objects(monkeypatch):
    calls = []
    monkeypatch.setattr(server, "_replace_page_com", lambda *args: calls.append(args))
    page_xml = '<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"/>'
    asyncio.run(server.create_server().call_tool(
        "replace_page", {"page_id": "p", "title": "t", "outlines": []},
    ))
    asyncio.run(server.create_server().call_tool(
        "replace_page_xml", {"page_id": "p", "page_xml": page_xml},
    ))
    structured, raw = calls
    assert structured[0] == "p" and structured[2] == server.UNSUPPORTED_PAGE_OBJECT_TAGS
    assert raw == ("p", page_xml)


def test_replace_page_ignores_the_read_only_outline_height(monkeypatch):
    calls = []
    monkeypatch.setattr(server, "_replace_page_com", lambda *args: calls.append(args))
    read_outline = {"height": 134.27, "items": [{"type": "paragraph", "text": "x"}]}
    asyncio.run(server.create_server().call_tool(
        "replace_page", {"page_id": "p", "title": "t", "outlines": [read_outline]},
    ))
    ((_, page_xml, _),) = calls
    assert "Size" not in page_xml and "134.27" not in page_xml


def test_write_tool_schemas_have_no_outline_height():
    for tool_name in ("replace_page", "append_page"):
        assert "height" not in _input_schema(tool_name)["$defs"]["Outline"]["properties"]


def test_get_page_reports_nested_and_unsupported_content_compactly(monkeypatch):
    xml = (
        '<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">'
        "<one:Title><one:OE><one:T>Notes</one:T></one:OE></one:Title>"
        "<one:Outline><one:OEChildren>"
        '<one:OE style="font-size:20.0pt"><one:T>parent</one:T>'
        "<one:OEChildren><one:OE><one:T>child</one:T></one:OE></one:OEChildren></one:OE>"
        '<one:OE><one:InkWord recognizedText="scribble"><one:CallbackID callbackID="{1}"/></one:InkWord></one:OE>'
        "</one:OEChildren></one:Outline>"
        '<one:InkDrawing><one:Position x="1.0" y="2.0"/><one:CallbackID callbackID="{2}"/></one:InkDrawing>'
        "</one:Page>"
    )
    monkeypatch.setattr(server, "_get_page_com", lambda page_id, include_binary: xml)
    assert asyncio.run(server.get_page("p")) == {
        "title": "Notes",
        "outlines": [{"items": [
            {"type": "paragraph", "text": "parent", "font_size": 20.0,
             "children": [{"type": "paragraph", "text": "child"}]},
            {"type": "unsupported", "kind": "ink", "text": "scribble"},
        ]}],
        "unsupported": [{"kind": "ink", "position": {"x": 1.0, "y": 2.0}}],
    }
