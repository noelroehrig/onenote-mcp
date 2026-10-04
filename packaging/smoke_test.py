"""Smoke-test a built onenote-mcp.exe over MCP stdio.

Starts the exe twice. The default start runs initialize and tools/list, checks
that the exe exposes exactly the tools of the installed source package, and
calls ping. A second start with ONENOTE_DISABLE_RAW_XML=1 must expose exactly
the 9 tools without the raw-XML escape hatches. With --require-onenote the
default start also requires a live OneNote: ping must report
onenote_responsive and get_notebooks must succeed. Exits non-zero on a failed
check, a timeout, or a Python traceback on the server's stderr.

Usage: python packaging/smoke_test.py dist/onenote-mcp.exe [--require-onenote]
"""

import argparse
import asyncio
import json
import sys
import tempfile
import time
import traceback
from pathlib import Path

from mcp import ClientSession, StdioServerParameters, stdio_client
from mcp.types import CallToolResult

from onenote_mcp.server import create_server

TIMEOUT_S = 90
RESTRICTED_ENV = {"ONENOTE_DISABLE_RAW_XML": "1"}
RESTRICTED_TOOL_COUNT = 9


class SmokeTestError(Exception):
    """A check against the built exe failed."""


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise SmokeTestError(message)


def _require_tools(actual: set[str], expected: set[str], label: str) -> None:
    _require(
        actual == expected,
        f"{label}: tool set differs from source: missing={sorted(expected - actual)}, "
        f"unexpected={sorted(actual - expected)}",
    )


def _payload(result: CallToolResult, tool: str):
    """Return a tool's result value: structured content (unwrapping FastMCP's
    {"result": ...}) or, for tools without an output schema, the JSON text."""
    texts = [c.text for c in result.content if hasattr(c, "text")]
    if result.isError:
        raise SmokeTestError(f"{tool} returned an error: {'; '.join(texts)}")
    structured = result.structuredContent
    if structured is not None:
        return structured["result"] if set(structured) == {"result"} else structured
    _require(len(texts) == 1, f"{tool} returned {len(texts)} text blocks, expected 1")
    return json.loads(texts[0])


async def _source_tool_names(raw_xml_tools: bool) -> set[str]:
    return {tool.name for tool in await create_server(raw_xml_tools).list_tools()}


async def _query(exe: Path, env: dict[str, str] | None, tool_calls: list[str], errlog) -> dict:
    """Start the exe once, list its tools and call each of *tool_calls* without
    arguments. Checks happen afterwards, so a failed check is not wrapped in the
    MCP client's task-group exceptions."""
    params = StdioServerParameters(command=str(exe), args=[], env=env)
    started = time.perf_counter()
    async with stdio_client(params, errlog=errlog) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            observed = {
                "startup_s": time.perf_counter() - started,
                "tools": {tool.name for tool in (await session.list_tools()).tools},
            }
            for name in tool_calls:
                observed[name] = await session.call_tool(name, {})
    return observed


async def _query_both(exe: Path, require_onenote: bool, errlog) -> tuple[dict, dict]:
    calls = ["ping", "get_notebooks"] if require_onenote else ["ping"]
    default = await _query(exe, None, calls, errlog)
    restricted = await _query(exe, RESTRICTED_ENV, [], errlog)
    return default, restricted


def _check(default: dict, restricted: dict, require_onenote: bool) -> None:
    print(f"initialize: answered after {default['startup_s']:.1f}s")

    _require_tools(default["tools"], asyncio.run(_source_tool_names(True)), "default start")
    print(f"tools/list: {len(default['tools'])} tools match the source package")

    ping = _payload(default["ping"], "ping")
    _require(ping.get("server") == "ok", f"ping did not report server ok: {ping}")
    print(f"ping: {ping}")

    if require_onenote:
        _require(ping.get("onenote_responsive") is True, "OneNote is not responsive")
        notebooks = _payload(default["get_notebooks"], "get_notebooks")
        _require(isinstance(notebooks, list), f"get_notebooks returned {notebooks!r}")
        print(f"get_notebooks: {[nb['name'] for nb in notebooks]}")

    restricted_tools = restricted["tools"]
    _require_tools(restricted_tools, asyncio.run(_source_tool_names(False)), "ONENOTE_DISABLE_RAW_XML=1")
    _require(
        len(restricted_tools) == RESTRICTED_TOOL_COUNT,
        f"ONENOTE_DISABLE_RAW_XML=1: expected {RESTRICTED_TOOL_COUNT} tools, got {len(restricted_tools)}",
    )
    print(f"tools/list with ONENOTE_DISABLE_RAW_XML=1: {len(restricted_tools)} tools match the source package")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("exe", type=Path, help="path to onenote-mcp.exe")
    parser.add_argument(
        "--require-onenote",
        action="store_true",
        help="also require a responsive OneNote and a working get_notebooks",
    )
    args = parser.parse_args()
    if not args.exe.is_file():
        parser.error(f"executable not found: {args.exe}")

    with tempfile.TemporaryFile(mode="w+", encoding="utf-8", errors="replace") as errlog:
        failure = None
        try:
            default, restricted = asyncio.run(
                asyncio.wait_for(_query_both(args.exe, args.require_onenote, errlog), TIMEOUT_S)
            )
            _check(default, restricted, args.require_onenote)
        except asyncio.TimeoutError:
            failure = f"timed out after {TIMEOUT_S}s"
        except SmokeTestError as exc:
            failure = str(exc)
        except Exception:
            # e.g. the exe crashed on startup; the server stderr below says why.
            failure = f"MCP client error:\n{traceback.format_exc()}"

        errlog.seek(0)
        stderr = errlog.read()

    print("---- server stderr ----")
    print(stderr or "(empty)")
    if "Traceback (most recent call last)" in stderr:
        failure = failure or "Python traceback on server stderr"

    if failure:
        print(f"SMOKE TEST FAILED: {failure}", file=sys.stderr)
        return 1
    print("SMOKE TEST PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
