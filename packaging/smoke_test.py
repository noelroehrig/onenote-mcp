"""Smoke-test a built onenote-mcp.exe over MCP stdio.

Runs initialize and tools/list, checks that the exe exposes exactly the tools of
the installed source package, and calls ping. With --require-onenote it also
requires a live OneNote: ping must report onenote_responsive and get_notebooks
must succeed. Exits non-zero on a failed check, a timeout, or a Python
traceback on the server's stderr.

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

from onenote_mcp.server import mcp as source_server

TIMEOUT_S = 90


class SmokeTestError(Exception):
    """A check against the built exe failed."""


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise SmokeTestError(message)


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


async def _query(exe: Path, require_onenote: bool, errlog) -> dict:
    """Talk to the exe and collect its responses. Checks happen afterwards, so a
    failed check is not wrapped in the MCP client's task-group exceptions."""
    params = StdioServerParameters(command=str(exe), args=[])
    started = time.perf_counter()
    async with stdio_client(params, errlog=errlog) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            observed = {
                "startup_s": time.perf_counter() - started,
                "tools": {tool.name for tool in (await session.list_tools()).tools},
                "ping": await session.call_tool("ping", {}),
            }
            if require_onenote:
                observed["get_notebooks"] = await session.call_tool("get_notebooks", {})
    return observed


def _check(observed: dict, expected_tools: set[str], require_onenote: bool) -> None:
    print(f"initialize: answered after {observed['startup_s']:.1f}s")

    tools = observed["tools"]
    _require(
        tools == expected_tools,
        f"tool set differs from source: missing={sorted(expected_tools - tools)}, "
        f"unexpected={sorted(tools - expected_tools)}",
    )
    print(f"tools/list: {len(tools)} tools match the source package")

    ping = _payload(observed["ping"], "ping")
    _require(ping.get("server") == "ok", f"ping did not report server ok: {ping}")
    print(f"ping: {ping}")

    if require_onenote:
        _require(ping.get("onenote_responsive") is True, "OneNote is not responsive")
        notebooks = _payload(observed["get_notebooks"], "get_notebooks")
        _require(isinstance(notebooks, list), f"get_notebooks returned {notebooks!r}")
        print(f"get_notebooks: {[nb['name'] for nb in notebooks]}")


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

    expected_tools = {tool.name for tool in asyncio.run(source_server.list_tools())}

    with tempfile.TemporaryFile(mode="w+", encoding="utf-8", errors="replace") as errlog:
        failure = None
        try:
            observed = asyncio.run(
                asyncio.wait_for(_query(args.exe, args.require_onenote, errlog), TIMEOUT_S)
            )
            _check(observed, expected_tools, args.require_onenote)
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
