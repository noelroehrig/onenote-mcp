"""MCP stdio test harness for OneNote server E2E tests.

Spawns ``python -m onenote_mcp.server`` as a subprocess, connects to it via
the official MCP Python SDK's stdio client, and exposes a synchronous
``call_tool`` helper suitable for use in ordinary (non-async) pytest tests.

The async MCP client runs on a dedicated ``asyncio`` event loop in a
background daemon thread, keeping pytest fixtures simple and synchronous.

Usage
-----
Declare the fixture in your test module or conftest::

    @pytest.mark.e2e
    def test_smoke(mcp_session):
        result = mcp_session.call_tool("get_notebooks")
        assert isinstance(result, list)

Where ``result`` is the dict/str returned by the tool.  For tools that return
``dict``, the value is already parsed from the JSON text content the server
sends.  For tools that return ``str``, it is the raw string.
"""

from __future__ import annotations

import asyncio
import json
import sys
import threading
from concurrent.futures import Future
from pathlib import Path
from typing import Any

import pytest

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_REPO_ROOT = Path(__file__).parent.parent.parent
_PYTHON = str(_REPO_ROOT / ".venv" / "Scripts" / "python.exe")


def _parse_tool_result(result: Any) -> Any:
    """Extract a usable Python value from a ``CallToolResult``.

    Prefers ``structuredContent`` (the faithful representation of dict/list
    return values) so list-returning tools aren't collapsed to their first
    element. FastMCP wraps non-object returns (lists, scalars) under a single
    ``{"result": ...}`` key, which we unwrap. Falls back to parsing the text
    content blocks. Raises ``RuntimeError`` if the server signalled an error.
    """
    if result.isError:
        # Collect error text from content blocks
        texts = [c.text for c in result.content if hasattr(c, "text")]
        raise RuntimeError(f"MCP tool returned error: {'; '.join(texts)}")

    structured = getattr(result, "structuredContent", None)
    if structured is not None:
        if isinstance(structured, dict) and set(structured.keys()) == {"result"}:
            return structured["result"]
        return structured

    if not result.content:
        return None

    first = result.content[0]
    if not hasattr(first, "text"):
        # Return non-text content as-is (ImageContent, etc.)
        return first

    text = first.text
    try:
        return json.loads(text)
    except (json.JSONDecodeError, TypeError):
        return text


# ---------------------------------------------------------------------------
# Background-thread event loop + async client
# ---------------------------------------------------------------------------


class McpTestSession:
    """Synchronous facade over an async ``ClientSession`` running on a
    background thread's event loop.

    Do not instantiate directly — use the ``mcp_session`` pytest fixture.
    """

    def __init__(self, loop: asyncio.AbstractEventLoop, session: Any) -> None:
        self._loop = loop
        self._session = session

    def call_tool(self, name: str, arguments: dict[str, Any] | None = None) -> Any:
        """Call an MCP tool and return the parsed result (synchronous)."""
        future: Future[Any] = asyncio.run_coroutine_threadsafe(
            self._session.call_tool(name, arguments or {}),
            self._loop,
        )
        raw = future.result(timeout=60)
        return _parse_tool_result(raw)


def _run_loop_in_thread(loop: asyncio.AbstractEventLoop) -> None:
    asyncio.set_event_loop(loop)
    loop.run_forever()


async def _build_session(python_exe: str) -> tuple[Any, Any]:
    """Start the server subprocess and return ``(stdio_cm, session)``."""
    from mcp.client.session import ClientSession
    from mcp.client.stdio import StdioServerParameters, stdio_client

    params = StdioServerParameters(
        command=python_exe,
        args=["-m", "onenote_mcp.server"],
        cwd=str(_REPO_ROOT),
    )
    # We enter the context manager manually so we can yield from a sync fixture.
    cm = stdio_client(params)
    read_stream, write_stream = await cm.__aenter__()
    session = ClientSession(read_stream, write_stream)
    await session.__aenter__()
    await session.initialize()
    return cm, session


async def _teardown_session(cm: Any, session: Any) -> None:
    try:
        await session.__aexit__(None, None, None)
    except Exception:
        pass
    try:
        await cm.__aexit__(None, None, None)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# pytest fixture
# ---------------------------------------------------------------------------


@pytest.fixture(scope="session")
def mcp_session(onenote_available):  # noqa: F811
    """Session-scoped fixture: one MCP server process for the whole suite.

    Depends on ``onenote_available`` — the test is skipped automatically if
    OneNote desktop is not reachable.

    Yields a :class:`McpTestSession` with a synchronous ``call_tool`` method.
    """
    loop = asyncio.new_event_loop()
    thread = threading.Thread(target=_run_loop_in_thread, args=(loop,), daemon=True)
    thread.start()

    # Build the session on the background loop
    future: Future[tuple[Any, Any]] = asyncio.run_coroutine_threadsafe(
        _build_session(_PYTHON), loop
    )
    try:
        cm, session = future.result(timeout=30)
    except Exception as exc:
        loop.call_soon_threadsafe(loop.stop)
        thread.join(timeout=5)
        pytest.fail(f"Failed to start MCP server: {exc}")

    yield McpTestSession(loop, session)

    # Teardown
    teardown_future: Future[None] = asyncio.run_coroutine_threadsafe(
        _teardown_session(cm, session), loop
    )
    teardown_future.result(timeout=10)
    loop.call_soon_threadsafe(loop.stop)
    thread.join(timeout=5)
