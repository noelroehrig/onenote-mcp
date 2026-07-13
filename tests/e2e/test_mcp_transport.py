"""Smoke test: MCP stdio transport wire-up through server.py.

Calls ``get_notebooks`` through the real stdio transport and asserts the
basic shape of the response.  This proves the MCP wire works end-to-end
without any mocking.
"""

import pytest

from tests.e2e.mcp_client import mcp_session  # noqa: F401 — re-export fixture


@pytest.mark.e2e
def test_get_notebooks_via_mcp_transport(mcp_session):
    """get_notebooks over real stdio MCP returns a non-empty notebook skeleton."""
    result = mcp_session.call_tool("get_notebooks")

    assert isinstance(result, list), (
        f"Expected list from get_notebooks, got {type(result).__name__}: {result!r}"
    )
    assert len(result) > 0, (
        "Expected at least one notebook — open a notebook in OneNote desktop"
    )
    first = result[0]
    assert "id" in first and "name" in first
    assert "sections" in first and isinstance(first["sections"], list), (
        f"Each notebook must carry a 'sections' list, got keys: {list(first.keys())}"
    )
