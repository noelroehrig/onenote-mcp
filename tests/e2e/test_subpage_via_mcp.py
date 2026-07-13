"""E2E test for sub-page creation via create_page(parent_page_id=...).

Goes through server.py over the MCP stdio transport; requires OneNote desktop.
"""
import xml.etree.ElementTree as ET

import pytest

from onenote_mcp import com
from tests.e2e.mcp_client import mcp_session  # noqa: F401 — re-export for fixture

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _ordered_pages(section_id: str) -> list[tuple[str, str | None]]:
    """Return (page_id, pageLevel) in section order, from the COM hierarchy."""
    root = ET.fromstring(com.list_hierarchy())
    for section in root.iter(f"{{{_ONE_NS}}}Section"):
        if section.get("ID") == section_id:
            return [
                (p.get("ID"), p.get("pageLevel"))
                for p in section.findall(f"{{{_ONE_NS}}}Page")
            ]
    return []


def _page_level(section_id: str, page_id: str) -> str | None:
    return next((lvl for pid, lvl in _ordered_pages(section_id) if pid == page_id), None)


@pytest.mark.e2e
def test_create_page_as_subpage_indents_and_appends_last(mcp_session, claudespike_section_id):
    parent_id = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "test_subpage_parent_auto",
    })
    child1 = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "test_subpage_child1_auto",
        "parent_page_id": parent_id,
    })
    child2 = mcp_session.call_tool("create_page", {
        "section_id": claudespike_section_id,
        "title": "test_subpage_child2_auto",
        "parent_page_id": parent_id,
    })

    assert _page_level(claudespike_section_id, parent_id) in (None, "1")
    assert _page_level(claudespike_section_id, child1) == "2"
    assert _page_level(claudespike_section_id, child2) == "2"

    order = [pid for pid, _ in _ordered_pages(claudespike_section_id)]
    i_parent, i_c1, i_c2 = order.index(parent_id), order.index(child1), order.index(child2)
    # parent, then first child, then the SECOND child appended after it (last).
    assert i_parent < i_c1 < i_c2, f"expected parent<child1<child2, got {order}"


@pytest.mark.e2e
def test_create_page_unknown_parent_returns_bad_request(mcp_session, claudespike_section_id):
    with pytest.raises(RuntimeError, match="bad_request"):
        mcp_session.call_tool("create_page", {
            "section_id": claudespike_section_id,
            "title": "test_subpage_badparent_auto",
            "parent_page_id": "{00000000-0000-0000-0000-000000000000}{1}{B0}",
        })
