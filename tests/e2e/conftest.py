"""Shared fixtures for E2E tests — all require OneNote desktop running."""
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "src"))

from onenote_mcp import com
from tests.e2e.mcp_client import mcp_session  # noqa: F401 — re-export for all e2e files

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_NS = {"one": _ONE_NS}

# ---------------------------------------------------------------------------
# OneNote availability check (done once at session start)
# ---------------------------------------------------------------------------

def _check_onenote() -> tuple[bool, str]:
    """Return (available, reason_if_not)."""
    try:
        xml = com.list_hierarchy()
        root = ET.fromstring(xml)
        nbs = root.findall(f"{{{_ONE_NS}}}Notebook")
        if not nbs:
            return False, "No notebooks found — open at least one notebook in OneNote desktop."
        return True, ""
    except Exception as exc:
        return False, f"OneNote not reachable: {exc}"

_ONENOTE_AVAILABLE, _ONENOTE_REASON = _check_onenote()

# ---------------------------------------------------------------------------
# Session-scoped fixtures
# ---------------------------------------------------------------------------

@pytest.fixture(scope="session")
def onenote_available():
    """Skip the test if OneNote is not reachable."""
    if not _ONENOTE_AVAILABLE:
        pytest.skip(f"OneNote not available: {_ONENOTE_REASON}")


@pytest.fixture(scope="session")
def claudespike_section_id(onenote_available):
    """Return the ID of the first section in the ClaudeSpike notebook.

    Falls back to the first section of the first notebook if ClaudeSpike is absent.
    """
    xml = com.list_hierarchy()
    root = ET.fromstring(xml)

    # Try ClaudeSpike first
    for nb in root.findall(f"{{{_ONE_NS}}}Notebook"):
        if nb.get("name") == "ClaudeSpike":
            sections = nb.findall(f".//{{{_ONE_NS}}}Section")
            if sections:
                return sections[0].get("ID")

    # Fall back to first notebook / first section
    nb = root.findall(f"{{{_ONE_NS}}}Notebook")[0]
    sections = nb.findall(f".//{{{_ONE_NS}}}Section")
    if not sections:
        pytest.skip("No sections found in any notebook.")
    return sections[0].get("ID")
