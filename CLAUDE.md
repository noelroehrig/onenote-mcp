# CLAUDE.md

## Running commands

This repo uses a local `.venv`. Always call the venv's Python via its
**absolute path** (resolve it from the repo root) — bare `python`, bare
`pytest`, and relative forms like `.venv\Scripts\python.exe` or
`.\.venv\Scripts\python.exe` fail in this environment:

```
<repo-root>\.venv\Scripts\python.exe
```

### Common commands

| Task | Command (prefix with the absolute venv python) |
|------|------------------------------------------------|
| Unit tests | `-m pytest tests/unit/ -q` |
| E2E tests | `-m pytest -m e2e tests/e2e/ -v` |
| Collect tests | `-m pytest tests --collect-only` |
| Run a script | `scripts/<script>.py` |
| Run the server | `-m onenote_mcp.server` |

## Architecture

```
src/onenote_mcp/
  server.py    MCP tool definitions (FastMCP, stdio transport)
  com.py       COM wrapper: worker threads, timeouts, image-handle cache,
               notebook allowlist
  builders.py  model <-> OneNote XML conversion
  models.py    Pydantic models = the JSON schema of the structured tools
  images.py    pixel size from image header bytes (proportional image scaling)
```

Layering: `server.py` → `com.py` → `builders.py`/`models.py`, and `com.py` →
`images.py`. COM calls run on
fresh daemon threads with hard timeouts (`_run_com`); functions named `_*_impl`
run on the worker thread and must not call CoInitialize themselves.

## Tests

- Unit tests need no OneNote and must stay that way (fake `app` objects stand
  in for COM where needed).
- E2E tests drive the real MCP server against the real OneNote desktop app and
  create pages. Open the **ClaudeSpike** notebook in OneNote first; the suite
  skips automatically when OneNote is unreachable.
