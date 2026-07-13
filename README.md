# OneNote MCP Server

A local MCP (Model Context Protocol) server that lets Claude Desktop — or any MCP
client — read and edit pages in the **OneNote desktop app** on Windows through
OneNote's COM automation API.

No Azure app registration, no Microsoft Graph, no cloud authentication: the server
talks directly to the OneNote application running on your machine.

## Foreword

I'm a software developer but built this MCP server by heavily relying an AI Coding Agent since I'm not a python developer. I needed a reliable MCP server to edit and build OneNote pages and couldn't find a proper existing solution, therefore I built it myself. Other solutions seemed to focus on pulling out page content to leverage search and summarization with an LLM.

I tried to keep Token usage as low as possible and built in some bare minimum security by [restricting notebook access](#configuration). I still recommend using regular backups before allowing an LLM to access your Notebooks.

I hope this helps someone in a similar situation like me! Feel free to use the content of this Repo in any way you like.

## Requirements

- Windows 10/11
- OneNote desktop (2013, 2016, or Microsoft 365)
- Python 3.10+

## Installation

```
py -m venv .venv
.venv\Scripts\python.exe -m pip install -e .
```

## Claude Desktop configuration

Add this to `%APPDATA%\Claude\claude_desktop_config.json` (adjust the path to
where you cloned the repo):

```json
{
  "mcpServers": {
    "onenote": {
      "command": "C:\\Projects\\Personal\\OneNoteSpike\\.venv\\Scripts\\python.exe",
      "args": ["-m", "onenote_mcp.server"]
    }
  }
}
```

Alternatively, installation creates a console script, so `command` can simply be
`"C:\\Projects\\Personal\\OneNoteSpike\\.venv\\Scripts\\onenote-mcp.exe"` with no
`args` at all.

Restart Claude Desktop; the OneNote tools should appear in the tools list.
OneNote desktop must be running (or startable) with at least one notebook open.

## Installing on another machine

No clone needed — install straight from GitHub into a fresh venv:

```
py -m venv %USERPROFILE%\onenote-mcp
%USERPROFILE%\onenote-mcp\Scripts\python.exe -m pip install git+https://github.com/noelroehrig/onenote-mcp
```

Then point Claude Desktop at the installed console script. On a machine that
isn't yours, restrict the server to a dedicated notebook right away:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "C:\\Users\\<name>\\onenote-mcp\\Scripts\\onenote-mcp.exe",
      "env": { "ONENOTE_ALLOWED_NOTEBOOKS": "SharedNotebook" }
    }
  }
}
```

Verify the whole chain: restart Claude Desktop and ask Claude to call the
`ping` tool — `{"server": "ok", "onenote_responsive": true}` means the server,
COM bridge, and OneNote are all talking to each other.

Note: the Microsoft Store "OneNote for Windows 10" app has no COM interface;
the desktop edition (2016 or Microsoft 365) is required.

## Tools

### Navigation

| Tool | What it does |
|------|--------------|
| `get_notebooks` | Open notebooks with their sections (no pages) — the small navigation skeleton. |
| `list_pages` | The pages of one section as `{id, name}` pairs. |

The intended flow: `get_notebooks` → pick a section → `list_pages(section_id)` →
pick a page → read/write tools below.

### Reading

| Tool | What it does |
|------|--------------|
| `get_page` | A page as structured JSON: paragraphs, headings, lists, inline and floating images. Image bytes are **not** included — images carry compact `mcpref:` handles instead. |
| `get_image_data` | The base64 bytes for a single image handle, fetched on demand. |
| `validate_handles` | Check whether image handles are still writable, without writing anything. |

### Writing

| Tool | What it does |
|------|--------------|
| `create_page` | Create a page in a section; optionally as a sub-page of an existing page. Returns the new page id. |
| `replace_page` | Replace a page's entire content from structured JSON (outlines + floating images). |
| `append_page` | Append one structured content block to an existing page, preserving current content. |

### Health

| Tool | What it does |
|------|--------------|
| `ping` | Fast health check: is the server alive, and is OneNote responding? Works even while another call is stuck on a wedged OneNote. |

### Raw-XML escape hatches

Prefer the structured tools above; these exist for direct schema control and
debugging.

| Tool | What it does |
|------|--------------|
| `list_hierarchy_xml` | The entire notebook → section → page tree as raw OneNote XML (can be very large). |
| `get_page_xml` | The full `<one:Page>` XML of a page, with images as `mcpref:` handles. |
| `replace_page_xml` | Replace a page from a caller-supplied `<one:Page>` document. |
| `append_page_xml` | Append a single raw XML element (e.g. `<one:Outline>`) to a page. |

## Image handles (`mcpref:`)

Reading a page never ships image bytes to the client. Each image is represented
by a short opaque handle such as `mcpref:a1b2c3d4e5f6`, together with its size and
position. This keeps reads of image-heavy pages small and fast.

- To **preserve** images when rewriting a page, copy the handles (or the whole
  `images` list from `get_page`) verbatim into `replace_page` / `append_page`.
  The server resolves handles back to real bytes just before writing to OneNote.
- To **inspect** an image's pixels, call `get_image_data` with its handle.
- Handles live for the lifetime of the server process. If a handle has gone
  stale (e.g. after a restart), re-read the source page with `get_page` to mint
  fresh ones.

## Error codes

Tool errors start with a machine-readable code so clients can react without
parsing prose:

| Code | Meaning |
|------|---------|
| `timeout` | OneNote did not respond within the per-operation deadline — it is likely showing a modal dialog or syncing. Dismiss any dialog and retry; `ping` tells you when it has recovered. |
| `backend_error` | OneNote returned a COM failure (bad id, locked content, …). |
| `bad_request` | The request itself was invalid (unknown image handle, malformed content, …). |

## Configuration

All configuration is via environment variables (set them in the `env` block of
the Claude Desktop server entry if needed):

| Variable | Default | Meaning |
|----------|---------|---------|
| `ONENOTE_ALLOWED_NOTEBOOKS` | *(unset — no restriction)* | Comma-separated notebook names, e.g. `ClaudeSpike, Mathe 5a`. When set, all other notebooks are invisible and untouchable: they are hidden from notebook listings and every read/write against their sections or pages is rejected. |
| `ONENOTE_READ_TIMEOUT` | `20` | Timeout (s) for page/hierarchy reads and image fetches |
| `ONENOTE_WRITE_TIMEOUT` | `25` | Timeout (s) for page creates and writes |
| `ONENOTE_PING_TIMEOUT` | `8` | Timeout (s) for the `ping` health probe |
| `ONENOTE_IMAGE_CACHE_MB` | `200` | In-memory cap for cached image bytes; least-recently-used entries are evicted beyond it |

## Development

Source layout:

```
src/onenote_mcp/
  server.py    MCP tool definitions (FastMCP, stdio transport)
  com.py       COM wrapper: threading, timeouts, image-handle cache
  builders.py  model <-> OneNote XML conversion
  models.py    Pydantic models = the JSON schema of the structured tools
```

Run the server manually:

```
.venv\Scripts\python.exe -m onenote_mcp.server
```

### Tests

Unit tests (no OneNote required):

```
.venv\Scripts\python.exe -m pytest tests/unit/ -q
```

End-to-end tests — these drive the real MCP server against the real OneNote
desktop app and create pages in a test notebook. Open the **ClaudeSpike**
notebook in OneNote first (they fall back to the first open notebook otherwise):

```
.venv\Scripts\python.exe -m pytest -m e2e tests/e2e/ -v
```

## License

MIT — see [LICENSE](LICENSE).
