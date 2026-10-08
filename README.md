# OneNote MCP Server

A local MCP (Model Context Protocol) server that lets Claude Desktop, or any other
MCP client, read and edit pages in the **OneNote desktop app** on Windows through
OneNote's COM automation API.

No Azure app registration, no Microsoft Graph, no cloud authentication: the server
talks directly to the OneNote application running on your machine.

## Foreword

I'm a software developer, but not a Python developer, so I built this MCP server relying heavily on an AI coding agent. I needed a reliable MCP server to create and edit OneNote pages and couldn't find an existing solution that did this well, so I built my own. Other solutions seemed to focus on pulling page content out for search and summarization with an LLM.

I tried to keep token usage as low as possible and built in basic safeguards by [restricting notebook access](#configuration). I still recommend regular backups before you let an LLM access your notebooks.

I hope this helps someone in a similar situation. Feel free to use the contents of this repo in any way you like.

## Requirements

- Windows 10/11
- OneNote desktop (2016 or Microsoft 365). The Microsoft Store "OneNote for
  Windows 10" app has no COM interface and does not work.
- Python 3.10+, unless you use the [standalone executable](#standalone-executable)

## Installation

There are three ways to install the server:

- [Standalone executable](#standalone-executable): a single `onenote-mcp.exe`, no Python needed.
- [Install with pip](#install-with-pip): installs the `onenote-mcp` command into a Python venv.
- [From a clone](#from-a-clone): for development.

Then [check the setup](#check-the-setup).

### Standalone executable

Every [GitHub release](https://github.com/noelroehrig/onenote-mcp/releases) ships a
self-contained `onenote-mcp.exe` that needs no Python installation. The asset names
stay the same across releases; the version is the release tag.

- Executable: <https://github.com/noelroehrig/onenote-mcp/releases/latest/download/onenote-mcp.exe>
- Checksum: <https://github.com/noelroehrig/onenote-mcp/releases/latest/download/onenote-mcp.exe.sha256>

The exe is not code-signed. Windows SmartScreen may warn about an unrecognized app,
and some antivirus scanners flag self-extracting PyInstaller executables as a false
positive. Verify the checksum before you allow it; [Standalone executable
details](#standalone-executable-details) explains the trade-offs.

**Verify the checksum.** In PowerShell, in the download folder (prints `True` when
the file is intact):

```powershell
(Get-FileHash onenote-mcp.exe -Algorithm SHA256).Hash -eq (Get-Content onenote-mcp.exe.sha256).Split(' ')[0]
```

The `.sha256` file uses the `sha256sum` format, so `sha256sum -c onenote-mcp.exe.sha256`
works as well (for example in Git Bash).

**Configure Claude Desktop.** Move the exe to a permanent location, for example
`%LOCALAPPDATA%\onenote-mcp\`, and point `%APPDATA%\Claude\claude_desktop_config.json`
at it:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "C:\\Users\\<name>\\AppData\\Local\\onenote-mcp\\onenote-mcp.exe",
      "env": { "ONENOTE_ALLOWED_NOTEBOOKS": "SharedNotebook" }
    }
  }
}
```

`ONENOTE_ALLOWED_NOTEBOOKS` limits the server to the listed notebooks; see
[Configuration](#configuration) for all options.

### Install with pip

Install straight from GitHub into a fresh venv:

```
py -m venv %USERPROFILE%\onenote-mcp
%USERPROFILE%\onenote-mcp\Scripts\python.exe -m pip install git+https://github.com/noelroehrig/onenote-mcp
```

These are cmd commands; in PowerShell write `$env:USERPROFILE` instead of `%USERPROFILE%`.

Then point Claude Desktop at the installed `onenote-mcp.exe` console script. On a
machine with notebooks you care about, restrict the server to a dedicated notebook
right away:

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

### From a clone

```
git clone https://github.com/noelroehrig/onenote-mcp
cd onenote-mcp
py -m venv .venv
.venv\Scripts\python.exe -m pip install -e .
```

Add this to `%APPDATA%\Claude\claude_desktop_config.json`, with the absolute path to
your clone. Claude Desktop does not start servers from the repo directory, so relative
paths do not work.

```json
{
  "mcpServers": {
    "onenote": {
      "command": "C:\\Projects\\onenote-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-m", "onenote_mcp.server"]
    }
  }
}
```

Alternatively, set `command` to the absolute path of the console script
`.venv\Scripts\onenote-mcp.exe` inside the clone and leave out `args`.

### Other MCP clients

Any MCP client that starts stdio servers works the same way: use the `command`,
`args` and `env` values from the examples above.

### Check the setup

OneNote desktop must be running (or startable) with at least one notebook open.
Restart Claude Desktop; the OneNote tools should appear in the tools list. Then ask
Claude to call the `ping` tool: `{"server": "ok", "onenote_responsive": true, "config_error": null}`
means the server, the COM bridge and OneNote are all talking to each other. A
non-null `config_error` names a configuration problem (such as an unexpanded
`${...}` placeholder in `ONENOTE_ALLOWED_NOTEBOOKS`) that blocks the notebook
tools even though OneNote itself is reachable.

## Tools

### Navigation

| Tool | What it does |
| --- | --- |
| `get_notebooks` | Open notebooks with their sections (no pages): the small navigation skeleton. |
| `list_pages` | The pages of one section in section order, as `{id, name, level}`. `level` is 1 for a top-level page; a page's subpages are the pages that follow it with a higher level. |

The intended flow: `get_notebooks` → pick a section → `list_pages(section_id)` →
pick a page → read/write tools below.

### Reading

| Tool | What it does |
| --- | --- |
| `get_page` | A page as structured JSON: paragraphs, headings, lists, indented content, inline and floating images, plus read-only markers for [content the structured tools cannot write](#content-the-structured-tools-cannot-write). Image bytes are **not** included; images carry compact `mcpref:` handles instead. |
| `get_image_data` | The base64 bytes for a single image handle, fetched on demand. |
| `validate_handles` | Check whether image handles are still writable, without writing anything. Fetches and caches the bytes of handles not cached yet, so `true` means the bytes are in hand; `false` means the handle is unknown or OneNote no longer has the image. |

### Writing

| Tool | What it does |
| --- | --- |
| `create_page` | Create a page in a section, optionally as a sub-page of an existing page. Returns the new page id. |
| `replace_page` | Replace a page's entire content from structured JSON (outlines + floating images). Rejects read-only `unsupported` items and keeps page-level objects it cannot write in place, see below. |
| `append_page` | Append one structured content block to an existing page, preserving current content. Rejects read-only `unsupported` items. |

### Health

| Tool | What it does |
| --- | --- |
| `ping` | Fast health check: is the server alive, is OneNote responding, and is a configuration error blocking the notebook tools? Works even while another call is stuck on a wedged OneNote. |

### Raw-XML escape hatches

Prefer the structured tools above; these exist for direct schema control and debugging.
They cost significantly more tokens because of the XML overhead. Set
[`ONENOTE_DISABLE_RAW_XML`](#configuration) to `1` to remove them from the server.

| Tool | What it does |
| --- | --- |
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

## Content the structured tools cannot write

Tables, handwriting and drawings (ink), attached files, audio/video recordings and
content from newer OneNote versions have no structured item type. `get_page` still
reports them, so nothing disappears unnoticed:

- Inside an outline, each becomes a read-only item such as
  `{"type": "unsupported", "kind": "table", "text": "Name | Age\nAda | 36"}`.
  `kind` is `table`, `ink`, `file`, `media` or `unknown`; `text` is the readable
  text when OneNote provides it (table cells one row per line, recognized
  handwriting, a file's name).
- Objects placed directly on the page canvas, such as handwriting, are listed in
  the top-level `unsupported` list with their kind, position and size.

`replace_page` and `append_page` reject a payload that still contains an
`unsupported` item with `bad_request`, because rewriting the page would delete that
content. Remove the item deliberately only if deleting it is intended. The page-level
`unsupported` objects cannot be passed back at all, so `replace_page` keeps them in
place at their position: place new content where it does not overlap them.
`replace_page_xml` still replaces everything, since its caller supplies the full page.

## Error codes

Tool errors start with a machine-readable code so clients can react without parsing prose:

| Code | Meaning |
| --- | --- |
| `timeout` | OneNote did not respond within the per-operation deadline, most likely because it is showing a modal dialog or syncing. Dismiss any dialog and retry; `ping` tells you when it has recovered. |
| `backend_error` | OneNote returned a COM failure (bad id, locked content, …). |
| `bad_request` | The request itself was invalid (unknown image handle, malformed content, …). |
| `config_error` | The server's configuration is invalid, e.g. `ONENOTE_ALLOWED_NOTEBOOKS` still contains an unexpanded `${...}` placeholder. Fix the MCP client config and restart; `ping` shows the message. |
| `partial_write` | A replace wrote the new content, but some old page objects could not be deleted and are still on the page. The message lists each one's kind, object id and hresult. Read the page with `get_page` and decide how to handle them instead of retrying the replace. |

## Configuration

All configuration is via environment variables (set them in the `env` block of
the Claude Desktop server entry if needed):

| Variable | Default | Meaning |
| --- | --- | --- |
| `ONENOTE_ALLOWED_NOTEBOOKS` | _(unset: no restriction)_ | Comma-separated notebook names, e.g. `SharedNotebook, Project Notes`. When set, all other notebooks are invisible and untouchable: they are hidden from notebook listings and every read/write against their sections or pages is rejected. An empty value allows no notebook. A value that still contains `${` (a placeholder the client did not expand) is a configuration error: notebook tools fail with `config_error` and `ping` reports it. |
| `ONENOTE_READ_TIMEOUT` | `20` | Timeout (s) for page/hierarchy reads and image fetches. |
| `ONENOTE_WRITE_TIMEOUT` | `25` | Timeout (s) for page creates and writes. |
| `ONENOTE_PING_TIMEOUT` | `8` | Timeout (s) for the `ping` health probe. |
| `ONENOTE_IMAGE_CACHE_MB` | `200` | In-memory cap for cached image bytes; least-recently-used entries are evicted beyond it. |
| `ONENOTE_DISABLE_RAW_XML` | _(unset: raw-XML tools on)_ | `1` or `true` (any case): the four raw-XML tools (`list_hierarchy_xml`, `get_page_xml`, `replace_page_xml`, `append_page_xml`) are not registered and do not appear in the tool list. Unset, empty, `0` or `false`: all tools are available. Any other value stops the server at startup with an error. Read once at startup. |

## Standalone executable details

### Why a single file, and what it costs

The exe is a PyInstaller `--onefile` build because a single file is the easiest thing
to download, verify and reference from a config. A `--onedir` build starts faster and
usually trips antivirus less often, but it is a folder with the whole Python runtime
that has to be unzipped and kept together. The single file costs:

- **Startup time.** Every launch unpacks the bundled runtime into a temporary
  `%TEMP%\_MEI*` folder, which is removed on exit. The server answers after about
  1.5 s instead of about 0.7 s from a venv, and the first launch after a download can
  take longer while antivirus scans the file. Claude Desktop starts the server once
  per session, so this is paid once.
- **Antivirus false positives.** Self-extracting PyInstaller executables are a common
  antivirus false positive, and this exe is not code-signed. If your scanner
  quarantines it, check the SHA256 before allowing it, and consider reporting the false
  positive to the vendor. Windows SmartScreen may also warn about an unrecognized app
  when you start the exe from Explorer.

### COM inside the exe

comtypes generates Python wrappers for OneNote's COM type library the first time the
server connects to OneNote. The packages bundled in the exe are read-only, so comtypes
writes the wrappers to `%TEMP%\comtypes_cache\onenote-mcp-314` instead (the suffix is
the bundled Python version) and reuses them on later launches. They cannot be
generated at build time: the type library ships inside `ONENOTE.EXE`, which the build
machine does not have. In the exe, comtypes does not check whether the type library
changed, so if OneNote calls start failing after an Office update, delete that folder
and the wrappers are regenerated on the next start.

## Development

Source layout:

```
src/onenote_mcp/
  server.py    MCP tool definitions (FastMCP, stdio transport)
  com.py       COM wrapper: threading, timeouts, image-handle cache, notebook allowlist
  builders.py  model <-> OneNote XML conversion
  models.py    Pydantic models = the JSON schema of the structured tools
  images.py    pixel size from image header bytes (proportional image scaling)
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

End-to-end tests drive the real MCP server against the real OneNote desktop app.
They create pages and leave them in place. Open a notebook named **ClaudeSpike** in
OneNote first: without it, the tests write into the first section of the first open
notebook.

```
.venv\Scripts\python.exe -m pytest -m e2e tests/e2e/ -v
```

### Releases and the standalone exe

`.github/workflows/release.yml` runs the unit tests, builds `onenote-mcp.exe`,
smoke-tests it and keeps the exe as a workflow artifact for one day.

To release, set `version` in `pyproject.toml` and merge it to `main`. Then either:

- In GitHub, open Actions → Release → Run workflow on `main` and tick **Publish
  release**. After the tests and the smoke test pass, the run tags that commit
  (`v1.0.1` for version `1.0.1`) and publishes a GitHub release with the exe and its
  checksum. It refuses to publish from another branch or when the tag already exists.
- Or push the matching tag, e.g. `git tag v1.0.1` and `git push origin v1.0.1`. The
  build fails if the tag and the version differ.

A manual run without the checkbox only builds and tests.

To build locally, use a fresh venv so the build matches CI, with the PyInstaller
version pinned in the workflow:

```
py -m venv %TEMP%\onenote-mcp-build
%TEMP%\onenote-mcp-build\Scripts\python.exe -m pip install . pyinstaller==6.22.3
%TEMP%\onenote-mcp-build\Scripts\python.exe -m PyInstaller packaging/onenote-mcp.spec --noconfirm --clean
%TEMP%\onenote-mcp-build\Scripts\python.exe packaging/smoke_test.py dist/onenote-mcp.exe --require-onenote
```

The smoke test speaks MCP over stdio to the exe: it checks that the exe exposes the
same tools as the source and that `ping` answers, then starts it again with
`ONENOTE_DISABLE_RAW_XML=1` and checks for exactly the 9 tools without the raw-XML
ones. `--require-onenote` additionally requires a responsive OneNote and a working
`get_notebooks`; CI runs without it because the runner has no OneNote.

## License

MIT, see [LICENSE](LICENSE).
