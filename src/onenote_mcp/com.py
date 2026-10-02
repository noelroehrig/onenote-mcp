"""Thin COM wrapper around OneNote.Application using comtypes.

Strategy
--------
- Discover the OneNote EXE path at runtime from the registry (never hard-code).
- Load the typelib directly from the EXE (resource index \\3) via
  comtypes.typeinfo.LoadTypeLibEx, then pass the ITypeLib pointer to
  comtypes.client.GetModule.  This bypasses the broken Win32-only registry
  pointer that prevents 64-bit pywin32 from loading the typelib.
- Use comtypes.client.CreateObject(mod.Application) per public call to get a
  typed wrapper (cheap: just attaches to the already-running OneNote process).
- _run_com spawns a fresh daemon thread per call: CoInitialize/CoUninitialize
  happen inside the worker (STA apartment is per-thread).  A hard, per-operation
  timeout ensures a hung COM call surfaces as a structured OneNoteError instead
  of blocking forever.

Image handling
--------------
get_page reads pages WITHOUT binary image bytes by default (pageInfo=0).  In
that mode OneNote returns each image as a <one:CallbackID> reference.  We mint a
stable ``mcpref:`` handle for each and remember its (page_id, callbackID) source
so the bytes can be fetched lazily via GetBinaryPageContent only when a handle is
actually written back.  This keeps reads of image-heavy pages small and fast and
avoids shipping megabytes of base64 to the model on every read.  Resolved bytes
are kept in an LRU cache capped at ONENOTE_IMAGE_CACHE_MB (default 200 MB).

Access restriction
------------------
When the ONENOTE_ALLOWED_NOTEBOOKS environment variable is set (comma-separated
notebook display names), hierarchy listings only show those notebooks and every
id-scoped read/write verifies that its target lives inside one of them.  Unset
or blank means unrestricted.
"""

import hashlib
import os
import sys
import threading
import winreg
import re
import xml.etree.ElementTree as ET
from collections import OrderedDict

import comtypes
import comtypes.client
import comtypes.typeinfo

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
ET.register_namespace("one", _ONE_NS)

_MCPREF_PREFIX = "mcpref:"

# Process-lifetime caches.
#   _IMAGE_CACHE:    handle -> resolved base64 bytes (the writable payload).
#                    LRU-ordered and capped at _IMAGE_CACHE_MAX_BYTES; access
#                    it through _cache_put/_cache_get so eviction stays correct.
#   _HANDLE_SOURCES: handle -> (page_id, callback_id) for lazy byte resolution.
# A handle is resolvable if it appears in either map; _IMAGE_CACHE wins.
_IMAGE_CACHE: OrderedDict[str, str] = OrderedDict()
_HANDLE_SOURCES: dict[str, tuple[str, str]] = {}
_CACHE_LOCK = threading.Lock()


def _env_float(name: str, default: float) -> float:
    """Read a positive float from the environment, falling back to *default*."""
    raw = os.environ.get(name)
    if raw is None:
        return default
    try:
        value = float(raw)
        return value if value > 0 else default
    except ValueError:
        return default


# Per-operation hard timeouts (seconds).  Reads are tiny (~50 ms in practice) so
# a short ceiling fails fast on a wedged OneNote; writes (delete loop +
# UpdatePageContent) get more headroom; the health probe must return quickly
# even while another call is stuck.  All three are overridable via environment
# variables for unusual setups.
READ_TIMEOUT = _env_float("ONENOTE_READ_TIMEOUT", 20.0)
WRITE_TIMEOUT = _env_float("ONENOTE_WRITE_TIMEOUT", 25.0)
PING_TIMEOUT = _env_float("ONENOTE_PING_TIMEOUT", 8.0)

# Cap for _IMAGE_CACHE, in base64 characters (~ bytes).  When exceeded, least-
# recently-used entries are evicted — re-fetchable ones first (their bytes can
# be pulled again via GetBinaryPageContent); content-addressed entries from
# binary reads only as a last resort, since evicting those turns their handles
# stale until the source page is re-read.
_IMAGE_CACHE_MAX_BYTES = int(_env_float("ONENOTE_IMAGE_CACHE_MB", 200.0) * 1024 * 1024)

# ---- EXE / typelib discovery ------------------------------------------------

_CLSID_LOCAL_SERVER = r"CLSID\{DC67E480-C3CB-49F8-8232-60B0C2056C8E}\LocalServer32"


def _discover_onenote_exe() -> str:
    """Return the OneNote EXE path from HKCR LocalServer32.

    The registry value may be:
      bare:          C:\\...\\ONENOTE.EXE
      quoted:        "C:\\...\\ONENOTE.EXE"
      quoted + args: "C:\\...\\ONENOTE.EXE" /something

    Raises OneNoteError when the key is missing, i.e. OneNote desktop is not
    installed (the Store app registers no COM server).
    """
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, _CLSID_LOCAL_SERVER) as key:
            value, _ = winreg.QueryValueEx(key, None)
    except FileNotFoundError as exc:
        raise OneNoteError(
            "OneNote desktop is not registered for COM automation. Install the "
            "desktop edition (2016 or Microsoft 365); the Microsoft Store app "
            "has no COM interface."
        ) from exc
    value = value.strip()
    if value.startswith('"'):
        # Quoted path: "C:\...\ONENOTE.EXE" or "C:\...\ONENOTE.EXE" /args
        m = re.match(r'"([^"]+)"', value)
        if m:
            return m.group(1)
    # Unquoted: the value is either a bare path (may contain spaces) or
    # a path followed by args.  Look for ".EXE" (case-insensitive) as the
    # natural end of the executable path.
    m = re.match(r'(.+?\.exe)', value, re.IGNORECASE)
    if m:
        return m.group(1)
    # Last resort: return as-is
    return value


# Cache the module reference for the process lifetime.
_onenote_mod = None


def _ensure_module():
    """Load the OneNote typelib from the EXE and generate/cache the wrapper module."""
    global _onenote_mod
    if _onenote_mod is None:
        exe_path = _discover_onenote_exe()
        # Resource index \3 contains the OneNote typelib
        typelib_path = exe_path + "\\3"
        tlib = comtypes.typeinfo.LoadTypeLibEx(typelib_path)
        _onenote_mod = comtypes.client.GetModule(tlib)
    return _onenote_mod


# ---- Error type -------------------------------------------------------------

class OneNoteError(Exception):
    """Raised when a OneNote COM call fails.

    Carries a machine-readable ``code`` so callers (and the model) can tell a
    transient timeout apart from a hard backend error without parsing prose:

      timeout        OneNote did not respond within the per-operation deadline.
      backend_error  OneNote returned a COM failure (bad id, locked content, …).
      bad_request    The request itself was malformed (e.g. unknown handle).
    """

    def __init__(self, message: str, code: str = "backend_error") -> None:
        super().__init__(message)
        self.code = code

    def __str__(self) -> str:
        # Prefix the code so it survives transport as plain error text, e.g.
        # "timeout: OneNote did not respond within 20s …".
        return f"{self.code}: {super().__str__()}"


# ---- Internal helpers -------------------------------------------------------

def _app():
    """Return a typed OneNote.Application COM object."""
    mod = _ensure_module()
    return comtypes.client.CreateObject(mod.Application)


def _run_com(fn, *args, timeout=READ_TIMEOUT):
    """Run fn(*args) in a fresh daemon thread with a hard timeout.

    A new thread is spawned per call so that a stuck call can be abandoned
    safely (daemon threads die when the process exits).  CoInitialize and
    CoUninitialize are called inside the worker so that the STA apartment is
    set up correctly on the worker thread — COM apartments are per-thread.

    Returns the value returned by fn(*args) on success.
    Re-raises the original exception (preserving type and traceback) if fn
    raises.
    Raises OneNoteError(code="timeout") if the call does not complete within
    *timeout* seconds.  The worker thread is abandoned (it cannot be killed
    while blocked inside a COM call) but, being a daemon, will not keep the
    process alive.
    """
    _result = [None]
    _exc_info = [None]
    done = threading.Event()

    def worker():
        comtypes.CoInitialize()
        try:
            _result[0] = fn(*args)
        except BaseException:  # noqa: BLE001 — capture everything to re-raise on the caller
            _exc_info[0] = sys.exc_info()
        finally:
            comtypes.CoUninitialize()
            done.set()

    t = threading.Thread(target=worker, daemon=True)
    t.start()
    if not done.wait(timeout):
        raise OneNoteError(
            f"OneNote did not respond within {timeout:.0f}s — it is likely showing a "
            "dialog or syncing. Dismiss any open OneNote dialog and retry; call ping "
            "to check whether OneNote has recovered.",
            code="timeout",
        )

    if _exc_info[0] is not None:
        _, ev, tb = _exc_info[0]
        raise ev.with_traceback(tb)

    return _result[0]


# ---- Image handle helpers ---------------------------------------------------

def _mint_handle(seed: str) -> str:
    """Return a stable 12-char handle id for *seed* (content- or source-addressed)."""
    return hashlib.sha1(seed.encode()).hexdigest()[:12]


def _cache_put(handle: str, data: str) -> None:
    """Insert *data* under *handle*, evicting LRU entries beyond the size cap.

    Eviction prefers entries that also appear in _HANDLE_SOURCES: those can be
    re-fetched lazily from OneNote at write time, so dropping them loses only a
    round-trip.  Content-addressed entries (binary reads) are evicted only when
    re-fetchable ones don't free enough space — their handles then behave like
    stale handles until the source page is re-read.  The just-inserted entry is
    never evicted, even if it alone exceeds the cap.
    """
    with _CACHE_LOCK:
        _IMAGE_CACHE.pop(handle, None)
        _IMAGE_CACHE[handle] = data
        excess = sum(len(v) for v in _IMAGE_CACHE.values()) - _IMAGE_CACHE_MAX_BYTES
        for refetchable_only in (True, False):
            if excess <= 0:
                break
            for key in list(_IMAGE_CACHE):
                if key == handle:
                    continue
                if refetchable_only and key not in _HANDLE_SOURCES:
                    continue
                excess -= len(_IMAGE_CACHE.pop(key))
                if excess <= 0:
                    break


def _cache_get(handle: str) -> str | None:
    """Return cached bytes for *handle* (marking it most-recently-used), or None."""
    with _CACHE_LOCK:
        data = _IMAGE_CACHE.get(handle)
        if data is not None:
            _IMAGE_CACHE.move_to_end(handle)
        return data


def _strip_data_to_mcpref(xml: str) -> str:
    """Replace inline base64 in <one:Image>/<one:Data> with mcpref handles.

    Used on the *binary* read path (pageInfo=1).  For each <one:Image> whose
    <one:Data> holds non-empty text that is not already a handle, the bytes are
    cached content-addressed (sha1 of the bytes) so identical images share a
    handle, and the element text is replaced with ``mcpref:<id>``.
    """
    root = ET.fromstring(xml)
    image_tag = f"{{{_ONE_NS}}}Image"
    data_tag = f"{{{_ONE_NS}}}Data"
    for img_el in root.iter(image_tag):
        data_el = img_el.find(data_tag)
        if data_el is None:
            continue
        text = data_el.text or ""
        if not text or text.startswith(_MCPREF_PREFIX):
            continue
        handle = _mint_handle(text.strip())
        _cache_put(handle, text)
        data_el.text = f"{_MCPREF_PREFIX}{handle}"
    return ET.tostring(root, encoding="unicode", xml_declaration=False)


def _strip_callbacks_to_mcpref(xml: str, page_id: str) -> str:
    """Rewrite a no-binary (pageInfo=0) page so images carry mcpref handles.

    In no-binary mode OneNote emits ``<one:Image><one:CallbackID callbackID=.../>``
    with no <one:Data>.  For each image we mint a handle keyed on
    (page_id, callbackID), remember the source for lazy byte resolution, and add
    a ``<one:Data>mcpref:<id></one:Data>`` element so the structured parser
    (which looks for Image/Data) sees the handle.
    """
    root = ET.fromstring(xml)
    image_tag = f"{{{_ONE_NS}}}Image"
    data_tag = f"{{{_ONE_NS}}}Data"
    cb_tag = f"{{{_ONE_NS}}}CallbackID"
    for img_el in root.iter(image_tag):
        # If bytes are already present (mixed content), leave the binary path to it.
        existing_data = img_el.find(data_tag)
        if existing_data is not None and (existing_data.text or "").strip():
            continue
        cb_el = img_el.find(cb_tag)
        if cb_el is None:
            continue
        callback_id = cb_el.get("callbackID")
        if not callback_id:
            continue
        handle = _mint_handle(f"{page_id}|{callback_id}")
        _HANDLE_SOURCES[handle] = (page_id, callback_id)
        data_el = existing_data if existing_data is not None else ET.SubElement(img_el, data_tag)
        data_el.text = f"{_MCPREF_PREFIX}{handle}"
    return ET.tostring(root, encoding="unicode", xml_declaration=False)


def _resolve_handle_bytes(handle: str, app=None) -> str:
    """Return the base64 bytes for *handle*, fetching lazily if needed.

    Resolution order:
      1. _IMAGE_CACHE — bytes already in hand (binary read or prior fetch).
      2. _HANDLE_SOURCES — fetch via GetBinaryPageContent(page_id, callbackID),
         cache, and return.  Requires *app* (only available on the COM worker
         thread); without it a known-but-unfetched handle is treated as unknown.
    Raises OneNoteError(code="bad_request") if the handle cannot be resolved.
    """
    cached = _cache_get(handle)
    if cached is not None:
        return cached
    source = _HANDLE_SOURCES.get(handle)
    if source is not None and app is not None:
        page_id, callback_id = source
        data = app.GetBinaryPageContent(page_id, callback_id)
        _cache_put(handle, data)
        return data
    full = f"{_MCPREF_PREFIX}{handle}"
    raise OneNoteError(
        f"Unknown image handle {full!r}. "
        "Re-read the source page with get_page to refresh handles, then retry.",
        code="bad_request",
    )


def _resolve_mcpref_to_data(xml: str, app=None) -> str:
    """Replace mcpref handles in <one:Data> elements with their base64 bytes.

    Real base64 content (no mcpref: prefix) passes through untouched.  Handles
    are resolved via _resolve_handle_bytes — from cache, or lazily from OneNote
    when *app* is supplied.  Raises OneNoteError(code="bad_request") for a handle
    that cannot be resolved.
    """
    root = ET.fromstring(xml)
    data_tag = f"{{{_ONE_NS}}}Data"
    for data_el in root.iter(data_tag):
        text = data_el.text or ""
        if not text.startswith(_MCPREF_PREFIX):
            continue
        handle = text[len(_MCPREF_PREFIX):]
        data_el.text = _resolve_handle_bytes(handle, app)
    return ET.tostring(root, encoding="unicode", xml_declaration=False)


def validate_handles(handles: list[str]) -> dict[str, bool]:
    """Return {handle: resolvable} for each handle without fetching bytes.

    A handle is resolvable if its bytes are cached or its (page, callback)
    source is known.  Pure validation — no COM call — so it is safe and fast.
    """
    result: dict[str, bool] = {}
    for h in handles:
        key = h[len(_MCPREF_PREFIX):] if h.startswith(_MCPREF_PREFIX) else h
        result[h] = key in _IMAGE_CACHE or key in _HANDLE_SOURCES
    return result


# ---- Request guards -----------------------------------------------------------

def _parse_user_xml(xml_str: str, what: str) -> None:
    """Reject malformed caller-supplied XML before any COM work.

    ET.ParseError subclasses SyntaxError (not ValueError), so without this
    guard it would bypass the OneNoteError handling in the server layer and
    surface as a raw parse error with no machine-readable code.
    """
    try:
        ET.fromstring(xml_str)
    except ET.ParseError as exc:
        raise OneNoteError(
            f"Malformed {what} XML: {exc}. Send a single well-formed XML "
            "element with the xmlns:one namespace declared.",
            code="bad_request",
        ) from exc


def _allowed_notebooks() -> set[str] | None:
    """Return the set of allowed notebook display names, or None when unrestricted.

    Read from the ONENOTE_ALLOWED_NOTEBOOKS environment variable at call time:
    a comma-separated list of notebook names, e.g. "ClaudeSpike, Mathe 5a".
    Unset or blank means no restriction.
    """
    raw = os.environ.get("ONENOTE_ALLOWED_NOTEBOOKS")
    if raw is None:
        return None
    names = {part.strip() for part in raw.split(",") if part.strip()}
    return names or None


def _filter_notebooks_xml(xml: str) -> str:
    """Drop <one:Notebook> elements whose name is not allowlisted.

    Applied to hierarchy listings so restricted notebooks are simply invisible.
    No-op when no allowlist is configured.
    """
    allowed = _allowed_notebooks()
    if allowed is None:
        return xml
    root = ET.fromstring(xml)
    for nb_el in list(root.findall(f"{{{_ONE_NS}}}Notebook")):
        if nb_el.get("name") not in allowed:
            root.remove(nb_el)
    return ET.tostring(root, encoding="unicode", xml_declaration=False)


def _ensure_allowed(app, object_id: str, kind: str) -> None:
    """Raise unless *object_id* (a section or page id) is inside an allowed notebook.

    No-op — and no COM call — when no allowlist is configured.  Guards every
    id-scoped read/write so an id from a restricted notebook cannot be used
    even if known from an earlier unrestricted session.  get_image_data needs
    no guard of its own: handles can only be minted by guarded page reads.
    Runs on the COM worker thread (*app* is live).
    """
    allowed = _allowed_notebooks()
    if allowed is None:
        return
    root = ET.fromstring(app.GetHierarchy("", 4))  # scope 4 = full tree
    target = next((el for el in root.iter() if el.get("ID") == object_id), None)
    if target is None:
        raise OneNoteError(
            f"{kind} {object_id!r} was not found in any open notebook.",
            code="bad_request",
        )
    parent_of = {child: parent for parent in root.iter() for child in parent}
    nb_tag = f"{{{_ONE_NS}}}Notebook"
    cur = target
    while cur is not None and cur.tag != nb_tag:
        cur = parent_of.get(cur)
    if cur is None or cur.get("name") not in allowed:
        raise OneNoteError(
            f"{kind} {object_id!r} is not in an allowed notebook — access is "
            "restricted by ONENOTE_ALLOWED_NOTEBOOKS.",
            code="bad_request",
        )


# ---- COM implementation functions (run on worker thread via _run_com) -------
# These functions must NOT call CoInitialize/CoUninitialize — that is _run_com's
# responsibility.  Public functions that internally call other COM work must
# invoke the _impl variant directly (same thread / STA apartment).

def _list_hierarchy_impl() -> str:
    app = _app()
    # scope 4 = hsChildren: one call returns the full tree
    # out-param pbstrHierarchyXmlOut is returned as the Python return value
    return _filter_notebooks_xml(app.GetHierarchy("", 4))


def _ping_impl() -> str:
    """Lightweight liveness probe: notebooks only (scope 0 = hsSelf)."""
    app = _app()
    return app.GetHierarchy("", 0)


def _list_sections_impl() -> str:
    app = _app()
    # scope 3 = hsSections: notebooks + their sections, WITHOUT pages.
    return _filter_notebooks_xml(app.GetHierarchy("", 3))


def _list_pages_impl(section_id: str) -> str:
    app = _app()
    _ensure_allowed(app, section_id, "Section")
    # scope 4 anchored on a section returns just that section's pages.
    return app.GetHierarchy(section_id, 4)


def _get_page_impl(page_id: str, include_binary: bool = False) -> str:
    app = _app()
    _ensure_allowed(app, page_id, "Page")
    if include_binary:
        # pageInfo=1 = piBinaryData: embeds image bytes as base64 in <one:Data>
        xml = app.GetPageContent(page_id, 1)
        return _strip_data_to_mcpref(xml)
    # pageInfo=0 = piBasic: no binary; images come back as <one:CallbackID>.
    xml = app.GetPageContent(page_id, 0)
    return _strip_callbacks_to_mcpref(xml, page_id)


def _get_image_data_impl(handle: str) -> str:
    app = _app()
    key = handle[len(_MCPREF_PREFIX):] if handle.startswith(_MCPREF_PREFIX) else handle
    return _resolve_handle_bytes(key, app)


def _replace_page_impl(page_id: str, page_xml: str) -> None:
    """Full-replace implementation — runs on the worker thread.

    See replace_page docstring for full semantics.
    """
    app = _app()
    _ensure_allowed(app, page_id, "Page")

    _TITLE_TAG = f"{{{_ONE_NS}}}Title"
    current_xml = app.GetPageContent(page_id, 0)
    current_root = ET.fromstring(current_xml)

    # --- prepare write XML (before any deletion) ---
    root = ET.fromstring(page_xml)
    root.set("ID", page_id)
    # Strip all objectIDs so OneNote treats every element as new content.
    for el in root.iter():
        el.attrib.pop("objectID", None)

    # Preserve the existing Title objectID so OneNote updates the title
    # in-place rather than appending a duplicate title element.
    existing_title = current_root.find(_TITLE_TAG)
    if existing_title is not None:
        existing_title_id = existing_title.get("objectID")
        if existing_title_id:
            new_title = root.find(_TITLE_TAG)
            if new_title is not None:
                new_title.set("objectID", existing_title_id)

    serialized = ET.tostring(root, encoding="unicode", xml_declaration=False)
    # Resolve image handles to bytes NOW, while the source images still exist on
    # the page — a handle read with include_binary=False is fetched lazily via
    # GetBinaryPageContent, which would fail if we deleted the image first.
    serialized = _resolve_mcpref_to_data(serialized, app)

    # --- delete-pass: remove all existing top-level children via DeletePageContent ---
    for child in current_root:
        object_id = child.get("objectID")
        if not object_id:
            continue
        if child.tag == _TITLE_TAG:
            # Title is preserved and updated in-place in the write-pass.
            continue
        # Best-effort: some objects may not be deletable (e.g. locked content).
        try:
            # Third arg dateExpectedLastModified must be a float (OLE date);
            # 0.0 means no conflict check.
            app.DeletePageContent(page_id, object_id, 0.0)
        except comtypes.COMError:
            pass

    # --- write-pass: insert the new content ---
    # Pass 0.0 for dateExpectedLastModified (OLE date epoch = no conflict check)
    app.UpdatePageContent(serialized, 0.0)


def _plan_subpage_order(
    page_ids: list[str],
    page_levels: dict[str, str],
    new_page_id: str,
    parent_page_id: str,
) -> tuple[list[str], dict[str, str]]:
    """Pure: return (ordered_ids, levels) placing *new_page_id* as the LAST
    sub-page of *parent_page_id* (after the parent's existing sub-pages), at
    parent_level + 1.

    A page's sub-tree is the run of pages immediately following it whose level is
    greater than the parent's; the new page is inserted just past that run so it
    becomes the last child — matching "add to the end of this parent" rather than
    pushing in front of existing children.

    Raises OneNoteError(code="bad_request") if the parent is not in the section.
    """
    if parent_page_id not in page_levels:
        raise OneNoteError(
            f"Parent page {parent_page_id!r} is not in section — pass a page_id "
            "from this section's list_hierarchy.",
            code="bad_request",
        )
    ordered = [pid for pid in page_ids if pid != new_page_id]
    parent_level = int(page_levels.get(parent_page_id) or "1")
    levels = dict(page_levels)
    levels[new_page_id] = str(parent_level + 1)

    # Walk past the parent's existing sub-tree to find the end-of-children slot.
    insert_at = ordered.index(parent_page_id) + 1
    while insert_at < len(ordered) and int(levels.get(ordered[insert_at]) or "1") > parent_level:
        insert_at += 1
    ordered.insert(insert_at, new_page_id)
    return ordered, levels


def _place_as_subpage(app, section_id: str, new_page_id: str, parent_page_id: str) -> None:
    """Reposition *new_page_id* as a sub-page of *parent_page_id* via UpdateHierarchy."""
    page_tag = f"{{{_ONE_NS}}}Page"
    section_xml = app.GetHierarchy(section_id, 4)  # scope 4 = pages
    src = ET.fromstring(section_xml)
    pages = [p for p in src.iter(page_tag) if p.get("ID")]
    page_ids = [p.get("ID") for p in pages]
    page_levels = {p.get("ID"): (p.get("pageLevel") or "1") for p in pages}

    ordered, levels = _plan_subpage_order(page_ids, page_levels, new_page_id, parent_page_id)

    section_el = ET.Element(f"{{{_ONE_NS}}}Section")
    section_el.set("ID", section_id)
    for pid in ordered:
        page_el = ET.SubElement(section_el, page_tag)
        page_el.set("ID", pid)
        page_el.set("pageLevel", levels[pid])
    app.UpdateHierarchy(ET.tostring(section_el, encoding="unicode", xml_declaration=False))


def _create_page_impl(
    section_id: str,
    title: str | None = None,
    parent_page_id: str | None = None,
) -> str:
    app = _app()
    _ensure_allowed(app, section_id, "Section")
    if parent_page_id is not None:
        # Validate the parent exists BEFORE creating anything, so a bad parent
        # does not leave an orphan page behind.
        existing = ET.fromstring(app.GetHierarchy(section_id, 4))
        page_tag = f"{{{_ONE_NS}}}Page"
        if not any(p.get("ID") == parent_page_id for p in existing.iter(page_tag)):
            raise OneNoteError(
                f"Parent page {parent_page_id!r} is not in section — pass a page_id "
                "from this section's list_hierarchy.",
                code="bad_request",
            )
    # out-param pbstrPageID is returned as the Python return value
    new_page_id = app.CreateNewPage(section_id)
    if title is not None:
        # A freshly created page has only an (empty) title placeholder, so we can
        # set the title with a single UpdatePageContent — no delete-pass needed.
        page_el = ET.Element(f"{{{_ONE_NS}}}Page")
        page_el.set("ID", new_page_id)
        title_el = ET.SubElement(page_el, f"{{{_ONE_NS}}}Title")
        oe_el = ET.SubElement(title_el, f"{{{_ONE_NS}}}OE")
        t_el = ET.SubElement(oe_el, f"{{{_ONE_NS}}}T")
        t_el.text = title
        title_xml = ET.tostring(page_el, encoding="unicode", xml_declaration=False)
        app.UpdatePageContent(title_xml, 0.0)
    if parent_page_id is not None:
        # CreateNewPage appends at the end of the section; reorder + indent so the
        # page becomes the LAST sub-page under the requested parent.
        _place_as_subpage(app, section_id, new_page_id, parent_page_id)
    return new_page_id


def _append_page_impl(page_id: str, content_xml: str) -> None:
    app = _app()
    _ensure_allowed(app, page_id, "Page")
    content_el = ET.fromstring(content_xml)
    for el in content_el.iter():
        el.attrib.pop("objectID", None)
    page_tag = f"{{{_ONE_NS}}}Page"
    page_el = ET.Element(page_tag)
    page_el.set("ID", page_id)
    page_el.append(content_el)
    serialized = ET.tostring(page_el, encoding="unicode", xml_declaration=False)
    serialized = _resolve_mcpref_to_data(serialized, app)
    app.UpdatePageContent(serialized, 0.0)


# ---- Public API -------------------------------------------------------------

def ping() -> bool:
    """Return True if OneNote answers a lightweight call within PING_TIMEOUT.

    Never raises for an unresponsive backend — returns False instead — so a
    caller can distinguish "server alive, OneNote wedged" from "server down".
    """
    try:
        _run_com(_ping_impl, timeout=PING_TIMEOUT)
        return True
    except (OneNoteError, comtypes.COMError):
        return False


def list_hierarchy() -> str:
    """Return the full notebook -> section -> page tree as OneNote XML."""
    try:
        return _run_com(_list_hierarchy_impl, timeout=READ_TIMEOUT)
    except comtypes.COMError as exc:
        raise OneNoteError(f"GetHierarchy failed: hresult={exc.hresult:#010x}") from exc


def list_sections() -> str:
    """Return notebooks and their sections (no pages) as OneNote XML."""
    try:
        return _run_com(_list_sections_impl, timeout=READ_TIMEOUT)
    except comtypes.COMError as exc:
        raise OneNoteError(
            f"GetHierarchy (sections) failed: hresult={exc.hresult:#010x}"
        ) from exc


def list_pages(section_id: str) -> str:
    """Return a single section's pages as OneNote XML (anchored on section_id)."""
    try:
        return _run_com(_list_pages_impl, section_id, timeout=READ_TIMEOUT)
    except comtypes.COMError as exc:
        raise OneNoteError(
            f"GetHierarchy (section {section_id!r}) failed: hresult={exc.hresult:#010x}"
        ) from exc


def get_page(page_id: str, include_binary: bool = False) -> str:
    """Return page XML with image data replaced by mcpref handles.

    By default (include_binary=False) the page is read WITHOUT image bytes,
    which keeps image-heavy pages small and fast; the bytes are fetched lazily
    only if a handle is later written back.  Pass include_binary=True to embed
    base64 image data inline (slower, larger).
    """
    try:
        return _run_com(_get_page_impl, page_id, include_binary, timeout=READ_TIMEOUT)
    except comtypes.COMError as exc:
        raise OneNoteError(
            f"GetPageContent failed (page {page_id!r}): hresult={exc.hresult:#010x}"
        ) from exc


def get_image_data(handle: str) -> str:
    """Return the base64 image bytes for an mcpref *handle*.

    Resolves from cache or lazily via GetBinaryPageContent.  Raises
    OneNoteError(code="bad_request") if the handle is unknown/stale.
    """
    try:
        return _run_com(_get_image_data_impl, handle, timeout=READ_TIMEOUT)
    except comtypes.COMError as exc:
        raise OneNoteError(
            f"get_image_data failed for {handle!r}: hresult={exc.hresult:#010x}"
        ) from exc


def create_page(
    section_id: str,
    title: str | None = None,
    parent_page_id: str | None = None,
) -> str:
    """Create a blank page in *section_id* and return its new page ID.

    If *title* is given, seeds the page with a <one:Title> element.
    If *parent_page_id* is given, the new page is placed as a sub-page directly
    under that page (reordered to sit after it, indented one level).
    """
    try:
        return _run_com(
            _create_page_impl, section_id, title, parent_page_id, timeout=WRITE_TIMEOUT
        )
    except comtypes.COMError as exc:
        raise OneNoteError(
            f"create_page failed (section {section_id!r}): hresult={exc.hresult:#010x}"
        ) from exc


def replace_page(page_id: str, page_xml: str) -> None:
    """Full-replace the page identified by *page_id* with *page_xml*.

    The caller supplies a complete <one:Page> document. The root element's ID
    attribute is overwritten to match *page_id* before calling
    UpdatePageContent, so callers don't need to manage that detail.
    mcpref handles in <one:Data> elements are resolved to real base64 before
    the COM call.

    Implementation note: OneNote UpdatePageContent is an upsert keyed on
    objectID.  To achieve a true replace, this function first issues a
    delete-pass that calls DeletePageContent for every top-level child of the
    current page (except <one:Title>), then writes the new content.
    DeletePageContent handles all object types including images and ink — no
    special-casing is required.  Each deletion is wrapped in a try/except so
    a single un-deletable object cannot abort the whole replace.
    """
    _parse_user_xml(page_xml, "page")
    try:
        _run_com(_replace_page_impl, page_id, page_xml, timeout=WRITE_TIMEOUT)
    except comtypes.COMError as exc:
        raise OneNoteError(
            f"replace_page failed (page {page_id!r}): hresult={exc.hresult:#010x}"
        ) from exc


def append_page(page_id: str, content_xml: str) -> None:
    """Append a single XML element to the page identified by *page_id*.

    *content_xml* must be a single, formally correct OneNote XML element
    with the xmlns:one declaration on the element itself, e.g.:

        <one:Outline xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
          <one:OEChildren>
            <one:OE><one:T>New paragraph</one:T></one:OE>
          </one:OEChildren>
        </one:Outline>

    Valid top-level elements include <one:Outline>, <one:Image>, <one:Table>, etc.
    All objectID attributes are stripped so OneNote treats the element as new
    content to be added (not an update to an existing element).
    mcpref handles in <one:Data> elements are resolved to real base64 before
    the COM call — same as replace_page.
    """
    _parse_user_xml(content_xml, "content")
    try:
        _run_com(_append_page_impl, page_id, content_xml, timeout=WRITE_TIMEOUT)
    except comtypes.COMError as exc:
        raise OneNoteError(
            f"append_page failed (page {page_id!r}): hresult={exc.hresult:#010x}"
        ) from exc
