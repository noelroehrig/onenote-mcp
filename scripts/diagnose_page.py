"""Read-only diagnostic for a OneNote page.

Usage:
    .venv\\Scripts\\python.exe scripts\\diagnose_page.py [PAGE_ID]

If PAGE_ID is omitted, the script lists pages whose name contains "Manual"
(case-insensitive) and exits — pick one and rerun with its ID.

Reports:
  1. Payload size and latency of GetPageContent, with and without binary
     image data.
  2. Top-level children of <one:Page>.
  3. QuickStyleDef inventory and quickStyleIndex usage (heading resolution).
  4. Image inventory: floating vs inline placement, and the state of each
     image's <one:Data>.

The script touches OneNote read-only: GetHierarchy + GetPageContent.
"""
from __future__ import annotations

import sys
import time
import xml.etree.ElementTree as ET
from collections import Counter

# Make src/ importable when running from repo root.
import os
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(_REPO_ROOT, "src"))

from onenote_mcp import com  # noqa: E402

_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _local(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag


def _list_manual_pages() -> None:
    print("No PAGE_ID given. Listing pages with 'manual' in the name:")
    xml = com.list_hierarchy()
    root = ET.fromstring(xml)
    found = 0
    for page_el in root.iter(f"{{{_NS}}}Page"):
        name = page_el.get("name", "")
        if "manual" in name.lower():
            print(f"  {page_el.get('ID')}  {name!r}")
            found += 1
    if not found:
        print("  (none found — listing all pages instead, first 30)")
        for i, page_el in enumerate(root.iter(f"{{{_NS}}}Page")):
            if i >= 30:
                print("  ... (truncated)")
                break
            print(f"  {page_el.get('ID')}  {page_el.get('name')!r}")


def _summarize(page_id: str) -> None:
    print(f"=== Diagnosing page {page_id} ===\n")

    # --- Time the binary-data fetch --------------------------------------
    # com.get_page() also runs _strip_data_to_mcpref which we don't want here:
    # we want to see real base64 sizes. So call GetPageContent directly under
    # CoInitialize/CoUninitialize.
    print("[1] Timing GetPageContent(piBinaryData=1) — what get_page(include_binary=True) does")
    import comtypes
    comtypes.CoInitialize()
    try:
        app = com._app()
        t0 = time.perf_counter()
        full_xml = app.GetPageContent(page_id, 1)
        elapsed_full = time.perf_counter() - t0
    finally:
        comtypes.CoUninitialize()
    full_size = len(full_xml.encode("utf-8"))
    print(f"    elapsed: {elapsed_full*1000:.0f} ms")
    print(f"    XML size with binary data: {full_size:,} bytes "
          f"({full_size/1024:.1f} KiB)")

    # Also get the basic (no binary) variant for comparison.
    comtypes.CoInitialize()
    try:
        app = com._app()
        t0 = time.perf_counter()
        basic_xml = app.GetPageContent(page_id, 0)
        elapsed_basic = time.perf_counter() - t0
    finally:
        comtypes.CoUninitialize()
    basic_size = len(basic_xml.encode("utf-8"))
    print(f"    XML size WITHOUT binary data: {basic_size:,} bytes "
          f"({basic_size/1024:.1f} KiB), elapsed {elapsed_basic*1000:.0f} ms")
    print(f"    => binary data adds {(full_size - basic_size)/1024:.1f} KiB\n")

    root = ET.fromstring(full_xml)

    # --- Top-level children ----------------------------------------------
    print("[2] Top-level children of <one:Page>")
    top_counts: Counter[str] = Counter()
    for child in root:
        top_counts[_local(child.tag)] += 1
    for tag, n in top_counts.most_common():
        print(f"    {tag:24s}  x{n}")
    print()

    # --- QuickStyleDefs ---------------------------------------------------
    print("[3] QuickStyleDef inventory")
    qsd_tag = f"{{{_NS}}}QuickStyleDef"
    qsds = list(root.findall(qsd_tag))
    if not qsds:
        # Sometimes nested? Walk all descendants.
        qsds_all = list(root.iter(qsd_tag))
        print(f"    direct children: 0   (descendants found anywhere: {len(qsds_all)})")
    else:
        print(f"    {len(qsds)} QuickStyleDef element(s) as direct children of <one:Page>:")
        for qsd in qsds:
            idx = qsd.get("index")
            name = qsd.get("name")
            objid = qsd.get("objectID")
            font = qsd.get("font")
            size = qsd.get("fontSize")
            color = qsd.get("fontColor")
            bold = qsd.get("bold")
            print(f"      index={idx!s:3} name={name!r:18} objectID={'YES' if objid else 'no':4}  "
                  f"font={font!s} size={size!s} bold={bold!s} color={color!s}")
    print()

    # --- quickStyleIndex usage on OE elements -----------------------------
    print("[4] quickStyleIndex usage on <one:OE> elements")
    oe_tag = f"{{{_NS}}}OE"
    oe_with_idx: Counter[str] = Counter()
    oe_total = 0
    for oe in root.iter(oe_tag):
        oe_total += 1
        v = oe.get("quickStyleIndex")
        if v is not None:
            oe_with_idx[v] += 1
    print(f"    total OE elements: {oe_total}")
    if oe_with_idx:
        print(f"    OE with quickStyleIndex set: {sum(oe_with_idx.values())}")
        for idx, n in sorted(oe_with_idx.items()):
            print(f"      index={idx!r}: {n} OE(s)")
    else:
        print("    (none)")
    print()

    # --- Image inventory ---------------------------------------------------
    print("[5] Image inventory (floating vs inline)")
    img_tag = f"{{{_NS}}}Image"
    data_tag = f"{{{_NS}}}Data"
    page_tag = f"{{{_NS}}}Page"

    # Build parent map for path tracing
    parent_map = {child: parent for parent in root.iter() for child in parent}

    def _path_to_root(el: ET.Element) -> str:
        parts: list[str] = []
        cur = el
        while cur is not None:
            parts.append(_local(cur.tag))
            cur = parent_map.get(cur)
        return " > ".join(reversed(parts))

    floating = 0
    inline = 0
    inline_paths: Counter[str] = Counter()
    image_data_states: Counter[str] = Counter()

    for img_el in root.iter(img_tag):
        # Floating = direct child of <one:Page>
        parent = parent_map.get(img_el)
        if parent is not None and parent.tag == page_tag:
            floating += 1
            location = "FLOATING (Page > Image)"
        else:
            inline += 1
            location = "INLINE"
            inline_paths[_path_to_root(img_el)] += 1

        data_el = img_el.find(data_tag)
        if data_el is None:
            state = "no <one:Data> child"
        else:
            text = data_el.text or ""
            if not text:
                state = "<one:Data> empty"
            elif text.startswith("mcpref:"):
                state = "<one:Data> = mcpref handle"
            else:
                state = f"<one:Data> = base64 ({len(text):,} chars)"
        image_data_states[f"{location} | {state}"] += 1

    print(f"    floating images (direct child of Page): {floating}")
    print(f"    inline images   (anywhere else):         {inline}")
    if inline_paths:
        print("    inline image paths from Page root:")
        for path, n in inline_paths.most_common():
            print(f"      x{n:<3} {path}")
    print("    state of <one:Data> per image:")
    for k, n in image_data_states.most_common():
        print(f"      x{n:<3} {k}")
    print()

    # --- Final headline summary ------------------------------------------
    print("=== Summary ===")
    print(f"  XML payload (with images):  {full_size/1024:.1f} KiB")
    print(f"  Read latency:               {elapsed_full*1000:.0f} ms")
    print(f"  QuickStyleDef defs at root: {len(root.findall(qsd_tag))}")
    print(f"  OEs claiming a heading:     {sum(oe_with_idx.values())}")
    print(f"  Inline images (in outlines): {inline}")
    print(f"  Floating images (page-level): {floating}")


def main() -> None:
    if len(sys.argv) < 2:
        _list_manual_pages()
        return
    _summarize(sys.argv[1])


if __name__ == "__main__":
    main()
