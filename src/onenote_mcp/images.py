"""Read an image's pixel dimensions from its header bytes (PNG, JPEG, GIF, BMP)."""

from __future__ import annotations

_PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
# JPEG start-of-frame markers carry the frame size; C4, C8 and CC are other segments.
_JPEG_SOF_MARKERS = frozenset(range(0xC0, 0xD0)) - {0xC4, 0xC8, 0xCC}


def pixel_size(data: bytes) -> tuple[int, int] | None:
    """Return (width, height) in pixels, or None for an unknown format or a zero dimension."""
    if data.startswith(_PNG_SIGNATURE) and data[12:16] == b"IHDR":
        size = (int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big"))
    elif data[:6] in (b"GIF87a", b"GIF89a"):
        size = (int.from_bytes(data[6:8], "little"), int.from_bytes(data[8:10], "little"))
    elif data.startswith(b"BM") and len(data) >= 26:
        # BITMAPINFOHEADER; a negative height marks a top-down bitmap.
        size = (
            abs(int.from_bytes(data[18:22], "little", signed=True)),
            abs(int.from_bytes(data[22:26], "little", signed=True)),
        )
    elif data.startswith(b"\xff\xd8"):
        size = _jpeg_size(data)
    else:
        return None
    if size is None or 0 in size:
        return None
    return size


def _jpeg_size(data: bytes) -> tuple[int, int] | None:
    """Walk the JPEG segments up to the first start-of-frame and read its size."""
    i = 2
    while i + 9 <= len(data):
        if data[i] != 0xFF:
            return None
        marker = data[i + 1]
        if marker == 0xFF:  # fill byte
            i += 1
            continue
        if marker == 0x01 or 0xD0 <= marker <= 0xD7:  # markers without a length
            i += 2
            continue
        if marker in _JPEG_SOF_MARKERS:
            # Segment layout: FF Cn, length (2), precision (1), height (2), width (2).
            return int.from_bytes(data[i + 7:i + 9], "big"), int.from_bytes(data[i + 5:i + 7], "big")
        i += 2 + int.from_bytes(data[i + 2:i + 4], "big")
    return None
