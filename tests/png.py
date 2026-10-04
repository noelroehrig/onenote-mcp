"""Minimal PNG images of a chosen pixel size, for tests that need real image bytes."""
import base64
import struct
import zlib


def png(width: int, height: int) -> bytes:
    """Return a valid black RGB PNG of *width* x *height* pixels."""
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    rows = b"".join(b"\x00" + b"\x00\x00\x00" * width for _ in range(height))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(rows))
        + chunk(b"IEND", b"")
    )


def png_base64(width: int, height: int) -> str:
    return base64.b64encode(png(width, height)).decode()
