"""Unit tests for onenote_mcp.images.pixel_size."""
import struct

from onenote_mcp.images import pixel_size
from tests.png import png as _png


def _jpeg(width: int, height: int) -> bytes:
    """SOI, an APP0 and a DHT segment, then a baseline start-of-frame."""
    app0 = b"\xff\xe0" + struct.pack(">H", 16) + b"JFIF\x00" + b"\x00" * 9
    dht = b"\xff\xc4" + struct.pack(">H", 5) + b"\x00\x00\x00"
    sof0 = b"\xff\xc0" + struct.pack(">HBHHB", 11, 8, height, width, 1) + b"\x01\x11\x00"
    return b"\xff\xd8" + app0 + dht + sof0 + b"\xff\xd9"


def test_png():
    assert pixel_size(_png(200, 100)) == (200, 100)


def test_jpeg_skips_segments_before_the_frame():
    assert pixel_size(_jpeg(640, 480)) == (640, 480)


def test_gif():
    assert pixel_size(b"GIF89a" + struct.pack("<HH", 32, 16) + b"\x00" * 8) == (32, 16)


def test_bmp_top_down_height_is_positive():
    header = b"BM" + b"\x00" * 12 + struct.pack("<Iii", 40, 30, -20)
    assert pixel_size(header) == (30, 20)


def test_unknown_format():
    assert pixel_size(b"not an image") is None


def test_zero_dimension():
    assert pixel_size(_png(0, 10)) is None


def test_truncated_jpeg():
    assert pixel_size(_jpeg(640, 480)[:20]) is None
