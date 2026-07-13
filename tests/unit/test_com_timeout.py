"""Unit tests for _run_com in onenote_mcp.com.

These tests exercise _run_com directly with plain Python callables — no COM,
no OneNote, no registry.  They validate the three required behaviors:
  1. Timeout path: hung fn → OneNoteError with descriptive message, returned fast.
  2. Error pass-through: fn that raises → same exception type re-raised.
  3. Success pass-through: fn that returns a value → value returned unchanged.
"""
import time
import pytest
from onenote_mcp.com import _run_com, OneNoteError


# ---------------------------------------------------------------------------
# Timeout path
# ---------------------------------------------------------------------------

def test_run_com_timeout_raises_one_note_error():
    """fn that sleeps 5s with timeout=0.5 raises OneNoteError well under 5s."""
    def slow_fn():
        time.sleep(5)
        return "should not reach here"

    start = time.perf_counter()
    with pytest.raises(OneNoteError):
        _run_com(slow_fn, timeout=0.5)
    elapsed = time.perf_counter() - start

    assert elapsed < 2.0, f"Expected fast timeout but took {elapsed:.2f}s"


def test_run_com_timeout_message_mentions_dialog():
    """Timeout error message mentions 'did not respond' and 'dialog'."""
    def slow_fn():
        time.sleep(5)

    with pytest.raises(OneNoteError) as exc_info:
        _run_com(slow_fn, timeout=0.5)

    msg = str(exc_info.value)
    assert "did not respond" in msg, f"Expected 'did not respond' in: {msg!r}"
    assert "dialog" in msg, f"Expected 'dialog' in: {msg!r}"


def test_run_com_timeout_carries_machine_readable_code():
    """A timeout surfaces a structured code='timeout' (not a generic error)."""
    def slow_fn():
        time.sleep(5)

    with pytest.raises(OneNoteError) as exc_info:
        _run_com(slow_fn, timeout=0.5)

    assert exc_info.value.code == "timeout"
    assert str(exc_info.value).startswith("timeout: ")


# ---------------------------------------------------------------------------
# Error pass-through
# ---------------------------------------------------------------------------

def test_run_com_re_raises_original_exception_type():
    """fn that raises a custom exception → same exception type re-raised (not OneNoteError)."""
    class MyCustomError(Exception):
        pass

    def failing_fn():
        raise MyCustomError("something went wrong")

    with pytest.raises(MyCustomError, match="something went wrong"):
        _run_com(failing_fn, timeout=5.0)


def test_run_com_re_raises_value_error():
    """fn that raises ValueError → ValueError re-raised unchanged."""
    def failing_fn():
        raise ValueError("bad value")

    with pytest.raises(ValueError, match="bad value"):
        _run_com(failing_fn, timeout=5.0)


def test_run_com_does_not_swallow_exception_as_one_note_error():
    """A re-raised exception must NOT be wrapped in OneNoteError."""
    class DistinctError(Exception):
        pass

    def failing_fn():
        raise DistinctError("distinct")

    # Must not raise OneNoteError
    with pytest.raises(DistinctError):
        _run_com(failing_fn, timeout=5.0)


# ---------------------------------------------------------------------------
# Success pass-through
# ---------------------------------------------------------------------------

def test_run_com_returns_value_unchanged():
    """fn returning a plain value → value returned without modification."""
    def successful_fn():
        return {"key": "value", "count": 42}

    result = _run_com(successful_fn, timeout=5.0)
    assert result == {"key": "value", "count": 42}


def test_run_com_returns_string():
    """fn returning a str → same string returned."""
    def str_fn():
        return "<xml>hello</xml>"

    result = _run_com(str_fn, timeout=5.0)
    assert result == "<xml>hello</xml>"


def test_run_com_returns_none():
    """fn that returns None (void functions) → None returned."""
    def void_fn():
        return None

    result = _run_com(void_fn, timeout=5.0)
    assert result is None


def test_run_com_passes_args_to_fn():
    """args passed to _run_com are forwarded to fn in order."""
    def add(a, b):
        return a + b

    result = _run_com(add, 3, 7, timeout=5.0)
    assert result == 10
