"""Unit tests for the pure sub-page ordering logic (_plan_subpage_order).

No COM needed — exercises the reorder/indent planning in isolation.
"""
import pytest

try:
    from onenote_mcp.com import _plan_subpage_order, OneNoteError
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not _COM_AVAILABLE,
    reason="com.py module-level imports unavailable in this environment",
)


def test_moves_new_page_after_parent_and_indents():
    ids = ["parent", "filler", "new"]
    levels = {"parent": "1", "filler": "1", "new": "1"}
    ordered, new_levels = _plan_subpage_order(ids, levels, "new", "parent")
    # 'filler' is a sibling (level 1), so the parent has no children yet — the new
    # page lands directly after the parent.
    assert ordered == ["parent", "new", "filler"]
    assert new_levels["new"] == "2"
    # Other pages keep their levels.
    assert new_levels["parent"] == "1"
    assert new_levels["filler"] == "1"


def test_new_page_becomes_last_child_after_existing_subpages():
    """The new sub-page is appended AFTER the parent's existing children, not in
    front of them (matches 'add to the end of this parent')."""
    ids = ["parent", "child1", "child2", "sibling", "new"]
    levels = {"parent": "1", "child1": "2", "child2": "2", "sibling": "1", "new": "1"}
    ordered, new_levels = _plan_subpage_order(ids, levels, "new", "parent")
    assert ordered == ["parent", "child1", "child2", "new", "sibling"]
    assert new_levels["new"] == "2"


def test_new_page_inserted_after_deeper_nested_descendants():
    """Walks past grandchildren too, so the new direct child is genuinely last."""
    ids = ["parent", "child1", "grandchild", "new"]
    levels = {"parent": "1", "child1": "2", "grandchild": "3", "new": "1"}
    ordered, new_levels = _plan_subpage_order(ids, levels, "new", "parent")
    assert ordered == ["parent", "child1", "grandchild", "new"]
    assert new_levels["new"] == "2"


def test_indents_one_level_below_a_nested_parent():
    ids = ["top", "parent", "new"]
    levels = {"top": "1", "parent": "2", "new": "1"}
    ordered, new_levels = _plan_subpage_order(ids, levels, "new", "parent")
    assert ordered == ["top", "parent", "new"]
    assert new_levels["new"] == "3"


def test_parent_already_last_keeps_new_adjacent():
    ids = ["a", "parent", "new"]
    levels = {"a": "1", "parent": "1", "new": "1"}
    ordered, _ = _plan_subpage_order(ids, levels, "new", "parent")
    assert ordered == ["a", "parent", "new"]


def test_unknown_parent_raises_bad_request():
    ids = ["a", "new"]
    levels = {"a": "1", "new": "1"}
    with pytest.raises(OneNoteError) as exc:
        _plan_subpage_order(ids, levels, "new", "does-not-exist")
    assert exc.value.code == "bad_request"
