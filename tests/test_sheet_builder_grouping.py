"""Tests for group_line_items_by_section helper."""

from hypothesis import given
from hypothesis import strategies as st

from excel_model.models._sheet_builder import group_line_items_by_section
from excel_model.spec import LineItemDef


def _make_li(key: str, section: str) -> LineItemDef:
    return LineItemDef(
        key=key,
        label=key,
        formula_type="literal",
        formula_params={},
        is_subtotal=False,
        is_total=False,
        section=section,
        format="",
    )


def test_empty_input() -> None:
    order, items = group_line_items_by_section(())
    assert order == []
    assert items == {}


def test_single_section() -> None:
    li_a = _make_li("a", "Revenue")
    li_b = _make_li("b", "Revenue")
    order, items = group_line_items_by_section((li_a, li_b))
    assert order == ["Revenue"]
    assert items == {"Revenue": [li_a, li_b]}


def test_multiple_sections_preserves_order() -> None:
    li1 = _make_li("r1", "Revenue")
    li2 = _make_li("c1", "Costs")
    li3 = _make_li("r2", "Revenue")
    order, items = group_line_items_by_section((li1, li2, li3))
    assert order == ["Revenue", "Costs"]
    assert items["Revenue"] == [li1, li3]
    assert items["Costs"] == [li2]


def test_empty_section_name() -> None:
    li = _make_li("x", "")
    order, items = group_line_items_by_section((li,))
    assert order == [""]
    assert items[""] == [li]


# Property-based test
_section_strategy = st.text(min_size=0, max_size=20)
_li_strategy = st.builds(
    LineItemDef,
    key=st.text(min_size=1, max_size=10),
    label=st.text(min_size=1, max_size=10),
    formula_type=st.just("literal"),
    formula_params=st.just({}),
    is_subtotal=st.just(False),
    is_total=st.just(False),
    section=_section_strategy,
    format=st.just(""),
)


@given(line_items=st.lists(_li_strategy, max_size=30))
def test_all_items_preserved(line_items: list[LineItemDef]) -> None:
    order, items = group_line_items_by_section(tuple(line_items))
    # Every item appears exactly once across all groups
    flat = [li for section_lis in items.values() for li in section_lis]
    assert len(flat) == len(line_items)
    # Section order has no duplicates
    assert len(order) == len(set(order))
    # Every section in items is in order and vice versa
    assert set(order) == set(items.keys())
