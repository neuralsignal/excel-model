"""Property-based tests using hypothesis."""

from hypothesis import given, settings
from hypothesis import strategies as st

from excel_model.formula_engine import CellContext, render_formula
from excel_model.time_engine import generate_periods

# Strategy for valid formula types
FORMULA_TYPES_WITH_SIMPLE_PARAMS = [
    ("constant", {"value": 42}),
    ("growth_projected", {"growth_assumption": "RevenueGrowthRate"}),
    ("subtraction", {"minuend_key": "a", "subtrahend_key": "b"}),
    ("ratio", {"numerator_key": "a", "denominator_key": "b"}),
    ("growth_rate", {"value_key": "a"}),
    ("variance", {"plan_key": "a", "actual_key": "b"}),
    ("variance_pct", {"plan_key": "a", "actual_key": "b"}),
    ("sum_of_rows", {"addend_keys": ["a", "b"]}),
]


def make_ctx(period_index: int, n_history: int, col: int) -> CellContext:
    from excel_model.named_ranges import get_col_letter

    col_letter = get_col_letter(col)
    prior_col_letter = get_col_letter(col - 1) if col > 2 else ""
    return CellContext(
        period_index=period_index,
        n_history=n_history,
        row=10,
        col=col,
        col_letter=col_letter,
        prior_col_letter=prior_col_letter,
        named_ranges={"RevenueGrowthRate": "RevenueGrowthRate", "WACC": "WACC"},
        row_map={"a": 10, "b": 11, "c": 12, "revenue": 5, "fcf": 8},
        inputs_row_map={"a": 3, "b": 4, "revenue": 3},
        scenario_prefix="",
        first_proj_col_letter="",
        last_proj_col_letter="",
        entity_col_range="",
    )


class TestFormulaEngineProperties:
    @given(
        period_index=st.integers(min_value=0, max_value=10),
        n_history=st.integers(min_value=0, max_value=5),
        col=st.integers(min_value=2, max_value=15),
    )
    @settings(max_examples=50)
    def test_growth_projected_returns_formula_string(self, period_index, n_history, col):
        ctx = make_ctx(period_index=period_index + n_history, n_history=n_history, col=col)
        result = render_formula("growth_projected", {"growth_assumption": "RevenueGrowthRate"}, ctx)
        # Must be a string starting with "="
        assert isinstance(result, str)
        assert result.startswith("=")

    @given(col=st.integers(min_value=2, max_value=26))
    def test_subtraction_returns_formula_string(self, col):
        ctx = make_ctx(period_index=2, n_history=0, col=col)
        result = render_formula("subtraction", {"minuend_key": "a", "subtrahend_key": "b"}, ctx)
        assert isinstance(result, str)
        assert result.startswith("=")

    @given(value=st.floats(min_value=-1e9, max_value=1e9, allow_nan=False, allow_infinity=False))
    def test_constant_returns_value(self, value):
        ctx = make_ctx(period_index=0, n_history=0, col=2)
        result = render_formula("constant", {"value": value}, ctx)
        # Constant should return the value, not a formula string
        assert result == value

    @given(col=st.integers(min_value=2, max_value=26))
    def test_pct_of_revenue_returns_formula(self, col):
        ctx = make_ctx(period_index=2, n_history=0, col=col)
        result = render_formula("pct_of_revenue", {"revenue_key": "revenue", "rate_assumption": "WACC"}, ctx)
        assert isinstance(result, str)
        assert result.startswith("=")
        assert "WACC" in result


class TestGeneratePeriodsProperties:
    @given(
        n_periods=st.integers(min_value=1, max_value=20),
        n_history=st.integers(min_value=0, max_value=10),
    )
    def test_output_length_is_n_history_plus_n_periods(self, n_periods, n_history):
        periods = generate_periods("2025", n_periods, n_history, "annual")
        assert len(periods) == n_periods + n_history

    @given(
        n_periods=st.integers(min_value=1, max_value=10),
        n_history=st.integers(min_value=0, max_value=5),
    )
    def test_history_flags_are_correct(self, n_periods, n_history):
        periods = generate_periods("2025", n_periods, n_history, "annual")
        for i, p in enumerate(periods):
            if i < n_history:
                assert p.is_history is True
            else:
                assert p.is_history is False

    @given(
        n_periods=st.integers(min_value=1, max_value=10),
        n_history=st.integers(min_value=0, max_value=5),
    )
    def test_indices_are_sequential(self, n_periods, n_history):
        periods = generate_periods("2025", n_periods, n_history, "annual")
        for i, p in enumerate(periods):
            assert p.index == i

    @given(
        n_periods=st.integers(min_value=1, max_value=10),
        n_history=st.integers(min_value=0, max_value=5),
    )
    def test_quarterly_length(self, n_periods, n_history):
        periods = generate_periods("2025-Q1", n_periods, n_history, "quarterly")
        assert len(periods) == n_periods + n_history

    @given(
        n_periods=st.integers(min_value=1, max_value=10),
        n_history=st.integers(min_value=0, max_value=5),
    )
    def test_monthly_length(self, n_periods, n_history):
        periods = generate_periods("2025-01", n_periods, n_history, "monthly")
        assert len(periods) == n_periods + n_history
