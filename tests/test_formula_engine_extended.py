"""Tests for new formula types: sum_subtraction, rank, index_to_base, bar_chart_text, fixed npv_sum."""
from excel_model.formula_engine import CellContext, render_formula


def make_ctx(**overrides) -> CellContext:
    defaults = dict(
        period_index=2,
        n_history=2,
        row=10,
        col=4,
        col_letter="D",
        prior_col_letter="C",
        named_ranges={"WACC": "WACC", "TGR": "TGR"},
        row_map={"nopat": 8, "capex": 9, "nwc_change": 10, "fcf": 11,
                 "pv_fcf": 12, "pv_terminal": 13, "revenue": 5, "ebitda": 6},
        inputs_row_map={},
        scenario_prefix="",
        first_proj_col_letter="D",
        last_proj_col_letter="H",
        entity_col_range="$B$5:$D$5",
    )
    defaults.update(overrides)
    return CellContext(**defaults)


class TestSumSubtraction:
    def test_basic(self):
        ctx = make_ctx()
        result = render_formula(
            "sum_subtraction",
            {"addend_key": "nopat", "subtrahend_keys": ["capex", "nwc_change"]},
            ctx,
        )
        assert result == "=$D$8-$D$9-$D$10"

    def test_single_subtrahend(self):
        ctx = make_ctx()
        result = render_formula(
            "sum_subtraction",
            {"addend_key": "nopat", "subtrahend_keys": ["capex"]},
            ctx,
        )
        assert result == "=$D$8-$D$9"


class TestNpvSumFixed:
    def test_spans_projection_range(self):
        ctx = make_ctx(
            first_proj_col_letter="D",
            last_proj_col_letter="H",
        )
        result = render_formula(
            "npv_sum",
            {"pv_fcf_key": "pv_fcf", "pv_terminal_key": "pv_terminal"},
            ctx,
        )
        # Should span D to H for pv_fcf, and use H for pv_terminal
        assert "SUM(D$12:H$12)" in result
        assert "H$13" in result


class TestRank:
    def test_basic(self):
        ctx = make_ctx(col_letter="B", entity_col_range="$B$5:$D$5")
        result = render_formula(
            "rank",
            {"value_key": "revenue"},
            ctx,
        )
        assert "RANK" in result
        assert "$B$5" in result
        assert "$B$5:$D$5" in result

    def test_range_uses_value_row_not_current_row(self):
        """RANK range must reference the value_key row, not the formula's own row."""
        ctx = make_ctx(
            row=12,  # formula lives on row 12
            col_letter="B",
            entity_col_range="$B$12:$D$12",  # anchored to row 12 (wrong if used directly)
        )
        # revenue is at row 5 in row_map
        result = render_formula("rank", {"value_key": "revenue"}, ctx)
        assert "$B$5:$D$5" in result, f"Expected value-row range $B$5:$D$5, got {result}"
        assert "$B$5" in result  # cell ref to revenue value


class TestIndexToBase:
    def test_basic(self):
        ctx = make_ctx(col_letter="C")
        result = render_formula(
            "index_to_base",
            {"value_key": "revenue", "base_entity_key": "company_a", "_base_col_letter": "B"},
            ctx,
        )
        assert result == "=$C$5/$B$5"


class TestBarChartText:
    def test_basic(self):
        ctx = make_ctx(col_letter="B", entity_col_range="$B$5:$D$5")
        result = render_formula(
            "bar_chart_text",
            {"value_key": "revenue"},
            ctx,
        )
        assert 'REPT("█"' in result
        assert "MAX($B$5:$D$5)" in result
        assert "*20" in result
