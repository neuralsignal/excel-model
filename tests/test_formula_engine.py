"""Tests for formula_engine.py."""

import pytest

from excel_model.formula_engine import CellContext, render_formula


def make_ctx(
    period_index: int = 2,
    n_history: int = 2,
    row: int = 10,
    col: int = 4,
    col_letter: str = "D",
    prior_col_letter: str = "C",
    named_ranges: dict | None = None,
    row_map: dict | None = None,
    inputs_row_map: dict | None = None,
    scenario_prefix: str = "",
    first_proj_col_letter: str = "",
    last_proj_col_letter: str = "",
    entity_col_range: str = "",
    driver_names: frozenset[str] = frozenset(),
) -> CellContext:
    return CellContext(
        period_index=period_index,
        n_history=n_history,
        row=row,
        col=col,
        col_letter=col_letter,
        prior_col_letter=prior_col_letter,
        named_ranges=named_ranges or {"RevenueGrowthRate": "RevenueGrowthRate"},
        row_map=row_map or {"revenue": 10, "cogs": 11, "gross_profit": 12},
        inputs_row_map=inputs_row_map or {"revenue": 3, "opex": 4},
        scenario_prefix=scenario_prefix,
        first_proj_col_letter=first_proj_col_letter,
        last_proj_col_letter=last_proj_col_letter,
        entity_col_range=entity_col_range,
        driver_names=driver_names,
    )


class TestConstant:
    def test_integer(self):
        ctx = make_ctx()
        result = render_formula("constant", {"value": 42}, ctx)
        assert result == 42

    def test_float(self):
        ctx = make_ctx()
        result = render_formula("constant", {"value": 3.14}, ctx)
        assert result == 3.14

    def test_zero(self):
        ctx = make_ctx()
        result = render_formula("constant", {"value": 0}, ctx)
        assert result == 0


class TestGrowthProjected:
    def test_first_projection_uses_prior_col(self):
        # period_index == n_history == 2: first projection period
        ctx = make_ctx(period_index=2, n_history=2, col_letter="D", prior_col_letter="C")
        result = render_formula(
            "growth_projected",
            {"growth_assumption": "RevenueGrowthRate"},
            ctx,
        )
        assert result == "=$C$10*(1+RevenueGrowthRate)"

    def test_second_projection_uses_prior_projection(self):
        ctx = make_ctx(period_index=3, n_history=2, col_letter="E", prior_col_letter="D")
        result = render_formula(
            "growth_projected",
            {"growth_assumption": "RevenueGrowthRate"},
            ctx,
        )
        assert result == "=$D$10*(1+RevenueGrowthRate)"

    def test_scenario_prefix(self):
        ctx = make_ctx(period_index=2, n_history=2, scenario_prefix="Bull")
        result = render_formula(
            "growth_projected",
            {"growth_assumption": "RevenueGrowthRate"},
            ctx,
        )
        assert "BullRevenueGrowthRate" in result


class TestInputRef:
    def test_history_period_references_inputs(self):
        ctx = make_ctx(
            period_index=0,
            n_history=2,
            col_letter="B",
            prior_col_letter="",
            inputs_row_map={"revenue": 3},
        )
        result = render_formula(
            "input_ref",
            {
                "line_item_key": "revenue",
                "projected_type": "growth_projected",
                "growth_assumption": "RevenueGrowthRate",
            },
            ctx,
        )
        assert result.startswith("=Inputs!")
        assert "$B$3" in result

    def test_projection_period_routes_to_projected_type(self):
        ctx = make_ctx(
            period_index=2,
            n_history=2,
            col_letter="D",
            prior_col_letter="C",
            inputs_row_map={"revenue": 3},
        )
        result = render_formula(
            "input_ref",
            {
                "line_item_key": "revenue",
                "projected_type": "growth_projected",
                "growth_assumption": "RevenueGrowthRate",
            },
            ctx,
        )
        # Should delegate to growth_projected
        assert "RevenueGrowthRate" in result
        assert "Inputs" not in result


class TestPctOfRevenue:
    def test_basic(self):
        ctx = make_ctx(col_letter="D", row_map={"revenue": 10, "cogs": 11})
        result = render_formula(
            "pct_of_revenue",
            {"revenue_key": "revenue", "rate_assumption": "COGSMargin"},
            ctx,
        )
        assert result == "=$D$10*COGSMargin"


class TestSumOfRows:
    def test_two_addends(self):
        ctx = make_ctx(
            col_letter="D",
            row_map={"a": 10, "b": 11, "c": 12},
        )
        result = render_formula(
            "sum_of_rows",
            {"addend_keys": ["a", "b"]},
            ctx,
        )
        assert result == "=$D$10+$D$11"

    def test_three_addends(self):
        ctx = make_ctx(
            col_letter="D",
            row_map={"a": 10, "b": 11, "c": 12},
        )
        result = render_formula(
            "sum_of_rows",
            {"addend_keys": ["a", "b", "c"]},
            ctx,
        )
        assert result == "=$D$10+$D$11+$D$12"


class TestSubtraction:
    def test_basic(self):
        ctx = make_ctx(col_letter="D", row_map={"revenue": 10, "cogs": 11, "gross_profit": 12})
        result = render_formula(
            "subtraction",
            {"minuend_key": "revenue", "subtrahend_key": "cogs"},
            ctx,
        )
        assert result == "=$D$10-$D$11"


class TestRatio:
    def test_basic(self):
        ctx = make_ctx(col_letter="D", row_map={"num": 10, "den": 11})
        result = render_formula(
            "ratio",
            {"numerator_key": "num", "denominator_key": "den"},
            ctx,
        )
        assert result == "=$D$10/$D$11"


class TestGrowthRate:
    def test_with_prior(self):
        ctx = make_ctx(col_letter="D", prior_col_letter="C", row_map={"revenue": 10})
        result = render_formula("growth_rate", {"value_key": "revenue"}, ctx)
        assert "($D$10/$C$10)-1" in result

    def test_first_period_no_prior(self):
        ctx = make_ctx(col_letter="B", prior_col_letter="", row_map={"revenue": 10})
        result = render_formula("growth_rate", {"value_key": "revenue"}, ctx)
        assert result == "=0"


class TestDiscountedPv:
    def test_first_projection(self):
        # period_index=2, n_history=2 → projection_index=1
        ctx = make_ctx(period_index=2, n_history=2, col_letter="D", row_map={"fcf": 15})
        result = render_formula(
            "discounted_pv",
            {"cashflow_key": "fcf", "rate_assumption": "WACC"},
            ctx,
        )
        assert result == "=$D$15/(1+WACC)^1"

    def test_third_projection(self):
        # period_index=4, n_history=2 → projection_index=3
        ctx = make_ctx(period_index=4, n_history=2, col_letter="F", row_map={"fcf": 15})
        result = render_formula(
            "discounted_pv",
            {"cashflow_key": "fcf", "rate_assumption": "WACC"},
            ctx,
        )
        assert result == "=$F$15/(1+WACC)^3"


class TestTerminalValue:
    def test_basic(self):
        ctx = make_ctx(col_letter="F", row_map={"fcf": 15})
        result = render_formula(
            "terminal_value",
            {"cashflow_key": "fcf", "growth_assumption": "TGR", "rate_assumption": "WACC"},
            ctx,
        )
        assert result == "=$F$15*(1+TGR)/(WACC-TGR)"


class TestVariance:
    def test_basic(self):
        ctx = make_ctx(col_letter="D", row_map={"revenue_plan": 10, "revenue_actual": 11})
        result = render_formula(
            "variance",
            {"plan_key": "revenue_plan", "actual_key": "revenue_actual"},
            ctx,
        )
        assert result == "=$D$11-$D$10"


class TestVariancePct:
    def test_basic(self):
        ctx = make_ctx(col_letter="D", row_map={"revenue_plan": 10, "revenue_actual": 11})
        result = render_formula(
            "variance_pct",
            {"plan_key": "revenue_plan", "actual_key": "revenue_actual"},
            ctx,
        )
        assert "ABS" in result
        assert "$D$11" in result
        assert "$D$10" in result


class TestCustom:
    def test_basic_substitution(self):
        ctx = make_ctx(col_letter="D", row_map={"revenue": 10})
        result = render_formula(
            "custom",
            {"formula": "{col_letter}10+1"},
            ctx,
        )
        assert result == "=D10+1"


class TestCustomPrevColLetter:
    def test_prev_col_letter_replaced(self):
        ctx = make_ctx(col_letter="E", prior_col_letter="D", row_map={"total_contract": 5})
        result = render_formula(
            "custom",
            {"formula": "={prev_col_letter}{total_contract_row}*0.5"},
            ctx,
        )
        assert result == "=D5*0.5"

    def test_prev_col_letter_empty_first_period(self):
        ctx = make_ctx(col_letter="B", prior_col_letter="", row_map={"total_contract": 5})
        result = render_formula(
            "custom",
            {"formula": "={prev_col_letter}{total_contract_row}*0.5"},
            ctx,
        )
        # Empty prior_col_letter replaces to empty string
        assert result == "=5*0.5"


class TestCustomScenarioPrefix:
    def test_single_assumption_gets_prefixed(self):
        ctx = make_ctx(
            col_letter="D",
            scenario_prefix="Standard",
            named_ranges={"PatientCount": "PatientCount"},
            row_map={"revenue": 10},
        )
        result = render_formula(
            "custom",
            {"formula": "={col_letter}{revenue_row}*PatientCount"},
            ctx,
        )
        assert "StandardPatientCount" in result
        assert result == "=D10*StandardPatientCount"

    def test_multiple_assumptions_prefixed(self):
        ctx = make_ctx(
            col_letter="D",
            scenario_prefix="Bull",
            named_ranges={"PatientCount": "PatientCount", "PricePerPatient": "PricePerPatient"},
            row_map={"revenue": 10},
        )
        result = render_formula(
            "custom",
            {"formula": "=PatientCount*PricePerPatient"},
            ctx,
        )
        assert "BullPatientCount" in result
        assert "BullPricePerPatient" in result

    def test_no_prefix_when_empty(self):
        ctx = make_ctx(
            col_letter="D",
            scenario_prefix="",
            named_ranges={"PatientCount": "PatientCount"},
            row_map={"revenue": 10},
        )
        result = render_formula(
            "custom",
            {"formula": "=PatientCount*2"},
            ctx,
        )
        assert result == "=PatientCount*2"

    def test_longer_name_replaced_before_shorter_substring(self):
        ctx = make_ctx(
            col_letter="D",
            scenario_prefix="Std",
            named_ranges={"Patient": "Patient", "PatientCount": "PatientCount"},
            row_map={"revenue": 10},
        )
        result = render_formula(
            "custom",
            {"formula": "=PatientCount+Patient"},
            ctx,
        )
        # PatientCount replaced first (longer), so Patient inside PatientCount
        # is NOT double-prefixed
        assert result == "=StdPatientCount+StdPatient"


class TestResolveNameWithDrivers:
    def test_driver_gets_prefixed(self):
        ctx = make_ctx(
            period_index=2,
            n_history=2,
            scenario_prefix="Bull",
            named_ranges={"RevenueGrowthRate": "RevenueGrowthRate", "PatientCount": "PatientCount"},
        )
        # Manually create ctx with driver_names
        ctx = CellContext(
            period_index=2,
            n_history=2,
            row=10,
            col=4,
            col_letter="D",
            prior_col_letter="C",
            named_ranges={"RevenueGrowthRate": "RevenueGrowthRate", "PatientCount": "PatientCount"},
            row_map={"revenue": 10},
            inputs_row_map={},
            scenario_prefix="Bull",
            first_proj_col_letter="",
            last_proj_col_letter="",
            entity_col_range="",
            driver_names=frozenset({"PatientCount"}),
        )
        result = render_formula(
            "growth_projected",
            {"growth_assumption": "PatientCount"},
            ctx,
        )
        assert "BullPatientCount" in result

    def test_assumption_stays_bare_when_driver_names_set(self):
        ctx = CellContext(
            period_index=2,
            n_history=2,
            row=10,
            col=4,
            col_letter="D",
            prior_col_letter="C",
            named_ranges={"RevenueGrowthRate": "RevenueGrowthRate", "PatientCount": "PatientCount"},
            row_map={"revenue": 10},
            inputs_row_map={},
            scenario_prefix="Bull",
            first_proj_col_letter="",
            last_proj_col_letter="",
            entity_col_range="",
            driver_names=frozenset({"PatientCount"}),
        )
        result = render_formula(
            "growth_projected",
            {"growth_assumption": "RevenueGrowthRate"},
            ctx,
        )
        # RevenueGrowthRate is an assumption, not a driver — should NOT be prefixed
        assert "RevenueGrowthRate" in result
        assert "BullRevenueGrowthRate" not in result

    def test_legacy_mode_prefixes_all(self):
        ctx = CellContext(
            period_index=2,
            n_history=2,
            row=10,
            col=4,
            col_letter="D",
            prior_col_letter="C",
            named_ranges={"RevenueGrowthRate": "RevenueGrowthRate"},
            row_map={"revenue": 10},
            inputs_row_map={},
            scenario_prefix="Bull",
            first_proj_col_letter="",
            last_proj_col_letter="",
            entity_col_range="",
            driver_names=frozenset(),  # empty = legacy mode
        )
        result = render_formula(
            "growth_projected",
            {"growth_assumption": "RevenueGrowthRate"},
            ctx,
        )
        assert "BullRevenueGrowthRate" in result

    def test_no_prefix_when_scenario_prefix_empty(self):
        ctx = CellContext(
            period_index=2,
            n_history=2,
            row=10,
            col=4,
            col_letter="D",
            prior_col_letter="C",
            named_ranges={"PatientCount": "PatientCount"},
            row_map={"revenue": 10},
            inputs_row_map={},
            scenario_prefix="",
            first_proj_col_letter="",
            last_proj_col_letter="",
            entity_col_range="",
            driver_names=frozenset({"PatientCount"}),
        )
        result = render_formula(
            "growth_projected",
            {"growth_assumption": "PatientCount"},
            ctx,
        )
        assert "PatientCount" in result
        # Should NOT have any prefix
        assert result.count("PatientCount") == 1


class TestCustomFormulaWithDrivers:
    def test_only_driver_names_prefixed(self):
        ctx = CellContext(
            period_index=0,
            n_history=0,
            row=10,
            col=4,
            col_letter="D",
            prior_col_letter="C",
            named_ranges={"PatientCount": "PatientCount", "CROPerPatient": "CROPerPatient"},
            row_map={"revenue": 10},
            inputs_row_map={},
            scenario_prefix="Standard",
            first_proj_col_letter="",
            last_proj_col_letter="",
            entity_col_range="",
            driver_names=frozenset({"PatientCount"}),
        )
        result = render_formula(
            "custom",
            {"formula": "=PatientCount*CROPerPatient"},
            ctx,
        )
        assert "StandardPatientCount" in result
        # CROPerPatient is an assumption — should NOT be prefixed
        assert "CROPerPatient" in result
        assert "StandardCROPerPatient" not in result

    def test_assumption_names_left_bare_in_mixed_formula(self):
        ctx = CellContext(
            period_index=0,
            n_history=0,
            row=10,
            col=4,
            col_letter="D",
            prior_col_letter="C",
            named_ranges={
                "PatientCount": "PatientCount",
                "PerPatientPrice": "PerPatientPrice",
                "CROPerPatient": "CROPerPatient",
            },
            row_map={"total_contract": 5, "revenue": 10},
            inputs_row_map={},
            scenario_prefix="Premium",
            first_proj_col_letter="",
            last_proj_col_letter="",
            entity_col_range="",
            driver_names=frozenset({"PatientCount", "PerPatientPrice"}),
        )
        result = render_formula(
            "custom",
            {"formula": "=${col_letter}${total_contract_row}/CROPerPatient"},
            ctx,
        )
        assert "CROPerPatient" in result
        assert "PremiumCROPerPatient" not in result


class TestInvalidFormulaType:
    def test_raises(self):
        ctx = make_ctx()
        with pytest.raises(ValueError):
            render_formula("nonexistent_type", {}, ctx)
