"""Tests for validator.py."""

import polars as pl

from excel_model.loader import InputData
from excel_model.spec import (
    AssumptionDef,
    InputsDef,
    LineItemDef,
    MetadataDef,
    ModelSpec,
)
from excel_model.validator import validate_inputs_against_spec, validate_spec


def make_minimal_spec(**overrides) -> ModelSpec:
    defaults = dict(
        model_type="p_and_l",
        title="Test",
        currency="CHF",
        granularity="annual",
        start_period="2025",
        n_periods=3,
        n_history_periods=0,
        assumptions=(),
        drivers=(),
        line_items=(),
        metadata=MetadataDef(preparer="", date="", version="1.0"),
        scenarios=(),
        column_groups=(),
        inputs=InputsDef(source="", period_col="period", sheet="", value_cols={}),
        entities=(),
    )
    defaults.update(overrides)
    return ModelSpec(**defaults)


class TestValidateSpec:
    def test_valid_spec_returns_empty(self):
        spec = make_minimal_spec()
        errors = validate_spec(spec)
        assert errors == []

    def test_invalid_model_type(self):
        spec = make_minimal_spec(model_type="unknown_type")
        errors = validate_spec(spec)
        assert any("model_type" in e for e in errors)

    def test_empty_title(self):
        spec = make_minimal_spec(title="")
        errors = validate_spec(spec)
        assert any("title" in e for e in errors)

    def test_invalid_granularity(self):
        spec = make_minimal_spec(granularity="weekly")
        errors = validate_spec(spec)
        assert any("granularity" in e for e in errors)

    def test_n_periods_zero(self):
        spec = make_minimal_spec(n_periods=0)
        errors = validate_spec(spec)
        assert any("n_periods" in e for e in errors)

    def test_n_history_negative(self):
        spec = make_minimal_spec(n_history_periods=-1)
        errors = validate_spec(spec)
        assert any("n_history_periods" in e for e in errors)

    def test_duplicate_assumption_names(self):
        dup = (
            AssumptionDef(name="Rate", label="Rate", value=0.1, format="percent", group="G"),
            AssumptionDef(name="Rate", label="Rate2", value=0.2, format="percent", group="G"),
        )
        spec = make_minimal_spec(assumptions=dup)
        errors = validate_spec(spec)
        assert any("Duplicate assumption" in e for e in errors)

    def test_duplicate_line_item_keys(self):
        dup = (
            LineItemDef(
                key="revenue",
                label="Revenue",
                formula_type="constant",
                formula_params={"value": 1},
                is_subtotal=False,
                is_total=False,
                section="",
                format="",
            ),
            LineItemDef(
                key="revenue",
                label="Revenue 2",
                formula_type="constant",
                formula_params={"value": 2},
                is_subtotal=False,
                is_total=False,
                section="",
                format="",
            ),
        )
        spec = make_minimal_spec(line_items=dup)
        errors = validate_spec(spec)
        assert any("Duplicate line item" in e for e in errors)

    def test_unknown_formula_type(self):
        li = (
            LineItemDef(
                key="x",
                label="X",
                formula_type="nonexistent",
                formula_params={},
                is_subtotal=False,
                is_total=False,
                section="",
                format="",
            ),
        )
        spec = make_minimal_spec(line_items=li)
        errors = validate_spec(spec)
        assert any("formula_type" in e for e in errors)

    def test_scenario_model_without_scenarios(self):
        spec = make_minimal_spec(model_type="scenario", scenarios=())
        errors = validate_spec(spec)
        assert any("scenario" in e.lower() for e in errors)

    def test_bva_model_without_column_groups(self):
        spec = make_minimal_spec(model_type="budget_vs_actuals", column_groups=())
        errors = validate_spec(spec)
        assert any("column_groups" in e for e in errors)

    def test_invalid_assumption_format(self):
        bad = (AssumptionDef(name="X", label="X", value=1.0, format="bogus", group="G"),)
        spec = make_minimal_spec(assumptions=bad)
        errors = validate_spec(spec)
        assert any("format" in e for e in errors)


class TestValidateInputsAgainstSpec:
    def test_valid_inputs(self):
        spec = make_minimal_spec(
            inputs=InputsDef(
                source="",
                period_col="period",
                sheet="",
                value_cols={"revenue": "revenue"},
            )
        )
        df = pl.DataFrame({"period": ["2025"], "revenue": [100.0]})
        inputs = InputData(df=df, period_col="period", value_cols=["revenue"])
        errors = validate_inputs_against_spec(spec, inputs)
        assert errors == []

    def test_missing_value_col(self):
        spec = make_minimal_spec(
            inputs=InputsDef(
                source="",
                period_col="period",
                sheet="",
                value_cols={"revenue": "revenue_column"},
            )
        )
        df = pl.DataFrame({"period": ["2025"], "other": [100.0]})
        inputs = InputData(df=df, period_col="period", value_cols=["other"])
        errors = validate_inputs_against_spec(spec, inputs)
        assert any("revenue_column" in e for e in errors)
