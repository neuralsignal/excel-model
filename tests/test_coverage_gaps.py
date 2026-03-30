"""Tests covering untested branches in formula_engine, validator, and time_engine (#60)."""

import polars as pl
import pytest
from hypothesis import given
from hypothesis import strategies as st

from excel_model.exceptions import FormulaInjectionError
from excel_model.formula_engine import CellContext, _row_ref, render_formula
from excel_model.loader import InputData
from excel_model.spec import (
    DriverDef,
    InputsDef,
    MetadataDef,
    ModelSpec,
)
from excel_model.time_engine import _parse_monthly, _parse_quarterly
from excel_model.validator import (
    validate_custom_formula,
    validate_inputs_against_spec,
    validate_spec,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_ctx(**overrides: object) -> CellContext:
    defaults = dict(
        period_index=2,
        n_history=2,
        row=10,
        col=4,
        col_letter="D",
        prior_col_letter="C",
        named_ranges={},
        row_map={"revenue": 5},
        inputs_row_map={},
        scenario_prefix="",
        first_proj_col_letter="D",
        last_proj_col_letter="H",
        entity_col_range="$B$5:$D$5",
        driver_names=frozenset(),
    )
    defaults.update(overrides)
    return CellContext(**defaults)


def _make_spec(**overrides: object) -> ModelSpec:
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


# ===========================================================================
# formula_engine — _row_ref KeyError (line 62)
# ===========================================================================


class TestRowRefMissingKey:
    def test_missing_key_raises_key_error(self) -> None:
        with pytest.raises(KeyError, match="missing"):
            _row_ref("missing", "B", {})

    @given(key=st.text(min_size=1, max_size=20))
    def test_arbitrary_missing_key_raises(self, key: str) -> None:
        with pytest.raises(KeyError):
            _row_ref(key, "B", {})


# ===========================================================================
# formula_engine — _render_growth_projected with prior_key (lines 119-120, 124-125)
# ===========================================================================


class TestGrowthProjectedWithPriorKey:
    def test_prior_key_first_projection_period(self) -> None:
        """Line 119-120: prior_key in row_map on the first projection period."""
        ctx = _make_ctx(
            period_index=2,
            n_history=2,
            row=10,
            prior_col_letter="C",
            row_map={"revenue": 5, "base_rev": 7},
        )
        result = render_formula(
            "growth_projected",
            {"growth_assumption": "GrowthRate", "prior_key": "base_rev"},
            ctx,
        )
        # Should reference base_rev row (7), not current row (10)
        assert "$C$7" in result
        assert "GrowthRate" in result

    def test_prior_key_subsequent_projection_period(self) -> None:
        """Line 124-125: prior_key in row_map on a subsequent projection period."""
        ctx = _make_ctx(
            period_index=3,
            n_history=2,
            row=10,
            prior_col_letter="C",
            row_map={"revenue": 5, "base_rev": 7},
        )
        result = render_formula(
            "growth_projected",
            {"growth_assumption": "GrowthRate", "prior_key": "base_rev"},
            ctx,
        )
        assert "$C$7" in result
        assert "GrowthRate" in result


# ===========================================================================
# validator — DDE pipe pattern (line 47)
# ===========================================================================


class TestFormulaInjectionDdePipe:
    def test_dde_pipe_pattern_raises(self) -> None:
        # Use a word that doesn't match _DANGEROUS_FORMULA_PATTERNS but triggers _DDE_PIPE_RE
        with pytest.raises(FormulaInjectionError, match="pipe-based DDE"):
            validate_custom_formula("=STUFF|'/c calc'!A0", "item")


# ===========================================================================
# validator — empty currency (line 119), empty start_period (line 125)
# ===========================================================================


class TestValidateSpecEmptyFields:
    def test_empty_currency(self) -> None:
        spec = _make_spec(currency="")
        errors = validate_spec(spec)
        assert any("currency" in e for e in errors)

    def test_empty_start_period(self) -> None:
        spec = _make_spec(start_period="")
        errors = validate_spec(spec)
        assert any("start_period" in e for e in errors)


# ===========================================================================
# validator — invalid driver format (line 188)
# ===========================================================================


class TestValidateDriverInvalidFormat:
    def test_invalid_driver_format(self) -> None:
        spec = _make_spec(
            drivers=(DriverDef(name="Volume", label="Vol", value=100, format="bogus", group="V"),),
        )
        errors = validate_spec(spec)
        assert any("Driver" in e and "format" in e for e in errors)

    @given(fmt=st.text(min_size=1, max_size=10).filter(lambda s: s not in {"number", "percent", "currency", "integer"}))
    def test_arbitrary_invalid_driver_format(self, fmt: str) -> None:
        spec = _make_spec(
            drivers=(DriverDef(name="Volume", label="Vol", value=100, format=fmt, group="V"),),
        )
        errors = validate_spec(spec)
        assert any("format" in e for e in errors)


# ===========================================================================
# validator — period_col not in DataFrame columns (line 307)
# ===========================================================================


class TestValidateInputsPeriodColMissing:
    def test_period_col_missing_from_dataframe(self) -> None:
        spec = _make_spec(
            inputs=InputsDef(source="", period_col="date", sheet="", value_cols={}),
        )
        df = pl.DataFrame({"other_col": [1, 2, 3]})
        inputs = InputData(df=df, period_col="missing_period", value_cols=["other_col"])
        errors = validate_inputs_against_spec(spec, inputs)
        assert any("period_col" in e for e in errors)


# ===========================================================================
# time_engine — _parse_quarterly invalid (line 66), _parse_monthly invalid (line 79)
# ===========================================================================


class TestParseQuarterlyInvalid:
    def test_invalid_format_raises(self) -> None:
        with pytest.raises(ValueError, match="Cannot parse quarterly"):
            _parse_quarterly("bad-format")

    @given(
        s=st.text(min_size=1, max_size=20).filter(
            lambda s: not s.strip().startswith("Q") and not s.strip()[:4].isdigit()
        )
    )
    def test_arbitrary_invalid_quarterly(self, s: str) -> None:
        with pytest.raises(ValueError, match="Cannot parse quarterly"):
            _parse_quarterly(s)


class TestParseMonthlyInvalid:
    def test_invalid_format_raises(self) -> None:
        with pytest.raises(ValueError, match="Cannot parse monthly"):
            _parse_monthly("bad-format")

    @given(s=st.from_regex(r"[a-z]{3,5}", fullmatch=True))
    def test_arbitrary_invalid_monthly(self, s: str) -> None:
        with pytest.raises(ValueError, match="Cannot parse monthly"):
            _parse_monthly(s)
