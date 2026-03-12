"""Tests for excel_writer.py — integration tests that write real xlsx files."""

from pathlib import Path

import pytest
from openpyxl import load_workbook

from excel_model.excel_writer import build_workbook
from excel_model.spec import (
    AssumptionDef,
    InputsDef,
    LineItemDef,
    MetadataDef,
    ModelSpec,
    ScenarioDef,
)
from excel_model.style import StyleConfig


def make_p_and_l_spec() -> ModelSpec:
    return ModelSpec(
        model_type="p_and_l",
        title="Test P&L",
        currency="CHF",
        granularity="annual",
        start_period="2025",
        n_periods=3,
        n_history_periods=2,
        assumptions=(
            AssumptionDef(
                name="RevenueGrowthRate", label="Revenue Growth Rate", value=0.10, format="percent", group="Growth"
            ),
            AssumptionDef(name="COGSMargin", label="COGS Margin", value=0.45, format="percent", group="Margins"),
        ),
        line_items=(
            LineItemDef(
                key="revenue",
                label="Revenue",
                formula_type="growth_projected",
                formula_params={"growth_assumption": "RevenueGrowthRate"},
                is_subtotal=False,
                is_total=False,
                section="Revenue",
                format="",
            ),
            LineItemDef(
                key="cogs",
                label="  COGS",
                formula_type="pct_of_revenue",
                formula_params={"revenue_key": "revenue", "rate_assumption": "COGSMargin"},
                is_subtotal=False,
                is_total=False,
                section="Cost",
                format="",
            ),
            LineItemDef(
                key="gross_profit",
                label="Gross Profit",
                formula_type="subtraction",
                formula_params={"minuend_key": "revenue", "subtrahend_key": "cogs"},
                is_subtotal=True,
                is_total=False,
                section="Profit",
                format="",
            ),
        ),
        metadata=MetadataDef(preparer="Test", date="2026-01-01", version="1.0"),
        scenarios=(),
        column_groups=(),
        inputs=InputsDef(source="", period_col="period", sheet="", value_cols={}),
        entities=(),
        drivers=(),
    )


@pytest.fixture
def p_and_l_spec():
    return make_p_and_l_spec()


@pytest.fixture
def style():
    return StyleConfig(
        header_fill_hex="1F3864",
        header_font_color="FFFFFF",
        subtotal_fill_hex="D6E4F0",
        total_fill_hex="AED6F1",
        history_col_fill_hex="F2F2F2",
        section_header_fill_hex="E8F4FD",
        font_name="Calibri",
        font_size=10,
        number_format_currency="#,##0",
        number_format_percent="0.0%",
        number_format_integer="#,##0",
        number_format_number="#,##0.00",
    )


class TestBuildWorkbook:
    def test_creates_file(self, p_and_l_spec, style, tmp_path):
        output = str(tmp_path / "test.xlsx")
        build_workbook(p_and_l_spec, None, output, style)
        assert Path(output).exists()

    def test_has_three_sheets(self, p_and_l_spec, style, tmp_path):
        output = str(tmp_path / "test.xlsx")
        build_workbook(p_and_l_spec, None, output, style)
        wb = load_workbook(output)
        assert "Assumptions" in wb.sheetnames
        assert "Inputs" in wb.sheetnames
        assert "Model" in wb.sheetnames

    def test_named_ranges_present(self, p_and_l_spec, style, tmp_path):
        output = str(tmp_path / "test.xlsx")
        build_workbook(p_and_l_spec, None, output, style)
        wb = load_workbook(output)
        assert "RevenueGrowthRate" in wb.defined_names
        assert "COGSMargin" in wb.defined_names

    def test_assumptions_values_in_cells(self, p_and_l_spec, style, tmp_path):
        output = str(tmp_path / "test.xlsx")
        build_workbook(p_and_l_spec, None, output, style)
        wb = load_workbook(output)
        ws = wb["Assumptions"]
        # Values should appear in column C (col 3)
        values = [ws.cell(row=r, column=3).value for r in range(1, 15)]
        assert 0.10 in values or any(v == 0.10 for v in values)

    def test_model_has_title(self, p_and_l_spec, style, tmp_path):
        output = str(tmp_path / "test.xlsx")
        build_workbook(p_and_l_spec, None, output, style)
        wb = load_workbook(output)
        ws = wb["Model"]
        assert ws["A1"].value == "Test P&L"

    def test_dcf_model(self, style, tmp_path):
        dcf_spec = ModelSpec(
            model_type="dcf",
            title="Test DCF",
            currency="CHF",
            granularity="annual",
            start_period="2025",
            n_periods=5,
            n_history_periods=0,
            assumptions=(AssumptionDef(name="WACC", label="WACC", value=0.10, format="percent", group="Valuation"),),
            line_items=(
                LineItemDef(
                    key="revenue",
                    label="Revenue",
                    formula_type="growth_projected",
                    formula_params={"growth_assumption": "WACC"},
                    is_subtotal=False,
                    is_total=False,
                    section="Income",
                    format="",
                ),
            ),
            metadata=MetadataDef(preparer="", date="", version="1.0"),
            scenarios=(),
            column_groups=(),
            inputs=InputsDef(source="", period_col="period", sheet="", value_cols={}),
            entities=(),
            drivers=(),
        )
        output = str(tmp_path / "dcf.xlsx")
        build_workbook(dcf_spec, None, output, style)
        wb = load_workbook(output)
        assert "Model" in wb.sheetnames
        assert "WACC" in wb.defined_names

    def test_scenario_model(self, style, tmp_path):
        scenario_spec = ModelSpec(
            model_type="scenario",
            title="Scenario Test",
            currency="CHF",
            granularity="annual",
            start_period="2025",
            n_periods=3,
            n_history_periods=0,
            assumptions=(
                AssumptionDef(
                    name="RevenueGrowthRate", label="Rev Growth", value=0.10, format="percent", group="Growth"
                ),
            ),
            line_items=(
                LineItemDef(
                    key="revenue",
                    label="Revenue",
                    formula_type="growth_projected",
                    formula_params={"growth_assumption": "RevenueGrowthRate"},
                    is_subtotal=False,
                    is_total=False,
                    section="Revenue",
                    format="",
                ),
            ),
            metadata=MetadataDef(preparer="", date="", version="1.0"),
            scenarios=(
                ScenarioDef(name="base", label="Base Case", assumption_overrides={}, driver_overrides={}),
                ScenarioDef(
                    name="bull",
                    label="Bull Case",
                    assumption_overrides={"RevenueGrowthRate": 0.20},
                    driver_overrides={},
                ),
            ),
            column_groups=(),
            inputs=InputsDef(source="", period_col="period", sheet="", value_cols={}),
            entities=(),
            drivers=(),
        )
        output = str(tmp_path / "scenario.xlsx")
        build_workbook(scenario_spec, None, output, style)
        wb = load_workbook(output)
        assert "Model" in wb.sheetnames
        assert "BullRevenueGrowthRate" in wb.defined_names
