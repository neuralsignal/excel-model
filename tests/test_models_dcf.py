"""Tests for DCF model builder."""
import pytest
from openpyxl import Workbook

from excel_model.models.dcf import build_dcf
from excel_model.spec import (
    AssumptionDef,
    InputsDef,
    LineItemDef,
    MetadataDef,
    ModelSpec,
)
from excel_model.style import StyleConfig
from excel_model.time_engine import generate_periods


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


@pytest.fixture
def dcf_spec():
    return ModelSpec(
        model_type="dcf",
        title="DCF Test",
        currency="CHF",
        granularity="annual",
        start_period="2025",
        n_periods=5,
        n_history_periods=0,
        assumptions=(
            AssumptionDef(name="WACC", label="Discount Rate", value=0.10, format="percent", group="Valuation"),
            AssumptionDef(name="TGR", label="Terminal Growth Rate", value=0.02, format="percent", group="Valuation"),
            AssumptionDef(name="RevenueGrowthRate", label="Revenue Growth", value=0.10, format="percent", group="Growth"),
        ),
        line_items=(
            LineItemDef(
                key="revenue", label="Revenue",
                formula_type="growth_projected",
                formula_params={"growth_assumption": "RevenueGrowthRate"},
                is_subtotal=False, is_total=False, section="Income", format="",
            ),
            LineItemDef(
                key="fcf", label="Free Cash Flow",
                formula_type="constant",
                formula_params={"value": 0},
                is_subtotal=True, is_total=False, section="FCF", format="",
            ),
            LineItemDef(
                key="pv_fcf", label="PV of FCF",
                formula_type="discounted_pv",
                formula_params={"cashflow_key": "fcf", "rate_assumption": "WACC"},
                is_subtotal=False, is_total=False, section="Valuation", format="",
            ),
            LineItemDef(
                key="terminal_value", label="Terminal Value",
                formula_type="terminal_value",
                formula_params={"cashflow_key": "fcf", "growth_assumption": "TGR", "rate_assumption": "WACC"},
                is_subtotal=False, is_total=False, section="Valuation", format="",
            ),
        ),
        metadata=MetadataDef(preparer="", date="", version="1.0"),
        scenarios=(),
        column_groups=(),
        inputs=InputsDef(source="", period_col="period", sheet="", value_cols={}),
        entities=(),
        drivers=(),
    )


def test_dcf_creates_sheets(dcf_spec, style):
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    periods = generate_periods("2025", 5, 0, "annual")
    build_dcf(wb, dcf_spec, None, style, periods)
    assert "Assumptions" in wb.sheetnames
    assert "Inputs" in wb.sheetnames
    assert "Model" in wb.sheetnames


def test_dcf_discounted_pv_formula(dcf_spec, style):
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    periods = generate_periods("2025", 5, 0, "annual")
    build_dcf(wb, dcf_spec, None, style, periods)

    ws = wb["Model"]
    # Should have a discounted_pv formula with WACC
    found = False
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "WACC" in cell.value and "^" in cell.value:
                found = True
    assert found, "No discounted PV formula found"


def test_dcf_terminal_value_formula(dcf_spec, style):
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    periods = generate_periods("2025", 5, 0, "annual")
    build_dcf(wb, dcf_spec, None, style, periods)

    ws = wb["Model"]
    found = False
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "TGR" in cell.value:
                found = True
    assert found, "No terminal value formula (with TGR) found"
