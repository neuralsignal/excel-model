"""Tests for Budget vs Actuals model builder."""
import pytest
from openpyxl import Workbook

from excel_model.models.budget_vs_actuals import build_budget_vs_actuals
from excel_model.spec import (
    AssumptionDef,
    ColumnGroupDef,
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
def bva_spec():
    return ModelSpec(
        model_type="budget_vs_actuals",
        title="BvA Test",
        currency="CHF",
        granularity="monthly",
        start_period="2025-01",
        n_periods=3,
        n_history_periods=0,
        assumptions=(
            AssumptionDef(name="RevenueGrowthRate", label="Rev Growth",
                          value=0.08, format="percent", group="Budget"),
        ),
        line_items=(
            LineItemDef(
                key="revenue_plan",
                label="Revenue (Plan)",
                formula_type="constant",
                formula_params={"value": 1000},
                is_subtotal=False, is_total=False, section="Revenue", format="",
            ),
            LineItemDef(
                key="revenue_actual",
                label="Revenue (Actual)",
                formula_type="constant",
                formula_params={"value": 1050},
                is_subtotal=False, is_total=False, section="Revenue", format="",
            ),
            LineItemDef(
                key="revenue",
                label="Revenue Variance",
                formula_type="variance",
                formula_params={"plan_key": "revenue_plan", "actual_key": "revenue_actual"},
                is_subtotal=False, is_total=False, section="Revenue", format="",
            ),
        ),
        metadata=MetadataDef(preparer="", date="", version="1.0"),
        scenarios=(),
        column_groups=(
            ColumnGroupDef(key="plan", label="Plan", color_hex="D6E4F0"),
            ColumnGroupDef(key="actual", label="Actual", color_hex="FDFEFE"),
            ColumnGroupDef(key="variance", label="Variance", color_hex="FEF9E7"),
        ),
        inputs=InputsDef(source="", period_col="period", sheet="", value_cols={}),
        entities=(),
        drivers=(),
    )


def test_bva_creates_sheets(bva_spec, style):
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    periods = generate_periods("2025-01", 3, 0, "monthly")
    build_budget_vs_actuals(wb, bva_spec, None, style, periods)
    assert "Assumptions" in wb.sheetnames
    assert "Inputs" in wb.sheetnames
    assert "Model" in wb.sheetnames


def test_bva_model_has_sub_columns(bva_spec, style):
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    periods = generate_periods("2025-01", 3, 0, "monthly")
    build_budget_vs_actuals(wb, bva_spec, None, style, periods)
    ws = wb["Model"]
    # Row 3 should have sub-column labels (Plan, Actual, Variance)
    sub_labels = [ws.cell(row=3, column=c).value for c in range(2, 12)]
    assert "Plan" in sub_labels
    assert "Actual" in sub_labels
    assert "Variance" in sub_labels
