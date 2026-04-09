"""Shared fixtures for excel_model tests."""

import pytest

from excel_model.spec import (
    AssumptionDef,
    InputsDef,
    LineItemDef,
    MetadataDef,
    ModelSpec,
)
from excel_model.style import StyleConfig


@pytest.fixture
def basic_assumption():
    return AssumptionDef(
        name="RevenueGrowthRate",
        label="Revenue Growth Rate",
        value=0.10,
        format="percent",
        group="Growth",
    )


@pytest.fixture
def basic_line_items():
    return (
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
            label="  Cost of Goods Sold",
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
    )


@pytest.fixture
def basic_spec(basic_assumption, basic_line_items):
    return ModelSpec(
        model_type="p_and_l",
        title="Test P&L",
        currency="CHF",
        granularity="annual",
        start_period="2025",
        n_periods=3,
        n_history_periods=2,
        assumptions=(basic_assumption,),
        drivers=(),
        line_items=basic_line_items,
        metadata=MetadataDef(preparer="Test", date="2026-01-01", version="1.0"),
        scenarios=(),
        column_groups=(),
        inputs=InputsDef(source="", period_col="period", sheet="", value_cols={}),
        entities=(),
    )


@pytest.fixture
def sample_style():
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
        alt_row_fill_hex="F2F2F2",
    )
