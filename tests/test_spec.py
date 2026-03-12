"""Tests for spec.py dataclasses."""

from dataclasses import FrozenInstanceError

import pytest

from excel_model.spec import (
    AssumptionDef,
    ColumnGroupDef,
    InputsDef,
    LineItemDef,
    MetadataDef,
    ScenarioDef,
)


def test_assumption_def_creation():
    a = AssumptionDef(
        name="RevenueGrowthRate",
        label="Revenue Growth Rate",
        value=0.10,
        format="percent",
        group="Growth",
    )
    assert a.name == "RevenueGrowthRate"
    assert a.value == 0.10
    assert a.format == "percent"
    assert a.group == "Growth"


def test_assumption_def_frozen():
    a = AssumptionDef(name="X", label="X", value=1.0, format="number", group="G")
    with pytest.raises(FrozenInstanceError):  # frozen dataclass raises FrozenInstanceError
        a.name = "Y"  # type: ignore


def test_line_item_def_creation():
    li = LineItemDef(
        key="revenue",
        label="Revenue",
        formula_type="growth_projected",
        formula_params={"growth_assumption": "RevenueGrowthRate"},
        is_subtotal=False,
        is_total=False,
        section="Revenue",
        format="",
    )
    assert li.key == "revenue"
    assert li.formula_type == "growth_projected"
    assert li.formula_params["growth_assumption"] == "RevenueGrowthRate"
    assert not li.is_subtotal
    assert not li.is_total


def test_line_item_def_frozen():
    li = LineItemDef(
        key="k",
        label="L",
        formula_type="constant",
        formula_params={},
        is_subtotal=False,
        is_total=False,
        section="",
        format="",
    )
    with pytest.raises(FrozenInstanceError):
        li.key = "new"  # type: ignore


def test_scenario_def_creation():
    s = ScenarioDef(
        name="bull",
        label="Bull Case",
        assumption_overrides={"RevenueGrowthRate": 0.20},
        driver_overrides={},
    )
    assert s.name == "bull"
    assert s.assumption_overrides["RevenueGrowthRate"] == 0.20


def test_scenario_def_frozen():
    s = ScenarioDef(name="base", label="Base", assumption_overrides={}, driver_overrides={})
    with pytest.raises(FrozenInstanceError):
        s.name = "other"  # type: ignore


def test_column_group_def():
    cg = ColumnGroupDef(key="plan", label="Plan", color_hex="D6E4F0")
    assert cg.key == "plan"
    assert cg.color_hex == "D6E4F0"


def test_inputs_def():
    inp = InputsDef(
        source="data.xlsx",
        period_col="period",
        sheet="Sheet1",
        value_cols={"revenue": "Revenue"},
    )
    assert inp.source == "data.xlsx"
    assert inp.value_cols["revenue"] == "Revenue"


def test_metadata_def():
    m = MetadataDef(preparer="Test User", date="2026-01-01", version="1.0")
    assert m.preparer == "Test User"
    assert m.version == "1.0"


def test_model_spec_creation(basic_spec):
    assert basic_spec.model_type == "p_and_l"
    assert basic_spec.n_periods == 3
    assert basic_spec.n_history_periods == 2
    assert len(basic_spec.assumptions) == 1
    assert len(basic_spec.line_items) == 3


def test_model_spec_frozen(basic_spec):
    with pytest.raises(FrozenInstanceError):
        basic_spec.model_type = "dcf"  # type: ignore


def test_model_spec_assumptions_are_tuple(basic_spec):
    assert isinstance(basic_spec.assumptions, tuple)
    assert isinstance(basic_spec.line_items, tuple)
    assert isinstance(basic_spec.scenarios, tuple)
    assert isinstance(basic_spec.column_groups, tuple)
