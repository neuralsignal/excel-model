"""Tests for excel_model.spec_loader."""

import pytest
from hypothesis import given
from hypothesis import strategies as st
from strictyaml import StrictYAMLError

from excel_model.spec import InputsDef, ModelSpec
from excel_model.spec_loader import _build_inputs, load_spec

VALID_P_AND_L_YAML = """\
model_type: p_and_l
title: Test P&L
currency: CHF
granularity: annual
start_period: "2025"
n_periods: 3
n_history_periods: 2
"""


def test_valid_p_and_l_spec(tmp_path):
    """Load a valid spec and check key fields."""
    spec_file = tmp_path / "spec.yaml"
    spec_file.write_text(VALID_P_AND_L_YAML)
    spec = load_spec(str(spec_file))
    assert isinstance(spec, ModelSpec)
    assert spec.model_type == "p_and_l"
    assert spec.title == "Test P&L"
    assert spec.currency == "CHF"
    assert spec.n_periods == 3
    assert spec.n_history_periods == 2


def test_missing_required_field_raises(tmp_path):
    """YAML missing n_periods raises StrictYAMLError."""
    yaml_text = """\
model_type: p_and_l
title: Test
currency: CHF
granularity: annual
start_period: "2025"
n_history_periods: 2
"""
    spec_file = tmp_path / "bad_spec.yaml"
    spec_file.write_text(yaml_text)
    with pytest.raises(StrictYAMLError):
        load_spec(str(spec_file))


def test_invalid_model_type_raises(tmp_path):
    """Invalid model_type raises StrictYAMLError."""
    yaml_text = """\
model_type: invalid_type
title: Test
currency: CHF
granularity: annual
start_period: "2025"
n_periods: 3
n_history_periods: 2
"""
    spec_file = tmp_path / "bad_spec.yaml"
    spec_file.write_text(yaml_text)
    with pytest.raises(StrictYAMLError):
        load_spec(str(spec_file))


def test_invalid_assumption_format_raises(tmp_path):
    """Invalid assumption format raises StrictYAMLError."""
    yaml_text = """\
model_type: p_and_l
title: Test
currency: CHF
granularity: annual
start_period: "2025"
n_periods: 3
n_history_periods: 2
assumptions:
  - name: growth_rate
    label: Growth Rate
    value: 0.10
    format: bad_format
"""
    spec_file = tmp_path / "bad_spec.yaml"
    spec_file.write_text(yaml_text)
    with pytest.raises(StrictYAMLError):
        load_spec(str(spec_file))


def test_optional_fields_default_empty(tmp_path):
    """Spec with no assumptions key results in empty assumptions tuple."""
    spec_file = tmp_path / "spec.yaml"
    spec_file.write_text(VALID_P_AND_L_YAML)
    spec = load_spec(str(spec_file))
    assert len(spec.assumptions) == 0


def test_value_types(tmp_path):
    """Float value 0.10 parses as float; integer value 21 parses as int."""
    yaml_text = """\
model_type: p_and_l
title: Test
currency: CHF
granularity: annual
start_period: "2025"
n_periods: 3
n_history_periods: 2
assumptions:
  - name: rate
    label: Rate
    value: 0.10
    format: percent
  - name: years
    label: Years
    value: 21
    format: integer
"""
    spec_file = tmp_path / "spec.yaml"
    spec_file.write_text(yaml_text)
    spec = load_spec(str(spec_file))
    assumptions = list(spec.assumptions)
    assert isinstance(assumptions[0].value, float)
    assert isinstance(assumptions[1].value, int)


def test_drivers_section_parsed(tmp_path):
    """YAML with drivers section parses correctly."""
    yaml_text = """\
model_type: scenario
title: Driver Test
currency: EUR
granularity: annual
start_period: "2026"
n_periods: 1
n_history_periods: 0
drivers:
  - name: PatientCount
    label: Patient Count
    value: 2500
    format: integer
    group: Volume
"""
    spec_file = tmp_path / "spec.yaml"
    spec_file.write_text(yaml_text)
    spec = load_spec(str(spec_file))
    assert len(spec.drivers) == 1
    assert spec.drivers[0].name == "PatientCount"
    assert spec.drivers[0].value == 2500


def test_no_drivers_returns_empty_tuple(tmp_path):
    """YAML without drivers returns empty tuple."""
    spec_file = tmp_path / "spec.yaml"
    spec_file.write_text(VALID_P_AND_L_YAML)
    spec = load_spec(str(spec_file))
    assert spec.drivers == ()


def test_driver_overrides_parsed(tmp_path):
    """driver_overrides in scenarios parses correctly."""
    yaml_text = """\
model_type: scenario
title: Test
currency: EUR
granularity: annual
start_period: "2026"
n_periods: 1
n_history_periods: 0
drivers:
  - name: PatientCount
    label: PC
    value: 100
    format: integer
scenarios:
  - name: base
    label: Base
  - name: premium
    label: Premium
    driver_overrides:
      PatientCount: 500
"""
    spec_file = tmp_path / "spec.yaml"
    spec_file.write_text(yaml_text)
    spec = load_spec(str(spec_file))
    assert spec.scenarios[0].driver_overrides == {}
    assert spec.scenarios[1].driver_overrides == {"PatientCount": 500}


def test_missing_spec_file_raises_file_not_found() -> None:
    """load_spec with a non-existent path raises FileNotFoundError."""
    with pytest.raises(FileNotFoundError, match="Model spec file not found"):
        load_spec("/nonexistent/path/spec.yaml")


def test_column_groups_section_parsed(tmp_path) -> None:
    """YAML with column_groups section parses correctly."""
    yaml_text = """\
model_type: budget_vs_actuals
title: BVA Test
currency: USD
granularity: annual
start_period: "2025"
n_periods: 2
n_history_periods: 1
column_groups:
  - key: budget
    label: Budget
    color_hex: "4472C4"
  - key: actual
    label: Actual
    color_hex: "70AD47"
"""
    spec_file = tmp_path / "spec.yaml"
    spec_file.write_text(yaml_text)
    spec = load_spec(str(spec_file))
    assert len(spec.column_groups) == 2
    assert spec.column_groups[0].key == "budget"
    assert spec.column_groups[0].label == "Budget"
    assert spec.column_groups[0].color_hex == "4472C4"
    assert spec.column_groups[1].key == "actual"


def test_entities_section_parsed(tmp_path) -> None:
    """YAML with entities section parses correctly."""
    yaml_text = """\
model_type: comparison
title: Entity Test
currency: EUR
granularity: annual
start_period: "2025"
n_periods: 3
n_history_periods: 0
entities:
  - key: unit_a
    label: Unit A
  - key: unit_b
    label: Unit B
"""
    spec_file = tmp_path / "spec.yaml"
    spec_file.write_text(yaml_text)
    spec = load_spec(str(spec_file))
    assert len(spec.entities) == 2
    assert spec.entities[0].key == "unit_a"
    assert spec.entities[0].label == "Unit A"
    assert spec.entities[1].key == "unit_b"


def test_inputs_section_parsed(tmp_path) -> None:
    """YAML with inputs block parses correctly."""
    yaml_text = """\
model_type: budget_vs_actuals
title: Inputs Test
currency: USD
granularity: annual
start_period: "2025"
n_periods: 2
n_history_periods: 1
inputs:
  source: data/actuals.csv
  period_col: month
  sheet: Sheet1
  value_cols:
    revenue: Revenue
    cost: Cost
"""
    spec_file = tmp_path / "spec.yaml"
    spec_file.write_text(yaml_text)
    spec = load_spec(str(spec_file))
    assert spec.inputs.source == "data/actuals.csv"
    assert spec.inputs.period_col == "month"
    assert spec.inputs.sheet == "Sheet1"
    assert spec.inputs.value_cols == {"revenue": "Revenue", "cost": "Cost"}


def test_metadata_section_parsed(tmp_path) -> None:
    """YAML with metadata block parses correctly."""
    yaml_text = """\
model_type: p_and_l
title: Metadata Test
currency: CHF
granularity: annual
start_period: "2025"
n_periods: 3
n_history_periods: 2
metadata:
  preparer: Jane Doe
  date: "2025-01-15"
  version: "2.0"
"""
    spec_file = tmp_path / "spec.yaml"
    spec_file.write_text(yaml_text)
    spec = load_spec(str(spec_file))
    assert spec.metadata.preparer == "Jane Doe"
    assert spec.metadata.date == "2025-01-15"
    assert spec.metadata.version == "2.0"


@given(
    source=st.text(min_size=1, max_size=50),
    period_col=st.text(min_size=1, max_size=20),
    sheet=st.text(min_size=0, max_size=30),
)
def test_build_inputs_returns_inputs_def(source: str, period_col: str, sheet: str) -> None:
    """_build_inputs always returns an InputsDef for valid input dicts."""
    raw = {"source": source, "period_col": period_col, "sheet": sheet}
    result = _build_inputs(raw)
    assert isinstance(result, InputsDef)
    assert result.source == source
    assert result.period_col == period_col
    assert result.sheet == sheet
    assert result.value_cols == {}
