"""Tests for excel_model.spec_loader."""

import pytest
from strictyaml import StrictYAMLError

from excel_model.spec import ModelSpec
from excel_model.spec_loader import load_spec

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
