"""Tests for excel_model.cli — build, validate, describe commands."""

import json
from unittest.mock import patch

import pytest
from click.testing import CliRunner

from excel_model.cli import _build_description, _render_description_text, main

VALID_P_AND_L_YAML = """\
model_type: p_and_l
title: Test P&L
currency: CHF
granularity: annual
start_period: "2025"
n_periods: 3
n_history_periods: 2
assumptions:
  - name: RevenueGrowthRate
    label: Revenue Growth Rate
    value: 0.10
    format: percent
    group: Growth
line_items:
  - key: revenue
    label: Revenue
    formula_type: growth_projected
    formula_params:
      growth_assumption: RevenueGrowthRate
    is_subtotal: false
    is_total: false
    section: Revenue
"""

MINIMAL_YAML = """\
model_type: p_and_l
title: Minimal
currency: CHF
granularity: annual
start_period: "2025"
n_periods: 3
n_history_periods: 0
"""

INVALID_SPEC_YAML = """\
model_type: p_and_l
title: Test
currency: CHF
granularity: annual
start_period: "2025"
n_periods: 3
n_history_periods: 2
line_items:
  - key: revenue
    label: Revenue
    formula_type: bogus_formula_type
    is_subtotal: false
    is_total: false
    section: Revenue
"""

SCENARIO_YAML = """\
model_type: scenario
title: Scenario Test
currency: CHF
granularity: annual
start_period: "2025"
n_periods: 3
n_history_periods: 0
assumptions:
  - name: GrowthRate
    label: Growth Rate
    value: 0.10
    format: percent
    group: Growth
scenarios:
  - name: base
    label: Base Case
  - name: bull
    label: Bull Case
    assumption_overrides:
      GrowthRate: 0.20
"""


@pytest.fixture
def runner():
    return CliRunner()


@pytest.fixture
def spec_file(tmp_path):
    p = tmp_path / "spec.yaml"
    p.write_text(VALID_P_AND_L_YAML)
    return p


# ---------------------------------------------------------------------------
# build command
# ---------------------------------------------------------------------------


class TestBuildBatchSuccess:
    def test_build_batch_outputs_json_ok(self, runner, spec_file, tmp_path):
        out_path = tmp_path / "out.xlsx"
        result = runner.invoke(
            main,
            ["build", "--spec", str(spec_file), "--output", str(out_path), "--mode", "batch"],
        )
        assert result.exit_code == 0, result.output
        payload = json.loads(result.output.strip())
        assert payload["status"] == "ok"
        assert "output" in payload
        assert out_path.exists()


class TestBuildInteractiveSuccess:
    def test_build_interactive_prints_info(self, runner, spec_file, tmp_path):
        out_path = tmp_path / "out.xlsx"
        result = runner.invoke(
            main,
            ["build", "--spec", str(spec_file), "--output", str(out_path), "--mode", "interactive"],
        )
        assert result.exit_code == 0, result.output
        assert "Loading model spec" in result.output
        assert "Validating model spec" in result.output
        assert "Building workbook" in result.output
        assert "Workbook saved to" in result.output


class TestBuildSpecLoadFailure:
    def test_batch_file_not_found(self, runner, tmp_path):
        """Simulate FileNotFoundError from load_spec via mock."""
        out_path = tmp_path / "out.xlsx"
        # Create a valid file so click's exists=True passes, but mock load_spec
        dummy = tmp_path / "dummy.yaml"
        dummy.write_text("placeholder")
        with patch(
            "excel_model.spec_loader.load_spec",
            side_effect=FileNotFoundError("spec not found"),
        ):
            result = runner.invoke(
                main,
                ["build", "--spec", str(dummy), "--output", str(out_path), "--mode", "batch"],
            )
        assert result.exit_code != 0
        payload = json.loads(result.output.strip())
        assert payload["status"] == "error"
        assert "Failed to load spec" in payload["message"]

    def test_interactive_spec_load_error(self, runner, tmp_path):
        dummy = tmp_path / "dummy.yaml"
        dummy.write_text("placeholder")
        out_path = tmp_path / "out.xlsx"
        with patch(
            "excel_model.spec_loader.load_spec",
            side_effect=ValueError("bad spec"),
        ):
            result = runner.invoke(
                main,
                ["build", "--spec", str(dummy), "--output", str(out_path), "--mode", "interactive"],
            )
        assert result.exit_code != 0
        assert "ERROR" in result.output


class TestBuildValidationFailure:
    def test_batch_validation_error(self, runner, tmp_path):
        spec_f = tmp_path / "spec.yaml"
        spec_f.write_text(INVALID_SPEC_YAML)
        out_path = tmp_path / "out.xlsx"
        result = runner.invoke(
            main,
            ["build", "--spec", str(spec_f), "--output", str(out_path), "--mode", "batch"],
        )
        assert result.exit_code != 0
        payload = json.loads(result.output.strip())
        assert payload["status"] == "error"
        assert "Spec validation failed" in payload["message"]


class TestBuildStyleFailure:
    def test_batch_style_config_error(self, runner, spec_file, tmp_path):
        from excel_model.exceptions import StyleConfigError

        out_path = tmp_path / "out.xlsx"
        with patch(
            "excel_model.cli.load_style",
            side_effect=StyleConfigError("bad style"),
        ):
            result = runner.invoke(
                main,
                ["build", "--spec", str(spec_file), "--output", str(out_path), "--mode", "batch"],
            )
        assert result.exit_code != 0
        payload = json.loads(result.output.strip())
        assert payload["status"] == "error"
        assert "style" in payload["message"].lower()

    def test_interactive_style_config_error(self, runner, spec_file, tmp_path):
        """Line 85: interactive build with StyleConfigError."""
        from excel_model.exceptions import StyleConfigError

        out_path = tmp_path / "out.xlsx"
        with patch(
            "excel_model.cli.load_style",
            side_effect=StyleConfigError("bad style config"),
        ):
            result = runner.invoke(
                main,
                ["build", "--spec", str(spec_file), "--output", str(out_path), "--mode", "interactive"],
            )
        assert result.exit_code != 0
        assert "bad style config" in result.output


class TestBuildWorkbookFailure:
    def test_batch_excel_model_error(self, runner, spec_file, tmp_path):
        from excel_model.exceptions import ExcelModelError

        out_path = tmp_path / "out.xlsx"
        with patch(
            "excel_model.excel_writer.build_workbook",
            side_effect=ExcelModelError("mock build error"),
        ):
            result = runner.invoke(
                main,
                ["build", "--spec", str(spec_file), "--output", str(out_path), "--mode", "batch"],
            )
        assert result.exit_code != 0
        payload = json.loads(result.output.strip())
        assert payload["status"] == "error"
        assert "mock build error" in payload["message"]

    def test_batch_value_error(self, runner, spec_file, tmp_path):
        out_path = tmp_path / "out.xlsx"
        with patch(
            "excel_model.excel_writer.build_workbook",
            side_effect=ValueError("value err"),
        ):
            result = runner.invoke(
                main,
                ["build", "--spec", str(spec_file), "--output", str(out_path), "--mode", "batch"],
            )
        assert result.exit_code != 0
        payload = json.loads(result.output.strip())
        assert payload["status"] == "error"
        assert "value err" in payload["message"]

    def test_interactive_workbook_error(self, runner, spec_file, tmp_path):
        out_path = tmp_path / "out.xlsx"
        with patch(
            "excel_model.excel_writer.build_workbook",
            side_effect=ValueError("build failed"),
        ):
            result = runner.invoke(
                main,
                ["build", "--spec", str(spec_file), "--output", str(out_path), "--mode", "interactive"],
            )
        assert result.exit_code != 0


class TestBuildWithData:
    def test_batch_with_data_file_not_found(self, runner, spec_file, tmp_path):
        """--data pointing to nonexistent file is caught by click exists=True."""
        out_path = tmp_path / "out.xlsx"
        result = runner.invoke(
            main,
            [
                "build",
                "--spec",
                str(spec_file),
                "--output",
                str(out_path),
                "--data",
                str(tmp_path / "no_such_file.csv"),
                "--mode",
                "batch",
            ],
        )
        assert result.exit_code != 0

    def test_batch_data_load_error(self, runner, spec_file, tmp_path):
        """Data loading fails with ValueError."""
        dummy_data = tmp_path / "data.csv"
        dummy_data.write_text("a,b\n1,2\n")
        out_path = tmp_path / "out.xlsx"
        with patch(
            "excel_model.loader.load",
            side_effect=ValueError("bad data"),
        ):
            result = runner.invoke(
                main,
                [
                    "build",
                    "--spec",
                    str(spec_file),
                    "--output",
                    str(out_path),
                    "--data",
                    str(dummy_data),
                    "--mode",
                    "batch",
                ],
            )
        assert result.exit_code != 0
        payload = json.loads(result.output.strip())
        assert payload["status"] == "error"
        assert "Failed to load input data" in payload["message"]

    def test_interactive_build_with_valid_data(self, runner, spec_file, tmp_path):
        """Lines 101-107: interactive build with valid --data file."""
        import polars as pl

        from excel_model.loader import InputData

        dummy_data = tmp_path / "data.csv"
        dummy_data.write_text("period,revenue\n2023,100\n2024,200\n")
        out_path = tmp_path / "out.xlsx"
        mock_inputs = InputData(
            df=pl.DataFrame({"period": ["2023", "2024"], "revenue": [100, 200]}),
            period_col="period",
            value_cols=["revenue"],
        )
        with (
            patch("excel_model.loader.load", return_value=mock_inputs),
            patch("excel_model.validator.validate_inputs_against_spec", return_value=[]),
        ):
            result = runner.invoke(
                main,
                [
                    "build",
                    "--spec",
                    str(spec_file),
                    "--output",
                    str(out_path),
                    "--data",
                    str(dummy_data),
                    "--mode",
                    "interactive",
                ],
            )
        assert result.exit_code == 0, result.output
        assert "Loaded 2 rows" in result.output
        assert "Workbook saved to" in result.output

    def test_interactive_build_data_validation_error(self, runner, spec_file, tmp_path):
        """Line 107: interactive build where input data validation fails."""
        import polars as pl

        from excel_model.loader import InputData

        dummy_data = tmp_path / "data.csv"
        dummy_data.write_text("period,revenue\n2023,100\n")
        out_path = tmp_path / "out.xlsx"
        mock_inputs = InputData(
            df=pl.DataFrame({"period": ["2023"], "revenue": [100]}),
            period_col="period",
            value_cols=["revenue"],
        )
        with (
            patch("excel_model.loader.load", return_value=mock_inputs),
            patch(
                "excel_model.validator.validate_inputs_against_spec",
                return_value=["missing column: cost"],
            ),
        ):
            result = runner.invoke(
                main,
                [
                    "build",
                    "--spec",
                    str(spec_file),
                    "--output",
                    str(out_path),
                    "--data",
                    str(dummy_data),
                    "--mode",
                    "interactive",
                ],
            )
        assert result.exit_code != 0
        assert "Input data validation failed" in result.output


# ---------------------------------------------------------------------------
# validate command
# ---------------------------------------------------------------------------


class TestValidateSuccess:
    def test_valid_spec_prints_ok(self, runner, spec_file):
        result = runner.invoke(main, ["validate", "--spec", str(spec_file)])
        assert result.exit_code == 0
        assert "OK" in result.output


class TestValidateFailure:
    def test_invalid_spec_exits_1(self, runner, tmp_path):
        spec_f = tmp_path / "spec.yaml"
        spec_f.write_text(INVALID_SPEC_YAML)
        result = runner.invoke(main, ["validate", "--spec", str(spec_f)])
        assert result.exit_code != 0

    def test_spec_load_error_exits_1(self, runner, tmp_path):
        dummy = tmp_path / "dummy.yaml"
        dummy.write_text("placeholder")
        with patch(
            "excel_model.spec_loader.load_spec",
            side_effect=ValueError("parse fail"),
        ):
            result = runner.invoke(main, ["validate", "--spec", str(dummy)])
        assert result.exit_code != 0
        assert "ERROR" in result.output


class TestValidateWithData:
    def test_data_not_found(self, runner, spec_file, tmp_path):
        result = runner.invoke(
            main,
            ["validate", "--spec", str(spec_file), "--data", str(tmp_path / "missing.csv")],
        )
        assert result.exit_code != 0

    def test_data_load_error_adds_to_errors(self, runner, spec_file, tmp_path):
        dummy_data = tmp_path / "data.csv"
        dummy_data.write_text("a,b\n1,2\n")
        with patch(
            "excel_model.loader.load",
            side_effect=ValueError("bad data file"),
        ):
            result = runner.invoke(
                main,
                ["validate", "--spec", str(spec_file), "--data", str(dummy_data)],
            )
        assert result.exit_code != 0
        assert "bad data file" in result.output

    def test_validate_with_valid_data(self, runner, spec_file, tmp_path):
        """Lines 162-165: validate with both --spec and a valid --data file."""
        import polars as pl

        from excel_model.loader import InputData

        dummy_data = tmp_path / "data.csv"
        dummy_data.write_text("period,revenue\n2023,100\n")
        mock_inputs = InputData(
            df=pl.DataFrame({"period": ["2023"], "revenue": [100]}),
            period_col="period",
            value_cols=["revenue"],
        )
        with (
            patch("excel_model.loader.load", return_value=mock_inputs),
            patch("excel_model.validator.validate_inputs_against_spec", return_value=[]),
        ):
            result = runner.invoke(
                main,
                ["validate", "--spec", str(spec_file), "--data", str(dummy_data)],
            )
        assert result.exit_code == 0
        assert "OK" in result.output


# ---------------------------------------------------------------------------
# describe command
# ---------------------------------------------------------------------------


class TestDescribeJson:
    def test_describe_json_output(self, runner, spec_file):
        result = runner.invoke(main, ["describe", "--spec", str(spec_file), "--format", "json"])
        assert result.exit_code == 0, result.output
        payload = json.loads(result.output)
        assert payload["model_type"] == "p_and_l"
        assert payload["title"] == "Test P&L"
        assert payload["currency"] == "CHF"
        assert payload["n_periods"] == 3
        assert isinstance(payload["period_labels"], list)
        assert isinstance(payload["sections"], dict)
        assert isinstance(payload["assumptions_count"], int)


class TestDescribeText:
    def test_describe_text_output(self, runner, spec_file):
        result = runner.invoke(main, ["describe", "--spec", str(spec_file), "--format", "text"])
        assert result.exit_code == 0, result.output
        assert "Model: Test P&L" in result.output
        assert "Type:  p_and_l" in result.output
        assert "Currency: CHF" in result.output
        assert "Projection" in result.output
        assert "Validation: OK" in result.output

    def test_describe_text_with_history(self, runner, spec_file):
        result = runner.invoke(main, ["describe", "--spec", str(spec_file), "--format", "text"])
        assert result.exit_code == 0
        assert "History (2)" in result.output

    def test_describe_text_no_history(self, runner, tmp_path):
        spec_f = tmp_path / "spec.yaml"
        spec_f.write_text(MINIMAL_YAML)
        result = runner.invoke(main, ["describe", "--spec", str(spec_f), "--format", "text"])
        assert result.exit_code == 0
        # n_history_periods=0 means the History line is not printed
        assert "History (" not in result.output


class TestDescribeWithScenarios:
    def test_describe_text_shows_scenarios(self, runner, tmp_path):
        spec_f = tmp_path / "spec.yaml"
        spec_f.write_text(SCENARIO_YAML)
        result = runner.invoke(main, ["describe", "--spec", str(spec_f), "--format", "text"])
        assert result.exit_code == 0, result.output
        assert "Scenarios" in result.output
        assert "Base Case" in result.output
        assert "Bull Case" in result.output

    def test_describe_json_shows_scenarios(self, runner, tmp_path):
        spec_f = tmp_path / "spec.yaml"
        spec_f.write_text(SCENARIO_YAML)
        result = runner.invoke(main, ["describe", "--spec", str(spec_f), "--format", "json"])
        assert result.exit_code == 0, result.output
        payload = json.loads(result.output)
        assert len(payload["scenarios"]) == 2


class TestDescribeFailure:
    def test_describe_spec_load_error(self, runner, tmp_path):
        dummy = tmp_path / "dummy.yaml"
        dummy.write_text("placeholder")
        with patch(
            "excel_model.spec_loader.load_spec",
            side_effect=ValueError("bad spec"),
        ):
            result = runner.invoke(main, ["describe", "--spec", str(dummy), "--format", "json"])
        assert result.exit_code != 0


class TestDescribeGeneratePeriodsError:
    def test_describe_bad_start_period(self, runner, tmp_path):
        """describe with invalid start_period propagates ValueError from generate_periods."""
        bad_period_yaml = """\
model_type: p_and_l
title: Bad Period
currency: CHF
granularity: annual
start_period: "not-a-date"
n_periods: 3
n_history_periods: 0
"""
        spec_f = tmp_path / "spec.yaml"
        spec_f.write_text(bad_period_yaml)
        result = runner.invoke(main, ["describe", "--spec", str(spec_f), "--format", "json"])
        assert result.exit_code != 0


class TestDescribeValidationErrors:
    def test_describe_text_shows_validation_errors(self, runner, tmp_path):
        spec_f = tmp_path / "spec.yaml"
        spec_f.write_text(INVALID_SPEC_YAML)
        result = runner.invoke(main, ["describe", "--spec", str(spec_f), "--format", "text"])
        assert result.exit_code == 0
        assert "Validation errors" in result.output


# ---------------------------------------------------------------------------
# _build_description helper
# ---------------------------------------------------------------------------


class TestBuildDescription:
    def _load_and_build(self, tmp_path, yaml_text, errors=None):
        from excel_model.spec_loader import load_spec
        from excel_model.time_engine import generate_periods

        spec_f = tmp_path / "spec.yaml"
        spec_f.write_text(yaml_text)
        spec = load_spec(str(spec_f))
        periods = generate_periods(
            start_period=spec.start_period,
            n_periods=spec.n_periods,
            n_history=spec.n_history_periods,
            granularity=spec.granularity,
        )
        return _build_description(spec, periods, errors if errors is not None else [])

    def test_returns_expected_keys(self, tmp_path):
        desc = self._load_and_build(tmp_path, VALID_P_AND_L_YAML)
        expected_keys = {
            "model_type",
            "title",
            "currency",
            "granularity",
            "start_period",
            "n_history_periods",
            "n_periods",
            "total_periods",
            "period_labels",
            "history_labels",
            "projection_labels",
            "metadata",
            "assumptions_count",
            "assumption_groups",
            "line_items_count",
            "sections",
            "scenarios",
            "column_groups",
            "sheets_to_create",
            "validation_errors",
            "inputs",
        }
        assert set(desc.keys()) == expected_keys

    def test_groups_assumptions_by_group(self, tmp_path):
        desc = self._load_and_build(tmp_path, VALID_P_AND_L_YAML)
        assert "Growth" in desc["assumption_groups"]
        assert desc["assumption_groups"]["Growth"][0]["name"] == "RevenueGrowthRate"

    def test_groups_line_items_by_section(self, tmp_path):
        desc = self._load_and_build(tmp_path, VALID_P_AND_L_YAML)
        assert "Revenue" in desc["sections"]
        items = desc["sections"]["Revenue"]
        assert items[0]["key"] == "revenue"
        assert items[0]["is_subtotal"] is False

    def test_splits_history_and_projection_periods(self, tmp_path):
        desc = self._load_and_build(tmp_path, VALID_P_AND_L_YAML)
        assert len(desc["history_labels"]) == 2
        assert len(desc["projection_labels"]) == 3
        assert desc["total_periods"] == 5

    def test_validation_errors_passed_through(self, tmp_path):
        desc = self._load_and_build(tmp_path, MINIMAL_YAML, errors=["err1", "err2"])
        assert desc["validation_errors"] == ["err1", "err2"]

    def test_scenarios_included(self, tmp_path):
        desc = self._load_and_build(tmp_path, SCENARIO_YAML)
        assert len(desc["scenarios"]) == 2
        assert desc["scenarios"][1]["assumption_overrides"] == {"GrowthRate": 0.20}


# ---------------------------------------------------------------------------
# _render_description_text helper
# ---------------------------------------------------------------------------


class TestRenderDescriptionText:
    def _minimal_description(self, **overrides):
        base = {
            "title": "Test",
            "model_type": "p_and_l",
            "currency": "USD",
            "granularity": "annual",
            "n_history_periods": 0,
            "n_periods": 1,
            "history_labels": [],
            "projection_labels": ["2025"],
            "assumptions_count": 0,
            "assumption_groups": {},
            "line_items_count": 0,
            "sections": {},
            "scenarios": [],
            "validation_errors": [],
        }
        base.update(overrides)
        return base

    def test_basic_rendering(self, tmp_path):
        from excel_model.spec_loader import load_spec
        from excel_model.time_engine import generate_periods

        spec_f = tmp_path / "spec.yaml"
        spec_f.write_text(VALID_P_AND_L_YAML)
        spec = load_spec(str(spec_f))
        periods = generate_periods(
            start_period=spec.start_period,
            n_periods=spec.n_periods,
            n_history=spec.n_history_periods,
            granularity=spec.granularity,
        )
        desc = _build_description(spec, periods, [])
        text = _render_description_text(desc)
        assert "Model: Test P&L" in text
        assert "Type:  p_and_l" in text
        assert "Currency: CHF" in text
        assert "Validation: OK" in text
        assert "History (2)" in text
        assert "Projection (3)" in text

    def test_renders_validation_errors(self):
        desc = self._minimal_description(validation_errors=["err1", "err2"])
        text = _render_description_text(desc)
        assert "Validation errors (2)" in text
        assert "  - err1" in text
        assert "  - err2" in text
        assert "Validation: OK" not in text

    def test_renders_validation_ok(self):
        text = _render_description_text(self._minimal_description())
        assert "Validation: OK" in text
        assert "Validation errors" not in text

    def test_renders_scenarios(self):
        desc = self._minimal_description(
            scenarios=[
                {"name": "base", "label": "Base", "assumption_overrides": {}},
                {"name": "bull", "label": "Bull", "assumption_overrides": {"Growth": 0.2}},
            ],
        )
        text = _render_description_text(desc)
        assert "Scenarios (2)" in text
        assert "  Base" in text
        assert "Bull (overrides: Growth=0.2)" in text

    def test_no_history_skips_history_line(self):
        text = _render_description_text(self._minimal_description())
        assert "History" not in text

    def test_with_history_shows_history_line(self):
        desc = self._minimal_description(n_history_periods=2, history_labels=["2023", "2024"])
        text = _render_description_text(desc)
        assert "History (2): 2023, 2024" in text
