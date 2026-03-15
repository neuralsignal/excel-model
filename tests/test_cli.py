"""Tests for excel_model.cli — build, validate, describe commands."""

import json
from unittest.mock import patch

import pytest
from click.testing import CliRunner

from excel_model.cli import main

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


class TestDescribeValidationErrors:
    def test_describe_text_shows_validation_errors(self, runner, tmp_path):
        spec_f = tmp_path / "spec.yaml"
        spec_f.write_text(INVALID_SPEC_YAML)
        result = runner.invoke(main, ["describe", "--spec", str(spec_f), "--format", "text"])
        assert result.exit_code == 0
        assert "Validation errors" in result.output
