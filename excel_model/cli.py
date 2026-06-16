"""CLI entry point for excel-model."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import TYPE_CHECKING

import click

from excel_model.config import load_style
from excel_model.describe import build_description, render_description_text
from excel_model.excel_writer import build_workbook
from excel_model.exceptions import ExcelModelError, StyleConfigError
from excel_model.input_validator import validate_inputs_against_spec
from excel_model.loader import load
from excel_model.spec_loader import load_spec
from excel_model.time_engine import generate_periods
from excel_model.validator import validate_spec

if TYPE_CHECKING:
    from excel_model.config import StyleConfig
    from excel_model.loader import InputData
    from excel_model.spec import ModelSpec


@click.group()
def main() -> None:
    """YAML-driven Excel financial model generator.

    Security: File path arguments (--spec, --data, --style, --output) are passed
    directly to the filesystem. Do not accept untrusted user input for these
    arguments without prior path validation and sanitization.
    """


def _load_and_validate_spec(spec_path: str) -> ModelSpec:
    try:
        loaded_spec = load_spec(spec_path)
    except (FileNotFoundError, ValueError, KeyError) as e:
        raise ExcelModelError(f"Failed to load spec: {e}") from e
    errors = validate_spec(loaded_spec)
    if errors:
        raise ExcelModelError("Spec validation failed:\n" + "\n".join(f"  - {e}" for e in errors))
    return loaded_spec


def _load_style_config(style_path: str | None) -> StyleConfig:
    try:
        return load_style(style_path)
    except StyleConfigError as e:
        raise ExcelModelError(f"Failed to load style config: {e}") from e


def _load_input_data(spec: ModelSpec, data_path: str) -> InputData:
    try:
        value_cols = list(spec.inputs.value_cols.values())
        inputs = load(
            source_path=data_path,
            period_col=spec.inputs.period_col,
            value_cols=value_cols,
            sheet=spec.inputs.sheet,
        )
    except (FileNotFoundError, ValueError) as e:
        raise ExcelModelError(f"Failed to load input data: {e}") from e
    input_errors = validate_inputs_against_spec(spec, inputs)
    if input_errors:
        raise ExcelModelError("Input data validation failed:\n" + "\n".join(f"  - {e}" for e in input_errors))
    return inputs


@main.command()
@click.option("--spec", required=True, type=click.Path(exists=True), help="Path to model spec YAML")
@click.option("--output", required=True, type=click.Path(), help="Path for output .xlsx file")
@click.option(
    "--style",
    required=False,
    type=click.Path(exists=True),
    help="Path to style config YAML (uses bundled defaults if omitted)",
)
@click.option("--data", required=False, type=click.Path(exists=True), help="Path to input data file")
@click.option(
    "--mode",
    required=True,
    type=click.Choice(["batch", "interactive"]),
    help="batch = JSON to stdout; interactive = verbose narrative",
)
def build(spec: str, output: str, style: str | None, data: str | None, mode: str) -> None:
    """Build an Excel financial model from a YAML spec."""

    def emit_error(message: str) -> None:
        if mode == "batch":
            click.echo(json.dumps({"status": "error", "message": message}))
        else:
            click.echo(f"ERROR: {message}", err=True)
        sys.exit(1)

    def emit_info(message: str) -> None:
        if mode == "interactive":
            click.echo(message)

    try:
        emit_info(f"Loading model spec: {spec}")
        loaded_spec = _load_and_validate_spec(spec)
        emit_info("Validating model spec...")
        emit_info(f"  Model type: {loaded_spec.model_type}")
        emit_info(f"  Title: {loaded_spec.title}")
        emit_info(f"  Currency: {loaded_spec.currency}")
        emit_info(
            f"  Periods: {loaded_spec.n_history_periods} history + {loaded_spec.n_periods} projection ({loaded_spec.granularity})"
        )
        emit_info(f"  Assumptions: {len(loaded_spec.assumptions)}")
        emit_info(f"  Line items: {len(loaded_spec.line_items)}")

        emit_info(f"Loading style config: {style or '(bundled defaults)'}")
        loaded_style = _load_style_config(style)

        inputs = None
        if data:
            emit_info(f"Loading input data: {data}")
            inputs = _load_input_data(loaded_spec, data)
            emit_info(f"  Loaded {len(inputs.df)} rows")

        emit_info("Building workbook...")
        build_workbook(spec=loaded_spec, inputs=inputs, output_path=output, style=loaded_style)
    except ExcelModelError as e:
        emit_error(str(e))
        return  # pragma: no cover
    except (ValueError, KeyError, FileNotFoundError) as e:
        emit_error(f"Failed to build workbook: {e}")
        return  # pragma: no cover

    output_path = str(Path(output).resolve())
    emit_info(f"Workbook saved to: {output_path}")

    if mode == "batch":
        click.echo(json.dumps({"status": "ok", "output": output_path}))


@main.command()
@click.option("--spec", required=True, type=click.Path(exists=True), help="Path to model spec YAML")
@click.option(
    "--data", required=False, type=click.Path(exists=True), help="Optional input data file to validate column mapping"
)
def validate(spec: str, data: str | None) -> None:
    """Validate a model spec YAML file."""
    try:
        loaded_spec = load_spec(spec)
    except (FileNotFoundError, ValueError, KeyError) as e:
        click.echo(f"ERROR: {e}")
        sys.exit(1)

    errors = validate_spec(loaded_spec)

    if data:
        try:
            value_cols = list(loaded_spec.inputs.value_cols.values())
            inputs = load(
                source_path=data,
                period_col=loaded_spec.inputs.period_col,
                value_cols=value_cols,
                sheet=loaded_spec.inputs.sheet,
            )
            input_errors = validate_inputs_against_spec(loaded_spec, inputs)
            errors.extend(input_errors)
        except (FileNotFoundError, ValueError) as e:
            errors.append(f"Input data: {e}")

    if errors:
        for err in errors:
            click.echo(err)
        sys.exit(1)
    else:
        click.echo("OK")


@main.command()
@click.option("--spec", required=True, type=click.Path(exists=True), help="Path to model spec YAML")
@click.option("--format", "output_format", required=True, type=click.Choice(["text", "json"]), help="Output format")
def describe(spec: str, output_format: str) -> None:
    """Dry-run description of what build would produce."""
    try:
        loaded_spec = load_spec(spec)
    except (FileNotFoundError, ValueError, KeyError) as e:
        click.echo(f"ERROR: Failed to load spec: {e}", err=True)
        sys.exit(1)

    errors = validate_spec(loaded_spec)

    periods = generate_periods(
        start_period=loaded_spec.start_period,
        n_periods=loaded_spec.n_periods,
        n_history=loaded_spec.n_history_periods,
        granularity=loaded_spec.granularity,
    )

    description = build_description(loaded_spec, periods, errors)

    if output_format == "json":
        click.echo(json.dumps(description, indent=2))
    else:
        click.echo(render_description_text(description))
