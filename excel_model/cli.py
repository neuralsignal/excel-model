"""CLI entry point for excel-model."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import TYPE_CHECKING, Any

import click

from excel_model.config import load_style
from excel_model.exceptions import ExcelModelError, StyleConfigError

if TYPE_CHECKING:
    from excel_model.spec import ModelSpec
    from excel_model.time_engine import Period


@click.group()
def main() -> None:
    """YAML-driven Excel financial model generator.

    Security: File path arguments (--spec, --data, --style, --output) are passed
    directly to the filesystem. Do not accept untrusted user input for these
    arguments without prior path validation and sanitization.
    """


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

    # Load spec
    emit_info(f"Loading model spec: {spec}")
    try:
        from excel_model.spec_loader import load_spec

        loaded_spec = load_spec(spec)
    except (FileNotFoundError, ValueError, KeyError) as e:
        emit_error(f"Failed to load spec: {e}")
        return  # unreachable, but keeps type checker happy

    # Validate spec
    emit_info("Validating model spec...")
    from excel_model.validator import validate_spec

    errors = validate_spec(loaded_spec)
    if errors:
        emit_error("Spec validation failed:\n" + "\n".join(f"  - {e}" for e in errors))
    emit_info(f"  Model type: {loaded_spec.model_type}")
    emit_info(f"  Title: {loaded_spec.title}")
    emit_info(f"  Currency: {loaded_spec.currency}")
    emit_info(
        f"  Periods: {loaded_spec.n_history_periods} history + {loaded_spec.n_periods} projection ({loaded_spec.granularity})"
    )
    emit_info(f"  Assumptions: {len(loaded_spec.assumptions)}")
    emit_info(f"  Line items: {len(loaded_spec.line_items)}")

    # Load style
    emit_info(f"Loading style config: {style or '(bundled defaults)'}")
    try:
        loaded_style = load_style(style)
    except StyleConfigError as e:
        emit_error(f"Failed to load style config: {e}")
        return

    # Load input data (optional)
    inputs = None
    if data:
        emit_info(f"Loading input data: {data}")
        try:
            from excel_model.loader import load

            value_cols = list(loaded_spec.inputs.value_cols.values())
            inputs = load(
                source_path=data,
                period_col=loaded_spec.inputs.period_col,
                value_cols=value_cols,
                sheet=loaded_spec.inputs.sheet,
            )
            emit_info(f"  Loaded {len(inputs.df)} rows")

            from excel_model.validator import validate_inputs_against_spec

            input_errors = validate_inputs_against_spec(loaded_spec, inputs)
            if input_errors:
                emit_error("Input data validation failed:\n" + "\n".join(f"  - {e}" for e in input_errors))
        except (FileNotFoundError, ValueError) as e:
            emit_error(f"Failed to load input data: {e}")

    # Build workbook
    emit_info("Building workbook...")
    try:
        from excel_model.excel_writer import build_workbook

        build_workbook(spec=loaded_spec, inputs=inputs, output_path=output, style=loaded_style)
    except ExcelModelError as e:
        emit_error(f"Failed to build workbook: {e}")
    except (ValueError, KeyError, FileNotFoundError) as e:
        emit_error(f"Failed to build workbook: {e}")

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
    # Load spec
    try:
        from excel_model.spec_loader import load_spec

        loaded_spec = load_spec(spec)
    except (FileNotFoundError, ValueError, KeyError) as e:
        click.echo(f"ERROR: {e}")
        sys.exit(1)

    # Validate spec
    from excel_model.validator import validate_spec

    errors = validate_spec(loaded_spec)

    # Optionally validate input data columns
    if data:
        try:
            from excel_model.loader import load

            value_cols = list(loaded_spec.inputs.value_cols.values())
            inputs = load(
                source_path=data,
                period_col=loaded_spec.inputs.period_col,
                value_cols=value_cols,
                sheet=loaded_spec.inputs.sheet,
            )
            from excel_model.validator import validate_inputs_against_spec

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


def _build_description(spec: ModelSpec, periods: list[Period], errors: list[str]) -> dict[str, Any]:
    assumption_groups: dict[str, list[dict[str, Any]]] = {}
    for a in spec.assumptions:
        assumption_groups.setdefault(a.group, []).append({"name": a.name, "value": a.value, "format": a.format})
    sections: dict[str, list[dict[str, Any]]] = {}
    for li in spec.line_items:
        sections.setdefault(li.section, []).append(
            {
                "key": li.key,
                "label": li.label,
                "formula_type": li.formula_type,
                "is_subtotal": li.is_subtotal,
                "is_total": li.is_total,
            }
        )
    return {
        "model_type": spec.model_type,
        "title": spec.title,
        "currency": spec.currency,
        "granularity": spec.granularity,
        "start_period": spec.start_period,
        "n_history_periods": spec.n_history_periods,
        "n_periods": spec.n_periods,
        "total_periods": len(periods),
        "period_labels": [p.label for p in periods],
        "history_labels": [p.label for p in periods if p.is_history],
        "projection_labels": [p.label for p in periods if not p.is_history],
        "metadata": {"preparer": spec.metadata.preparer, "date": spec.metadata.date, "version": spec.metadata.version},
        "assumptions_count": len(spec.assumptions),
        "assumption_groups": assumption_groups,
        "line_items_count": len(spec.line_items),
        "sections": sections,
        "scenarios": [
            {"name": s.name, "label": s.label, "assumption_overrides": dict(s.assumption_overrides)}
            for s in spec.scenarios
        ],
        "column_groups": [{"key": cg.key, "label": cg.label} for cg in spec.column_groups],
        "sheets_to_create": ["Assumptions", "Inputs", "Model"],
        "validation_errors": errors,
        "inputs": {
            "source": spec.inputs.source,
            "period_col": spec.inputs.period_col,
            "value_cols": dict(spec.inputs.value_cols),
        },
    }


def _render_description_text(description: dict[str, Any]) -> str:
    lines: list[str] = [
        f"Model: {description['title']}",
        f"Type:  {description['model_type']}",
        f"Currency: {description['currency']}",
        "",
        f"Periods ({description['granularity']}):",
    ]
    if description["n_history_periods"] > 0:
        lines.append(f"  History ({description['n_history_periods']}): {', '.join(description['history_labels'])}")
    lines.append(f"  Projection ({description['n_periods']}): {', '.join(description['projection_labels'])}")
    lines += ["", f"Assumptions ({description['assumptions_count']}):"]
    for group, assumptions in description["assumption_groups"].items():
        lines.append(f"  [{group}]")
        for a in assumptions:
            lines.append(f"    {a['name']}: {a['value']} ({a['format']})")
    lines += ["", f"Line Items ({description['line_items_count']}):"]
    for section, items in description["sections"].items():
        if section:
            lines.append(f"  [{section}]")
        for li in items:
            marker = " [subtotal]" if li["is_subtotal"] else (" [total]" if li["is_total"] else "")
            lines.append(f"    {li['label'].strip()}: {li['formula_type']}{marker}")
    if description["scenarios"]:
        lines += ["", f"Scenarios ({len(description['scenarios'])}):"]
        for s in description["scenarios"]:
            overrides = ", ".join(f"{k}={v}" for k, v in s["assumption_overrides"].items())
            override_str = f" (overrides: {overrides})" if overrides else ""
            lines.append(f"  {s['label']}{override_str}")
    if description["validation_errors"]:
        lines += ["", f"Validation errors ({len(description['validation_errors'])}):"]
        for e in description["validation_errors"]:
            lines.append(f"  - {e}")
    else:
        lines += ["", "Validation: OK"]
    return "\n".join(lines)


@main.command()
@click.option("--spec", required=True, type=click.Path(exists=True), help="Path to model spec YAML")
@click.option("--format", "output_format", required=True, type=click.Choice(["text", "json"]), help="Output format")
def describe(spec: str, output_format: str) -> None:
    """Dry-run description of what build would produce."""
    try:
        from excel_model.spec_loader import load_spec

        loaded_spec = load_spec(spec)
    except (FileNotFoundError, ValueError, KeyError) as e:
        click.echo(f"ERROR: Failed to load spec: {e}", err=True)
        sys.exit(1)

    from excel_model.validator import validate_spec

    errors = validate_spec(loaded_spec)
    from excel_model.time_engine import generate_periods

    periods = generate_periods(
        start_period=loaded_spec.start_period,
        n_periods=loaded_spec.n_periods,
        n_history=loaded_spec.n_history_periods,
        granularity=loaded_spec.granularity,
    )

    description = _build_description(loaded_spec, periods, errors)

    if output_format == "json":
        click.echo(json.dumps(description, indent=2))
    else:
        click.echo(_render_description_text(description))
