"""Load a YAML model spec file into a ModelSpec dataclass.

Uses strictyaml for validated parsing (raises StrictYAMLError with line/col on
schema violations) and dacite for dataclass construction.
"""
from pathlib import Path
from typing import Any

from strictyaml import load as syaml_load  # noqa: F401 (re-exported for callers)

from excel_model.spec import (
    AssumptionDef,
    ColumnGroupDef,
    DriverDef,
    EntityDef,
    InputsDef,
    LineItemDef,
    MetadataDef,
    ModelSpec,
    ScenarioDef,
)
from excel_model.spec_schema import SPEC_SCHEMA


def _build_assumption(raw: dict[str, Any]) -> AssumptionDef:
    return AssumptionDef(
        name=raw["name"],
        label=raw["label"],
        value=raw["value"],
        format=raw["format"],
        group=raw["group"],
    )


def _build_driver(raw: dict[str, Any]) -> DriverDef:
    return DriverDef(
        name=raw["name"],
        label=raw["label"],
        value=raw["value"],
        format=raw["format"],
        group=raw["group"],
    )


def _build_line_item(raw: dict[str, Any]) -> LineItemDef:
    return LineItemDef(
        key=raw["key"],
        label=raw["label"],
        formula_type=raw["formula_type"],
        formula_params=dict(raw["formula_params"]) if "formula_params" in raw else {},
        is_subtotal=bool(raw["is_subtotal"]),
        is_total=bool(raw["is_total"]),
        section=raw["section"],
        format=raw.get("format", ""),
    )


def _build_scenario(raw: dict[str, Any]) -> ScenarioDef:
    return ScenarioDef(
        name=raw["name"],
        label=raw["label"],
        assumption_overrides=dict(raw["assumption_overrides"]) if "assumption_overrides" in raw else {},
        driver_overrides=dict(raw["driver_overrides"]) if "driver_overrides" in raw else {},
    )


def _build_column_group(raw: dict[str, Any]) -> ColumnGroupDef:
    return ColumnGroupDef(
        key=raw["key"],
        label=raw["label"],
        color_hex=raw["color_hex"],
    )


def _build_entity(raw: dict[str, Any]) -> EntityDef:
    return EntityDef(
        key=raw["key"],
        label=raw["label"],
    )


def _build_inputs(raw: dict[str, Any] | None) -> InputsDef:
    if raw is None:
        return InputsDef(source="", period_col="period", sheet="", value_cols={})
    return InputsDef(
        source=raw["source"],
        period_col=raw["period_col"],
        sheet=raw["sheet"],
        value_cols=dict(raw["value_cols"]) if "value_cols" in raw else {},
    )


def _build_metadata(raw: dict[str, Any] | None) -> MetadataDef:
    if raw is None:
        return MetadataDef(preparer="", date="", version="1.0")
    return MetadataDef(
        preparer=raw["preparer"],
        date=raw["date"],
        version=raw["version"],
    )


def load_spec(path: str) -> ModelSpec:
    """Load a YAML model spec and return a ModelSpec.

    Raises StrictYAMLError with line/column info if the YAML is invalid or
    violates the schema. Raises FileNotFoundError if the file does not exist.
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Model spec file not found: {path}")

    text = p.read_text(encoding="utf-8")
    validated = syaml_load(text, SPEC_SCHEMA)
    data: dict[str, Any] = validated.data

    assumptions = tuple(
        _build_assumption(a) for a in (data.get("assumptions") or [])
    )
    drivers = tuple(
        _build_driver(d) for d in (data.get("drivers") or [])
    )
    line_items = tuple(
        _build_line_item(li) for li in (data.get("line_items") or [])
    )
    scenarios = tuple(
        _build_scenario(s) for s in (data.get("scenarios") or [])
    )
    column_groups = tuple(
        _build_column_group(cg) for cg in (data.get("column_groups") or [])
    )
    entities = tuple(
        _build_entity(e) for e in (data.get("entities") or [])
    )
    inputs = _build_inputs(data.get("inputs"))
    metadata = _build_metadata(data.get("metadata"))

    return ModelSpec(
        model_type=data["model_type"],
        title=data["title"],
        currency=data["currency"],
        granularity=data["granularity"],
        start_period=str(data["start_period"]),
        n_periods=int(data["n_periods"]),
        n_history_periods=int(data["n_history_periods"]),
        assumptions=assumptions,
        drivers=drivers,
        line_items=line_items,
        metadata=metadata,
        scenarios=scenarios,
        column_groups=column_groups,
        inputs=inputs,
        entities=entities,
    )
