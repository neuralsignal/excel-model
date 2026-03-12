"""Frozen dataclasses for model spec definitions."""
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class AssumptionDef:
    name: str           # Excel named range name (CamelCase, no spaces)
    label: str          # Human-readable label
    value: float | int | str
    format: str         # number | percent | currency | integer
    group: str          # Group name for sheet layout


@dataclass(frozen=True)
class LineItemDef:
    key: str            # Unique identifier
    label: str          # Display label (indent = leading spaces in label)
    formula_type: str   # See FormulaType enum
    formula_params: dict[str, Any]
    is_subtotal: bool
    is_total: bool
    section: str
    format: str         # "" = currency (default), "percent", "number", "integer"


@dataclass(frozen=True)
class EntityDef:
    key: str            # Unique identifier
    label: str          # Display column header


@dataclass(frozen=True)
class DriverDef:
    name: str           # Excel named range name (CamelCase, no spaces)
    label: str          # Human-readable label
    value: float | int | str
    format: str         # number | percent | currency | integer
    group: str          # Group name for sheet layout


@dataclass(frozen=True)
class ScenarioDef:
    name: str
    label: str
    assumption_overrides: dict[str, Any]
    driver_overrides: dict[str, Any]


@dataclass(frozen=True)
class ColumnGroupDef:
    key: str
    label: str
    color_hex: str


@dataclass(frozen=True)
class InputsDef:
    source: str
    period_col: str
    sheet: str
    value_cols: dict[str, str]


@dataclass(frozen=True)
class MetadataDef:
    preparer: str
    date: str
    version: str


@dataclass(frozen=True)
class ModelSpec:
    model_type: str
    title: str
    currency: str
    granularity: str
    start_period: str
    n_periods: int
    n_history_periods: int
    assumptions: tuple[AssumptionDef, ...]
    drivers: tuple[DriverDef, ...]
    line_items: tuple[LineItemDef, ...]
    metadata: MetadataDef
    scenarios: tuple[ScenarioDef, ...]
    column_groups: tuple[ColumnGroupDef, ...]
    inputs: InputsDef
    entities: tuple[EntityDef, ...]
