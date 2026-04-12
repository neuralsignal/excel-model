"""Frozen dataclasses for model spec definitions."""

from collections.abc import Mapping
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class AssumptionDef:
    name: str  # Excel named range name (CamelCase, no spaces)
    label: str  # Human-readable label
    value: float | int | str
    format: str  # number | percent | currency | integer
    group: str  # Group name for sheet layout


@dataclass(frozen=True)
class LineItemDef:
    key: str  # Unique identifier
    label: str  # Display label (indent = leading spaces in label)
    formula_type: str  # See FormulaType enum
    formula_params: dict[str, Any]
    is_subtotal: bool
    is_total: bool
    section: str
    format: str  # "" = currency (default), "percent", "number", "integer"


@dataclass(frozen=True)
class EntityDef:
    key: str  # Unique identifier
    label: str  # Display column header


@dataclass(frozen=True)
class DriverDef:
    name: str  # Excel named range name (CamelCase, no spaces)
    label: str  # Human-readable label
    value: float | int | str
    format: str  # number | percent | currency | integer
    group: str  # Group name for sheet layout


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
class DataSheetDef:
    """Configuration for a tabular data sheet."""

    sheet_name: str
    title: str
    headers: tuple[str, ...]
    col_widths: tuple[float, ...]
    number_formats: Mapping[int, str]
    freeze_row: int


@dataclass(frozen=True)
class SumifsPivotDef:
    """Configuration for a SUMIFS pivot sheet.

    The ``value_col``, ``row_filter_cols``, and ``col_filter_col`` fields are
    Excel column letters (e.g. ``"AO"``, ``"AM"``) referring to columns on the
    ``data_sheet`` source worksheet:

    - ``value_col``: column containing the numeric values to sum.
    - ``row_filter_cols``: columns matched against each row's label values;
      len(row_filter_cols) must be <= len(row_label_headers).
    - ``col_filter_col``: column matched against each ``col_dim_values`` header.
    """

    sheet_name: str
    title: str
    row_label_headers: tuple[str, ...]
    col_dim_values: tuple[str | int | float, ...]
    data_sheet: str
    value_col: str
    row_filter_cols: tuple[str, ...]
    col_filter_col: str
    append_total: bool
    append_yoy: bool
    col_widths: tuple[float, ...]
    number_format_data: str
    number_format_pct: str
    freeze_row: int


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
