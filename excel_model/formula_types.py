"""FormulaType enum and CellContext dataclass — shared types for formula rendering."""

from dataclasses import dataclass
from enum import Enum


class FormulaType(Enum):
    INPUT_REF = "input_ref"
    GROWTH_PROJECTED = "growth_projected"
    PCT_OF_REVENUE = "pct_of_revenue"
    SUM_OF_ROWS = "sum_of_rows"
    SUBTRACTION = "subtraction"
    SUM_SUBTRACTION = "sum_subtraction"
    RATIO = "ratio"
    GROWTH_RATE = "growth_rate"
    DISCOUNTED_PV = "discounted_pv"
    TERMINAL_VALUE = "terminal_value"
    NPV_SUM = "npv_sum"
    VARIANCE = "variance"
    VARIANCE_PCT = "variance_pct"
    CONSTANT = "constant"
    CUSTOM = "custom"
    RANK = "rank"
    INDEX_TO_BASE = "index_to_base"
    BAR_CHART_TEXT = "bar_chart_text"


@dataclass(frozen=True)
class CellContext:
    """Everything formula_engine needs to know about the current cell position."""

    period_index: int  # 0-based index among all periods (history + projection)
    n_history: int  # number of history periods
    row: int  # 1-based Excel row of this cell
    col: int  # 1-based Excel column of this cell
    col_letter: str  # e.g., "C"
    prior_col_letter: str  # col_letter of period_index - 1 (empty string if first)
    named_ranges: dict[str, str]  # assumption_name → Excel named range name
    row_map: dict[str, int]  # line_item_key → Excel row number (Model sheet)
    inputs_row_map: dict[str, int]  # line_item_key → Inputs sheet row number
    scenario_prefix: str  # e.g., "Bull" for scenario models
    first_proj_col_letter: str  # e.g., "D" — first projection column
    last_proj_col_letter: str  # e.g., "H" — last projection column
    entity_col_range: str  # e.g., "$B$5:$H$5" — full row range for RANK/MAX formulas
    driver_names: frozenset[str]  # when non-empty, only these names get scenario-prefixed
