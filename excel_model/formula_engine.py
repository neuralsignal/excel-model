"""FormulaType enum and render_formula() — produces Excel formula strings."""

import re
from dataclasses import dataclass
from enum import Enum
from typing import Any


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
    driver_names: frozenset[str] = frozenset()  # when non-empty, only these names get scenario-prefixed


def _abs_col(col_letter: str, row: int) -> str:
    """Return absolute Excel reference like $C$5."""
    return f"${col_letter}${row}"


def _abs_inputs_ref(col_letter: str, row: int) -> str:
    """Return absolute Inputs sheet reference like Inputs!$C$5."""
    return f"Inputs!${col_letter}${row}"


def _entity_range_for_row(entity_col_range: str, row: int) -> str:
    """Rewrite entity_col_range (e.g., '$B$9:$D$9') to use a different row number."""
    # entity_col_range format: "$B$9:$D$9"
    left, right = entity_col_range.split(":")
    left_col = "".join(c for c in left if c.isalpha() or c == "$")
    right_col = "".join(c for c in right if c.isalpha() or c == "$")
    return f"{left_col}{row}:{right_col}{row}"


def _row_ref(key: str, col_letter: str, row_map: dict[str, int]) -> str:
    """Return absolute reference to a line item row in the current column."""
    if key not in row_map:
        raise KeyError(f"Line item key {key!r} not found in row_map. Available: {list(row_map)}")
    row = row_map[key]
    return f"${col_letter}${row}"


def _resolve_name(name: str, ctx: CellContext) -> str:
    """Build scenario-prefixed name when appropriate.

    - No scenario_prefix → bare name (non-scenario models).
    - driver_names empty → prefix ALL names (legacy scenario mode).
    - driver_names non-empty → prefix only names in driver_names; assumptions stay bare.
    """
    if not ctx.scenario_prefix:
        return name
    if not ctx.driver_names:
        # Legacy mode: prefix everything
        return f"{ctx.scenario_prefix}{name}"
    # New mode: only prefix driver names
    if name in ctx.driver_names:
        return f"{ctx.scenario_prefix}{name}"
    return name


def render_formula(
    formula_type: str,
    formula_params: dict[str, Any],
    ctx: CellContext,
) -> str | float | int:
    """Return an Excel formula string (starting with '=') or a literal value.

    Raises ValueError for unknown formula_type or KeyError for missing params.
    """
    ft = FormulaType(formula_type)  # raises ValueError if unknown

    is_history = ctx.period_index < ctx.n_history

    if ft == FormulaType.CONSTANT:
        return formula_params["value"]

    if ft == FormulaType.INPUT_REF:
        if is_history:
            # Reference the Inputs sheet for history periods.
            # If no inputs data is loaded (empty inputs_row_map), emit 0 as placeholder.
            key = formula_params["line_item_key"]
            if key and key in ctx.inputs_row_map:
                inputs_row = ctx.inputs_row_map[key]
                return f"=Inputs!${ctx.col_letter}${inputs_row}"
            # No input data available for this history period — emit 0 placeholder
            return 0
        else:
            # Route to projected_type
            projected_type = formula_params["projected_type"]
            projected_params = {k: v for k, v in formula_params.items() if k not in ("projected_type",)}
            return render_formula(projected_type, projected_params, ctx)

    if ft == FormulaType.GROWTH_PROJECTED:
        growth_name = formula_params["growth_assumption"]
        prior_key = formula_params.get("prior_key")  # genuinely optional
        rate_name = _resolve_name(growth_name, ctx)

        if ctx.period_index == 0 or ctx.period_index == ctx.n_history:
            # First period: can't use prior col (would be out of range or history)
            # Actually for projections, period_index == n_history is first projection
            # prior_col_letter points to the last history col
            if ctx.prior_col_letter:
                if prior_key and prior_key in ctx.row_map:
                    prior_row = ctx.row_map[prior_key]
                    return f"=${ctx.prior_col_letter}${prior_row}*(1+{rate_name})"
                return f"=${ctx.prior_col_letter}${ctx.row}*(1+{rate_name})"
            # No prior column — first ever period
            return f"=1*(1+{rate_name})"
        else:
            if prior_key and prior_key in ctx.row_map:
                prior_row = ctx.row_map[prior_key]
                return f"=${ctx.prior_col_letter}${prior_row}*(1+{rate_name})"
            return f"=${ctx.prior_col_letter}${ctx.row}*(1+{rate_name})"

    if ft == FormulaType.PCT_OF_REVENUE:
        revenue_key = formula_params["revenue_key"]
        rate_name = _resolve_name(formula_params["rate_assumption"], ctx)
        rev_row = ctx.row_map[revenue_key]
        return f"=${ctx.col_letter}${rev_row}*{rate_name}"

    if ft == FormulaType.SUM_OF_ROWS:
        addend_keys = formula_params["addend_keys"]
        refs = [f"${ctx.col_letter}${ctx.row_map[k]}" for k in addend_keys]
        return "=" + "+".join(refs)

    if ft == FormulaType.SUBTRACTION:
        minuend_key = formula_params["minuend_key"]
        subtrahend_key = formula_params["subtrahend_key"]
        min_row = ctx.row_map[minuend_key]
        sub_row = ctx.row_map[subtrahend_key]
        return f"=${ctx.col_letter}${min_row}-${ctx.col_letter}${sub_row}"

    if ft == FormulaType.SUM_SUBTRACTION:
        addend_key = formula_params["addend_key"]
        subtrahend_keys = formula_params["subtrahend_keys"]
        addend_ref = _row_ref(addend_key, ctx.col_letter, ctx.row_map)
        sub_refs = [_row_ref(k, ctx.col_letter, ctx.row_map) for k in subtrahend_keys]
        return f"={addend_ref}" + "".join(f"-{r}" for r in sub_refs)

    if ft == FormulaType.RATIO:
        num_key = formula_params["numerator_key"]
        den_key = formula_params["denominator_key"]
        num_row = ctx.row_map[num_key]
        den_row = ctx.row_map[den_key]
        return f"=${ctx.col_letter}${num_row}/${ctx.col_letter}${den_row}"

    if ft == FormulaType.GROWTH_RATE:
        value_key = formula_params["value_key"]
        val_row = ctx.row_map[value_key]
        if ctx.prior_col_letter:
            return f"=(${ctx.col_letter}${val_row}/${ctx.prior_col_letter}${val_row})-1"
        return "=0"

    if ft == FormulaType.DISCOUNTED_PV:
        cashflow_key = formula_params["cashflow_key"]
        rate_name = _resolve_name(formula_params["rate_assumption"], ctx)
        cf_row = ctx.row_map[cashflow_key]
        # Exponent = projection period number (1-based within projections)
        projection_index = ctx.period_index - ctx.n_history + 1
        return f"=${ctx.col_letter}${cf_row}/(1+{rate_name})^{projection_index}"

    if ft == FormulaType.TERMINAL_VALUE:
        cashflow_key = formula_params["cashflow_key"]
        growth_name = _resolve_name(formula_params["growth_assumption"], ctx)
        rate_name = _resolve_name(formula_params["rate_assumption"], ctx)
        cf_row = ctx.row_map[cashflow_key]
        # Gordon Growth Model: TV = FCF * (1 + TGR) / (WACC - TGR)
        return f"=${ctx.col_letter}${cf_row}*(1+{growth_name})/({rate_name}-{growth_name})"

    if ft == FormulaType.NPV_SUM:
        pv_fcf_key = formula_params["pv_fcf_key"]
        pv_terminal_key = formula_params["pv_terminal_key"]
        pv_fcf_row = ctx.row_map[pv_fcf_key]
        pv_terminal_row = ctx.row_map[pv_terminal_key]
        first = ctx.first_proj_col_letter
        last = ctx.last_proj_col_letter
        return f"=SUM({first}${pv_fcf_row}:{last}${pv_fcf_row})+{last}${pv_terminal_row}"

    if ft == FormulaType.VARIANCE:
        plan_key = formula_params["plan_key"]
        actual_key = formula_params["actual_key"]
        plan_row = ctx.row_map[plan_key]
        actual_row = ctx.row_map[actual_key]
        return f"=${ctx.col_letter}${actual_row}-${ctx.col_letter}${plan_row}"

    if ft == FormulaType.VARIANCE_PCT:
        plan_key = formula_params["plan_key"]
        actual_key = formula_params["actual_key"]
        plan_row = ctx.row_map[plan_key]
        actual_row = ctx.row_map[actual_key]
        return f"=(${ctx.col_letter}${actual_row}-${ctx.col_letter}${plan_row})/ABS(${ctx.col_letter}${plan_row})"

    if ft == FormulaType.CUSTOM:
        raw = formula_params["formula"]
        # Replace column letter tokens
        result = raw.replace("{col_letter}", ctx.col_letter)
        result = result.replace("{prev_col_letter}", ctx.prior_col_letter)
        # Replace row reference tokens
        for key, row in ctx.row_map.items():
            result = result.replace(f"{{{key}_row}}", str(row))
        # Apply scenario prefix to named ranges
        # Use regex word-boundary matching to avoid substring corruption
        if ctx.scenario_prefix:
            # Determine which names to prefix
            if ctx.driver_names:
                # New mode: only prefix driver names
                names_to_prefix = sorted(ctx.driver_names, key=len, reverse=True)
            else:
                # Legacy mode: prefix all named ranges
                names_to_prefix = sorted(ctx.named_ranges.keys(), key=len, reverse=True)
            for name in names_to_prefix:
                prefixed = f"{ctx.scenario_prefix}{name}"
                result = re.sub(rf"(?<![A-Za-z0-9_]){re.escape(name)}(?![A-Za-z0-9_])", prefixed, result)
        if not result.startswith("="):
            result = "=" + result
        return result

    if ft == FormulaType.RANK:
        value_key = formula_params["value_key"]
        val_row = ctx.row_map[value_key]
        cell_ref = f"${ctx.col_letter}${val_row}"
        value_range = _entity_range_for_row(ctx.entity_col_range, val_row)
        return f"=RANK({cell_ref},{value_range})"

    if ft == FormulaType.INDEX_TO_BASE:
        value_key = formula_params["value_key"]
        val_row = ctx.row_map[value_key]
        base_col = formula_params["_base_col_letter"]
        return f"=${ctx.col_letter}${val_row}/${base_col}${val_row}"

    if ft == FormulaType.BAR_CHART_TEXT:
        value_key = formula_params["value_key"]
        val_row = ctx.row_map[value_key]
        cell_ref = f"${ctx.col_letter}${val_row}"
        value_range = _entity_range_for_row(ctx.entity_col_range, val_row)
        return f'=REPT("█",{cell_ref}/MAX({value_range})*20)'

    raise ValueError(f"Unhandled formula type: {formula_type!r}")
