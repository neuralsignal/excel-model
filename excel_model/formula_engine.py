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
    driver_names: frozenset[str]  # when non-empty, only these names get scenario-prefixed


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


def _render_constant(
    params: dict[str, Any],
    ctx: CellContext,
) -> str | float | int:
    return params["value"]


def _render_input_ref(
    params: dict[str, Any],
    ctx: CellContext,
) -> str | float | int:
    is_history = ctx.period_index < ctx.n_history
    if is_history:
        key = params["line_item_key"]
        if key and key in ctx.inputs_row_map:
            inputs_row = ctx.inputs_row_map[key]
            return f"=Inputs!${ctx.col_letter}${inputs_row}"
        return 0
    projected_type = params["projected_type"]
    projected_params = {k: v for k, v in params.items() if k not in ("projected_type",)}
    return render_formula(projected_type, projected_params, ctx)


def _render_growth_projected(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    growth_name = params["growth_assumption"]
    prior_key = params.get("prior_key")
    rate_name = _resolve_name(growth_name, ctx)

    if ctx.period_index == 0 or ctx.period_index == ctx.n_history:
        if ctx.prior_col_letter:
            if prior_key and prior_key in ctx.row_map:
                prior_row = ctx.row_map[prior_key]
                return f"=${ctx.prior_col_letter}${prior_row}*(1+{rate_name})"
            return f"=${ctx.prior_col_letter}${ctx.row}*(1+{rate_name})"
        return f"=1*(1+{rate_name})"
    if prior_key and prior_key in ctx.row_map:
        prior_row = ctx.row_map[prior_key]
        return f"=${ctx.prior_col_letter}${prior_row}*(1+{rate_name})"
    return f"=${ctx.prior_col_letter}${ctx.row}*(1+{rate_name})"


def _render_pct_of_revenue(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    revenue_key = params["revenue_key"]
    rate_name = _resolve_name(params["rate_assumption"], ctx)
    rev_row = ctx.row_map[revenue_key]
    return f"=${ctx.col_letter}${rev_row}*{rate_name}"


def _render_sum_of_rows(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    addend_keys = params["addend_keys"]
    refs = [f"${ctx.col_letter}${ctx.row_map[k]}" for k in addend_keys]
    return "=" + "+".join(refs)


def _render_subtraction(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    minuend_key = params["minuend_key"]
    subtrahend_key = params["subtrahend_key"]
    min_row = ctx.row_map[minuend_key]
    sub_row = ctx.row_map[subtrahend_key]
    return f"=${ctx.col_letter}${min_row}-${ctx.col_letter}${sub_row}"


def _render_sum_subtraction(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    addend_key = params["addend_key"]
    subtrahend_keys = params["subtrahend_keys"]
    addend_ref = _row_ref(addend_key, ctx.col_letter, ctx.row_map)
    sub_refs = [_row_ref(k, ctx.col_letter, ctx.row_map) for k in subtrahend_keys]
    return f"={addend_ref}" + "".join(f"-{r}" for r in sub_refs)


def _render_ratio(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    num_key = params["numerator_key"]
    den_key = params["denominator_key"]
    num_row = ctx.row_map[num_key]
    den_row = ctx.row_map[den_key]
    return f"=${ctx.col_letter}${num_row}/${ctx.col_letter}${den_row}"


def _render_growth_rate(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    value_key = params["value_key"]
    val_row = ctx.row_map[value_key]
    if ctx.prior_col_letter:
        return f"=(${ctx.col_letter}${val_row}/${ctx.prior_col_letter}${val_row})-1"
    return "=0"


def _render_discounted_pv(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    cashflow_key = params["cashflow_key"]
    rate_name = _resolve_name(params["rate_assumption"], ctx)
    cf_row = ctx.row_map[cashflow_key]
    projection_index = ctx.period_index - ctx.n_history + 1
    return f"=${ctx.col_letter}${cf_row}/(1+{rate_name})^{projection_index}"


def _render_terminal_value(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    cashflow_key = params["cashflow_key"]
    growth_name = _resolve_name(params["growth_assumption"], ctx)
    rate_name = _resolve_name(params["rate_assumption"], ctx)
    cf_row = ctx.row_map[cashflow_key]
    return f"=${ctx.col_letter}${cf_row}*(1+{growth_name})/({rate_name}-{growth_name})"


def _render_npv_sum(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    pv_fcf_key = params["pv_fcf_key"]
    pv_terminal_key = params["pv_terminal_key"]
    pv_fcf_row = ctx.row_map[pv_fcf_key]
    pv_terminal_row = ctx.row_map[pv_terminal_key]
    first = ctx.first_proj_col_letter
    last = ctx.last_proj_col_letter
    return f"=SUM({first}${pv_fcf_row}:{last}${pv_fcf_row})+{last}${pv_terminal_row}"


def _render_variance(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    plan_key = params["plan_key"]
    actual_key = params["actual_key"]
    plan_row = ctx.row_map[plan_key]
    actual_row = ctx.row_map[actual_key]
    return f"=${ctx.col_letter}${actual_row}-${ctx.col_letter}${plan_row}"


def _render_variance_pct(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    plan_key = params["plan_key"]
    actual_key = params["actual_key"]
    plan_row = ctx.row_map[plan_key]
    actual_row = ctx.row_map[actual_key]
    return f"=(${ctx.col_letter}${actual_row}-${ctx.col_letter}${plan_row})/ABS(${ctx.col_letter}${plan_row})"


def _render_custom(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    from excel_model.validator import validate_custom_formula

    raw = params["formula"]
    line_item_key = params.get("_line_item_key", "<unknown>")
    validate_custom_formula(raw, line_item_key)
    result = raw.replace("{col_letter}", ctx.col_letter)
    result = result.replace("{prev_col_letter}", ctx.prior_col_letter)
    for key, row in ctx.row_map.items():
        result = result.replace(f"{{{key}_row}}", str(row))
    if ctx.scenario_prefix:
        if ctx.driver_names:
            names_to_prefix = sorted(ctx.driver_names, key=len, reverse=True)
        else:
            names_to_prefix = sorted(ctx.named_ranges.keys(), key=len, reverse=True)
        for name in names_to_prefix:
            prefixed = f"{ctx.scenario_prefix}{name}"
            result = re.sub(rf"(?<![A-Za-z0-9_]){re.escape(name)}(?![A-Za-z0-9_])", prefixed, result)
    if not result.startswith("="):
        result = "=" + result
    return result


def _render_rank(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    value_key = params["value_key"]
    val_row = ctx.row_map[value_key]
    cell_ref = f"${ctx.col_letter}${val_row}"
    value_range = _entity_range_for_row(ctx.entity_col_range, val_row)
    return f"=RANK({cell_ref},{value_range})"


def _render_index_to_base(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    value_key = params["value_key"]
    val_row = ctx.row_map[value_key]
    base_col = params["_base_col_letter"]
    return f"=${ctx.col_letter}${val_row}/${base_col}${val_row}"


def _render_bar_chart_text(
    params: dict[str, Any],
    ctx: CellContext,
) -> str:
    value_key = params["value_key"]
    val_row = ctx.row_map[value_key]
    cell_ref = f"${ctx.col_letter}${val_row}"
    value_range = _entity_range_for_row(ctx.entity_col_range, val_row)
    return f'=REPT("█",{cell_ref}/MAX({value_range})*20)'


_FORMULA_DISPATCH: dict[FormulaType, Any] = {
    FormulaType.CONSTANT: _render_constant,
    FormulaType.INPUT_REF: _render_input_ref,
    FormulaType.GROWTH_PROJECTED: _render_growth_projected,
    FormulaType.PCT_OF_REVENUE: _render_pct_of_revenue,
    FormulaType.SUM_OF_ROWS: _render_sum_of_rows,
    FormulaType.SUBTRACTION: _render_subtraction,
    FormulaType.SUM_SUBTRACTION: _render_sum_subtraction,
    FormulaType.RATIO: _render_ratio,
    FormulaType.GROWTH_RATE: _render_growth_rate,
    FormulaType.DISCOUNTED_PV: _render_discounted_pv,
    FormulaType.TERMINAL_VALUE: _render_terminal_value,
    FormulaType.NPV_SUM: _render_npv_sum,
    FormulaType.VARIANCE: _render_variance,
    FormulaType.VARIANCE_PCT: _render_variance_pct,
    FormulaType.CUSTOM: _render_custom,
    FormulaType.RANK: _render_rank,
    FormulaType.INDEX_TO_BASE: _render_index_to_base,
    FormulaType.BAR_CHART_TEXT: _render_bar_chart_text,
}


def render_formula(
    formula_type: str,
    formula_params: dict[str, Any],
    ctx: CellContext,
) -> str | float | int:
    """Return an Excel formula string (starting with '=') or a literal value.

    Raises ValueError for unknown formula_type or KeyError for missing params.
    """
    ft = FormulaType(formula_type)  # raises ValueError if unknown
    handler = _FORMULA_DISPATCH[ft]
    return handler(formula_params, ctx)
