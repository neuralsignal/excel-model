"""Budget vs Actuals sheet builder."""

from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from excel_model.formula_engine import CellContext, render_formula
from excel_model.loader import InputData
from excel_model.models._auxiliary_sheets import build_assumptions_sheet, build_inputs_sheet
from excel_model.models._sheet_builder import (
    HeaderLayout,
    apply_label_style,
    assign_row_map,
    build_model_header,
    compute_proj_col_range,
    group_line_items_by_section,
    write_section_header,
)
from excel_model.named_ranges import get_col_letter
from excel_model.spec import ModelSpec
from excel_model.style import (
    StyleConfig,
    apply_conditional_formatting,
    apply_header_style,
    get_number_format,
)
from excel_model.time_engine import Period


def build_budget_vs_actuals(
    wb: Workbook,
    spec: ModelSpec,
    inputs: InputData | None,
    style: StyleConfig,
    periods: list[Period],
) -> None:
    """Build Assumptions, Inputs, and Budget vs Actuals Model sheets."""
    build_assumptions_sheet(wb, spec, style, "")
    inputs_row_map = build_inputs_sheet(wb, spec, inputs, style, periods)
    _build_bva_model_sheet(wb, spec, style, periods, inputs_row_map)


def _write_bva_headers(
    ws: object,
    periods: list[Period],
    groups: tuple[object, ...],
    n_sub_cols: int,
    total_cols: int,
    style: StyleConfig,
) -> None:
    """Write period group headers (row 2) and sub-column labels (row 3)."""
    label_header = ws.cell(row=2, column=1, value="Line Item")  # type: ignore[union-attr]
    apply_header_style(label_header, style)
    for p_idx, period in enumerate(periods):
        base_col = 2 + p_idx * n_sub_cols
        end_col = base_col + n_sub_cols - 1
        ws.merge_cells(f"{get_column_letter(base_col)}2:{get_column_letter(end_col)}2")  # type: ignore[union-attr]
        ph = ws.cell(row=2, column=base_col, value=period.label)  # type: ignore[union-attr]
        apply_header_style(ph, style)

    sub_label_cell = ws.cell(row=3, column=1, value="")  # type: ignore[union-attr]
    apply_header_style(sub_label_cell, style)
    for p_idx, _period in enumerate(periods):
        base_col = 2 + p_idx * n_sub_cols
        for g_idx, group in enumerate(groups):
            cell = ws.cell(row=3, column=base_col + g_idx, value=group.label)  # type: ignore[union-attr]
            apply_header_style(cell, style)


def _build_bva_model_sheet(
    wb: Workbook,
    spec: ModelSpec,
    style: StyleConfig,
    periods: list[Period],
    inputs_row_map: dict[str, int],
) -> None:
    """BvA: Plan | Actual | Variance | Variance% columns per period."""
    ws = wb.create_sheet(title="Model")

    groups = spec.column_groups
    n_sub_cols = len(groups)
    total_cols = 1 + len(periods) * n_sub_cols

    first_proj_col_letter, last_proj_col_letter = compute_proj_col_range(periods, n_sub_cols, 2)

    build_model_header(ws, spec.title, total_cols, style, HeaderLayout("Line Item", 12, "B4"))
    _write_bva_headers(ws, periods, groups, n_sub_cols, total_cols, style)

    sections_order, sections_items = group_line_items_by_section(spec.line_items)
    row_map = assign_row_map(sections_order, sections_items, 4)

    # Write data
    current_row = 4
    for section in sections_order:
        if section:
            write_section_header(ws, section, current_row, total_cols, style)
            current_row += 1

        for li in sections_items[section]:
            label_cell = ws.cell(row=current_row, column=1, value=li.label)
            apply_label_style(label_cell, li, style)

            for p_idx, period in enumerate(periods):
                base_col = 2 + p_idx * n_sub_cols

                for g_idx, _group in enumerate(groups):
                    col_idx = base_col + g_idx
                    col_letter = get_col_letter(col_idx)
                    prior_col_letter = get_col_letter(col_idx - n_sub_cols) if p_idx > 0 else ""

                    params = dict(li.formula_params)
                    if li.formula_type == "input_ref":
                        params["line_item_key"] = li.key

                    ctx = CellContext(
                        period_index=period.index,
                        n_history=spec.n_history_periods,
                        row=current_row,
                        col=col_idx,
                        col_letter=col_letter,
                        prior_col_letter=prior_col_letter,
                        named_ranges={a.name: a.name for a in spec.assumptions},
                        row_map=row_map,
                        inputs_row_map=inputs_row_map,
                        scenario_prefix="",
                        first_proj_col_letter=first_proj_col_letter,
                        last_proj_col_letter=last_proj_col_letter,
                        entity_col_range="",
                        driver_names=frozenset(),
                    )

                    value = render_formula(li.formula_type, params, ctx)
                    cell = ws.cell(row=current_row, column=col_idx, value=value)

                    fmt = "percent" if "pct" in li.formula_type.lower() or "margin" in li.key.lower() else "currency"
                    cell.number_format = get_number_format(fmt, style)
                    cell.alignment = Alignment(horizontal="right")
                    apply_label_style(cell, li, style)

            # Apply conditional formatting to variance rows when positive_is_good is specified
            if li.formula_type in ("variance", "variance_pct") and "positive_is_good" in li.formula_params:
                positive_is_good = bool(li.formula_params["positive_is_good"])
                data_start = get_column_letter(2)
                data_end = get_column_letter(total_cols)
                cf_range = f"{data_start}{current_row}:{data_end}{current_row}"
                apply_conditional_formatting(ws, cf_range, positive_is_good, style)

            current_row += 1
