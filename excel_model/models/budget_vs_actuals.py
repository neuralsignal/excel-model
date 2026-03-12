"""Budget vs Actuals sheet builder."""
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from excel_model.formula_engine import CellContext, render_formula
from excel_model.loader import InputData
from excel_model.models._sheet_builder import build_assumptions_sheet, build_inputs_sheet
from excel_model.named_ranges import get_col_letter
from excel_model.spec import ModelSpec
from excel_model.style import (
    StyleConfig,
    apply_conditional_formatting,
    apply_header_style,
    apply_normal_style,
    apply_section_header_style,
    apply_subtotal_style,
    apply_total_style,
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
    n_sub_cols = len(groups)  # typically 3: plan, actual, variance
    n_period_cols = len(periods) * n_sub_cols
    total_cols = 1 + n_period_cols

    # Compute projection column range
    proj_periods = [p for p in periods if not p.is_history]
    if proj_periods:
        first_proj_col_letter = get_col_letter(2 + proj_periods[0].index * n_sub_cols)
        last_proj_col_letter = get_col_letter(2 + proj_periods[-1].index * n_sub_cols + n_sub_cols - 1)
    else:
        first_proj_col_letter = ""
        last_proj_col_letter = ""

    # Row 1: Title
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
    title_cell = ws["A1"]
    title_cell.value = spec.title
    apply_header_style(title_cell, style)
    ws.row_dimensions[1].height = 20

    # Row 2: Period group headers (merged across sub-cols)
    label_header = ws.cell(row=2, column=1, value="Line Item")
    apply_header_style(label_header, style)

    for p_idx, period in enumerate(periods):
        base_col = 2 + p_idx * n_sub_cols
        end_col = base_col + n_sub_cols - 1
        ws.merge_cells(f"{get_column_letter(base_col)}2:{get_column_letter(end_col)}2")
        ph = ws.cell(row=2, column=base_col, value=period.label)
        apply_header_style(ph, style)

    # Row 3: Sub-column labels (Plan | Actual | Variance | ...)
    sub_label_cell = ws.cell(row=3, column=1, value="")
    apply_header_style(sub_label_cell, style)
    for p_idx, _period in enumerate(periods):
        base_col = 2 + p_idx * n_sub_cols
        for g_idx, group in enumerate(groups):
            cell = ws.cell(row=3, column=base_col + g_idx, value=group.label)
            apply_header_style(cell, style)

    ws.column_dimensions["A"].width = 28
    for col_idx in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 12

    ws.freeze_panes = "B4"

    # Assign row numbers
    current_row = 4
    row_map: dict[str, int] = {}
    sections_order: list[str] = []
    sections_items: dict[str, list] = {}
    for li in spec.line_items:
        if li.section not in sections_items:
            sections_order.append(li.section)
            sections_items[li.section] = []
        sections_items[li.section].append(li)

    for section in sections_order:
        if section:
            current_row += 1
        for li in sections_items[section]:
            row_map[li.key] = current_row
            current_row += 1

    # Write data
    current_row = 4
    for section in sections_order:
        if section:
            ws.merge_cells(f"A{current_row}:{get_column_letter(total_cols)}{current_row}")
            sec_cell = ws[f"A{current_row}"]
            sec_cell.value = section
            apply_section_header_style(sec_cell, style)
            current_row += 1

        for li in sections_items[section]:
            label_cell = ws.cell(row=current_row, column=1, value=li.label)
            apply_normal_style(label_cell, style)
            if li.is_subtotal:
                apply_subtotal_style(label_cell, style)
            elif li.is_total:
                apply_total_style(label_cell, style)

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
                    )

                    value = render_formula(li.formula_type, params, ctx)
                    cell = ws.cell(row=current_row, column=col_idx, value=value)

                    fmt = "percent" if "pct" in li.formula_type.lower() or "margin" in li.key.lower() else "currency"
                    cell.number_format = get_number_format(fmt, style)
                    cell.alignment = Alignment(horizontal="right")

                    apply_normal_style(cell, style)
                    if li.is_subtotal:
                        apply_subtotal_style(cell, style)
                    elif li.is_total:
                        apply_total_style(cell, style)

            # Apply conditional formatting to variance rows when positive_is_good is specified
            if li.formula_type in ("variance", "variance_pct") and "positive_is_good" in li.formula_params:
                positive_is_good = bool(li.formula_params["positive_is_good"])
                data_start = get_column_letter(2)
                data_end = get_column_letter(total_cols)
                cf_range = f"{data_start}{current_row}:{data_end}{current_row}"
                apply_conditional_formatting(ws, cf_range, positive_is_good, style)

            current_row += 1
