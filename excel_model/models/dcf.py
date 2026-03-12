"""DCF sheet builder — has its own model sheet builder for NPV_SUM aggregation."""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from excel_model.formula_engine import CellContext, render_formula
from excel_model.loader import InputData
from excel_model.models._sheet_builder import build_assumptions_sheet, build_inputs_sheet
from excel_model.named_ranges import get_col_letter
from excel_model.spec import ModelSpec
from excel_model.style import (
    StyleConfig,
    apply_header_style,
    apply_history_col_style,
    apply_normal_style,
    apply_section_header_style,
    apply_subtotal_style,
    apply_total_style,
    get_number_format,
)
from excel_model.time_engine import Period


def build_dcf(
    wb: Workbook,
    spec: ModelSpec,
    inputs: InputData | None,
    style: StyleConfig,
    periods: list[Period],
) -> None:
    """Build Assumptions, Inputs, and DCF Model sheets."""
    build_assumptions_sheet(wb, spec, style, "")
    inputs_row_map = build_inputs_sheet(wb, spec, inputs, style, periods)
    _build_dcf_model_sheet(wb, spec, style, periods, inputs_row_map)


def _build_dcf_model_sheet(
    wb: Workbook,
    spec: ModelSpec,
    style: StyleConfig,
    periods: list[Period],
    inputs_row_map: dict[str, int],
) -> None:
    """DCF model sheet — renders NPV_SUM only in the first data column,
    aggregating PV FCFs across ALL projection columns."""
    ws = wb.create_sheet(title="Model")

    n_period_cols = len(periods)
    total_cols = 1 + n_period_cols

    # Compute projection column range
    proj_periods = [p for p in periods if not p.is_history]
    if proj_periods:
        first_proj_col_letter = get_col_letter(2 + proj_periods[0].index)
        last_proj_col_letter = get_col_letter(2 + proj_periods[-1].index)
    else:
        first_proj_col_letter = ""
        last_proj_col_letter = ""

    # Row 1: Title header
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
    title_cell = ws["A1"]
    title_cell.value = spec.title
    apply_header_style(title_cell, style)
    ws.row_dimensions[1].height = 20

    # Row 2: Period labels
    label_header = ws.cell(row=2, column=1, value="Line Item")
    apply_header_style(label_header, style)

    for col_idx, period in enumerate(periods, start=2):
        cell = ws.cell(row=2, column=col_idx, value=period.label)
        if period.is_history:
            apply_history_col_style(cell, style)
            cell.font = Font(name=style.font_name, size=style.font_size, bold=True)
        else:
            apply_header_style(cell, style)

    ws.column_dimensions["A"].width = 28
    for col_idx in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 14

    ws.freeze_panes = "B3"

    # First pass: assign row numbers
    current_row = 3
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

    # Second pass: write data
    current_row = 3
    for section in sections_order:
        if section:
            ws.merge_cells(f"A{current_row}:{get_column_letter(total_cols)}{current_row}")
            sec_cell = ws[f"A{current_row}"]
            sec_cell.value = section
            apply_section_header_style(sec_cell, style)
            current_row += 1

        for li in sections_items[section]:
            assert row_map[li.key] == current_row, f"Row mismatch for {li.key}"

            label_cell = ws.cell(row=current_row, column=1, value=li.label)
            apply_normal_style(label_cell, style)
            if li.is_subtotal:
                apply_subtotal_style(label_cell, style)
            elif li.is_total:
                apply_total_style(label_cell, style)

            is_npv_sum = li.formula_type == "npv_sum"

            if is_npv_sum:
                # NPV_SUM: write formula ONLY in first data column, spanning all projection cols
                first_data_col = 2
                col_letter = get_col_letter(first_data_col)

                ctx = CellContext(
                    period_index=0,
                    n_history=spec.n_history_periods,
                    row=current_row,
                    col=first_data_col,
                    col_letter=col_letter,
                    prior_col_letter="",
                    named_ranges={a.name: a.name for a in spec.assumptions},
                    row_map=row_map,
                    inputs_row_map=inputs_row_map,
                    scenario_prefix="",
                    first_proj_col_letter=first_proj_col_letter,
                    last_proj_col_letter=last_proj_col_letter,
                    entity_col_range="",
                )

                value = render_formula(li.formula_type, dict(li.formula_params), ctx)
                cell = ws.cell(row=current_row, column=first_data_col, value=value)
                cell.number_format = get_number_format("currency", style)
                cell.alignment = Alignment(horizontal="right")
                if li.is_total:
                    apply_total_style(cell, style)
                elif li.is_subtotal:
                    apply_subtotal_style(cell, style)
                else:
                    apply_normal_style(cell, style)
            else:
                # Standard per-column rendering
                for col_idx, period in enumerate(periods, start=2):
                    col_letter = get_col_letter(col_idx)
                    prior_col_letter = get_col_letter(col_idx - 1) if col_idx > 2 else ""

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
                    cell.number_format = get_number_format("currency", style)
                    cell.alignment = Alignment(horizontal="right")

                    if period.is_history:
                        apply_history_col_style(cell, style)
                    if li.is_subtotal:
                        apply_subtotal_style(cell, style)
                    elif li.is_total:
                        apply_total_style(cell, style)
                    else:
                        apply_normal_style(cell, style)
                        if period.is_history:
                            apply_history_col_style(cell, style)

            # Thin vertical border after last history col
            if spec.n_history_periods > 0:
                border_col = 1 + spec.n_history_periods + 1
                if border_col <= total_cols:
                    bc = ws.cell(row=current_row, column=border_col)
                    bc.border = Border(left=Side(style="thin"))

            current_row += 1
