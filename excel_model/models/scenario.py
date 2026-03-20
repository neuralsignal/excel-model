"""Scenario Analysis sheet builder — side-by-side column groups."""

from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from excel_model.formula_engine import CellContext, render_formula
from excel_model.loader import InputData
from excel_model.models._sheet_builder import (
    build_assumptions_sheet,
    build_drivers_sheet,
    build_inputs_sheet,
    group_line_items_by_section,
)
from excel_model.named_ranges import get_col_letter, register_named_range
from excel_model.spec import ModelSpec, ScenarioDef
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


def build_scenario(
    wb: Workbook,
    spec: ModelSpec,
    inputs: InputData | None,
    style: StyleConfig,
    periods: list[Period],
) -> None:
    """Build Assumptions, (optionally Drivers), Inputs, and Scenario Model sheets."""
    if spec.drivers:
        # New mode: separate Assumptions (bare names) and Drivers (prefixed per scenario)
        build_assumptions_sheet(wb, spec, style, "")
        build_drivers_sheet(wb, spec, style, spec.scenarios)
    else:
        # Legacy mode: all assumptions prefixed per scenario
        _build_scenario_assumptions(wb, spec, style)

    inputs_row_map = build_inputs_sheet(wb, spec, inputs, style, periods)
    _build_scenario_model_sheet(wb, spec, style, periods, inputs_row_map)


def _scenario_prefix(scenario: ScenarioDef) -> str:
    """Convert scenario name to CamelCase prefix for named ranges."""
    return scenario.name.capitalize()


def _build_scenario_assumptions(
    wb: Workbook,
    spec: ModelSpec,
    style: StyleConfig,
) -> None:
    """Build Assumptions sheet with one section per scenario."""
    ws = wb.create_sheet(title="Assumptions")

    ws.merge_cells("A1:D1")
    title_cell = ws["A1"]
    title_cell.value = f"{spec.title} — Scenario Assumptions"
    apply_header_style(title_cell, style)
    ws.row_dimensions[1].height = 20

    headers = ["Parameter", "Named Range", "Value", "Format"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        apply_header_style(cell, style)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 12

    current_row = 3

    for scenario in spec.scenarios:
        prefix = _scenario_prefix(scenario)

        # Scenario section header
        ws.merge_cells(f"A{current_row}:D{current_row}")
        sec_cell = ws[f"A{current_row}"]
        sec_cell.value = f"{scenario.label} Assumptions"
        apply_section_header_style(sec_cell, style)
        current_row += 1

        for assumption in spec.assumptions:
            range_name = f"{prefix}{assumption.name}"
            # Override value if specified for this scenario
            value = scenario.assumption_overrides.get(assumption.name, assumption.value)

            ws.cell(row=current_row, column=1, value=assumption.label)
            ws.cell(row=current_row, column=2, value=range_name)

            value_cell = ws.cell(row=current_row, column=3, value=value)
            value_cell.number_format = get_number_format(assumption.format, style)
            value_cell.alignment = Alignment(horizontal="right")

            ws.cell(row=current_row, column=4, value=assumption.format)

            register_named_range(wb, range_name, "Assumptions", current_row, 3)
            current_row += 1


def _build_scenario_model_sheet(
    wb: Workbook,
    spec: ModelSpec,
    style: StyleConfig,
    periods: list[Period],
    inputs_row_map: dict[str, int],
) -> None:
    """Scenario model: Base | Bull | Bear column groups per period."""
    ws = wb.create_sheet(title="Model")

    n_scenarios = len(spec.scenarios)
    n_period_groups = len(periods)
    n_sub_cols = n_scenarios
    total_cols = 1 + n_period_groups * n_sub_cols

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

    # Row 2: Period group headers
    label_header = ws.cell(row=2, column=1, value="Line Item")
    apply_header_style(label_header, style)
    for p_idx, period in enumerate(periods):
        base_col = 2 + p_idx * n_sub_cols
        end_col = base_col + n_sub_cols - 1
        ws.merge_cells(f"{get_column_letter(base_col)}2:{get_column_letter(end_col)}2")
        ph = ws.cell(row=2, column=base_col, value=period.label)
        apply_header_style(ph, style)

    # Row 3: Scenario labels
    ws.cell(row=3, column=1, value="")
    apply_header_style(ws.cell(row=3, column=1), style)
    for p_idx in range(len(periods)):
        base_col = 2 + p_idx * n_sub_cols
        for s_idx, scenario in enumerate(spec.scenarios):
            cell = ws.cell(row=3, column=base_col + s_idx, value=scenario.label)
            apply_header_style(cell, style)

    ws.column_dimensions["A"].width = 28
    for col_idx in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 13

    ws.freeze_panes = "B4"

    # Assign rows
    current_row = 4
    row_map: dict[str, int] = {}
    sections_order, sections_items = group_line_items_by_section(spec.line_items)

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

                for s_idx, scenario in enumerate(spec.scenarios):
                    col_idx = base_col + s_idx
                    col_letter = get_col_letter(col_idx)
                    prior_col_letter = get_col_letter(col_idx - n_sub_cols) if p_idx > 0 else ""
                    prefix = _scenario_prefix(scenario)

                    params = dict(li.formula_params)
                    if li.formula_type == "input_ref":
                        params["line_item_key"] = li.key

                    # Build named_ranges: assumptions + drivers
                    all_named = {a.name: a.name for a in spec.assumptions}
                    for d in spec.drivers:
                        all_named[d.name] = d.name

                    ctx = CellContext(
                        period_index=period.index,
                        n_history=spec.n_history_periods,
                        row=current_row,
                        col=col_idx,
                        col_letter=col_letter,
                        prior_col_letter=prior_col_letter,
                        named_ranges=all_named,
                        row_map=row_map,
                        inputs_row_map=inputs_row_map,
                        scenario_prefix=prefix,
                        first_proj_col_letter=first_proj_col_letter,
                        last_proj_col_letter=last_proj_col_letter,
                        entity_col_range="",
                        driver_names=frozenset(d.name for d in spec.drivers),
                    )

                    value = render_formula(li.formula_type, params, ctx)
                    cell = ws.cell(row=current_row, column=col_idx, value=value)
                    fmt = li.format if li.format else "currency"
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
