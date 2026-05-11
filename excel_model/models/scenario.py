"""Scenario Analysis sheet builder — side-by-side column groups."""

from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from excel_model.formula_engine import CellContext, render_formula
from excel_model.injection_guard import sanitize_cell_text
from excel_model.loader import InputData
from excel_model.models._auxiliary_sheets import (
    build_assumptions_sheet,
    build_drivers_sheet,
    build_inputs_sheet,
)
from excel_model.models._sheet_builder import (
    apply_label_style,
    assign_row_map,
    build_model_header,
    compute_proj_col_range,
    group_line_items_by_section,
    write_section_header,
    write_title_row,
)
from excel_model.named_ranges import get_col_letter, register_named_range
from excel_model.spec import ModelSpec, ScenarioDef
from excel_model.style import (
    StyleConfig,
    apply_conditional_formatting,
    apply_header_style,
    apply_section_header_style,
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

    write_title_row(ws, f"{spec.title} — Scenario Assumptions", 4, style)

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
        sec_cell.value = sanitize_cell_text(f"{scenario.label} Assumptions")
        apply_section_header_style(sec_cell, style)
        current_row += 1

        for assumption in spec.assumptions:
            range_name = f"{prefix}{assumption.name}"
            # Override value if specified for this scenario
            value = scenario.assumption_overrides.get(assumption.name, assumption.value)

            ws.cell(row=current_row, column=1, value=sanitize_cell_text(assumption.label))
            ws.cell(row=current_row, column=2, value=range_name)

            value_cell = ws.cell(row=current_row, column=3, value=value)
            value_cell.number_format = get_number_format(assumption.format, style)
            value_cell.alignment = Alignment(horizontal="right")

            ws.cell(row=current_row, column=4, value=assumption.format)

            register_named_range(wb, range_name, "Assumptions", current_row, 3)
            current_row += 1


def _write_scenario_headers(
    ws: object,
    periods: list[Period],
    spec: ModelSpec,
    n_sub_cols: int,
    total_cols: int,
    style: StyleConfig,
) -> None:
    """Write period group headers (row 2) and scenario labels (row 3)."""
    label_header = ws.cell(row=2, column=1, value="Line Item")  # type: ignore[union-attr]
    apply_header_style(label_header, style)
    for p_idx, period in enumerate(periods):
        base_col = 2 + p_idx * n_sub_cols
        end_col = base_col + n_sub_cols - 1
        ws.merge_cells(f"{get_column_letter(base_col)}2:{get_column_letter(end_col)}2")  # type: ignore[union-attr]
        ph = ws.cell(row=2, column=base_col, value=period.label)  # type: ignore[union-attr]
        apply_header_style(ph, style)

    ws.cell(row=3, column=1, value="")  # type: ignore[union-attr]
    apply_header_style(ws.cell(row=3, column=1), style)  # type: ignore[union-attr]
    for p_idx in range(len(periods)):
        base_col = 2 + p_idx * n_sub_cols
        for s_idx, scenario in enumerate(spec.scenarios):
            cell = ws.cell(row=3, column=base_col + s_idx, value=sanitize_cell_text(scenario.label))  # type: ignore[union-attr]
            apply_header_style(cell, style)


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
    n_sub_cols = n_scenarios
    total_cols = 1 + len(periods) * n_sub_cols

    first_proj_col_letter, last_proj_col_letter = compute_proj_col_range(periods, n_sub_cols, 2)

    build_model_header(ws, spec.title, total_cols, style, "Line Item", 13, "B4")
    _write_scenario_headers(ws, periods, spec, n_sub_cols, total_cols, style)

    sections_order, sections_items = group_line_items_by_section(spec.line_items)
    row_map = assign_row_map(sections_order, sections_items, 4)

    # Write data
    current_row = 4
    for section in sections_order:
        if section:
            write_section_header(ws, section, current_row, total_cols, style)
            current_row += 1

        for li in sections_items[section]:
            label_cell = ws.cell(row=current_row, column=1, value=sanitize_cell_text(li.label))
            apply_label_style(label_cell, li, style)

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
                    apply_label_style(cell, li, style)

            # Apply conditional formatting to variance rows when positive_is_good is specified
            if li.formula_type in ("variance", "variance_pct") and "positive_is_good" in li.formula_params:
                positive_is_good = bool(li.formula_params["positive_is_good"])
                data_start = get_column_letter(2)
                data_end = get_column_letter(total_cols)
                cf_range = f"{data_start}{current_row}:{data_end}{current_row}"
                apply_conditional_formatting(ws, cf_range, positive_is_good, style)

            current_row += 1
