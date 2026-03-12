"""Shared utilities for building Assumptions and Inputs sheets."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from excel_model.loader import InputData
from excel_model.named_ranges import register_named_range
from excel_model.spec import AssumptionDef, ModelSpec, ScenarioDef
from excel_model.style import (
    StyleConfig,
    apply_assumption_sheet_validation,
    apply_header_style,
    apply_normal_style,
    apply_section_header_style,
    get_number_format,
)
from excel_model.time_engine import Period


def build_assumptions_sheet(
    wb: Workbook,
    spec: ModelSpec,
    style: StyleConfig,
    scenario_prefix: str,
) -> dict[str, int]:
    """Build the Assumptions sheet.

    Returns a dict of assumption_name → row number (1-based).
    Registers named ranges for all assumptions.
    """
    sheet_name = "Assumptions"
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(title=sheet_name)

    # Row 1: Title header
    ws.merge_cells("A1:D1")
    title_cell = ws["A1"]
    title_cell.value = f"{spec.title} — Assumptions"
    apply_header_style(title_cell, style)
    ws.row_dimensions[1].height = 20

    # Row 2: Column headers
    headers = ["Parameter", "Named Range", "Value", "Format"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        apply_header_style(cell, style)

    # Column widths
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 12

    current_row = 3
    assumption_rows: dict[str, int] = {}

    # Group assumptions by group
    groups: dict[str, list[AssumptionDef]] = {}
    for assumption in spec.assumptions:
        groups.setdefault(assumption.group, []).append(assumption)

    for group_name, assumptions in groups.items():
        # Section header row
        ws.merge_cells(f"A{current_row}:D{current_row}")
        section_cell = ws[f"A{current_row}"]
        section_cell.value = group_name
        apply_section_header_style(section_cell, style)
        current_row += 1

        group_start_row = current_row
        for assumption in assumptions:
            range_name = f"{scenario_prefix}{assumption.name}" if scenario_prefix else assumption.name
            ws.cell(row=current_row, column=1, value=assumption.label)
            ws.cell(row=current_row, column=2, value=range_name)

            value_cell = ws.cell(row=current_row, column=3, value=assumption.value)
            value_cell.number_format = get_number_format(assumption.format, style)
            value_cell.alignment = Alignment(horizontal="right")

            apply_normal_style(ws.cell(row=current_row, column=1), style)
            apply_normal_style(ws.cell(row=current_row, column=2), style)
            apply_normal_style(value_cell, style)

            ws.cell(row=current_row, column=4, value=assumption.format)
            apply_normal_style(ws.cell(row=current_row, column=4), style)

            assumption_rows[assumption.name] = current_row

            # Register named range — column C = column index 3
            register_named_range(wb, range_name, sheet_name, current_row, 3)

            current_row += 1

        # Apply data validation to this group's assumption rows
        apply_assumption_sheet_validation(
            ws=ws,
            assumptions=assumptions,
            value_col=3,
            format_col=4,
            start_row=group_start_row,
        )

    return assumption_rows


def build_drivers_sheet(
    wb: Workbook,
    spec: ModelSpec,
    style: StyleConfig,
    scenarios: tuple[ScenarioDef, ...],
) -> dict[str, int]:
    """Build the Drivers sheet with one section per scenario.

    Returns a dict of driver_name → row number (1-based) for the last section written.
    Registers scenario-prefixed named ranges for each driver.
    """
    sheet_name = "Drivers"
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(title=sheet_name)

    ws.merge_cells("A1:D1")
    title_cell = ws["A1"]
    title_cell.value = f"{spec.title} — Drivers"
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
    driver_rows: dict[str, int] = {}

    for scenario in scenarios:
        prefix = scenario.name.capitalize()

        # Scenario section header
        ws.merge_cells(f"A{current_row}:D{current_row}")
        sec_cell = ws[f"A{current_row}"]
        sec_cell.value = f"{scenario.label} Drivers"
        apply_section_header_style(sec_cell, style)
        current_row += 1

        for driver in spec.drivers:
            range_name = f"{prefix}{driver.name}"
            # Override value if specified in driver_overrides
            value = scenario.driver_overrides.get(driver.name, driver.value)

            ws.cell(row=current_row, column=1, value=driver.label)
            ws.cell(row=current_row, column=2, value=range_name)

            value_cell = ws.cell(row=current_row, column=3, value=value)
            value_cell.number_format = get_number_format(driver.format, style)
            value_cell.alignment = Alignment(horizontal="right")

            apply_normal_style(ws.cell(row=current_row, column=1), style)
            apply_normal_style(ws.cell(row=current_row, column=2), style)
            apply_normal_style(value_cell, style)

            ws.cell(row=current_row, column=4, value=driver.format)
            apply_normal_style(ws.cell(row=current_row, column=4), style)

            driver_rows[driver.name] = current_row

            register_named_range(wb, range_name, sheet_name, current_row, 3)
            current_row += 1

    return driver_rows


def build_inputs_sheet(
    wb: Workbook,
    spec: ModelSpec,
    inputs: InputData | None,
    style: StyleConfig,
    periods: list[Period],
) -> dict[str, int]:
    """Build the Inputs sheet.

    Returns a dict of line_item_key → row number (1-based).
    """
    sheet_name = "Inputs"
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(title=sheet_name)

    history_periods = [p for p in periods if p.is_history]
    n_cols = 1 + len(history_periods)  # label col + one per history period

    # Row 1: Title header
    ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
    title_cell = ws["A1"]
    title_cell.value = "Historical Input Data"
    apply_header_style(title_cell, style)
    ws.row_dimensions[1].height = 20

    if not history_periods:
        ws["A2"] = "No history periods defined."
        return {}

    # Row 2: Column headers — "Line Item" + period labels
    header_cell = ws.cell(row=2, column=1, value="Line Item")
    apply_header_style(header_cell, style)
    for col_idx, period in enumerate(history_periods, start=2):
        cell = ws.cell(row=2, column=col_idx, value=period.label)
        apply_header_style(cell, style)

    ws.column_dimensions["A"].width = 28

    if inputs is None:
        ws["A3"] = "(No input data provided)"
        return {}

    inputs_row_map: dict[str, int] = {}
    current_row = 3

    # One row per value_col mapping
    for line_item_key, source_col in spec.inputs.value_cols.items():
        label_cell = ws.cell(row=current_row, column=1, value=line_item_key)
        apply_normal_style(label_cell, style)

        inputs_row_map[line_item_key] = current_row

        # Fill history data from inputs DataFrame
        for col_idx, period in enumerate(history_periods, start=2):
            period_val = None
            if inputs.period_col in inputs.df.columns and source_col in inputs.df.columns:
                rows_for_period = inputs.df.filter(inputs.df[inputs.period_col] == period.label)
                if len(rows_for_period) > 0:
                    period_val = rows_for_period[source_col][0]
            cell = ws.cell(row=current_row, column=col_idx, value=period_val)
            apply_normal_style(cell, style)

        current_row += 1

    return inputs_row_map


def write_model_header_row(
    ws,
    periods: list[Period],
    style: StyleConfig,
    n_header_cols_before: int = 0,
) -> None:
    """Write the period label row (row 2 in model sheet).

    Columns: "Line Item" + one column per period.
    History columns get history fill.
    """
    from excel_model.style import apply_history_col_style

    label_cell = ws.cell(row=2, column=1, value="Line Item")
    apply_header_style(label_cell, style)

    for col_idx, period in enumerate(periods, start=2):
        cell = ws.cell(row=2, column=col_idx, value=period.label)
        if period.is_history:
            apply_history_col_style(cell, style)
            cell.font = Font(name=style.font_name, size=style.font_size, bold=True)
        else:
            apply_header_style(cell, style)
