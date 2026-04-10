"""Shared utilities for building Assumptions, Inputs, and Model sheets."""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from excel_model.loader import InputData
from excel_model.named_ranges import get_col_letter, register_named_range
from excel_model.spec import AssumptionDef, LineItemDef, ModelSpec, ScenarioDef
from excel_model.style import (
    StyleConfig,
    apply_assumption_sheet_validation,
    apply_header_style,
    apply_history_col_style,
    apply_normal_style,
    apply_section_header_style,
    apply_subtotal_style,
    apply_total_style,
    get_number_format,
)
from excel_model.time_engine import Period


@dataclass(frozen=True)
class SheetRenderContext:
    """Shared rendering state passed to model sheet cell-writing functions."""

    row_map: dict[str, int]
    inputs_row_map: dict[str, int]
    first_proj_col_letter: str
    last_proj_col_letter: str
    n_history: int
    named_ranges: dict[str, str]
    style: StyleConfig


def build_model_header(
    ws: Worksheet,
    title: str,
    total_cols: int,
    style: StyleConfig,
    label_col_header: str,
    data_col_width: int,
    freeze_cell: str,
) -> None:
    """Build the standard model sheet header rows shared by all builders.

    Handles: row-1 title merge + style, row-2 label column header,
    column-A width, data-column widths, and freeze panes.
    """
    # Row 1: Title
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
    title_cell = ws["A1"]
    title_cell.value = title
    apply_header_style(title_cell, style)
    ws.row_dimensions[1].height = 20

    # Row 2, column 1: Label header
    label_header = ws.cell(row=2, column=1, value=label_col_header)
    apply_header_style(label_header, style)

    # Column widths
    ws.column_dimensions["A"].width = 28
    for col_idx in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = data_col_width

    # Freeze panes
    ws.freeze_panes = freeze_cell


def group_line_items_by_section(
    line_items: tuple[LineItemDef, ...],
) -> tuple[list[str], dict[str, list[LineItemDef]]]:
    """Group line items by section, preserving insertion order.

    Returns (sections_order, sections_items) where sections_order is the list
    of unique section names in first-seen order and sections_items maps each
    section name to its line items.
    """
    sections_order: list[str] = []
    sections_items: dict[str, list[LineItemDef]] = {}
    for li in line_items:
        if li.section not in sections_items:
            sections_order.append(li.section)
            sections_items[li.section] = []
        sections_items[li.section].append(li)
    return sections_order, sections_items


def write_title_row(ws: Worksheet, title: str, total_cols: int, style: StyleConfig) -> None:
    """Write merged title row (row 1) with header styling."""
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
    cell = ws["A1"]
    cell.value = title
    apply_header_style(cell, style)
    ws.row_dimensions[1].height = 20


def assign_row_map(
    sections_order: list[str],
    sections_items: dict[str, list[LineItemDef]],
    start_row: int,
) -> dict[str, int]:
    """First-pass row number assignment for all line items."""
    current_row = start_row
    row_map: dict[str, int] = {}
    for section in sections_order:
        if section:
            current_row += 1
        for li in sections_items[section]:
            row_map[li.key] = current_row
            current_row += 1
    return row_map


def write_section_header(
    ws: Worksheet,
    section: str,
    row: int,
    total_cols: int,
    style: StyleConfig,
) -> None:
    """Write a merged section header row."""
    ws.merge_cells(f"A{row}:{get_column_letter(total_cols)}{row}")
    ws[f"A{row}"].value = section
    apply_section_header_style(ws[f"A{row}"], style)


def apply_label_style(cell: Cell, li: LineItemDef, style: StyleConfig) -> None:
    """Apply normal/subtotal/total style to a line item label cell."""
    apply_normal_style(cell, style)
    if li.is_subtotal:
        apply_subtotal_style(cell, style)
    elif li.is_total:
        apply_total_style(cell, style)


def apply_data_cell_style(cell: Cell, li: LineItemDef, style: StyleConfig, is_history: bool) -> None:
    """Apply appropriate style to a data cell based on line item type and history status."""
    if is_history:
        apply_history_col_style(cell, style)
    if li.is_subtotal:
        apply_subtotal_style(cell, style)
    elif li.is_total:
        apply_total_style(cell, style)
    else:
        apply_normal_style(cell, style)
        if is_history:
            apply_history_col_style(cell, style)


def compute_proj_col_range(
    periods: list[Period],
    col_multiplier: int,
    col_offset: int,
) -> tuple[str, str]:
    """Compute first/last projection column letters for formula context."""
    proj = [p for p in periods if not p.is_history]
    if not proj:
        return "", ""
    return (
        get_col_letter(col_offset + proj[0].index * col_multiplier),
        get_col_letter(col_offset + proj[-1].index * col_multiplier + col_multiplier - 1),
    )


def set_column_widths(
    ws: Worksheet,
    total_cols: int,
    label_width: float,
    data_width: float,
) -> None:
    """Set column A width and uniform data column widths."""
    ws.column_dimensions["A"].width = label_width
    for col_idx in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = data_width


def write_history_border(ws: Worksheet, row: int, n_history: int, total_cols: int) -> None:
    """Write thin vertical border at the first projection column."""
    if n_history > 0:
        border_col = 2 + n_history
        if border_col <= total_cols:
            ws.cell(row=row, column=border_col).border = Border(left=Side(style="thin"))


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

    write_title_row(ws, f"{spec.title} — Assumptions", 4, style)

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
        write_section_header(ws, group_name, current_row, 4, style)
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

    write_title_row(ws, f"{spec.title} — Drivers", 4, style)

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

        write_section_header(ws, f"{scenario.label} Drivers", current_row, 4, style)
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

    write_title_row(ws, "Historical Input Data", n_cols, style)

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
