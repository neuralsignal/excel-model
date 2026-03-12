"""Comparison model builder — entities as columns (not time periods)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from excel_model.formula_engine import CellContext, render_formula
from excel_model.loader import InputData
from excel_model.models._sheet_builder import build_assumptions_sheet
from excel_model.named_ranges import get_col_letter
from excel_model.spec import ModelSpec
from excel_model.style import (
    StyleConfig,
    apply_header_style,
    apply_normal_style,
    apply_section_header_style,
    apply_subtotal_style,
    apply_total_style,
    get_number_format,
)


def build_comparison(
    wb: Workbook,
    spec: ModelSpec,
    inputs: InputData | None,
    style: StyleConfig,
) -> None:
    """Build Assumptions and Comparison Model sheets."""
    build_assumptions_sheet(wb, spec, style, "")
    _build_comparison_model_sheet(wb, spec, style)


def _build_comparison_model_sheet(
    wb: Workbook,
    spec: ModelSpec,
    style: StyleConfig,
) -> None:
    """Comparison model: entities as columns, line items as rows."""
    ws = wb.create_sheet(title="Model")

    entities = spec.entities
    n_entities = len(entities)
    total_cols = 1 + n_entities  # label col + one per entity

    # Row 1: Title header
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
    title_cell = ws["A1"]
    title_cell.value = spec.title
    apply_header_style(title_cell, style)
    ws.row_dimensions[1].height = 20

    # Row 2: "Metric" | entity labels
    label_header = ws.cell(row=2, column=1, value="Metric")
    apply_header_style(label_header, style)

    for e_idx, entity in enumerate(entities):
        cell = ws.cell(row=2, column=2 + e_idx, value=entity.label)
        apply_header_style(cell, style)

    ws.column_dimensions["A"].width = 28
    for col_idx in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 16

    ws.freeze_panes = "B3"

    # Build entity_col_range for RANK/MAX formulas (row placeholder replaced per row)
    first_entity_col = get_col_letter(2)
    last_entity_col = get_col_letter(1 + n_entities)

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
            current_row += 1  # section header row
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

            # Build entity_col_range for this row (used by RANK/MAX formulas)
            entity_col_range = f"${first_entity_col}${current_row}:${last_entity_col}${current_row}"

            for e_idx, _entity in enumerate(entities):
                col_idx = 2 + e_idx
                col_letter = get_col_letter(col_idx)

                params = dict(li.formula_params)

                # For index_to_base, inject the base entity's column letter
                if li.formula_type == "index_to_base":
                    base_entity_key = params.get("base_entity_key", "")
                    base_col_idx = _entity_col_index(spec, base_entity_key)
                    if base_col_idx is not None:
                        params["_base_col_letter"] = get_col_letter(base_col_idx)

                ctx = CellContext(
                    period_index=0,
                    n_history=0,
                    row=current_row,
                    col=col_idx,
                    col_letter=col_letter,
                    prior_col_letter="",
                    named_ranges={a.name: a.name for a in spec.assumptions},
                    row_map=row_map,
                    inputs_row_map={},
                    scenario_prefix="",
                    first_proj_col_letter="",
                    last_proj_col_letter="",
                    entity_col_range=entity_col_range,
                )

                value = render_formula(li.formula_type, params, ctx)

                cell = ws.cell(row=current_row, column=col_idx, value=value)

                # Format: use percent for ratio-like items, currency otherwise
                fmt = "percent" if li.formula_type in ("ratio", "index_to_base") else "currency"
                cell.number_format = get_number_format(fmt, style)
                cell.alignment = Alignment(horizontal="right")

                apply_normal_style(cell, style)
                if li.is_subtotal:
                    apply_subtotal_style(cell, style)
                elif li.is_total:
                    apply_total_style(cell, style)

            current_row += 1


def _entity_col_index(spec: ModelSpec, entity_key: str) -> int | None:
    """Return the 1-based column index for the given entity key, or None."""
    for i, entity in enumerate(spec.entities):
        if entity.key == entity_key:
            return 2 + i
    return None
