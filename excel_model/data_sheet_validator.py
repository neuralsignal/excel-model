"""Validation for DataSheetDef and SumifsPivotDef."""

import re

from openpyxl.utils import column_index_from_string

from excel_model.spec import DataSheetDef, SumifsPivotDef

_VALID_COL_LETTER_RE = re.compile(r"^[A-Z]{1,3}$")
_SAFE_SHEET_NAME_RE = re.compile(r"^[A-Za-z0-9_ ()\-]+$")


def _validate_column_letter(col: str, field_name: str) -> list[str]:
    """Validate that col is a valid Excel column letter (A–XFD)."""
    errors: list[str] = []
    if not _VALID_COL_LETTER_RE.match(col):
        errors.append(f"{field_name}: {col!r} is not a valid Excel column letter (A–XFD)")
        return errors
    try:
        column_index_from_string(col)
    except ValueError:
        errors.append(f"{field_name}: {col!r} is not a valid Excel column letter (A–XFD)")
    return errors


def _validate_sheet_name_safety(name: str, field_name: str) -> list[str]:
    """Reject sheet names with characters that could enable formula injection."""
    errors: list[str] = []
    if not name:
        errors.append(f"{field_name} must not be empty")
        return errors
    if not _SAFE_SHEET_NAME_RE.match(name):
        errors.append(
            f"{field_name}: {name!r} contains unsafe characters. "
            f"Only letters, digits, underscores, spaces, hyphens, and parentheses are allowed."
        )
    if len(name) > 31:
        errors.append(f"{field_name}: sheet name must be <= 31 characters, got {len(name)}")
    return errors


def validate_data_sheet_def(spec: DataSheetDef) -> list[str]:
    """Return a list of error strings for a DataSheetDef. Empty list means valid."""
    errors: list[str] = []
    errors.extend(_validate_sheet_name_safety(spec.sheet_name, "sheet_name"))
    if not spec.headers:
        errors.append("headers must not be empty")
    if spec.headers and len(spec.col_widths) != len(spec.headers):
        errors.append(f"col_widths length ({len(spec.col_widths)}) must match headers length ({len(spec.headers)})")
    if spec.freeze_row < 0:
        errors.append(f"freeze_row must be >= 0, got {spec.freeze_row}")
    for idx in spec.number_formats:
        if spec.headers and (idx < 0 or idx >= len(spec.headers)):
            errors.append(f"number_formats key {idx} is out of range (0–{len(spec.headers) - 1})")
    return errors


def _compute_expected_columns(spec: SumifsPivotDef) -> int:
    """Compute total column count from a SumifsPivotDef."""
    n = len(spec.row_label_headers) + len(spec.col_dim_values)
    if spec.append_total:
        n += 1
    if spec.append_yoy and len(spec.col_dim_values) > 1:
        n += len(spec.col_dim_values) - 1
    return n


def validate_sumifs_pivot_def(spec: SumifsPivotDef) -> list[str]:
    """Return a list of error strings for a SumifsPivotDef. Empty list means valid."""
    errors: list[str] = []
    errors.extend(_validate_sheet_name_safety(spec.sheet_name, "sheet_name"))
    errors.extend(_validate_sheet_name_safety(spec.data_sheet, "data_sheet"))
    if not spec.row_label_headers:
        errors.append("row_label_headers must not be empty")
    if not spec.row_filter_cols:
        errors.append("row_filter_cols must not be empty")
    if spec.row_filter_cols and spec.row_label_headers and len(spec.row_filter_cols) > len(spec.row_label_headers):
        errors.append(
            f"row_filter_cols length ({len(spec.row_filter_cols)}) "
            f"must not exceed row_label_headers length ({len(spec.row_label_headers)})"
        )
    errors.extend(_validate_column_letter(spec.value_col, "value_col"))
    for i, col in enumerate(spec.row_filter_cols):
        errors.extend(_validate_column_letter(col, f"row_filter_cols[{i}]"))
    errors.extend(_validate_column_letter(spec.col_filter_col, "col_filter_col"))
    expected_cols = _compute_expected_columns(spec)
    if len(spec.col_widths) != expected_cols:
        errors.append(f"col_widths length ({len(spec.col_widths)}) must match total column count ({expected_cols})")
    if spec.freeze_row < 0:
        errors.append(f"freeze_row must be >= 0, got {spec.freeze_row}")
    return errors
