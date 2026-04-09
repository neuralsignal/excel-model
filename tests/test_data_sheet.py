"""Tests for excel_model.data_sheet — write_data_sheet and write_sumifs_pivot."""

from __future__ import annotations

import pytest
from hypothesis import given, settings
from hypothesis import strategies as st
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from excel_model.config import load_style
from excel_model.data_sheet import write_data_sheet, write_sumifs_pivot


@pytest.fixture
def style():
    return load_style(None)  # bundled defaults


@pytest.fixture
def wb():
    workbook = Workbook()
    del workbook["Sheet"]
    return workbook


# ---------------------------------------------------------------------------
# write_data_sheet tests
# ---------------------------------------------------------------------------


def test_creates_sheet_with_correct_name(wb, style):
    write_data_sheet(
        wb=wb, sheet_name="TEST_SHEET",
        headers=["A", "B"], rows=[[1, 2], [3, 4]],
        style=style, title="Test",
        col_widths=[15.0, 15.0], number_formats={}, freeze_row=2,
    )
    assert "TEST_SHEET" in wb.sheetnames


def test_row_count(wb, style):
    data_rows = [[i, i * 2] for i in range(10)]
    ws = write_data_sheet(
        wb=wb, sheet_name="RC",
        headers=["X", "Y"], rows=data_rows,
        style=style, title="Row Count Test",
        col_widths=[10.0, 10.0], number_formats={}, freeze_row=2,
    )
    assert ws.max_row == 2 + len(data_rows)


def test_header_values(wb, style):
    headers = ["Hub", "2023", "2024", "Total"]
    write_data_sheet(
        wb=wb, sheet_name="HDR",
        headers=headers, rows=[["DE Berlin", 100, 200, 300]],
        style=style, title="Header Test",
        col_widths=[20.0, 12.0, 12.0, 14.0], number_formats={}, freeze_row=2,
    )
    ws = wb["HDR"]
    actual = [ws.cell(row=2, column=c).value for c in range(1, len(headers) + 1)]
    assert actual == headers


def test_number_format_applied(wb, style):
    write_data_sheet(
        wb=wb, sheet_name="FMT",
        headers=["Name", "Amount"], rows=[["Vendor A", 12345.67]],
        style=style, title="Format Test",
        col_widths=[20.0, 14.0], number_formats={1: "#,##0"}, freeze_row=2,
    )
    ws = wb["FMT"]
    assert ws.cell(row=3, column=2).number_format == "#,##0"


def test_freeze_panes_default(wb, style):
    ws = write_data_sheet(
        wb=wb, sheet_name="FRZ",
        headers=["A"], rows=[[1]],
        style=style, title="Freeze Test",
        col_widths=[10.0], number_formats={}, freeze_row=2,
    )
    assert ws.freeze_panes == "A3"


def test_freeze_panes_custom(wb, style):
    ws = write_data_sheet(
        wb=wb, sheet_name="FRZ2",
        headers=["A"], rows=[[1]],
        style=style, title="Freeze Test",
        col_widths=[10.0], number_formats={}, freeze_row=1,
    )
    assert ws.freeze_panes == "A2"


def test_col_widths(wb, style):
    widths = [25.0, 12.5, 18.0]
    ws = write_data_sheet(
        wb=wb, sheet_name="WID",
        headers=["A", "B", "C"], rows=[],
        style=style, title="Width Test",
        col_widths=widths, number_formats={}, freeze_row=2,
    )
    for i, expected in enumerate(widths, start=1):
        assert ws.column_dimensions[get_column_letter(i)].width == expected


def test_title_in_row_1(wb, style):
    ws = write_data_sheet(
        wb=wb, sheet_name="TTL",
        headers=["X"], rows=[[42]],
        style=style, title="My Title",
        col_widths=[10.0], number_formats={}, freeze_row=2,
    )
    assert ws["A1"].value == "My Title"


def test_empty_rows(wb, style):
    ws = write_data_sheet(
        wb=wb, sheet_name="EMPTY",
        headers=["Col1", "Col2"], rows=[],
        style=style, title="Empty",
        col_widths=[10.0, 10.0], number_formats={}, freeze_row=2,
    )
    assert ws.max_row == 2


def test_returns_worksheet(wb, style):
    from openpyxl.worksheet.worksheet import Worksheet
    result = write_data_sheet(
        wb=wb, sheet_name="RET",
        headers=["A"], rows=[[1]],
        style=style, title="Return Test",
        col_widths=[10.0], number_formats={}, freeze_row=2,
    )
    assert isinstance(result, Worksheet)


# ---------------------------------------------------------------------------
# Property test for write_data_sheet (printable text only)
# ---------------------------------------------------------------------------

_printable_text = st.text(
    alphabet=st.characters(
        whitelist_categories=("L", "N", "P", "S"),  # letters, numbers, punctuation, symbols
    ),
    min_size=1, max_size=20,
)


@given(
    headers=st.lists(_printable_text, min_size=1, max_size=8),
    n_rows=st.integers(min_value=0, max_value=30),
)
@settings(max_examples=50)
def test_write_data_sheet_row_count_property(headers, n_rows):
    style = load_style(None)
    workbook = Workbook()
    del workbook["Sheet"]
    rows = [[None] * len(headers) for _ in range(n_rows)]
    col_widths = [10.0] * len(headers)
    ws = write_data_sheet(
        wb=workbook, sheet_name="PROP",
        headers=headers, rows=rows,
        style=style, title="Property Test",
        col_widths=col_widths, number_formats={}, freeze_row=2,
    )
    assert ws.max_row == 2 + n_rows


# ---------------------------------------------------------------------------
# write_sumifs_pivot tests
# ---------------------------------------------------------------------------


def _make_sumifs_sheet(wb, style, **overrides):
    """Helper with sensible defaults for sumifs pivot tests."""
    defaults = dict(
        sheet_name="SUMIFS_TEST",
        title="Test Pivot",
        style=style,
        row_label_headers=["Hub"],
        row_labels=[["DE Berlin"], ["DE Munich"], ["DE Hamburg"]],
        col_dim_values=[2023, 2024, 2025],
        data_sheet="TRANSACTIONS_LNFW",
        value_col="AO",
        row_filter_cols=["AM"],
        col_filter_col="AJ",
        append_total=True,
        append_yoy=False,
        col_widths=[28.0, 14.0, 14.0, 14.0, 16.0],
        number_format_data="#,##0",
        number_format_pct="0.0%",
        freeze_row=2,
    )
    defaults.update(overrides)
    return write_sumifs_pivot(wb=wb, **defaults)


def test_sumifs_creates_sheet(wb, style):
    _make_sumifs_sheet(wb, style, sheet_name="MY_PIVOT")
    assert "MY_PIVOT" in wb.sheetnames


def test_sumifs_row_count(wb, style):
    row_labels = [["Hub A"], ["Hub B"], ["Hub C"]]
    ws = _make_sumifs_sheet(wb, style, row_labels=row_labels)
    assert ws.max_row == 2 + len(row_labels)


def test_sumifs_formula_in_data_cell(wb, style):
    """Cell B3 (first data col, first data row) contains a SUMIFS formula."""
    ws = _make_sumifs_sheet(wb, style)
    cell_value = ws.cell(row=3, column=2).value
    assert isinstance(cell_value, str)
    assert cell_value.startswith("=SUMIFS(")
    assert "TRANSACTIONS_LNFW" in cell_value
    assert "$AO:$AO" in cell_value
    assert "$AM:$AM" in cell_value
    assert "$AJ:$AJ" in cell_value


def test_sumifs_formula_row_ref_pattern(wb, style):
    """Row 3 formula uses $A3 (absolute col, relative row); row 4 uses $A4."""
    ws = _make_sumifs_sheet(wb, style)
    formula_row3 = ws.cell(row=3, column=2).value
    formula_row4 = ws.cell(row=4, column=2).value
    assert "$A3" in formula_row3
    assert "$A4" in formula_row4


def test_sumifs_formula_col_ref_pattern(wb, style):
    """Column 2 formula uses B$2 (relative col); column 3 uses C$2."""
    ws = _make_sumifs_sheet(wb, style)
    formula_col2 = ws.cell(row=3, column=2).value  # col B
    formula_col3 = ws.cell(row=3, column=3).value  # col C
    assert "B$2" in formula_col2
    assert "C$2" in formula_col3


def test_sumifs_row_labels_written(wb, style):
    row_labels = [["DE Kiel-NB"], ["DE Braunschweig"]]
    ws = _make_sumifs_sheet(wb, style, row_labels=row_labels)
    assert ws.cell(row=3, column=1).value == "DE Kiel-NB"
    assert ws.cell(row=4, column=1).value == "DE Braunschweig"


def test_sumifs_multi_label_cols(wb, style):
    """Two label columns (Hub × Account) both appear in SUMIFS criteria."""
    ws = write_sumifs_pivot(
        wb=wb, sheet_name="BY_CAT",
        title="By Category", style=style,
        row_label_headers=["Hub", "Account"],
        row_labels=[["DE Berlin", "Maintenance"], ["DE Berlin", "Licenses"]],
        col_dim_values=[2023, 2024],
        data_sheet="TRANSACTIONS_LNFW",
        value_col="AO",
        row_filter_cols=["AM", "AN"],
        col_filter_col="AJ",
        append_total=True, append_yoy=False,
        col_widths=[28, 22, 14, 14, 16],
        number_format_data="#,##0", number_format_pct="0.0%",
        freeze_row=2,
    )
    formula = ws.cell(row=3, column=3).value  # first data col (C)
    assert "$AM:$AM" in formula
    assert "$AN:$AN" in formula
    assert "$A3" in formula
    assert "$B3" in formula


def test_sumifs_static_label_not_in_formula(wb, style):
    """Category col (index 1) is static — not in row_filter_cols — so absent from SUMIFS."""
    ws = write_sumifs_pivot(
        wb=wb, sheet_name="VENDORS",
        title="Top Vendors", style=style,
        row_label_headers=["Vendor", "Category"],
        row_labels=[["Microsoft 365 (via reseller)", "Software/SaaS"]],
        col_dim_values=[2023, 2024],
        data_sheet="TRANSACTIONS_LNFW",
        value_col="AO",
        row_filter_cols=["AQ"],       # only vendor_canonical; Category not filtered
        col_filter_col="AJ",
        append_total=True, append_yoy=False,
        col_widths=[35, 18, 14, 14, 16],
        number_format_data="#,##0", number_format_pct="0.0%",
        freeze_row=2,
    )
    formula = ws.cell(row=3, column=3).value  # first data col
    # AQ referenced (vendor filter)
    assert "$AQ:$AQ" in formula
    # col B (Category) NOT referenced in formula
    assert "$B3" not in formula
    # Category label is written statically
    assert ws.cell(row=3, column=2).value == "Software/SaaS"


def test_sumifs_total_column(wb, style):
    """Total column formula is SUM over data cols."""
    ws = _make_sumifs_sheet(wb, style, col_dim_values=[2023, 2024, 2025], append_total=True)
    # 1 label col + 3 data cols → total at col 5
    total_cell = ws.cell(row=3, column=5).value
    assert isinstance(total_cell, str)
    assert total_cell.startswith("=SUM(")
    assert "B3:D3" in total_cell


def test_sumifs_no_total(wb, style):
    """When append_total=False, total col not present."""
    ws = _make_sumifs_sheet(wb, style, col_dim_values=[2023, 2024], append_total=False, col_widths=[28, 14, 14])
    # 1 label + 2 data = 3 cols max
    assert ws.max_column == 3


def test_sumifs_yoy_columns(wb, style):
    """YoY columns contain IF formula with ABS."""
    ws = write_sumifs_pivot(
        wb=wb, sheet_name="YOY",
        title="YoY Test", style=style,
        row_label_headers=["Hub"],
        row_labels=[["DE Berlin"]],
        col_dim_values=[2023, 2024, 2025],
        data_sheet="TRANSACTIONS_LNFW",
        value_col="AO", row_filter_cols=["AM"], col_filter_col="AJ",
        append_total=True, append_yoy=True,
        col_widths=[28, 14, 14, 14, 16, 12, 12],
        number_format_data="#,##0", number_format_pct="0.0%",
        freeze_row=2,
    )
    # layout: A=Hub, B=2023, C=2024, D=2025, E=Total, F=YoY23→24, G=YoY24→25
    yoy_cell = ws.cell(row=3, column=6).value
    assert isinstance(yoy_cell, str)
    assert yoy_cell.startswith("=IF(")
    assert "ABS(" in yoy_cell


def test_sumifs_no_yoy(wb, style):
    """When append_yoy=False, no YoY columns after Total."""
    ws = _make_sumifs_sheet(
        wb, style,
        col_dim_values=[2023, 2024],
        append_total=True, append_yoy=False,
        col_widths=[28, 14, 14, 16],
    )
    # 1 label + 2 data + 1 total = 4 cols
    assert ws.max_column == 4


def test_sumifs_col_headers_in_row2(wb, style):
    """Year values appear in row 2 of data columns."""
    ws = _make_sumifs_sheet(wb, style, col_dim_values=[2023, 2024, 2025])
    assert ws.cell(row=2, column=2).value == 2023
    assert ws.cell(row=2, column=3).value == 2024
    assert ws.cell(row=2, column=4).value == 2025


def test_sumifs_freeze_panes(wb, style):
    ws = _make_sumifs_sheet(wb, style, freeze_row=2)
    assert ws.freeze_panes == "A3"


def test_sumifs_number_format_on_data_cells(wb, style):
    ws = _make_sumifs_sheet(wb, style, number_format_data="#,##0")
    assert ws.cell(row=3, column=2).number_format == "#,##0"


def test_sumifs_number_format_on_yoy_cells(wb, style):
    ws = write_sumifs_pivot(
        wb=wb, sheet_name="FMTYOY",
        title="Fmt YoY", style=style,
        row_label_headers=["Hub"],
        row_labels=[["DE Berlin"]],
        col_dim_values=[2023, 2024],
        data_sheet="TRANSACTIONS_LNFW",
        value_col="AO", row_filter_cols=["AM"], col_filter_col="AJ",
        append_total=False, append_yoy=True,
        col_widths=[28, 14, 14, 12],
        number_format_data="#,##0", number_format_pct="0.0%",
        freeze_row=2,
    )
    # layout: A=Hub, B=2023, C=2024, D=YoY23→24
    yoy_cell = ws.cell(row=3, column=4)
    assert yoy_cell.number_format == "0.0%"
