"""Register assumption cells as Excel named ranges using openpyxl."""

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName


def get_col_letter(col: int) -> str:
    """Convert 1-based column index to letter(s): 1→A, 26→Z, 27→AA."""
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


def register_named_range(wb: Workbook, name: str, sheet_name: str, row: int, col: int) -> None:
    """Register a single cell as an Excel named range.

    Uses absolute reference: 'SheetName'!$C$5
    """
    col_letter = get_col_letter(col)
    # Sheet names with spaces must be quoted
    if " " in sheet_name:
        ref = f"'{sheet_name}'!${col_letter}${row}"
    else:
        ref = f"{sheet_name}!${col_letter}${row}"

    defined_name = DefinedName(name=name, attr_text=ref)
    wb.defined_names[name] = defined_name
