# excel-model

[![CI](https://github.com/neuralsignal/excel-model/actions/workflows/ci.yml/badge.svg)](https://github.com/neuralsignal/excel-model/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

YAML-driven Excel financial model generator.

Build professional financial models (P&L, DCF, Budget vs Actuals, Scenario Analysis) from declarative YAML specs. Generates `.xlsx` workbooks with named ranges, styled sheets, and Excel formulas using openpyxl.

**[Documentation](https://neuralsignal.github.io/excel-model/)**

## Installation

```bash
pip install excel-model
```

Or for development:

```bash
pixi install
```

## Quick Start

### CLI

```bash
# Build a P&L model
excel-model build --spec model.yaml --output model.xlsx --mode batch

# Validate a spec
excel-model validate --spec model.yaml

# Describe what a spec would produce
excel-model describe --spec model.yaml --format text
```

### Python API

```python
from excel_model.spec_loader import load_spec
from excel_model.validator import validate_spec
from excel_model.excel_writer import build_workbook
from excel_model.config import load_style

spec = load_spec("model.yaml")
errors = validate_spec(spec)
assert not errors

style = load_style(None)  # uses bundled defaults
build_workbook(spec=spec, inputs=None, output_path="model.xlsx", style=style)
```

### Standalone Data Sheets (no spec required)

```python
from openpyxl import Workbook
from excel_model.config import load_style
from excel_model.data_sheet import write_data_sheet, write_sumifs_pivot

wb = Workbook()
del wb["Sheet"]
style = load_style(None)

# Simple tabular data sheet
write_data_sheet(
    wb=wb, sheet_name="DATA",
    headers=["Name", "Amount"], rows=[["Vendor A", 12345]],
    style=style, title="My Data",
    col_widths=[20.0, 14.0], number_formats={1: "#,##0"}, freeze_row=2,
)

# SUMIFS pivot sheet (formulas reference a source data sheet)
write_sumifs_pivot(
    wb=wb, sheet_name="PIVOT",
    title="Revenue by Hub", style=style,
    row_label_headers=["Hub"], row_labels=[["Berlin"], ["Munich"]],
    col_dim_values=[2023, 2024, 2025],
    data_sheet="DATA", value_col="B",
    row_filter_cols=["A"], col_filter_col="C",
    append_total=True, append_yoy=True,
    col_widths=[28.0, 14.0, 14.0, 14.0, 16.0, 12.0, 12.0],
    number_format_data="#,##0", number_format_pct="0.0%", freeze_row=2,
)

wb.save("output.xlsx")
```

## Model Types

| Type | Description |
|------|-------------|
| `p_and_l` | Profit & Loss statement |
| `dcf` | Discounted Cash Flow valuation |
| `budget_vs_actuals` | Budget vs Actuals comparison |
| `scenario` | Multi-scenario analysis (Base/Bull/Bear) |
| `comparison` | Cross-entity comparison |

## Formula Types

21 built-in formula types including `growth_projected`, `pct_of_revenue`, `sum_of_rows`, `subtraction`, `ratio`, `discounted_pv`, `terminal_value`, `npv_sum`, `variance`, `variance_pct`, `constant`, `custom`, and more.

**Custom formula security:** The `custom` formula type rejects formulas containing dangerous patterns (DDE, WEBSERVICE, IMPORTDATA, CALL, EXEC, FILTERXML, REGISTER.ID, etc.) to prevent Excel formula injection attacks. Standard Excel functions like SUM, IF, ROUND, and MAX are allowed.

## Configuration

Style config controls Excel formatting (colors, fonts, number formats). A bundled default is included; override with `--style`:

```yaml
header_fill_hex: "1F3864"
header_font_color: "FFFFFF"
font_name: "Calibri"
font_size: 10
number_format_currency: '#,##0'
number_format_percent: '0.0%'
```

## Security Note

File path arguments (`--spec`, `--data`, `--style`, `--output`) are passed directly to the filesystem without path containment checks. This is safe for the default CLI usage where the authenticated user controls their own filesystem. If you wrap this tool in a web API or automated pipeline that accepts user-controlled path inputs, you must validate that resolved paths stay within an allowed base directory before invoking the CLI to prevent path traversal attacks.

## Looking for Financial Modeling Input

This library was built by a software engineer, not a financial analyst. The model structures, formula types, and default assumptions reflect a developer's interpretation of common financial models.

If you work in finance, FP&A, investment banking, or accounting, your input would be incredibly valuable:

- **Are the formula types correct?** Do `growth_projected`, `pct_of_revenue`, `discounted_pv`, and `terminal_value` follow standard conventions?
- **Missing model patterns?** Are there common financial model structures (e.g., waterfall, three-statement, LBO) that should be supported?
- **Named range conventions** -- do the Excel named range naming patterns match what analysts expect?
- **Number formatting** -- are the default currency/percent/integer formats appropriate for professional models?
- **Scenario analysis** -- does the base/bull/bear override pattern match how scenarios are typically structured?

Please open an [issue](https://github.com/neuralsignal/excel-model/issues) with the `type:feat` label, or start a discussion. All feedback is welcome, from quick corrections to detailed model reviews.

## License

MIT
