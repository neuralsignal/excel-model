# excel-model

YAML-driven Excel financial model generator.

Build professional financial models from declarative YAML specs. Generates `.xlsx` workbooks with named ranges, styled sheets, and Excel formulas.

## Model Types

| Type | Description |
|------|-------------|
| `p_and_l` | Profit & Loss statement with growth projections |
| `dcf` | Discounted Cash Flow valuation |
| `budget_vs_actuals` | Monthly/quarterly variance analysis |
| `scenario` | Multi-scenario (Base/Bull/Bear) side-by-side analysis |
| `comparison` | Cross-entity comparison with rankings |
| `custom` | Custom P&L-style model (uses P&L builder) |

## Quick Start

### Installation

```bash
pip install excel-model
```

### CLI

```bash
# Build a model
excel-model build --spec model.yaml --output model.xlsx --mode batch

# Build with input data
excel-model build --spec model.yaml --output model.xlsx --mode batch --data actuals.csv

# Validate a spec
excel-model validate --spec model.yaml

# Validate spec and input data column mapping
excel-model validate --spec model.yaml --data actuals.csv

# Describe what a spec would produce (dry run)
excel-model describe --spec model.yaml --format text
```

`--mode` accepts `batch` (JSON to stdout) or `interactive` (verbose narrative).

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

## Standalone Data Sheets

Build standalone tabular data sheets and SUMIFS pivot sheets independent of a full model spec:

```python
from openpyxl import Workbook
from excel_model.config import load_style
from excel_model import DataSheetDef, SumifsPivotDef, build_data_sheet, build_sumifs_pivot

wb = Workbook()
del wb["Sheet"]
style = load_style(None)

# Simple tabular data sheet
data_spec = DataSheetDef(
    sheet_name="DATA", title="My Data",
    headers=("Name", "Amount"),
    col_widths=(20.0, 14.0), number_formats={1: "#,##0"}, freeze_row=2,
)
build_data_sheet(wb=wb, spec=data_spec, rows=[["Vendor A", 12345]], style=style)

# SUMIFS pivot sheet (formulas reference a source data sheet)
pivot_spec = SumifsPivotDef(
    sheet_name="PIVOT", title="Revenue by Hub",
    row_label_headers=("Hub",), col_dim_values=(2023, 2024, 2025),
    data_sheet="DATA", value_col="B",
    row_filter_cols=("A",), col_filter_col="C",
    append_total=True, append_yoy=True,
    col_widths=(28.0, 14.0, 14.0, 14.0, 16.0, 12.0, 12.0),
    number_format_data="#,##0", number_format_pct="0.0%", freeze_row=2,
)
build_sumifs_pivot(wb=wb, spec=pivot_spec, row_labels=[["Berlin"], ["Munich"]], style=style)

wb.save("output.xlsx")
```

## Configuration

Style config controls Excel formatting (colors, fonts, number formats). A bundled default is included; override with `--style path/to/style.yaml`.

User values are deep-merged on top of the bundled defaults — specify only the keys you want to override:

```yaml
header_fill_hex: "1F3864"
header_font_color: "FFFFFF"
subtotal_fill_hex: "D6E4F0"
total_fill_hex: "AED6F1"
history_col_fill_hex: "F2F2F2"
section_header_fill_hex: "E8F4FD"
alt_row_fill_hex: "F2F2F2"
font_name: "Calibri"
font_size: 10
number_format_currency: '#,##0'
number_format_percent: '0.0%'
number_format_integer: '#,##0'
number_format_number: '#,##0.00'
```

## Security

The `custom` formula type validates user-supplied formulas before writing them to the workbook. Formulas containing dangerous patterns that could enable Excel formula injection (CMD, DDE, DDEAUTO, WEBSERVICE, IMPORTDATA, IMPORTFEED, IMPORTHTML, IMPORTRANGE, IMPORTXML, CALL, EXEC, FILTERXML, HYPERLINK, REGISTER.ID, RTD, INDIRECT, ENCODEURL, pipe-based DDE invocations, and UNC paths) are rejected with a `FormulaInjectionError`. Standard Excel functions (SUM, IF, ROUND, etc.) are allowed.

See the [API Reference](api/spec.md) for full Python API documentation.
