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

## Configuration

Style config controls Excel formatting (colors, fonts, number formats). A bundled default is included; override with `--style path/to/style.yaml`.

All keys are required when providing a custom style file; user values are deep-merged on top of the bundled defaults so you only need to specify what you want to override:

```yaml
header_fill_hex: "1F3864"
header_font_color: "FFFFFF"
subtotal_fill_hex: "D9E1F2"
total_fill_hex: "BDD7EE"
history_col_fill_hex: "F2F2F2"
section_header_fill_hex: "E2EFDA"
font_name: "Calibri"
font_size: 10
number_format_currency: '#,##0'
number_format_percent: '0.0%'
number_format_integer: '#,##0'
number_format_number: '#,##0.00'
```

## Security

The `custom` formula type validates user-supplied formulas before writing them to the workbook. Formulas containing dangerous patterns that could enable Excel formula injection (CMD, DDE, DDEAUTO, WEBSERVICE, IMPORTDATA, IMPORTFEED, IMPORTHTML, IMPORTRANGE, IMPORTXML, CALL, EXEC, FILTERXML, HYPERLINK, REGISTER.ID, RTD, and pipe-based DDE invocations) are rejected with a `FormulaInjectionError`. Standard Excel functions (SUM, IF, ROUND, etc.) are allowed.

See the [API Reference](api/spec.md) for full Python API documentation.
