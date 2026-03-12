# excel-model

YAML-driven Excel financial model generator.

Build professional financial models (P&L, DCF, Budget vs Actuals, Scenario Analysis) from declarative YAML specs. Generates `.xlsx` workbooks with named ranges, styled sheets, and Excel formulas using openpyxl.

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

## License

MIT
