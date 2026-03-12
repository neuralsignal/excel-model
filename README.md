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
