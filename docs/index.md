# excel-model

YAML-driven Excel financial model generator.

Build professional financial models from declarative YAML specs. Generates `.xlsx` workbooks with named ranges, styled sheets, and Excel formulas.

## Model Types

- **P&L** — Profit & Loss statement with growth projections
- **DCF** — Discounted Cash Flow valuation
- **Budget vs Actuals** — Monthly/quarterly variance analysis
- **Scenario** — Multi-scenario (Base/Bull/Bear) side-by-side analysis
- **Comparison** — Cross-entity comparison with rankings

## Quick Start

```bash
pip install excel-model
excel-model build --spec model.yaml --output model.xlsx --mode batch
```

See the [API Reference](api/spec.md) for Python API usage.
