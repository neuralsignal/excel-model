# Changelog

All notable changes to this project will be documented in this file.

## [0.1.0] - 2026-03-12

### Added

- Initial release
- 5 model types: P&L, DCF, Budget vs Actuals, Scenario Analysis, Cross-entity Comparison
- 21 formula types including growth_projected, pct_of_revenue, sum_of_rows, discounted_pv, terminal_value, custom
- YAML-driven model specs with strictyaml schema validation
- Named range registration for all assumptions and drivers
- Scenario analysis with per-scenario assumption/driver overrides
- Drivers concept: separate operational levers from structural assumptions
- Multi-format input data loading (CSV, Excel, Parquet, JSON, YAML, Markdown)
- Click CLI with build, validate, and describe commands
- Bundled default style config with deep-merge user overrides
- Period generation: annual, quarterly, monthly, auto-detect
- Conditional formatting for variance rows
- Data validation on assumptions sheets
- Property-based testing with hypothesis
- 191 tests, all passing
