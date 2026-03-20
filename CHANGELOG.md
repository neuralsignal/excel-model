# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Security

* Reject dangerous patterns in custom Excel formulas (DDE, WEBSERVICE, IMPORTDATA, CALL, EXEC, etc.) to prevent formula injection attacks. Validation runs at both spec validation and formula rendering time as defense-in-depth. ([#28](https://github.com/neuralsignal/excel-model/issues/28))

## [0.1.1](https://github.com/neuralsignal/excel-model/compare/v0.1.0...v0.1.1) (2026-03-17)


### Bug Fixes

* regenerate pixi.lock in release-please PR ([30bab7d](https://github.com/neuralsignal/excel-model/commit/30bab7d61ca39d68bee7fc05f550c5873bb63532))
* regenerate pixi.lock in release-please PR ([f883e54](https://github.com/neuralsignal/excel-model/commit/f883e5443676c9c76281055e181e14673bad0bf7))
* replace assert with ExcelModelError in model sheet builders ([#6](https://github.com/neuralsignal/excel-model/issues/6)) ([58dba01](https://github.com/neuralsignal/excel-model/commit/58dba0161c27c164c4166a338267e66158cf7b8d))


### Documentation

* add docs link, badges, and financial modeling input section ([d5a7af2](https://github.com/neuralsignal/excel-model/commit/d5a7af2eac179d575f9f6ad69f6c5751d29de5d5))
* add docs link, badges, and financial modeling input section ([fcd2ded](https://github.com/neuralsignal/excel-model/commit/fcd2ded0528d5339a73a801bf3ec6adac0ef2c5a))
* sync documentation with codebase ([dca4d0d](https://github.com/neuralsignal/excel-model/commit/dca4d0d109857bd7ffc959e677c13654ec3d93a9))
* sync documentation with codebase changes ([ce4a78c](https://github.com/neuralsignal/excel-model/commit/ce4a78c93d68ccf19ac0627a52c3004258163d0c))

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
