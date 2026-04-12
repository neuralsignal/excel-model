# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Features

* refactor data sheet builders to spec-driven architecture: `DataSheetDef`/`SumifsPivotDef` spec dataclasses, `build_data_sheet`/`build_sumifs_pivot` builder functions, input validation with formula injection guards, config-driven alternating row styling ([#88](https://github.com/neuralsignal/excel-model/pull/88))

### Bug Fixes

* single-quote `data_sheet` references in SUMIFS formulas so names containing spaces produce valid Excel syntax (e.g. `'My Data'!$AO:$AO`); also annotate `DataSheetDef.number_formats` as `Mapping[int, str]` and narrow `SumifsPivotDef.col_dim_values` to `tuple[str | int | float, ...]` ([#89](https://github.com/neuralsignal/excel-model/pull/89))

## [0.1.3](https://github.com/neuralsignal/excel-model/compare/v0.1.2...v0.1.3) (2026-04-05)


### Bug Fixes

* add HYPERLINK and RTD to formula injection blocklist ([#53](https://github.com/neuralsignal/excel-model/issues/53)) ([13b2b50](https://github.com/neuralsignal/excel-model/commit/13b2b50f8f3f14d893cf2efaed90668f104a2f96))
* auto-fix CI failures (attempt 1) ([5abe174](https://github.com/neuralsignal/excel-model/commit/5abe174e6c449e637b37d812b49fae28900c6b72))
* block INDIRECT, ENCODEURL, and UNC paths in formula injection filter ([#72](https://github.com/neuralsignal/excel-model/issues/72)) ([79f6aeb](https://github.com/neuralsignal/excel-model/commit/79f6aeba1525ac1b0de942aeca8b1e4019583f7a))
* pin requests &gt;=2.33.0 to resolve CVE-2026-25645 ([#67](https://github.com/neuralsignal/excel-model/issues/67)) ([73b87af](https://github.com/neuralsignal/excel-model/commit/73b87afcf6b10dc5424ffeabde31b63cd9d77c21))
* pin requests &gt;=2.33.0 to resolve CVE-2026-25645 (predictable temp ([abb1807](https://github.com/neuralsignal/excel-model/commit/abb18077240f9dda711a20369e83044a41f361b0))
* propagate ValueError from generate_periods in describe command ([#78](https://github.com/neuralsignal/excel-model/issues/78)) ([d666ac4](https://github.com/neuralsignal/excel-model/commit/d666ac45a1fb3f1d9778d188ed9855d410e5afe4))
* resolve merge conflicts with main (PR [#79](https://github.com/neuralsignal/excel-model/issues/79) refactoring) ([d4c31b1](https://github.com/neuralsignal/excel-model/commit/d4c31b1c904c00712ece09fe7c0751d4a5c5df7d))


### Documentation

* sync documentation with codebase ([d243529](https://github.com/neuralsignal/excel-model/commit/d2435299764d38276afd3d7c61db2b4bf52f1315))
* sync documentation with codebase ([6c91c08](https://github.com/neuralsignal/excel-model/commit/6c91c088f780869c5216441d894559cac801ef55))
* sync documentation with codebase changes ([077d624](https://github.com/neuralsignal/excel-model/commit/077d624ad2316484d430fe51d18990f5f1f99eb4))
* sync documentation with codebase changes ([1b654b9](https://github.com/neuralsignal/excel-model/commit/1b654b9e1008877fbae80c1037daadcf1827d5a1))

## [0.1.2](https://github.com/neuralsignal/excel-model/compare/v0.1.1...v0.1.2) (2026-03-20)


### Bug Fixes

* merge main, fix DRY violation, and update docs for formula injection ([8347faa](https://github.com/neuralsignal/excel-model/commit/8347faaa72bb8e439d778cc71cbc39aba450b041))
* reject dangerous patterns in custom Excel formulas ([5786891](https://github.com/neuralsignal/excel-model/commit/5786891164bd5dc8ab5ec62c90e0e47ee4240d14))
* reject dangerous patterns in custom Excel formulas ([78d64c0](https://github.com/neuralsignal/excel-model/commit/78d64c06e9ce8216d6cc07d1b0264efa2119bbf5)), closes [#28](https://github.com/neuralsignal/excel-model/issues/28)

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
