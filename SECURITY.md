# Security

## Reporting a Vulnerability

Report suspected security issues by opening a private GitHub security advisory
on this repository. Do not disclose vulnerabilities in public issues.

## Tracked Advisories (no fix available)

Advisories in this list affect a dependency of this project but have no
upstream fix at the time of the most recent audit. They are tracked here so
that, once an upstream fix ships, the relevant version pin can be added to
`pixi.toml` and the entry removed from this list.

### CVE-2026-4539 — pygments 2.19.2 (ReDoS in `AdlLexer`)

- **Severity:** Low (local access required)
- **Package:** `pygments` 2.19.2 (transitive dev dependency via `pytest`,
  `mkdocs-material`)
- **Description:** Inefficient regular expression in `AdlLexer`
  (`pygments/lexers/archetype.py`) enables regular expression denial of
  service when syntax-highlighting Archetype Definition Language (ADL)
  files.
- **Project impact:** None. `excel-model` does not invoke `AdlLexer` or
  process ADL files; `pygments` is pulled in only by dev tooling.
- **Audit date:** 2026-03-30
- **Upstream tracker:** https://github.com/pygments/pygments/issues
- **Remediation once fixed:** pin `pygments >= <fixed-version>` in
  `pixi.toml` and remove this entry.
