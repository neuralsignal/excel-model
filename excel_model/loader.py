"""Multi-format data loader: xlsx/csv/parquet/json/yaml/md → InputData."""

import re
from dataclasses import dataclass
from pathlib import Path

import polars as pl
import yaml


@dataclass(frozen=True)
class InputData:
    df: pl.DataFrame  # Columns: period_col + value_cols
    period_col: str
    value_cols: list[str]


def _load_df(source_path: str, sheet: str) -> pl.DataFrame:
    """Load raw DataFrame from file based on extension."""
    path = Path(source_path)
    if not path.exists():
        raise FileNotFoundError(f"Input data file not found: {source_path}")

    suffix = path.suffix.lower()

    if suffix in (".xlsx", ".xls"):
        if sheet:
            return pl.read_excel(source_path, sheet_name=sheet)
        return pl.read_excel(source_path)

    if suffix == ".csv":
        return pl.read_csv(source_path)

    if suffix == ".parquet":
        return pl.read_parquet(source_path)

    if suffix == ".json":
        return pl.read_json(source_path)

    if suffix in (".yaml", ".yml"):
        with path.open() as f:
            data = yaml.safe_load(f)
        if isinstance(data, list):
            return pl.DataFrame(data)
        raise ValueError(f"YAML file must contain a list of records, got: {type(data)}")

    if suffix == ".md":
        return _load_markdown_table(source_path)

    raise ValueError(f"Unsupported file extension: {suffix!r}")


def _load_markdown_table(source_path: str) -> pl.DataFrame:
    """Parse the first markdown pipe table found in the file."""
    path = Path(source_path)
    text = path.read_text()

    lines = text.splitlines()
    table_lines: list[str] = []
    in_table = False

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("|"):
            in_table = True
            table_lines.append(stripped)
        elif in_table:
            break

    if len(table_lines) < 2:
        raise ValueError(f"No valid markdown table found in: {source_path}")

    # First line = header, second line = separator (|---|---|), rest = data
    header_line = table_lines[0]
    headers = [h.strip() for h in header_line.split("|") if h.strip()]

    data_lines = [line for line in table_lines[2:] if not re.fullmatch(r"[\s|:-]+", line)]
    rows: list[dict] = []
    for line in data_lines:
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) != len(headers):
            raise ValueError(
                f"Malformed markdown table row (expected {len(headers)} columns, got {len(cells)}): {line!r}"
            )
        rows.append(dict(zip(headers, cells, strict=True)))

    if not rows:
        raise ValueError(f"Markdown table has no data rows: {source_path}")

    return pl.DataFrame(rows)


def load(
    source_path: str,
    period_col: str,
    value_cols: list[str],
    sheet: str,
) -> InputData:
    """Load input data from file, validate required columns, return InputData.

    Fails fast if any required column is missing.
    """
    df = _load_df(source_path, sheet)
    required_cols = [period_col] + value_cols
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Input data missing required columns: {missing}. Available columns: {df.columns}")
    return InputData(df=df, period_col=period_col, value_cols=value_cols)
