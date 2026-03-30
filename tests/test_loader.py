"""Tests for loader.py."""

import json
from unittest.mock import patch

import polars as pl
import pytest
import yaml as pyyaml
from hypothesis import given, settings
from hypothesis import strategies as st

from excel_model.loader import InputData, _load_df, _load_markdown_table, load


class TestLoad:
    def test_load_csv(self, tmp_path):
        csv_file = tmp_path / "data.csv"
        csv_file.write_text("period,revenue,opex\n2023,1000,500\n2024,1100,520\n")
        result = load(str(csv_file), "period", ["revenue", "opex"], "")
        assert isinstance(result, InputData)
        assert result.period_col == "period"
        assert result.value_cols == ["revenue", "opex"]
        assert len(result.df) == 2

    def test_load_json(self, tmp_path):
        json_file = tmp_path / "data.json"
        data = [{"period": "2023", "revenue": 1000}, {"period": "2024", "revenue": 1100}]
        json_file.write_text(json.dumps(data))
        result = load(str(json_file), "period", ["revenue"], "")
        assert len(result.df) == 2

    def test_load_yaml(self, tmp_path):
        yaml_file = tmp_path / "data.yaml"
        data = [{"period": "2023", "revenue": 1000}, {"period": "2024", "revenue": 1100}]
        with yaml_file.open("w") as f:
            pyyaml.dump(data, f)
        result = load(str(yaml_file), "period", ["revenue"], "")
        assert len(result.df) == 2

    def test_load_parquet(self, tmp_path):
        parquet_file = tmp_path / "data.parquet"
        df = pl.DataFrame({"period": ["2023", "2024"], "revenue": [1000, 1100]})
        df.write_parquet(str(parquet_file))
        result = load(str(parquet_file), "period", ["revenue"], "")
        assert len(result.df) == 2

    def test_missing_file_raises(self):
        with pytest.raises(FileNotFoundError):
            load("/nonexistent/file.csv", "period", ["revenue"], "")

    def test_missing_column_raises(self, tmp_path):
        csv_file = tmp_path / "data.csv"
        csv_file.write_text("period,other\n2023,100\n")
        with pytest.raises(ValueError, match="missing required columns"):
            load(str(csv_file), "period", ["revenue"], "")

    def test_unsupported_extension(self, tmp_path):
        f = tmp_path / "data.txt"
        f.write_text("period,revenue\n2023,100\n")
        with pytest.raises(ValueError, match="Unsupported file extension"):
            load(str(f), "period", ["revenue"], "")


class TestLoadMarkdownTable:
    def test_basic_table(self, tmp_path):
        md_file = tmp_path / "data.md"
        md_file.write_text(
            "# Data\n\n"
            "| period | revenue | opex |\n"
            "|--------|---------|------|\n"
            "| 2023   | 1000    | 500  |\n"
            "| 2024   | 1100    | 520  |\n"
        )
        df = _load_markdown_table(str(md_file))
        assert len(df) == 2
        assert "period" in df.columns
        assert "revenue" in df.columns

    def test_no_table_raises(self, tmp_path):
        md_file = tmp_path / "data.md"
        md_file.write_text("# Just a heading\n\nNo table here.\n")
        with pytest.raises(ValueError, match="No valid markdown table"):
            _load_markdown_table(str(md_file))

    def test_load_md_through_load(self, tmp_path):
        md_file = tmp_path / "data.md"
        md_file.write_text("| period | revenue |\n|----|----|\n| 2023 | 1000 |\n")
        result = load(str(md_file), "period", ["revenue"], "")
        assert len(result.df) == 1

    def test_table_terminated_by_non_pipe_line(self, tmp_path):
        md_file = tmp_path / "data.md"
        md_file.write_text(
            "| period | revenue |\n|--------|--------|\n| 2023   | 1000   |\nSome trailing text\nMore text\n"
        )
        df = _load_markdown_table(str(md_file))
        assert len(df) == 1
        assert df["period"][0] == "2023"

    def test_malformed_row_raises(self, tmp_path):
        md_file = tmp_path / "data.md"
        md_file.write_text("| period | revenue | opex |\n|--------|---------|------|\n| 2023   | 1000    |\n")
        with pytest.raises(ValueError, match="Malformed markdown table row"):
            _load_markdown_table(str(md_file))

    def test_no_data_rows_raises(self, tmp_path):
        md_file = tmp_path / "data.md"
        md_file.write_text("| period | revenue |\n|--------|--------|\n")
        with pytest.raises(ValueError, match="no data rows"):
            _load_markdown_table(str(md_file))


class TestLoadXlsxWithSheet:
    def test_load_xlsx_with_sheet(self, tmp_path):
        xlsx_file = tmp_path / "data.xlsx"
        xlsx_file.write_bytes(b"")  # file must exist for the path check
        expected_df = pl.DataFrame({"period": ["2023"], "cost": [500]})
        with patch("excel_model.loader.pl.read_excel", return_value=expected_df) as mock_read:
            df = _load_df(str(xlsx_file), sheet="MySheet")
            mock_read.assert_called_once_with(str(xlsx_file), sheet_name="MySheet")
        assert "cost" in df.columns

    def test_load_xlsx_without_sheet(self, tmp_path):
        xlsx_file = tmp_path / "data.xlsx"
        xlsx_file.write_bytes(b"")
        expected_df = pl.DataFrame({"period": ["2023"], "revenue": [1000]})
        with patch("excel_model.loader.pl.read_excel", return_value=expected_df) as mock_read:
            _load_df(str(xlsx_file), sheet="")
            mock_read.assert_called_once_with(str(xlsx_file))


class TestLoadYamlNonList:
    def test_yaml_non_list_raises(self, tmp_path):
        f = tmp_path / "data.yaml"
        f.write_text("key: value\n")
        with pytest.raises(ValueError, match="list of records"):
            _load_df(str(f), sheet="")


# --- Property-based tests for _load_markdown_table ---

# Strategy: generate valid markdown tables
_header_st = st.text(
    alphabet=st.characters(whitelist_categories=("L", "N"), whitelist_characters="_"),
    min_size=1,
    max_size=10,
)


def _build_md_table(headers: list[str], rows: list[list[str]]) -> str:
    """Build a markdown pipe table string."""
    header_line = "| " + " | ".join(headers) + " |"
    sep_line = "| " + " | ".join("---" for _ in headers) + " |"
    data_lines = ["| " + " | ".join(row) + " |" for row in rows]
    return "\n".join([header_line, sep_line] + data_lines) + "\n"


_cell_st = st.text(
    alphabet=st.characters(whitelist_categories=("L", "N"), whitelist_characters="_ .-"),
    min_size=1,
    max_size=8,
).filter(lambda s: "|" not in s)


@given(
    headers=st.lists(_header_st, min_size=1, max_size=5, unique=True),
    data=st.data(),
)
@settings(max_examples=30)
def test_load_markdown_table_roundtrip(tmp_path_factory, headers, data):
    num_rows = data.draw(st.integers(min_value=1, max_value=5))
    rows = [data.draw(st.lists(_cell_st, min_size=len(headers), max_size=len(headers))) for _ in range(num_rows)]
    md = _build_md_table(headers, rows)
    md_file = tmp_path_factory.mktemp("md") / "table.md"
    md_file.write_text(md)
    df = _load_markdown_table(str(md_file))
    assert list(df.columns) == headers
    assert len(df) == num_rows


@given(
    headers=st.lists(_header_st, min_size=2, max_size=5, unique=True),
)
@settings(max_examples=20)
def test_load_markdown_table_no_data_rows_property(tmp_path_factory, headers):
    header_line = "| " + " | ".join(headers) + " |"
    sep_line = "| " + " | ".join("---" for _ in headers) + " |"
    md = header_line + "\n" + sep_line + "\n"
    md_file = tmp_path_factory.mktemp("md") / "table.md"
    md_file.write_text(md)
    with pytest.raises(ValueError, match="no data rows"):
        _load_markdown_table(str(md_file))
