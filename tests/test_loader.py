"""Tests for loader.py."""
import json

import polars as pl
import pytest
import yaml as pyyaml

from excel_model.loader import InputData, _load_markdown_table, load


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
        md_file.write_text(
            "| period | revenue |\n"
            "|----|----|\n"
            "| 2023 | 1000 |\n"
        )
        result = load(str(md_file), "period", ["revenue"], "")
        assert len(result.df) == 1
