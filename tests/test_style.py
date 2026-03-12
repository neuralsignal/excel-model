"""Tests for style.py."""

from dataclasses import FrozenInstanceError

import pytest
import yaml
from openpyxl import Workbook

from excel_model.style import apply_header_style, get_number_format, load_style_config

SAMPLE_STYLE_DATA = {
    "header_fill_hex": "1F3864",
    "header_font_color": "FFFFFF",
    "subtotal_fill_hex": "D6E4F0",
    "total_fill_hex": "AED6F1",
    "history_col_fill_hex": "F2F2F2",
    "section_header_fill_hex": "E8F4FD",
    "font_name": "Calibri",
    "font_size": 10,
    "number_format_currency": "#,##0",
    "number_format_percent": "0.0%",
    "number_format_integer": "#,##0",
    "number_format_number": "#,##0.00",
}


def test_style_config_creation(sample_style):
    assert sample_style.header_fill_hex == "1F3864"
    assert sample_style.font_name == "Calibri"
    assert sample_style.font_size == 10


def test_style_config_frozen(sample_style):
    with pytest.raises(FrozenInstanceError):
        sample_style.font_size = 12  # type: ignore


def test_load_style_config(tmp_path, sample_style):
    style_path = tmp_path / "style.yaml"
    with style_path.open("w") as f:
        yaml.dump(SAMPLE_STYLE_DATA, f)

    loaded = load_style_config(str(style_path))
    assert loaded.header_fill_hex == "1F3864"
    assert loaded.font_name == "Calibri"
    assert loaded.font_size == 10


def test_load_style_config_missing_file():
    with pytest.raises(FileNotFoundError):
        load_style_config("/nonexistent/path/style.yaml")


def test_load_style_config_missing_keys(tmp_path):
    incomplete = {"header_fill_hex": "FFFFFF"}
    style_path = tmp_path / "style.yaml"
    with style_path.open("w") as f:
        yaml.dump(incomplete, f)
    with pytest.raises(ValueError, match="missing required keys"):
        load_style_config(str(style_path))


def test_get_number_format_currency(sample_style):
    assert get_number_format("currency", sample_style) == "#,##0"


def test_get_number_format_percent(sample_style):
    assert get_number_format("percent", sample_style) == "0.0%"


def test_get_number_format_integer(sample_style):
    assert get_number_format("integer", sample_style) == "#,##0"


def test_get_number_format_number(sample_style):
    assert get_number_format("number", sample_style) == "#,##0.00"


def test_get_number_format_invalid(sample_style):
    with pytest.raises(ValueError, match="Unknown format_type"):
        get_number_format("invalid_format", sample_style)


def test_apply_header_style(sample_style):
    wb = Workbook()
    ws = wb.active
    cell = ws["A1"]
    apply_header_style(cell, sample_style)
    # Cell should have fill applied
    assert cell.fill is not None
    assert cell.font is not None
    assert cell.font.bold is True
