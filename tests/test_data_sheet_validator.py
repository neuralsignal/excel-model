"""Tests for data_sheet_validator.py — uncovered branches and property-based tests."""

from __future__ import annotations

from hypothesis import given, settings
from hypothesis import strategies as st

from excel_model.data_sheet_validator import (
    _validate_column_letter,
    _validate_sheet_name_safety,
    validate_data_sheet_def,
    validate_sumifs_pivot_def,
)
from excel_model.spec import DataSheetDef, SumifsPivotDef


def _data_spec(**overrides) -> DataSheetDef:
    defaults = dict(
        sheet_name="TEST_SHEET",
        title="Test",
        headers=("A", "B"),
        col_widths=(15.0, 15.0),
        number_formats={},
        freeze_row=2,
    )
    defaults.update(overrides)
    return DataSheetDef(**defaults)


def _pivot_spec(**overrides) -> SumifsPivotDef:
    defaults = dict(
        sheet_name="SUMIFS_TEST",
        title="Test Pivot",
        row_label_headers=("Hub",),
        col_dim_values=(2023, 2024, 2025),
        data_sheet="TRANSACTIONS_LNFW",
        value_col="AO",
        row_filter_cols=("AM",),
        col_filter_col="AJ",
        append_total=True,
        append_yoy=False,
        col_widths=(28.0, 14.0, 14.0, 14.0, 16.0),
        number_format_data="#,##0",
        number_format_pct="0.0%",
        freeze_row=2,
    )
    defaults.update(overrides)
    return SumifsPivotDef(**defaults)


# ---------------------------------------------------------------------------
# _validate_column_letter — the except ValueError (lines 21-22) is defensive
# dead code: openpyxl.column_index_from_string accepts all ^[A-Z]{1,3}$
# strings without raising. We verify the regex gate catches bad input instead.
# ---------------------------------------------------------------------------


def test_column_letter_lowercase_rejected():
    errors = _validate_column_letter("ab", "test_col")
    assert any("not a valid Excel column letter" in e for e in errors)


def test_column_letter_digits_rejected():
    errors = _validate_column_letter("123", "test_col")
    assert any("not a valid Excel column letter" in e for e in errors)


def test_column_letter_four_chars_rejected():
    errors = _validate_column_letter("ABCD", "test_col")
    assert any("not a valid Excel column letter" in e for e in errors)


# ---------------------------------------------------------------------------
# _validate_sheet_name_safety — uncovered branches
# ---------------------------------------------------------------------------


def test_empty_sheet_name_rejected():
    errors = validate_data_sheet_def(_data_spec(sheet_name=""))
    assert any("must not be empty" in e for e in errors)


def test_sheet_name_too_long():
    errors = validate_data_sheet_def(_data_spec(sheet_name="A" * 32))
    assert any("<= 31 characters" in e for e in errors)


def test_sheet_name_exactly_31_chars_accepted():
    errors = validate_data_sheet_def(_data_spec(sheet_name="A" * 31))
    assert not any("<= 31 characters" in e for e in errors)


# ---------------------------------------------------------------------------
# validate_sumifs_pivot_def — uncovered branches
# ---------------------------------------------------------------------------


def test_sumifs_empty_row_label_headers():
    spec = _pivot_spec(
        row_label_headers=(),
        row_filter_cols=(),
        col_widths=(14.0, 14.0, 14.0, 16.0),
    )
    errors = validate_sumifs_pivot_def(spec)
    assert any("row_label_headers must not be empty" in e for e in errors)


def test_sumifs_negative_freeze_row():
    errors = validate_sumifs_pivot_def(_pivot_spec(freeze_row=-1))
    assert any("freeze_row must be >= 0" in e for e in errors)


# ---------------------------------------------------------------------------
# Property-based tests
# ---------------------------------------------------------------------------


@given(col=st.sampled_from(["A", "B", "Z", "AA", "AZ", "XF", "XFD"]))
def test_known_valid_column_letters_accepted(col: str) -> None:
    errors = _validate_column_letter(col, "test")
    assert errors == []


@given(col=st.from_regex(r"^[A-Z]{1,2}$", fullmatch=True))
@settings(max_examples=100)
def test_one_or_two_letter_columns_accepted(col: str) -> None:
    errors = _validate_column_letter(col, "test")
    assert errors == []


@given(col=st.from_regex(r"^[a-z0-9!@#$%^&*]+$", fullmatch=True))
@settings(max_examples=50)
def test_invalid_column_strings_rejected(col: str) -> None:
    errors = _validate_column_letter(col, "test")
    assert len(errors) > 0


@given(name=st.from_regex(r"^[A-Za-z0-9_ ()-]{1,31}$", fullmatch=True))
@settings(max_examples=100)
def test_safe_sheet_names_accepted(name: str) -> None:
    errors = _validate_sheet_name_safety(name, "test")
    assert errors == []


@given(name=st.text(min_size=32, max_size=60, alphabet=st.characters(whitelist_categories=("L", "N"))))
@settings(max_examples=50)
def test_long_sheet_names_rejected(name: str) -> None:
    errors = _validate_sheet_name_safety(name, "test")
    assert any("<= 31 characters" in e for e in errors)


@given(
    headers=st.lists(st.text(min_size=1, max_size=5, alphabet="ABCDE"), min_size=1, max_size=5),
    freeze_row=st.integers(min_value=0, max_value=100),
)
@settings(max_examples=50)
def test_valid_data_sheet_def_property(headers: list[str], freeze_row: int) -> None:
    col_widths = tuple(10.0 for _ in headers)
    spec = DataSheetDef(
        sheet_name="VALID",
        title="T",
        headers=tuple(headers),
        col_widths=col_widths,
        number_formats={},
        freeze_row=freeze_row,
    )
    errors = validate_data_sheet_def(spec)
    assert errors == []


@given(freeze_row=st.integers(min_value=-1000, max_value=-1))
@settings(max_examples=20)
def test_negative_freeze_row_always_rejected(freeze_row: int) -> None:
    errors = validate_data_sheet_def(_data_spec(freeze_row=freeze_row))
    assert any("freeze_row must be >= 0" in e for e in errors)
