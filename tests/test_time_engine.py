"""Tests for time_engine.py."""
import pytest

from excel_model.time_engine import detect_granularity, generate_periods


class TestDetectGranularity:
    def test_annual(self):
        assert detect_granularity("2025") == "annual"

    def test_quarterly_dash(self):
        assert detect_granularity("2025-Q1") == "quarterly"

    def test_quarterly_space(self):
        assert detect_granularity("Q3 2025") == "quarterly"

    def test_monthly_dash(self):
        assert detect_granularity("2025-01") == "monthly"

    def test_monthly_name(self):
        assert detect_granularity("Jan 2025") == "monthly"

    def test_monthly_name_dec(self):
        assert detect_granularity("Dec 2025") == "monthly"

    def test_invalid(self):
        with pytest.raises(ValueError, match="Cannot detect granularity"):
            detect_granularity("2025/01")


class TestGeneratePeriodsAnnual:
    def test_basic(self):
        periods = generate_periods("2025", 3, 2, "annual")
        assert len(periods) == 5
        labels = [p.label for p in periods]
        assert labels == ["2023", "2024", "2025", "2026", "2027"]

    def test_history_flags(self):
        periods = generate_periods("2025", 3, 2, "annual")
        assert periods[0].is_history is True
        assert periods[1].is_history is True
        assert periods[2].is_history is False
        assert periods[4].is_history is False

    def test_no_history(self):
        periods = generate_periods("2025", 3, 0, "annual")
        assert len(periods) == 3
        assert all(not p.is_history for p in periods)
        assert [p.label for p in periods] == ["2025", "2026", "2027"]

    def test_indices(self):
        periods = generate_periods("2025", 3, 2, "annual")
        for i, p in enumerate(periods):
            assert p.index == i


class TestGeneratePeriodsQuarterly:
    def test_basic(self):
        periods = generate_periods("2025-Q1", 4, 0, "quarterly")
        assert len(periods) == 4
        assert [p.label for p in periods] == ["Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025"]

    def test_year_rollover(self):
        periods = generate_periods("2025-Q3", 3, 0, "quarterly")
        assert [p.label for p in periods] == ["Q3 2025", "Q4 2025", "Q1 2026"]

    def test_q_space_format(self):
        periods = generate_periods("Q1 2025", 2, 0, "quarterly")
        assert len(periods) == 2
        assert periods[0].label == "Q1 2025"


class TestGeneratePeriodsMonthly:
    def test_basic(self):
        periods = generate_periods("2025-01", 3, 0, "monthly")
        assert len(periods) == 3
        assert [p.label for p in periods] == ["Jan 2025", "Feb 2025", "Mar 2025"]

    def test_year_rollover(self):
        periods = generate_periods("2025-11", 3, 0, "monthly")
        assert [p.label for p in periods] == ["Nov 2025", "Dec 2025", "Jan 2026"]

    def test_month_name_format(self):
        periods = generate_periods("Jan 2025", 2, 0, "monthly")
        assert periods[0].label == "Jan 2025"
        assert periods[1].label == "Feb 2025"

    def test_with_history(self):
        periods = generate_periods("2025-03", 3, 2, "monthly")
        assert len(periods) == 5
        assert periods[0].label == "Jan 2025"
        assert periods[1].label == "Feb 2025"
        assert periods[2].label == "Mar 2025"
        assert periods[0].is_history is True
        assert periods[2].is_history is False


class TestGeneratePeriodsAuto:
    def test_auto_annual(self):
        periods = generate_periods("2025", 2, 0, "auto")
        assert all(p.label.isdigit() for p in periods)

    def test_auto_quarterly(self):
        periods = generate_periods("2025-Q2", 2, 0, "auto")
        assert "Q" in periods[0].label

    def test_auto_monthly(self):
        periods = generate_periods("2025-06", 2, 0, "auto")
        assert "2025" in periods[0].label


def test_invalid_granularity():
    with pytest.raises(ValueError, match="Unknown granularity"):
        generate_periods("2025", 3, 0, "weekly")
