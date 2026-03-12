"""Period generation, auto-detect granularity, label formatting."""

import re
from dataclasses import dataclass


@dataclass(frozen=True)
class Period:
    label: str  # Display label: "2025", "Q1 2025", "Jan 2025"
    index: int  # 0-based position among all periods (history + projection)
    is_history: bool


_MONTH_NAMES = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
]


def detect_granularity(start_period: str) -> str:
    """Auto-detect granularity from start_period format.

    '2025'              -> annual
    '2025-Q1' or 'Q1 2025' -> quarterly
    '2025-01' or 'Jan 2025' -> monthly
    """
    s = start_period.strip()
    # Annual: bare 4-digit year
    if re.fullmatch(r"\d{4}", s):
        return "annual"
    # Quarterly: YYYY-Qn or Qn YYYY
    if re.fullmatch(r"\d{4}-Q[1-4]", s) or re.fullmatch(r"Q[1-4]\s+\d{4}", s):
        return "quarterly"
    # Monthly: YYYY-MM or MonName YYYY
    if re.fullmatch(r"\d{4}-\d{2}", s):
        return "monthly"
    month_names_pattern = "|".join(_MONTH_NAMES)
    if re.fullmatch(rf"(?:{month_names_pattern})\s+\d{{4}}", s):
        return "monthly"
    raise ValueError(f"Cannot detect granularity from start_period: {start_period!r}")


def _parse_annual_year(start_period: str) -> int:
    return int(start_period.strip())


def _parse_quarterly(start_period: str) -> tuple[int, int]:
    """Return (year, quarter_1based)."""
    s = start_period.strip()
    m = re.fullmatch(r"(\d{4})-Q([1-4])", s)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.fullmatch(r"Q([1-4])\s+(\d{4})", s)
    if m:
        return int(m.group(2)), int(m.group(1))
    raise ValueError(f"Cannot parse quarterly period: {start_period!r}")


def _parse_monthly(start_period: str) -> tuple[int, int]:
    """Return (year, month_1based)."""
    s = start_period.strip()
    m = re.fullmatch(r"(\d{4})-(\d{2})", s)
    if m:
        return int(m.group(1)), int(m.group(2))
    for i, name in enumerate(_MONTH_NAMES, start=1):
        m2 = re.fullmatch(rf"{name}\s+(\d{{4}})", s)
        if m2:
            return int(m2.group(1)), i
    raise ValueError(f"Cannot parse monthly period: {start_period!r}")


def _annual_label(year: int, offset: int) -> str:
    return str(year + offset)


def _quarterly_label(year: int, quarter: int, offset: int) -> str:
    total_q = (year * 4 + (quarter - 1)) + offset
    y = total_q // 4
    q = (total_q % 4) + 1
    return f"Q{q} {y}"


def _monthly_label(year: int, month: int, offset: int) -> str:
    total_m = (year * 12 + (month - 1)) + offset
    y = total_m // 12
    mo = (total_m % 12) + 1
    return f"{_MONTH_NAMES[mo - 1]} {y}"


def generate_periods(
    start_period: str,
    n_periods: int,
    n_history: int,
    granularity: str,
) -> list[Period]:
    """Generate history + projection periods.

    History periods come before start_period.
    Projection periods start at start_period.
    Total output length = n_history + n_periods.

    granularity: monthly | quarterly | annual | auto
    """
    if granularity == "auto":
        granularity = detect_granularity(start_period)

    total = n_history + n_periods
    periods: list[Period] = []

    if granularity == "annual":
        base_year = _parse_annual_year(start_period)
        for i in range(total):
            offset = i - n_history
            label = _annual_label(base_year, offset)
            periods.append(Period(label=label, index=i, is_history=(i < n_history)))

    elif granularity == "quarterly":
        base_year, base_q = _parse_quarterly(start_period)
        for i in range(total):
            offset = i - n_history
            label = _quarterly_label(base_year, base_q, offset)
            periods.append(Period(label=label, index=i, is_history=(i < n_history)))

    elif granularity == "monthly":
        base_year, base_month = _parse_monthly(start_period)
        for i in range(total):
            offset = i - n_history
            label = _monthly_label(base_year, base_month, offset)
            periods.append(Period(label=label, index=i, is_history=(i < n_history)))

    else:
        raise ValueError(f"Unknown granularity: {granularity!r}")

    return periods
