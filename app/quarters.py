from __future__ import annotations

import datetime as dt
import re
from typing import Iterable, Tuple

QUARTER_PATTERN = re.compile(r"^\s*(\d{4})\s*[Qq]\s*([1-4])\s*$")


def parse_quarter(value: str) -> Tuple[int, int]:
    match = QUARTER_PATTERN.match(value)
    if not match:
        raise ValueError(f"Invalid quarter value: {value}")
    year = int(match.group(1))
    quarter = int(match.group(2))
    return year, quarter


def quarter_to_str(year: int, quarter: int) -> str:
    if quarter < 1 or quarter > 4:
        raise ValueError("Quarter must be in [1, 4].")
    return f"{year}Q{quarter}"


def quarter_sort_key(value: str) -> int:
    year, quarter = parse_quarter(value)
    return year * 10 + quarter


def shift_quarter(value: str, offset: int) -> str:
    year, quarter = parse_quarter(value)
    base = year * 4 + (quarter - 1)
    shifted = base + offset
    shifted_year = shifted // 4
    shifted_quarter = shifted % 4 + 1
    return quarter_to_str(shifted_year, shifted_quarter)


def quarter_from_date(date_value: dt.date) -> str:
    quarter = (date_value.month - 1) // 3 + 1
    return quarter_to_str(date_value.year, quarter)


def latest_completed_quarter(today: dt.date | None = None) -> str:
    if today is None:
        today = dt.date.today()
    current_quarter = quarter_from_date(today)
    return shift_quarter(current_quarter, -1)


def normalize_quarters(values: Iterable[str]) -> list[str]:
    cleaned: list[str] = []
    seen: set[str] = set()
    for raw in values:
        text = raw.strip()
        if not text:
            continue
        year, quarter = parse_quarter(text)
        normalized = quarter_to_str(year, quarter)
        if normalized in seen:
            continue
        seen.add(normalized)
        cleaned.append(normalized)
    return sorted(cleaned, key=quarter_sort_key)

