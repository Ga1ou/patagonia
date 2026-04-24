from __future__ import annotations

from typing import Any

from .quarters import quarter_sort_key, shift_quarter


def _to_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _mean(values: list[float]) -> float | None:
    if not values:
        return None
    return sum(values) / len(values)


def estimate_eps(target_quarter: str, final_eps_by_quarter: dict[str, float | None]) -> float | None:
    prev_quarter = shift_quarter(target_quarter, -1)
    same_quarter_last_year = shift_quarter(target_quarter, -4)
    recent_quarters = [shift_quarter(target_quarter, -offset) for offset in range(1, 5)]

    prev_eps = final_eps_by_quarter.get(prev_quarter)
    yoy_eps = final_eps_by_quarter.get(same_quarter_last_year)
    recent_values = [final_eps_by_quarter.get(q) for q in recent_quarters]
    recent_mean = _mean([v for v in recent_values if v is not None])

    weighted_sum = 0.0
    active_weight = 0.0

    if prev_eps is not None:
        weighted_sum += 0.5 * prev_eps
        active_weight += 0.5
    if yoy_eps is not None:
        weighted_sum += 0.3 * yoy_eps
        active_weight += 0.3
    if recent_mean is not None:
        weighted_sum += 0.2 * recent_mean
        active_weight += 0.2

    if active_weight == 0.0:
        return None
    return round(weighted_sum / active_weight, 3)


def estimate_missing_eps(records: list[dict[str, Any]]) -> dict[str, float]:
    ordered = sorted(records, key=lambda item: quarter_sort_key(item["quarter"]))
    final_eps_by_quarter: dict[str, float | None] = {}
    estimates: dict[str, float] = {}

    for record in ordered:
        quarter = record["quarter"]
        reported = _to_float(record.get("eps_reported"))
        estimated = _to_float(record.get("eps_estimated"))
        final_eps = reported if reported is not None else estimated

        if final_eps is None:
            predicted = estimate_eps(quarter, final_eps_by_quarter)
            if predicted is not None:
                estimates[quarter] = predicted
                final_eps = predicted

        final_eps_by_quarter[quarter] = final_eps

    return estimates

