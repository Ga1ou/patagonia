from __future__ import annotations

from datetime import date, datetime
from typing import Any

import pandas as pd
import yfinance as yf

from app.constants import CompanyProfile
from app.quarters import latest_completed_quarter, quarter_from_date, quarter_sort_key

from .base import DataProvider


class YahooFinanceProvider(DataProvider):
    REVENUE_KEYS = [
        "Total Revenue",
        "Operating Revenue",
        "Revenue",
    ]

    NET_INCOME_KEYS = [
        "Net Income",
        "Net Income Common Stockholders",
        "Net Income Including Noncontrolling Interests",
    ]

    def fetch_financials(self, profile: CompanyProfile) -> list[dict]:
        ticker = yf.Ticker(profile.yahoo_ticker)

        income_stmt = self._safe_quarterly_income_stmt(ticker)
        eps_by_quarter = self._safe_eps_map(ticker)

        by_quarter: dict[str, dict[str, Any]] = {}
        source_name = "yahoo_finance"
        now_text = datetime.now().isoformat(timespec="seconds")

        if income_stmt is not None and not income_stmt.empty:
            for column in income_stmt.columns:
                quarter = self._column_to_quarter(column)
                if quarter is None:
                    continue

                by_quarter[quarter] = {
                    "company_id": profile.company_id,
                    "company_name": profile.name,
                    "quarter": quarter,
                    "revenue": self._extract_metric(income_stmt, self.REVENUE_KEYS, column),
                    "net_income": self._extract_metric(income_stmt, self.NET_INCOME_KEYS, column),
                    "eps_reported": None,
                    "eps_estimated": None,
                    "source": source_name,
                    "fetched_at": now_text,
                }

        for quarter, eps in eps_by_quarter.items():
            if quarter in by_quarter:
                by_quarter[quarter]["eps_reported"] = eps
            else:
                by_quarter[quarter] = {
                    "company_id": profile.company_id,
                    "company_name": profile.name,
                    "quarter": quarter,
                    "revenue": None,
                    "net_income": None,
                    "eps_reported": eps,
                    "eps_estimated": None,
                    "source": source_name,
                    "fetched_at": now_text,
                }

        records = list(by_quarter.values())
        records.sort(key=lambda item: quarter_sort_key(item["quarter"]))
        return records

    def _safe_quarterly_income_stmt(self, ticker: yf.Ticker) -> pd.DataFrame | None:
        try:
            data = ticker.quarterly_income_stmt
            if isinstance(data, pd.DataFrame):
                return data
            return None
        except Exception:
            return None

    def _safe_eps_map(self, ticker: yf.Ticker) -> dict[str, float]:
        eps_by_quarter: dict[str, float] = {}
        try:
            earnings_dates = ticker.get_earnings_dates(limit=20)
        except Exception:
            earnings_dates = None

        if isinstance(earnings_dates, pd.DataFrame) and not earnings_dates.empty:
            for index, row in earnings_dates.iterrows():
                report_date = self._index_to_date(index)
                if report_date is None:
                    continue
                quarter = latest_completed_quarter(report_date)
                eps = self._to_float(
                    row.get("Reported EPS")
                    if "Reported EPS" in row
                    else row.get("EPS Actual")
                )
                if eps is not None and quarter not in eps_by_quarter:
                    eps_by_quarter[quarter] = eps

        return eps_by_quarter

    def _column_to_quarter(self, column: Any) -> str | None:
        if isinstance(column, pd.Timestamp):
            return quarter_from_date(column.date())
        if isinstance(column, datetime):
            return quarter_from_date(column.date())
        if isinstance(column, date):
            return quarter_from_date(column)
        if isinstance(column, str):
            try:
                parsed = pd.to_datetime(column)
                return quarter_from_date(parsed.date())
            except Exception:
                return None
        return None

    def _extract_metric(self, df: pd.DataFrame, keys: list[str], column: Any) -> float | None:
        for key in keys:
            if key in df.index:
                return self._to_float(df.at[key, column])

        lowered = {str(index).lower(): index for index in df.index}
        for key in keys:
            candidate = lowered.get(key.lower())
            if candidate is not None:
                return self._to_float(df.at[candidate, column])
        return None

    def _index_to_date(self, index: Any) -> date | None:
        if isinstance(index, pd.Timestamp):
            return index.date()
        if isinstance(index, datetime):
            return index.date()
        if isinstance(index, date):
            return index
        if isinstance(index, str):
            try:
                return pd.to_datetime(index).date()
            except Exception:
                return None
        return None

    def _to_float(self, value: Any) -> float | None:
        if value is None:
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

