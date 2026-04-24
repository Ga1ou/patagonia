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

    # 台股 Basic EPS 欄位名稱（yfinance 有時用不同 key）
    EPS_KEYS = [
        "Basic EPS",
        "Diluted EPS",
        "EPS",
    ]

    def fetch_financials(self, profile: CompanyProfile) -> list[dict]:
        ticker = yf.Ticker(profile.yahoo_ticker)

        income_stmt = self._safe_quarterly_income_stmt(ticker)
        shares_outstanding = self._safe_shares_outstanding(ticker)

        by_quarter: dict[str, dict[str, Any]] = {}
        source_name = "yahoo_finance"
        now_text = datetime.now().isoformat(timespec="seconds")

        if income_stmt is not None and not income_stmt.empty:
            for column in income_stmt.columns:
                quarter = self._column_to_quarter(column)
                if quarter is None:
                    continue

                revenue = self._extract_metric(income_stmt, self.REVENUE_KEYS, column)
                net_income = self._extract_metric(income_stmt, self.NET_INCOME_KEYS, column)

                # 優先從 income_stmt 讀 Basic EPS
                eps_reported = self._extract_metric(income_stmt, self.EPS_KEYS, column)

                # Fallback：淨利 ÷ 股數（台股單位：元/股）
                if eps_reported is None and net_income is not None and shares_outstanding:
                    eps_reported = round(net_income / shares_outstanding, 2)

                by_quarter[quarter] = {
                    "company_id": profile.company_id,
                    "company_name": profile.name,
                    "quarter": quarter,
                    "revenue": revenue,
                    "net_income": net_income,
                    "eps_reported": eps_reported,
                    "eps_estimated": None,
                    "pe_ratio": None,
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

    def _safe_shares_outstanding(self, ticker: yf.Ticker) -> float | None:
        """
        抓流通股數，用於 EPS fallback 計算。
        優先用 info['sharesOutstanding']，再試 fast_info。
        台股單位是「股」，需除以 1 (yfinance 回傳已是股數)。
        """
        try:
            info = ticker.info or {}
            shares = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
            if shares:
                return float(shares)
        except Exception:
            pass

        try:
            shares = ticker.fast_info.get("shares")
            if shares:
                return float(shares)
        except Exception:
            pass

        return None

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
                val = self._to_float(df.at[key, column])
                if val is not None:
                    return val

        lowered = {str(idx).lower(): idx for idx in df.index}
        for key in keys:
            candidate = lowered.get(key.lower())
            if candidate is not None:
                val = self._to_float(df.at[candidate, column])
                if val is not None:
                    return val

        return None

    def _to_float(self, value: Any) -> float | None:
        if value is None:
            return None
        try:
            f = float(value)
            return None if pd.isna(f) else f
        except (TypeError, ValueError):
            return None