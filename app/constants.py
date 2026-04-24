from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class CompanyProfile:
    company_id: str
    name: str
    yahoo_ticker: str


TRACKED_COMPANIES = [
    CompanyProfile(company_id="8299", name="群聯", yahoo_ticker="8299.TWO"),
    CompanyProfile(company_id="2330", name="台積電", yahoo_ticker="2330.TW"),
    CompanyProfile(company_id="2454", name="聯發科", yahoo_ticker="2454.TW"),
]

DEFAULT_TARGET_QUARTERS = ["2025Q3", "2025Q4", "2026Q1"]

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_ROOT / "data"
EXPORT_DIR = PROJECT_ROOT / "exports"
DB_PATH = DATA_DIR / "financial_records.db"

APP_TITLE = "Taiwan Financial Console"

