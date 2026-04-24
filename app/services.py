from __future__ import annotations

from datetime import datetime
from typing import Callable

from .constants import CompanyProfile
from .database import Database
from .estimators import estimate_missing_eps
from .quarters import latest_completed_quarter, normalize_quarters, quarter_sort_key
from .providers.base import DataProvider

ProgressCallback = Callable[[int, int, str], None]


class FinancialCollectorService:
    def __init__(
        self,
        db: Database,
        provider: DataProvider,
        company_profiles: list[CompanyProfile],
    ) -> None:
        self.db = db
        self.provider = provider
        self.company_profiles = {item.company_id: item for item in company_profiles}

    def collect(
        self,
        company_ids: list[str],
        target_quarters: list[str],
        include_latest: bool = True,
        on_progress: ProgressCallback | None = None,
    ) -> dict:
        normalized_targets = normalize_quarters(target_quarters)
        summary = {
            "targets": normalized_targets,
            "success": [],
            "failed": [],
        }

        total = len(company_ids)
        for idx, company_id in enumerate(company_ids, start=1):
            profile = self.company_profiles.get(company_id)
            if profile is None:
                summary["failed"].append({"company_id": company_id, "reason": "unknown company"})
                continue

            try:
                pulled = self.provider.fetch_financials(profile)
                pulled_map = {item["quarter"]: item for item in pulled}

                desired = set(normalized_targets)
                if include_latest:
                    desired.add(latest_completed_quarter())
                    if pulled_map:
                        latest_from_provider = max(pulled_map, key=quarter_sort_key)
                        desired.add(latest_from_provider)

                desired_list = sorted(desired, key=quarter_sort_key)
                now_text = datetime.now().isoformat(timespec="seconds")
                upsert_rows = []
                for quarter in desired_list:
                    if quarter in pulled_map:
                        row = pulled_map[quarter]
                    else:
                        row = {
                            "company_id": profile.company_id,
                            "company_name": profile.name,
                            "quarter": quarter,
                            "revenue": None,
                            "net_income": None,
                            "eps_reported": None,
                            "eps_estimated": None,
                            "source": "manual_required",
                            "fetched_at": now_text,
                        }
                    upsert_rows.append(row)

                self.db.upsert_records(upsert_rows)
                summary["success"].append(
                    {
                        "company_id": profile.company_id,
                        "quarter_count": len(upsert_rows),
                    }
                )

                if on_progress:
                    on_progress(idx, total, f"{profile.company_id} completed ({len(upsert_rows)} quarters)")
            except Exception as exc:
                summary["failed"].append({"company_id": profile.company_id, "reason": str(exc)})
                if on_progress:
                    on_progress(idx, total, f"{profile.company_id} failed: {exc}")

        return summary

    def auto_estimate(self, company_id: str) -> dict[str, float]:
        records = self.db.fetch_company_records(company_id)
        estimates = estimate_missing_eps(records)
        for quarter, value in estimates.items():
            self.db.update_estimated_eps(company_id=company_id, quarter=quarter, eps_estimated=value)
        return estimates

    def update_manual_eps(self, company_id: str, quarter: str, eps_value: float) -> None:
        self.db.update_estimated_eps(company_id=company_id, quarter=quarter, eps_estimated=eps_value)

