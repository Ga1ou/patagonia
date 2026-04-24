from __future__ import annotations

from abc import ABC, abstractmethod

from app.constants import CompanyProfile


class DataProvider(ABC):
    @abstractmethod
    def fetch_financials(self, profile: CompanyProfile) -> list[dict]:
        """Return quarterly records for the company."""

