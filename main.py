from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from app.constants import TRACKED_COMPANIES
from app.database import Database
from app.providers.yahoo_provider import YahooFinanceProvider
from app.services import FinancialCollectorService
from app.ui import MainWindow


def main() -> None:
    app = QApplication(sys.argv)

    database = Database()
    provider = YahooFinanceProvider()
    collector = FinancialCollectorService(
        db=database,
        provider=provider,
        company_profiles=TRACKED_COMPANIES,
    )

    window = MainWindow(db=database, collector=collector)
    window.show()

    exit_code = app.exec()
    database.close()
    raise SystemExit(exit_code)


if __name__ == "__main__":
    main()

