from __future__ import annotations

import csv
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Sequence

from .constants import DB_PATH, DATA_DIR, EXPORT_DIR

RECORD_COLUMNS = [
    "company_id",
    "company_name",
    "quarter",
    "revenue",
    "net_income",
    "eps_reported",
    "eps_estimated",
    "pe_ratio",
    "source",
    "fetched_at",
]


class Database:
    def __init__(self, db_path: str | Path = DB_PATH) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(self.db_path))
        self.conn.row_factory = sqlite3.Row
        self._init_schema()

    def _init_schema(self) -> None:
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS financial_records (
                company_id TEXT NOT NULL,
                company_name TEXT NOT NULL,
                quarter TEXT NOT NULL,
                revenue REAL,
                net_income REAL,
                eps_reported REAL,
                eps_estimated REAL,
                pe_ratio REAL,
                source TEXT NOT NULL DEFAULT 'unknown',
                fetched_at TEXT NOT NULL,
                PRIMARY KEY (company_id, quarter)
            );

            CREATE INDEX IF NOT EXISTS idx_records_company ON financial_records(company_id);
            CREATE INDEX IF NOT EXISTS idx_records_quarter ON financial_records(quarter);
            """
        )
        self._ensure_column("financial_records", "pe_ratio", "REAL")
        self.conn.commit()

    def _ensure_column(self, table_name: str, column_name: str, column_type: str) -> None:
        rows = self.conn.execute(f"PRAGMA table_info({table_name})").fetchall()
        existing_columns = {row["name"] for row in rows}
        if column_name in existing_columns:
            return
        self.conn.execute(
            f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}"
        )

    def upsert_records(self, records: Iterable[dict[str, Any]]) -> int:
        now_text = datetime.now().isoformat(timespec="seconds")
        rows: list[tuple[Any, ...]] = []
        for record in records:
            rows.append(
                (
                    record.get("company_id"),
                    record.get("company_name", ""),
                    record.get("quarter"),
                    record.get("revenue"),
                    record.get("net_income"),
                    record.get("eps_reported"),
                    record.get("eps_estimated"),
                    record.get("pe_ratio"),
                    record.get("source", "unknown"),
                    record.get("fetched_at", now_text),
                )
            )
        if not rows:
            return 0

        self.conn.executemany(
            """
            INSERT INTO financial_records (
                company_id,
                company_name,
                quarter,
                revenue,
                net_income,
                eps_reported,
                eps_estimated,
                pe_ratio,
                source,
                fetched_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(company_id, quarter) DO UPDATE SET
                company_name=excluded.company_name,
                revenue=excluded.revenue,
                net_income=excluded.net_income,
                eps_reported=excluded.eps_reported,
                eps_estimated=COALESCE(financial_records.eps_estimated, excluded.eps_estimated),
                pe_ratio=COALESCE(excluded.pe_ratio, financial_records.pe_ratio),
                source=excluded.source,
                fetched_at=excluded.fetched_at;
            """,
            rows,
        )
        self.conn.commit()
        return len(rows)

    def fetch_records(
        self,
        company_ids: Sequence[str] | None = None,
        quarters: Sequence[str] | None = None,
    ) -> list[dict[str, Any]]:
        query = f"SELECT {', '.join(RECORD_COLUMNS)} FROM financial_records WHERE 1=1"
        params: list[Any] = []

        if company_ids:
            placeholders = ", ".join("?" for _ in company_ids)
            query += f" AND company_id IN ({placeholders})"
            params.extend(company_ids)

        if quarters:
            placeholders = ", ".join("?" for _ in quarters)
            query += f" AND quarter IN ({placeholders})"
            params.extend(quarters)

        query += " ORDER BY company_id ASC, quarter ASC"
        rows = self.conn.execute(query, params).fetchall()
        return [dict(row) for row in rows]

    def fetch_company_records(self, company_id: str) -> list[dict[str, Any]]:
        rows = self.conn.execute(
            f"""
            SELECT {', '.join(RECORD_COLUMNS)}
            FROM financial_records
            WHERE company_id = ?
            ORDER BY quarter ASC
            """,
            (company_id,),
        ).fetchall()
        return [dict(row) for row in rows]

    def update_estimated_eps(self, company_id: str, quarter: str, eps_estimated: float) -> None:
        now_text = datetime.now().isoformat(timespec="seconds")
        cursor = self.conn.execute(
            """
            UPDATE financial_records
            SET eps_estimated = ?, fetched_at = ?
            WHERE company_id = ? AND quarter = ?
            """,
            (eps_estimated, now_text, company_id, quarter),
        )
        if cursor.rowcount == 0:
            self.conn.execute(
                """
                INSERT INTO financial_records (
                    company_id, company_name, quarter, source, fetched_at, eps_estimated
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (company_id, company_id, quarter, "manual_input", now_text, eps_estimated),
            )
        self.conn.commit()

    def update_pe_ratio(self, company_id: str, quarter: str, pe_ratio: float) -> None:
        now_text = datetime.now().isoformat(timespec="seconds")
        cursor = self.conn.execute(
            """
            UPDATE financial_records
            SET pe_ratio = ?, fetched_at = ?
            WHERE company_id = ? AND quarter = ?
            """,
            (pe_ratio, now_text, company_id, quarter),
        )
        if cursor.rowcount == 0:
            self.conn.execute(
                """
                INSERT INTO financial_records (
                    company_id, company_name, quarter, source, fetched_at, pe_ratio
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (company_id, company_id, quarter, "manual_input", now_text, pe_ratio),
            )
        self.conn.commit()

    def export_csv(
        self,
        export_path: str | Path,
        company_ids: Sequence[str] | None = None,
        quarters: Sequence[str] | None = None,
    ) -> Path:
        rows = self.fetch_records(company_ids=company_ids, quarters=quarters)
        path = Path(export_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", newline="", encoding="utf-8-sig") as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=RECORD_COLUMNS)
            writer.writeheader()
            writer.writerows(rows)
        return path

    def count_records(self) -> int:
        row = self.conn.execute("SELECT COUNT(*) AS count FROM financial_records").fetchone()
        return int(row["count"])

    def latest_sync_time(self) -> str | None:
        row = self.conn.execute("SELECT MAX(fetched_at) AS latest FROM financial_records").fetchone()
        latest = row["latest"]
        return str(latest) if latest else None

    def close(self) -> None:
        self.conn.close()
