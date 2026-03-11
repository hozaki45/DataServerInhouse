"""Data access abstraction layer.

Defines access patterns as methods, with SQLite implementation.
DynamoDB implementation can be added later without changing callers.
"""

from __future__ import annotations

import sqlite3
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Any

from src.config import DB_BACKEND, SQLITE_DB_PATH
from src.db_schema import init_db


class Repository(ABC):
    """Abstract repository defining access patterns for derivative price data."""

    @abstractmethod
    def bulk_insert(self, trade_date: str, records: list[dict[str, Any]]) -> int:
        """Insert multiple records for a given trade date. Returns count inserted."""

    @abstractmethod
    def get_by_date(self, trade_date: str) -> list[dict[str, Any]]:
        """Get all records for a given trade date."""

    @abstractmethod
    def get_by_date_and_underlying(
        self, trade_date: str, underlying_name: str
    ) -> list[dict[str, Any]]:
        """Get records for a given date and underlying asset name."""

    @abstractmethod
    def get_instrument_history(
        self, instrument_code: str, date_from: str | None = None, date_to: str | None = None
    ) -> list[dict[str, Any]]:
        """Get time series data for a specific instrument."""

    @abstractmethod
    def get_underlying_names(self) -> list[str]:
        """Get list of all unique underlying asset names."""

    @abstractmethod
    def log_import(
        self, file_name: str, trade_date: str, record_count: int, status: str = "success"
    ) -> None:
        """Record an import event."""

    @abstractmethod
    def get_imported_files(self) -> list[str]:
        """Get list of already imported file names."""

    @abstractmethod
    def get_import_log(self) -> list[dict[str, Any]]:
        """Get full import log."""

    @abstractmethod
    def close(self) -> None:
        """Close the connection."""


class SQLiteRepository(Repository):
    """SQLite implementation of the repository."""

    def __init__(self) -> None:
        self.conn = init_db(SQLITE_DB_PATH)
        self.conn.row_factory = sqlite3.Row

    def bulk_insert(self, trade_date: str, records: list[dict[str, Any]]) -> int:
        cursor = self.conn.cursor()
        inserted = 0
        for record in records:
            try:
                cursor.execute(
                    """INSERT OR IGNORE INTO derivative_prices
                    (trade_date, instrument_code, instrument_name, put_call,
                     contract_month, strike_price, settlement_price,
                     theoretical_price, underlying_price, volatility,
                     interest_rate, days_to_expiry, underlying_name)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        trade_date,
                        record["instrument_code"],
                        record.get("instrument_name"),
                        record.get("put_call"),
                        record.get("contract_month"),
                        record.get("strike_price"),
                        record.get("settlement_price"),
                        record.get("theoretical_price"),
                        record.get("underlying_price"),
                        record.get("volatility"),
                        record.get("interest_rate"),
                        record.get("days_to_expiry"),
                        record.get("underlying_name"),
                    ),
                )
                if cursor.rowcount > 0:
                    inserted += 1
            except sqlite3.Error:
                continue
        self.conn.commit()
        return inserted

    def get_by_date(self, trade_date: str) -> list[dict[str, Any]]:
        cursor = self.conn.execute(
            "SELECT * FROM derivative_prices WHERE trade_date = ? ORDER BY instrument_code",
            (trade_date,),
        )
        return [dict(row) for row in cursor.fetchall()]

    def get_by_date_and_underlying(
        self, trade_date: str, underlying_name: str
    ) -> list[dict[str, Any]]:
        cursor = self.conn.execute(
            """SELECT * FROM derivative_prices
            WHERE trade_date = ? AND underlying_name = ?
            ORDER BY instrument_code""",
            (trade_date, underlying_name),
        )
        return [dict(row) for row in cursor.fetchall()]

    def get_instrument_history(
        self, instrument_code: str, date_from: str | None = None, date_to: str | None = None
    ) -> list[dict[str, Any]]:
        query = "SELECT * FROM derivative_prices WHERE instrument_code = ?"
        params: list[str] = [instrument_code]
        if date_from:
            query += " AND trade_date >= ?"
            params.append(date_from)
        if date_to:
            query += " AND trade_date <= ?"
            params.append(date_to)
        query += " ORDER BY trade_date"
        cursor = self.conn.execute(query, params)
        return [dict(row) for row in cursor.fetchall()]

    def get_underlying_names(self) -> list[str]:
        cursor = self.conn.execute(
            "SELECT DISTINCT underlying_name FROM derivative_prices ORDER BY underlying_name"
        )
        return [row[0] for row in cursor.fetchall() if row[0]]

    def log_import(
        self, file_name: str, trade_date: str, record_count: int, status: str = "success"
    ) -> None:
        self.conn.execute(
            """INSERT INTO import_log (file_name, trade_date, record_count, imported_at, status)
            VALUES (?, ?, ?, ?, ?)""",
            (file_name, trade_date, record_count, datetime.now().isoformat(), status),
        )
        self.conn.commit()

    def get_imported_files(self) -> list[str]:
        cursor = self.conn.execute(
            "SELECT file_name FROM import_log WHERE status = 'success'"
        )
        return [row[0] for row in cursor.fetchall()]

    def get_import_log(self) -> list[dict[str, Any]]:
        cursor = self.conn.execute("SELECT * FROM import_log ORDER BY imported_at DESC")
        return [dict(row) for row in cursor.fetchall()]

    def close(self) -> None:
        self.conn.close()


def get_repository() -> Repository:
    """Factory function to get the configured repository backend."""
    if DB_BACKEND == "sqlite":
        return SQLiteRepository()
    elif DB_BACKEND == "dynamodb":
        raise NotImplementedError("DynamoDB repository not yet implemented")
    else:
        raise ValueError(f"Unknown DB_BACKEND: {DB_BACKEND}")
