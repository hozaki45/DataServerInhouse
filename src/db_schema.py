"""SQLite database schema definition and initialization."""

import sqlite3
from pathlib import Path

from src.config import DB_DIR, SQLITE_DB_PATH

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS derivative_prices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    trade_date TEXT NOT NULL,
    instrument_code TEXT NOT NULL,
    instrument_name TEXT,
    put_call TEXT,
    contract_month TEXT,
    strike_price REAL,
    settlement_price REAL,
    theoretical_price REAL,
    underlying_price REAL,
    volatility REAL,
    interest_rate REAL,
    days_to_expiry INTEGER,
    underlying_name TEXT,
    UNIQUE(trade_date, instrument_code)
);

CREATE INDEX IF NOT EXISTS idx_date_underlying
    ON derivative_prices(trade_date, underlying_name);

CREATE INDEX IF NOT EXISTS idx_underlying_month
    ON derivative_prices(underlying_name, contract_month);

CREATE TABLE IF NOT EXISTS import_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name TEXT NOT NULL,
    trade_date TEXT NOT NULL,
    record_count INTEGER NOT NULL,
    imported_at TEXT NOT NULL DEFAULT (datetime('now')),
    status TEXT NOT NULL DEFAULT 'success'
);
"""


def init_db(db_path: Path = SQLITE_DB_PATH) -> sqlite3.Connection:
    """Initialize the SQLite database and return a connection."""
    DB_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.executescript(SCHEMA_SQL)
    conn.commit()
    return conn
