"""CSV parser for JPX derivative theoretical price data.

Handles CP932-encoded CSV files with the format:
- 3 header/comment rows to skip
- 12 columns: instrument_code, instrument_name, put_call, contract_month,
  strike_price, settlement_price, theoretical_price, underlying_price,
  volatility, interest_rate, days_to_expiry, underlying_name
"""

from __future__ import annotations

import csv
import io
import re
from typing import Any

from src.config import CSV_COLUMNS, CSV_ENCODING, CSV_HEADER_ROWS


def parse_trade_date(filename: str) -> str:
    """Extract trade date from filename like 'rb20260310.csv' -> '2026-03-10'."""
    match = re.match(r"rb(\d{4})(\d{2})(\d{2})\.csv", filename)
    if not match:
        raise ValueError(f"Cannot extract date from filename: {filename}")
    year, month, day = match.groups()
    return f"{year}-{month}-{day}"


def _to_float(value: str) -> float | None:
    """Convert string to float, returning None for empty strings."""
    value = value.strip()
    if not value:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def _to_int(value: str) -> int | None:
    """Convert string to int, returning None for empty strings."""
    value = value.strip()
    if not value:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def parse_csv(data: bytes) -> list[dict[str, Any]]:
    """Parse JPX derivative price CSV data (CP932 bytes) into list of dicts."""
    text = data.decode(CSV_ENCODING)
    reader = csv.reader(io.StringIO(text))

    # Skip header rows
    for _ in range(CSV_HEADER_ROWS):
        next(reader, None)

    records = []
    for row in reader:
        if len(row) < 12:
            continue

        # Skip empty rows
        if not row[0].strip():
            continue

        record: dict[str, Any] = {
            "instrument_code": row[0].strip(),
            "instrument_name": row[1].strip(),
            "put_call": row[2].strip() or None,
            "contract_month": row[3].strip() or None,
            "strike_price": _to_float(row[4]),
            "settlement_price": _to_float(row[5]),
            "theoretical_price": _to_float(row[6]),
            "underlying_price": _to_float(row[7]),
            "volatility": _to_float(row[8]),
            "interest_rate": _to_float(row[9]),
            "days_to_expiry": _to_int(row[10]),
            "underlying_name": row[11].strip() if len(row) > 11 else None,
        }
        records.append(record)

    return records
