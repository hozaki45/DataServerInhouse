"""Data loader module - orchestrates CSV reading and DB insertion."""

from __future__ import annotations

import hashlib

from src.csv_parser import parse_csv, parse_trade_date
from src.repository import Repository
from src.storage import StorageBackend


def compute_file_hash(data: bytes) -> str:
    """Compute SHA-256 hash of file contents."""
    return hashlib.sha256(data).hexdigest()


def import_csv_file(
    filename: str, storage: StorageBackend, repo: Repository
) -> tuple[str, int]:
    """Import a single CSV file into the database.

    Returns (trade_date, record_count) tuple.
    Raises DuplicateDataError if the same data has already been imported.
    """
    trade_date = parse_trade_date(filename)
    data = storage.read_file(filename)
    file_hash = compute_file_hash(data)

    if repo.hash_exists(file_hash):
        raise DuplicateDataError(
            f"Same data already imported (hash match). "
            f"File '{filename}' is likely a holiday duplicate."
        )

    records = parse_csv(data)
    inserted = repo.bulk_insert(trade_date, records)
    repo.log_import(filename, trade_date, inserted, file_hash=file_hash)
    return trade_date, inserted


class DuplicateDataError(Exception):
    """Raised when a CSV file has the same content as a previously imported file."""


def import_all_new(storage: StorageBackend, repo: Repository) -> list[tuple[str, int]]:
    """Import all CSV files that haven't been imported yet.

    Returns list of (trade_date, record_count) tuples.
    """
    imported_files = set(repo.get_imported_files())
    csv_files = storage.list_csv_files()
    results = []

    for filename in csv_files:
        if filename in imported_files:
            continue
        try:
            result = import_csv_file(filename, storage, repo)
            results.append(result)
        except DuplicateDataError as e:
            print(f"  SKIPPED {filename}: {e}")
        except Exception as e:
            trade_date = "unknown"
            try:
                trade_date = parse_trade_date(filename)
            except ValueError:
                pass
            repo.log_import(filename, trade_date, 0, status=f"error: {e}")
            print(f"  ERROR importing {filename}: {e}")

    return results
