"""Data loader module - orchestrates CSV reading and DB insertion."""

from __future__ import annotations

from src.csv_parser import parse_csv, parse_trade_date
from src.repository import Repository
from src.storage import StorageBackend


def import_csv_file(
    filename: str, storage: StorageBackend, repo: Repository
) -> tuple[str, int]:
    """Import a single CSV file into the database.

    Returns (trade_date, record_count) tuple.
    """
    trade_date = parse_trade_date(filename)
    data = storage.read_file(filename)
    records = parse_csv(data)
    inserted = repo.bulk_insert(trade_date, records)
    repo.log_import(filename, trade_date, inserted)
    return trade_date, inserted


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
        except Exception as e:
            trade_date = "unknown"
            try:
                trade_date = parse_trade_date(filename)
            except ValueError:
                pass
            repo.log_import(filename, trade_date, 0, status=f"error: {e}")
            print(f"  ERROR importing {filename}: {e}")

    return results
