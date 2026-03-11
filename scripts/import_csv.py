"""Import CSV files into the database.

Usage:
    uv run python scripts/import_csv.py              # Import all new files
    uv run python scripts/import_csv.py rb20260310.csv  # Import specific file
"""

import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.data_loader import import_all_new, import_csv_file
from src.repository import get_repository
from src.storage import get_storage


def main() -> None:
    storage = get_storage()
    repo = get_repository()

    try:
        if len(sys.argv) > 1:
            # Import specific file
            filename = sys.argv[1]
            print(f"Importing {filename}...")
            trade_date, count = import_csv_file(filename, storage, repo)
            print(f"  Done: {count} records imported for {trade_date}")
        else:
            # Import all new files
            csv_files = storage.list_csv_files()
            imported = set(repo.get_imported_files())
            pending = [f for f in csv_files if f not in imported]

            print(f"Found {len(csv_files)} CSV files, {len(pending)} new to import")

            if not pending:
                print("Nothing to import.")
                return

            results = import_all_new(storage, repo)
            total = sum(count for _, count in results)
            print(f"\nImported {len(results)} files, {total} total records")
            for trade_date, count in results:
                print(f"  {trade_date}: {count} records")
    finally:
        repo.close()


if __name__ == "__main__":
    main()
