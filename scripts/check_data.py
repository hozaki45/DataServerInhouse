"""Data verification and summary script.

Usage:
    uv run python scripts/check_data.py
"""

import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.query import get_power_futures, summary_by_underlying
from src.repository import get_repository


def main() -> None:
    repo = get_repository()

    try:
        # Show import log
        log = repo.get_import_log()
        print("=" * 60)
        print("IMPORT LOG")
        print("=" * 60)
        if not log:
            print("  No imports found. Run import_csv.py first.")
            return
        for entry in log:
            print(
                f"  {entry['trade_date']} | {entry['file_name']} | "
                f"{entry['record_count']} records | {entry['status']}"
            )

        # Get the most recent trade date
        latest_date = log[0]["trade_date"]
        print(f"\n{'=' * 60}")
        print(f"DATA SUMMARY FOR {latest_date}")
        print("=" * 60)

        # Summary by underlying
        summary = summary_by_underlying(repo, latest_date)
        total_records = sum(s["total"] for s in summary)
        print(f"\nTotal records: {total_records}")
        print(f"Underlying assets: {len(summary)}")
        print(f"\n{'Underlying':<35} {'Total':>6} {'FUT':>5} {'CAL':>6} {'PUT':>6}")
        print("-" * 60)
        for s in summary:
            print(
                f"{s['underlying_name']:<35} {s['total']:>6} "
                f"{s['fut']:>5} {s['cal']:>6} {s['put']:>6}"
            )

        # Power futures detail
        print(f"\n{'=' * 60}")
        print("POWER FUTURES DETAIL")
        print("=" * 60)
        power = get_power_futures(repo, latest_date)
        if not power:
            print("  No power futures data found.")
        else:
            # Show only FUT (not options)
            power_fut = [r for r in power if not r.get("put_call")]
            print(f"\nPower futures contracts: {len(power_fut)}")
            print(
                f"\n{'Name':<30} {'Month':>8} {'Settlement':>12} "
                f"{'Theoretical':>12} {'Days':>5}"
            )
            print("-" * 70)
            for r in power_fut:
                settle = r.get("settlement_price")
                theo = r.get("theoretical_price")
                days = r.get("days_to_expiry")
                print(
                    f"{r['instrument_name']:<30} {r.get('contract_month', '') or '':>8} "
                    f"{settle if settle is not None else '':>12} "
                    f"{theo if theo is not None else '':>12} "
                    f"{days if days is not None else '':>5}"
                )
    finally:
        repo.close()


if __name__ == "__main__":
    main()
