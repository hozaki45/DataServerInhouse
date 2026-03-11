"""Query utilities for data analysis and reporting."""

from __future__ import annotations

from typing import Any

from src.repository import Repository


def summary_by_underlying(repo: Repository, trade_date: str) -> list[dict[str, Any]]:
    """Get summary statistics grouped by underlying asset for a given date."""
    records = repo.get_by_date(trade_date)

    stats: dict[str, dict[str, Any]] = {}
    for r in records:
        name = r.get("underlying_name") or "unknown"
        if name not in stats:
            stats[name] = {"underlying_name": name, "total": 0, "fut": 0, "cal": 0, "put": 0}
        stats[name]["total"] += 1
        pc = r.get("put_call")
        if pc == "CAL":
            stats[name]["cal"] += 1
        elif pc == "PUT":
            stats[name]["put"] += 1
        else:
            stats[name]["fut"] += 1

    return sorted(stats.values(), key=lambda x: x["underlying_name"])


def get_power_futures(repo: Repository, trade_date: str) -> list[dict[str, Any]]:
    """Get all power-related futures for a given date."""
    power_names = [
        name for name in repo.get_underlying_names()
        if "電力" in name
    ]
    results = []
    for name in power_names:
        records = repo.get_by_date_and_underlying(trade_date, name)
        results.extend(records)
    return results
