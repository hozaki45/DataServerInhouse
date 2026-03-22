"""Commodity-focused query utilities for multi-asset analysis.

Extends the basic query.py pattern with commodity-specific queries
for forward curves, cross-asset snapshots, and time series.
"""

from __future__ import annotations

from typing import Any

from src.asset_taxonomy import (
    ASSET_TAXONOMY,
    CATEGORY_META,
    COMMODITY_CATEGORIES,
    get_category_for_underlying,
    get_display_name,
)
from src.repository import Repository


def get_commodity_futures(
    repo: Repository, trade_date: str, category: str | None = None
) -> list[dict[str, Any]]:
    """Get all commodity futures (non-option) for a given date, optionally filtered by category."""
    names = repo.get_underlying_names()
    results = []
    for name in names:
        cat = get_category_for_underlying(name)
        if cat is None or cat not in COMMODITY_CATEGORIES:
            continue
        if category and cat != category:
            continue
        records = repo.get_by_date_and_underlying(trade_date, name)
        futures = [r for r in records if not r.get("put_call")]
        results.extend(futures)
    return results


def get_commodity_forward_curve(
    repo: Repository, trade_date: str, underlying_name: str
) -> list[dict[str, Any]]:
    """Build a forward curve for any commodity underlying.

    Returns list of dicts with month, settlement, theoretical, days_to_expiry,
    sorted by contract_month.
    """
    records = repo.get_by_date_and_underlying(trade_date, underlying_name)
    futures = [r for r in records if not r.get("put_call")]
    curve = []
    for r in futures:
        curve.append({
            "month": r.get("contract_month", ""),
            "settlement": r.get("settlement_price"),
            "theoretical": r.get("theoretical_price"),
            "days": r.get("days_to_expiry"),
            "name": r.get("instrument_name", ""),
            "instrument_code": r.get("instrument_code", ""),
        })
    curve.sort(key=lambda x: x["month"])
    return curve


def get_front_month_price(
    repo: Repository, trade_date: str, underlying_name: str
) -> dict[str, Any] | None:
    """Get the front-month (nearest expiry) futures price for a given underlying."""
    curve = get_commodity_forward_curve(repo, trade_date, underlying_name)
    # Filter to only those with a settlement price and a valid month
    valid = [c for c in curve if c["settlement"] is not None and c["month"]]
    if not valid:
        return None
    return valid[0]  # Already sorted by month, first is front-month


def get_cross_commodity_snapshot(
    repo: Repository,
    trade_date: str,
    prev_date: str | None = None,
) -> list[dict[str, Any]]:
    """Get front-month prices across all commodity asset classes with DoD changes.

    Returns a list of dicts suitable for a commodity overview table/chart.
    """
    # Build previous day map if available
    prev_map: dict[str, float] = {}
    if prev_date:
        for name, info in ASSET_TAXONOMY.items():
            if info["category"] not in COMMODITY_CATEGORIES:
                continue
            if info["subcategory"] == "power":
                continue  # Power handled separately
            front = get_front_month_price(repo, prev_date, name)
            if front and front["settlement"] is not None:
                prev_map[name] = front["settlement"]

    snapshot = []
    for name, info in ASSET_TAXONOMY.items():
        if info["category"] not in COMMODITY_CATEGORIES:
            continue
        if info["subcategory"] == "power":
            continue  # Power already has dedicated section

        front = get_front_month_price(repo, trade_date, name)
        if not front or front["settlement"] is None:
            continue

        price = front["settlement"]
        prev_price = prev_map.get(name)
        diff = None
        pct = None
        if prev_price is not None and prev_price != 0:
            diff = round(price - prev_price, 2)
            pct = round((price - prev_price) / prev_price * 100, 2)

        snapshot.append({
            "underlying_name": name,
            "display_en": info["display_en"],
            "display_ja": info["display_ja"],
            "category": info["category"],
            "subcategory": info["subcategory"],
            "unit": info.get("unit", ""),
            "settlement": price,
            "contract_month": front["month"],
            "instrument_name": front["name"],
            "prev_settlement": prev_price,
            "change_diff": diff,
            "change_pct": pct,
        })

    # Sort by category, then by display name
    category_order = {"energy": 0, "metals": 1, "industrial": 2, "agriculture": 3}
    snapshot.sort(key=lambda x: (category_order.get(x["category"], 99), x["display_en"]))
    return snapshot


def get_commodity_time_series(
    repo: Repository,
    underlying_name: str,
    contract_month: str,
    date_from: str | None = None,
    date_to: str | None = None,
) -> list[dict[str, Any]]:
    """Get historical prices for a specific underlying + contract month.

    Uses instrument_code-based lookup via get_instrument_history if possible,
    otherwise falls back to date-by-date queries.
    """
    # First, find the instrument_code for this underlying+month combo
    # by checking the latest available data
    log = repo.get_import_log()
    success_dates = sorted(
        [e["trade_date"] for e in log if e["status"] == "success"],
        reverse=True,
    )

    instrument_code = None
    for dt in success_dates[:5]:  # Check recent dates
        records = repo.get_by_date_and_underlying(dt, underlying_name)
        for r in records:
            if r.get("contract_month") == contract_month and not r.get("put_call"):
                instrument_code = r["instrument_code"]
                break
        if instrument_code:
            break

    if not instrument_code:
        return []

    history = repo.get_instrument_history(instrument_code, date_from, date_to)
    return [
        {
            "trade_date": r["trade_date"],
            "settlement": r.get("settlement_price"),
            "theoretical": r.get("theoretical_price"),
            "days": r.get("days_to_expiry"),
        }
        for r in history
        if not r.get("put_call")
    ]


def get_all_commodity_forward_curves(
    repo: Repository,
    trade_date: str,
    prev_date: str | None = None,
) -> dict[str, dict[str, Any]]:
    """Get forward curves for all non-power commodities.

    Returns dict keyed by underlying_name with curve data and metadata.
    """
    results = {}
    for name, info in ASSET_TAXONOMY.items():
        if info["category"] not in COMMODITY_CATEGORIES:
            continue
        if info["subcategory"] == "power":
            continue

        curve = get_commodity_forward_curve(repo, trade_date, name)
        if not curve:
            continue

        prev_curve = []
        if prev_date:
            prev_curve = get_commodity_forward_curve(repo, prev_date, name)

        results[name] = {
            "display_en": info["display_en"],
            "display_ja": info["display_ja"],
            "category": info["category"],
            "subcategory": info["subcategory"],
            "unit": info.get("unit", ""),
            "curve": curve,
            "prev_curve": prev_curve,
        }

    return results
