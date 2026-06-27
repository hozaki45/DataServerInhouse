"""Spread analysis computations for the dashboard "スプレッド分析" page.

Recreates the core analysis from `SampleSheet/Spread Calculator_v2.xlsx`
directly from the daily-updating SQLite DB:

  - Power forward curves (East/West/Chubu × Base/Peak), ¥/kWh
  - Spark spread = power price − gas generation cost (from LNG/JKM)
  - Clean spark spread = spark − CO2 cost
  - The raw inputs (power, JKM, interpolated FX) shipped to the page so the
    parameter sliders can recompute spark/clean spark client-side.

Scope notes (v1):
  - Dark spread (coal) is intentionally omitted: JPX has no Newcastle coal
    series. The extension points (COAL_* constants, DARK_SPREAD_MODE) are left
    in place so it can be enabled once a coal price source is available.
  - The FX series in the DB is quarterly USD/JPY futures; it is linearly
    interpolated to monthly here. The default spark formula treats JKM as
    yen-denominated (¥/MMBtu) and does NOT multiply by FX — FX is fetched and
    shipped anyway so the unit assumption is auditable and the formula can be
    switched in one place if JKM turns out to be USD-denominated.

All public functions return JSON-serializable structures.
"""

from __future__ import annotations

from typing import Any

from src.commodity_query import get_commodity_forward_curve
from src.repository import Repository

# ─────────────────────────────────────────────────────────────────────────────
# Parameters (v1 server defaults; also seed the page's interactive sliders)
# Mirrors the Excel "Parameters" sheet.
# ─────────────────────────────────────────────────────────────────────────────
GAS_THERMAL_EFF = 0.5          # ガス火力 熱効率 (combined cycle)
GAS_CONV_KWH_MMBTU = 293.07    # ガス単位換算 kWh/MMBtu
GAS_CO2_KG_PER_KWH = 0.4       # ガス CO2 排出係数 kg-CO2/kWh
CO2_PRICE_YEN_PER_T = 5000     # CO2 価格 円/t

# ── Coal / dark-spread extension points (v1 unused — JPX has no coal) ──
COAL_THERMAL_EFF = 0.4
COAL_CONV_KWH_PER_T = 6978
COAL_CO2_KG_PER_KWH = 0.9
DARK_SPREAD_MODE = "omit"      # "omit" | "manual" | "oil_proxy"
MANUAL_COAL_PRICE_YEN_PER_T: float | None = None

# ── Power curves to chart: display label → underlying_name keyword(s) ──
# Resolved against the DB at runtime so we never hardcode a fragile exact
# string. Each label maps to a list of substrings that must all appear in the
# underlying_name (and excludes weekly/yearly variants).
POWER_CURVES: dict[str, list[str]] = {
    "東・ベース": ["電力(東・ベース)"],
    "東・日中": ["電力(東・日中)"],
    "西・ベース": ["電力(西・ベース)"],
    "西・日中": ["電力(西・日中)"],
    "中部・ベース": ["電力(中部・ベース)"],
    "中部・日中": ["電力(中部・日中)"],
}

JKM_KEYWORD = "JKM"
FX_KEYWORD = "米ドル/日本円"


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
def _ym_to_ord(ym: str) -> int | None:
    """'YYYYMM' → ordinal month index (year*12 + month). None if unparseable."""
    if not ym or len(ym) < 6:
        return None
    try:
        y, m = int(ym[:4]), int(ym[4:6])
    except ValueError:
        return None
    if not (1 <= m <= 12):
        return None
    return y * 12 + m


def _resolve_underlying(repo: Repository, *required: str) -> str | None:
    """Find the single DB underlying_name containing all `required` substrings.

    Robust against leading/trailing whitespace quirks in the source data
    (e.g. the FX underlying is stored as ' 米ドル/日本円' with a leading space).
    Returns the first exact DB string that matches, or None.
    """
    for name in repo.get_underlying_names():
        if all(r in name for r in required):
            return name
    return None


def _curve_map(repo: Repository, date: str, underlying: str) -> dict[str, float]:
    """{ 'YYYYMM': settlement } for an underlying on a date (Nones dropped)."""
    out: dict[str, float] = {}
    for pt in get_commodity_forward_curve(repo, date, underlying):
        month, settle = pt.get("month"), pt.get("settlement")
        if month and settle is not None:
            out[month] = float(settle)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# FX interpolation: quarterly USD/JPY futures → monthly
# ─────────────────────────────────────────────────────────────────────────────
def interpolate_fx_to_monthly(
    fx_points: dict[str, float], target_months: list[str]
) -> dict[str, float]:
    """Linear-interpolate a sparse FX curve to the requested monthly contracts.

    `fx_points` maps 'YYYYMM' → rate (e.g. quarterly futures). Targets between
    two known points are linearly interpolated on the month ordinal; targets
    outside the known range are flat-extrapolated from the nearest endpoint.
    Returns { 'YYYYMM': rate } for every target that could be resolved.
    """
    known = sorted(
        ((o, r) for ym, r in fx_points.items() if (o := _ym_to_ord(ym)) is not None),
        key=lambda t: t[0],
    )
    if not known:
        return {}

    out: dict[str, float] = {}
    for ym in target_months:
        o = _ym_to_ord(ym)
        if o is None:
            continue
        if o <= known[0][0]:
            out[ym] = known[0][1]          # flat extrapolation (front)
        elif o >= known[-1][0]:
            out[ym] = known[-1][1]         # flat extrapolation (back)
        else:
            for (o0, r0), (o1, r1) in zip(known, known[1:]):
                if o0 <= o <= o1:
                    frac = (o - o0) / (o1 - o0) if o1 != o0 else 0.0
                    out[ym] = r0 + (r1 - r0) * frac
                    break
    return out


# ─────────────────────────────────────────────────────────────────────────────
# Forward curves
# ─────────────────────────────────────────────────────────────────────────────
def forward_curves(repo: Repository, date: str) -> dict[str, list[dict[str, Any]]]:
    """6 power forward curves: { label: [{month, price}, ...] } sorted by month."""
    curves: dict[str, list[dict[str, Any]]] = {}
    for label, keywords in POWER_CURVES.items():
        underlying = _resolve_underlying(repo, *keywords)
        if not underlying:
            curves[label] = []
            continue
        cmap = _curve_map(repo, date, underlying)
        curves[label] = [
            {"month": m, "price": round(cmap[m], 4)} for m in sorted(cmap)
        ]
    return curves


# ─────────────────────────────────────────────────────────────────────────────
# Gas generation cost & spark spread
# ─────────────────────────────────────────────────────────────────────────────
def co2_cost(co2_kg_per_kwh: float, co2_price_yen_per_t: float) -> float:
    """CO2 cost in ¥/kWh = factor[kg/kWh] × price[¥/t] / 1000[kg/t]."""
    return co2_kg_per_kwh * co2_price_yen_per_t / 1000.0


def gas_generation_cost_curve(repo: Repository, date: str) -> dict[str, Any]:
    """Gas generation cost (¥/kWh) per month from the JKM (LNG) curve.

    gas_cost[m] = (JKM[m] / GAS_THERMAL_EFF) / GAS_CONV_KWH_MMBTU

    Returns the gas_cost map plus the raw JKM curve and the FX curve
    interpolated onto JKM's months (FX is not used by the default formula but
    is shipped for audit / future USD-denominated switch).
    """
    jkm_name = _resolve_underlying(repo, JKM_KEYWORD)
    jkm = _curve_map(repo, date, jkm_name) if jkm_name else {}

    fx_name = _resolve_underlying(repo, FX_KEYWORD)
    fx_raw = _curve_map(repo, date, fx_name) if fx_name else {}
    fx_monthly = interpolate_fx_to_monthly(fx_raw, sorted(jkm)) if jkm else {}

    gas_cost = {
        m: round((px / GAS_THERMAL_EFF) / GAS_CONV_KWH_MMBTU, 4)
        for m, px in jkm.items()
    }
    return {
        "gas_cost": gas_cost,
        "jkm": {m: round(px, 4) for m, px in jkm.items()},
        "fx": {m: round(r, 4) for m, r in fx_monthly.items()},
        "jkm_underlying": jkm_name,
        "fx_underlying": fx_name,
    }


def spark_spread_curves(repo: Repository, date: str) -> dict[str, Any]:
    """Spark & clean-spark spread (¥/kWh) per power curve, over the months
    where both the power curve and the gas cost exist.

    spark[m]       = power_price[m] − gas_cost[m]
    clean_spark[m] = spark[m] − co2_cost
    """
    gas = gas_generation_cost_curve(repo, date)
    gas_cost = gas["gas_cost"]
    co2 = co2_cost(GAS_CO2_KG_PER_KWH, CO2_PRICE_YEN_PER_T)
    fwd = forward_curves(repo, date)

    spark: dict[str, Any] = {}
    all_months: set[str] = set()
    for label, curve in fwd.items():
        pmap = {p["month"]: p["price"] for p in curve}
        months = sorted(set(pmap) & set(gas_cost))
        if not months:
            spark[label] = {"months": [], "spark": [], "clean_spark": []}
            continue
        all_months.update(months)
        spark[label] = {
            "months": months,
            "spark": [round(pmap[m] - gas_cost[m], 4) for m in months],
            "clean_spark": [round(pmap[m] - gas_cost[m] - co2, 4) for m in months],
        }

    horizon = f"{min(all_months)}..{max(all_months)}" if all_months else None
    return {"spark": spark, "co2_cost": round(co2, 4), "horizon": horizon}


# ─────────────────────────────────────────────────────────────────────────────
# Dark spread (v1 = omit; extension point)
# ─────────────────────────────────────────────────────────────────────────────
def dark_spread_curves(repo: Repository, date: str) -> dict[str, Any] | None:
    """Coal-fired dark spread. v1: omitted (JPX has no Newcastle coal).

    Returns a stub describing the gap. To enable later, set DARK_SPREAD_MODE
    to "manual" (with MANUAL_COAL_PRICE_YEN_PER_T) or "oil_proxy" (Dubai crude
    with a heat-rate assumption) and implement the branch here.
    """
    if DARK_SPREAD_MODE == "omit":
        return {
            "mode": "omit",
            "note": "石炭(Newcastle)価格がJPXデータに無いため、Dark Spreadはv1では非対応。",
        }
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Aggregator → page payload
# ─────────────────────────────────────────────────────────────────────────────
def compute_spread_analysis(
    repo: Repository, latest_date: str, prev_date: str | None = None
) -> dict[str, Any]:
    """Full JSON-serializable payload for the スプレッド分析 page.

    Includes the computed curves AND the raw inputs (power, jkm) so the
    parameter sliders can recompute spark/clean spark in the browser.
    """
    fwd = forward_curves(repo, latest_date)
    gas = gas_generation_cost_curve(repo, latest_date)
    spark = spark_spread_curves(repo, latest_date)

    jkm_curve = [
        {"month": m, "price": gas["jkm"][m]} for m in sorted(gas["jkm"])
    ]

    return {
        "latest_date": latest_date,
        "prev_date": prev_date,
        "forward_curves": fwd,
        "spark": spark["spark"],
        "horizon": spark["horizon"],
        "co2_cost": spark["co2_cost"],
        "gas_cost": gas["gas_cost"],
        "jkm_curve": jkm_curve,
        "fx": gas["fx"],
        "dark_spread": dark_spread_curves(repo, latest_date),
        "params": {
            "gas_thermal_eff": GAS_THERMAL_EFF,
            "gas_conv_kwh_mmbtu": GAS_CONV_KWH_MMBTU,
            "gas_co2_kg_per_kwh": GAS_CO2_KG_PER_KWH,
            "co2_price_yen_per_t": CO2_PRICE_YEN_PER_T,
            "jkm_unit_assumption": "¥/MMBtu (FX not applied)",
            "jkm_underlying": gas["jkm_underlying"],
            "fx_underlying": gas["fx_underlying"],
        },
    }
