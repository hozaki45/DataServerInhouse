"""Generate static HTML dashboard for GitHub Pages.

Reads data from SQLite, outputs docs/index.html + data.json.

Usage:
    uv run python scripts/generate_site.py
"""

import json
import sys
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.query import get_power_futures, summary_by_underlying
from src.repository import get_repository
from src.commodity_query import (
    find_nearest_business_day_before,
    get_cross_commodity_snapshot,
    get_all_commodity_forward_curves,
    get_commodity_forward_curve,
    get_front_month_price,
)
from src.asset_taxonomy import CATEGORY_META, ASSET_TAXONOMY, COMMODITY_CATEGORIES
from src.spread_analysis import compute_spread_analysis

SITE_DIR = Path(__file__).resolve().parent.parent / "docs"

# Mapping of instrument name patterns to curve categories
CURVE_CATEGORIES = {
    "東・ベース(月次)": {"prefix": "FUT_EEB_", "exclude": ["W", "Y"], "underlying": "電力(東・ベース)"},
    "東・日中(月次)": {"prefix": "FUT_EEP_", "exclude": ["W", "Y"], "underlying": "電力(東・日中)"},
    "西・ベース(月次)": {"prefix": "FUT_EWB_", "exclude": ["W", "Y"], "underlying": "電力(西・ベース)"},
    "西・日中(月次)": {"prefix": "FUT_EWP_", "exclude": ["W", "Y"], "underlying": "電力(西・日中)"},
    "東・ベース(週間)": {"prefix": "FUT_EEBW_", "exclude": [], "underlying": "電力(東・週間ベース)"},
    "東・日中(週間)": {"prefix": "FUT_EEPW_", "exclude": [], "underlying": "電力(東・週間日中)"},
    "西・ベース(週間)": {"prefix": "FUT_EWBW_", "exclude": [], "underlying": "電力(西・週間ベース)"},
    "西・日中(週間)": {"prefix": "FUT_EWPW_", "exclude": [], "underlying": "電力(西・週間日中)"},
    "東・ベース(年度)": {"prefix": "FUT_EEBY_", "exclude": [], "underlying": "電力(東・年度ベース)"},
    "東・日中(年度)": {"prefix": "FUT_EEPY_", "exclude": [], "underlying": "電力(東・年度日中)"},
    "西・ベース(年度)": {"prefix": "FUT_EWBY_", "exclude": [], "underlying": "電力(西・年度ベース)"},
    "西・日中(年度)": {"prefix": "FUT_EWPY_", "exclude": [], "underlying": "電力(西・年度日中)"},
}


def classify_power_future(instrument_name: str) -> str | None:
    """Classify a power future into a curve category."""
    for cat_name, cfg in CURVE_CATEGORIES.items():
        if instrument_name.startswith(cfg["prefix"]):
            # Check exclusions
            suffix = instrument_name[len(cfg["prefix"]):]
            if any(ex in suffix for ex in cfg["exclude"]):
                continue
            return cat_name
    return None


def _build_prev_price_map(repo, prev_date: str | None) -> dict[str, float]:
    """Build instrument_name -> settlement_price map for previous date."""
    if not prev_date:
        return {}
    prev_power = get_power_futures(repo, prev_date)
    return {
        r["instrument_name"]: r.get("settlement_price")
        for r in prev_power
        if not r.get("put_call") and r.get("settlement_price") is not None
    }


def _calc_change(current: float | None, prev: float | None) -> dict:
    """Calculate absolute and percentage change."""
    if current is None or prev is None or prev == 0:
        return {"diff": None, "pct": None}
    diff = round(current - prev, 2)
    pct = round((current - prev) / prev * 100, 2)
    return {"diff": diff, "pct": pct}


def _build_month_map_prev(cat_key: str, prev_forward_curves: dict) -> dict:
    """Build month -> item map from previous day forward curves."""
    items = prev_forward_curves.get(cat_key, [])
    return {it["month"]: it for it in items if len(it.get("month", "")) == 6}


def generate_data_json(repo) -> dict:
    """Generate data.json from database for chart consumption."""
    log = repo.get_import_log()
    if not log:
        return {}

    success_dates = sorted(
        [e["trade_date"] for e in log if e["status"] == "success"],
        reverse=True,
    )
    latest_date = success_dates[0]
    prev_date = success_dates[1] if len(success_dates) > 1 else None

    # Previous day price map
    prev_map = _build_prev_price_map(repo, prev_date)

    # Power futures
    power = get_power_futures(repo, latest_date)
    power_fut = [r for r in power if not r.get("put_call")]

    # Build forward curves per category (with prev day comparison)
    forward_curves = {}
    # Also build prev day curves for overlay
    prev_forward_curves = {}
    if prev_date:
        prev_power = get_power_futures(repo, prev_date)
        prev_power_fut = [r for r in prev_power if not r.get("put_call")]
        for r in prev_power_fut:
            cat = classify_power_future(r["instrument_name"])
            if cat is None:
                continue
            if cat not in prev_forward_curves:
                prev_forward_curves[cat] = []
            prev_forward_curves[cat].append({
                "month": r.get("contract_month", ""),
                "settlement": r.get("settlement_price"),
            })
        for cat in prev_forward_curves:
            prev_forward_curves[cat].sort(key=lambda x: x["month"])

    for r in power_fut:
        cat = classify_power_future(r["instrument_name"])
        if cat is None:
            continue
        if cat not in forward_curves:
            forward_curves[cat] = []

        month = r.get("contract_month", "")
        settlement = r.get("settlement_price")
        prev_price = prev_map.get(r["instrument_name"])
        change = _calc_change(settlement, prev_price)

        forward_curves[cat].append({
            "month": month,
            "settlement": settlement,
            "theoretical": r.get("theoretical_price"),
            "days": r.get("days_to_expiry"),
            "name": r["instrument_name"],
            "prev_settlement": prev_price,
            "change_diff": change["diff"],
            "change_pct": change["pct"],
        })

    # Sort each curve by month/date
    for cat in forward_curves:
        forward_curves[cat].sort(key=lambda x: x["month"])

    # Overview chart: monthly curves for the 4 main types
    overview_chart = {}
    for cat_name in ["東・ベース(月次)", "東・日中(月次)", "西・ベース(月次)", "西・日中(月次)"]:
        if cat_name in forward_curves:
            overview_chart[cat_name] = [
                {"month": p["month"], "price": p["settlement"]}
                for p in forward_curves[cat_name]
                if p["settlement"] is not None
            ]

    # Previous day overview for comparison
    prev_overview_chart = {}
    for cat_name in ["東・ベース(月次)", "東・日中(月次)", "西・ベース(月次)", "西・日中(月次)"]:
        if cat_name in prev_forward_curves:
            prev_overview_chart[cat_name] = [
                {"month": p["month"], "price": p["settlement"]}
                for p in prev_forward_curves[cat_name]
                if p["settlement"] is not None
            ]

    # Top movers: biggest absolute changes in settlement price
    all_changes = []
    for r in power_fut:
        name = r["instrument_name"]
        settlement = r.get("settlement_price")
        prev_price = prev_map.get(name)
        change = _calc_change(settlement, prev_price)
        if change["diff"] is not None:
            all_changes.append({
                "name": name,
                "category": classify_power_future(name) or "",
                "month": r.get("contract_month", ""),
                "settlement": settlement,
                "prev_settlement": prev_price,
                "diff": change["diff"],
                "pct": change["pct"],
            })
    top_movers = sorted(all_changes, key=lambda x: abs(x["diff"]), reverse=True)[:15]

    # Summary by underlying
    summary = summary_by_underlying(repo, latest_date)

    # Import log
    import_dates = [entry["trade_date"] for entry in log if entry["status"] == "success"]

    # Commodity snapshot (non-power commodities with DoD changes)
    commodity_snapshot = get_cross_commodity_snapshot(repo, latest_date, prev_date)

    # Commodity forward curves
    commodity_curves = get_all_commodity_forward_curves(repo, latest_date, prev_date)
    # Simplify for JSON serialization
    commodity_curves_json = {}
    for name, cdata in commodity_curves.items():
        commodity_curves_json[name] = {
            "display_en": cdata["display_en"],
            "display_ja": cdata["display_ja"],
            "category": cdata["category"],
            "unit": cdata["unit"],
            "curve": cdata["curve"],
            "prev_curve": cdata["prev_curve"],
        }

    # ── Power Heatmap: price changes + E-W spreads ──
    heatmap_types = {
        "東・ベース": "東・ベース(月次)",
        "東・日中": "東・日中(月次)",
        "西・ベース": "西・ベース(月次)",
        "西・日中": "西・日中(月次)",
    }
    # Collect monthly data per type
    heatmap_price_changes: dict[str, list] = {}
    for label, cat_key in heatmap_types.items():
        items = forward_curves.get(cat_key, [])
        heatmap_price_changes[label] = [
            {
                "month": it["month"],
                "settlement": it["settlement"],
                "prev": it.get("prev_settlement"),
                "diff": it.get("change_diff"),
                "pct": it.get("change_pct"),
            }
            for it in items
            if len(it.get("month", "")) == 6  # monthly only
        ]

    # Compute E-W spreads
    def _build_month_map(cat_key: str) -> dict:
        items = forward_curves.get(cat_key, [])
        return {it["month"]: it for it in items if len(it.get("month", "")) == 6}

    east_base_map = _build_month_map("東・ベース(月次)")
    west_base_map = _build_month_map("西・ベース(月次)")
    east_peak_map = _build_month_map("東・日中(月次)")
    west_peak_map = _build_month_map("西・日中(月次)")

    prev_east_base_map = _build_month_map_prev("東・ベース(月次)", prev_forward_curves)
    prev_west_base_map = _build_month_map_prev("西・ベース(月次)", prev_forward_curves)
    prev_east_peak_map = _build_month_map_prev("東・日中(月次)", prev_forward_curves)
    prev_west_peak_map = _build_month_map_prev("西・日中(月次)", prev_forward_curves)

    heatmap_months = sorted(set(east_base_map.keys()) & set(west_base_map.keys()))

    spread_base = []
    spread_peak = []
    for m in heatmap_months:
        eb = east_base_map[m].get("settlement")
        wb = west_base_map[m].get("settlement")
        ep = east_peak_map.get(m, {}).get("settlement")
        wp = west_peak_map.get(m, {}).get("settlement")

        s_base = round(eb - wb, 2) if eb is not None and wb is not None else None
        s_peak = round(ep - wp, 2) if ep is not None and wp is not None else None

        # Previous day spreads
        peb = prev_east_base_map.get(m, {}).get("settlement")
        pwb = prev_west_base_map.get(m, {}).get("settlement")
        pep = prev_east_peak_map.get(m, {}).get("settlement")
        pwp = prev_west_peak_map.get(m, {}).get("settlement")

        ps_base = round(peb - pwb, 2) if peb is not None and pwb is not None else None
        ps_peak = round(pep - pwp, 2) if pep is not None and pwp is not None else None

        sc_base = round(s_base - ps_base, 2) if s_base is not None and ps_base is not None else None
        sc_peak = round(s_peak - ps_peak, 2) if s_peak is not None and ps_peak is not None else None

        spread_base.append({"month": m, "spread": s_base, "prev_spread": ps_base, "spread_change": sc_base})
        spread_peak.append({"month": m, "spread": s_peak, "prev_spread": ps_peak, "spread_change": sc_peak})

    power_heatmap = {
        "months": heatmap_months,
        "price_changes": heatmap_price_changes,
        "spreads": {"base": spread_base, "peak": spread_peak},
    }

    return {
        "latest_date": latest_date,
        "prev_date": prev_date,
        "generated_at": datetime.now().isoformat(),
        "total_records": sum(s["total"] for s in summary),
        "underlying_count": len(summary),
        "import_dates": import_dates,
        "power_futures_count": len(power_fut),
        "overview_chart": overview_chart,
        "prev_overview_chart": prev_overview_chart,
        "forward_curves": forward_curves,
        "prev_forward_curves": prev_forward_curves,
        "top_movers": top_movers,
        "power_heatmap": power_heatmap,
        "power_futures": [
            {
                "name": r["instrument_name"],
                "underlying": r.get("underlying_name", ""),
                "month": r.get("contract_month", ""),
                "settlement": r.get("settlement_price"),
                "theoretical": r.get("theoretical_price"),
                "days": r.get("days_to_expiry"),
                "category": classify_power_future(r["instrument_name"]),
                "prev_settlement": prev_map.get(r["instrument_name"]),
                "change_diff": _calc_change(
                    r.get("settlement_price"), prev_map.get(r["instrument_name"])
                )["diff"],
                "change_pct": _calc_change(
                    r.get("settlement_price"), prev_map.get(r["instrument_name"])
                )["pct"],
            }
            for r in power_fut
        ],
        "summary": summary,
        "commodity_snapshot": commodity_snapshot,
        "commodity_curves": commodity_curves_json,
        "commodity_count": len(commodity_snapshot),
    }


def _change_html(diff, pct) -> str:
    """Generate HTML for a day-over-day change cell."""
    if diff is None:
        return '<td class="num change">-</td>'
    sign = "+" if diff >= 0 else ""
    css = "positive" if diff > 0 else "negative" if diff < 0 else ""
    return (
        f'<td class="num change {css}">'
        f'{sign}{diff:.2f}'
        f'<span class="change-pct">({sign}{pct:.1f}%)</span>'
        f'</td>'
    )


def generate_html(data: dict) -> str:
    """Generate the HTML dashboard page with 2025 dark mode design."""
    latest = data.get("latest_date", "N/A")
    prev_date = data.get("prev_date", "")
    generated = data.get("generated_at", "")[:19].replace("T", " ")
    total = data.get("total_records", 0)
    underlying_count = data.get("underlying_count", 0)
    dates = data.get("import_dates", [])
    power_count = data.get("power_futures_count", 0)
    commodity_count = data.get("commodity_count", 0)

    # Power futures table rows (with change column)
    power_rows = ""
    for pf in data.get("power_futures", []):
        settle = pf["settlement"] if pf["settlement"] is not None else ""
        theo = pf["theoretical"] if pf["theoretical"] is not None else ""
        days = pf["days"] if pf["days"] is not None else ""
        cat = pf.get("category", "") or ""
        change_cell = _change_html(pf.get("change_diff"), pf.get("change_pct"))
        power_rows += f"""            <tr>
              <td>{pf['name']}</td>
              <td>{cat}</td>
              <td>{pf['month']}</td>
              <td class="num">{settle}</td>
              {change_cell}
              <td class="num">{theo}</td>
              <td class="num">{days}</td>
            </tr>\n"""

    # Top movers table rows
    movers_rows = ""
    for m in data.get("top_movers", []):
        sign = "+" if m["diff"] >= 0 else ""
        css = "positive" if m["diff"] > 0 else "negative"
        movers_rows += f"""            <tr>
              <td>{m['name']}</td>
              <td>{m['category']}</td>
              <td>{m['month']}</td>
              <td class="num">{m['settlement']:.2f}</td>
              <td class="num">{m['prev_settlement']:.2f}</td>
              <td class="num {css}">{sign}{m['diff']:.2f}</td>
              <td class="num {css}">{sign}{m['pct']:.1f}%</td>
            </tr>\n"""

    # Summary table rows
    summary_rows = ""
    for s in data.get("summary", []):
        summary_rows += f"""            <tr>
              <td>{s['underlying_name']}</td>
              <td class="num">{s['total']}</td>
              <td class="num">{s['fut']}</td>
              <td class="num">{s['cal']}</td>
              <td class="num">{s['put']}</td>
            </tr>\n"""

    # Build curve selector options
    curve_options = ""
    for cat_name in sorted(data.get("forward_curves", {}).keys()):
        curve_options += f'          <option value="{cat_name}">{cat_name}</option>\n'

    # Build commodity snapshot table rows
    commodity_rows = ""
    for cs in data.get("commodity_snapshot", []):
        price = cs["settlement"]
        unit = cs.get("unit", "")
        cat_meta = CATEGORY_META.get(cs["category"], {})
        cat_label = cat_meta.get("display_en", cs["category"])
        diff = cs.get("change_diff")
        pct = cs.get("change_pct")
        if diff is not None:
            sign = "+" if diff >= 0 else ""
            css_cls = "positive" if diff > 0 else "negative" if diff < 0 else ""
            diff_cell = f'<td class="num change {css_cls}">{sign}{diff:.2f}</td>'
            pct_cell = f'<td class="num change {css_cls}">{sign}{pct:.1f}%</td>'
        else:
            diff_cell = '<td class="num change">-</td>'
            pct_cell = '<td class="num change">-</td>'
        commodity_rows += f"""            <tr>
              <td>{cs['display_en']}<span class="commodity-ja">{cs['display_ja']}</span></td>
              <td><span class="cat-badge" style="background:{cat_meta.get('color', '#666')}22;color:{cat_meta.get('color', '#666')}">{cat_label}</span></td>
              <td>{cs.get('contract_month', '')}</td>
              <td class="num">{price:.2f}</td>
              <td class="num" style="opacity:0.7">{unit}</td>
              {diff_cell}
              {pct_cell}
            </tr>\n"""

    # Build commodity curve selector options
    commodity_curve_options = ""
    for cname in sorted(data.get("commodity_curves", {}).keys()):
        cdata = data["commodity_curves"][cname]
        commodity_curve_options += f'          <option value="{cname}">{cdata["display_en"]} ({cdata["display_ja"]})</option>\n'

    # Build category meta JSON for JS
    category_meta_json = json.dumps(
        {k: {"color": v["color"], "display_en": v["display_en"]} for k, v in CATEGORY_META.items()},
        ensure_ascii=False,
    )

    # Inline JSON for chart data (avoids fetch/CORS issues on file://)
    chart_data = {
        "overview_chart": data.get("overview_chart", {}),
        "prev_overview_chart": data.get("prev_overview_chart", {}),
        "forward_curves": data.get("forward_curves", {}),
        "prev_forward_curves": data.get("prev_forward_curves", {}),
        "commodity_curves": data.get("commodity_curves", {}),
        "commodity_snapshot": data.get("commodity_snapshot", []),
        "power_heatmap": data.get("power_heatmap", {}),
    }
    chart_data_json = json.dumps(chart_data, ensure_ascii=False)
    prev_label = prev_date or "N/A"

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>DataServer In-House — Market Analytics Dashboard</title>
  <link rel="stylesheet" href="style.css">
  <script src="https://cdn.jsdelivr.net/npm/apexcharts@3"></script>
</head>
<body>

<div class="bg-gradient"></div>

<div class="app-layout">

  <!-- Sidebar -->
  <nav class="sidebar">
    <div class="sidebar-logo">DS</div>
    <div class="sidebar-nav">
      <a href="#overview" class="nav-item active">
        <span>&#9670;</span>
        <span class="nav-tooltip">Overview</span>
      </a>
      <a href="#curves" class="nav-item">
        <span>&#9699;</span>
        <span class="nav-tooltip">Forward Curves</span>
      </a>
      <a href="#spread" class="nav-item">
        <span>&#8651;</span>
        <span class="nav-tooltip">East-West Spread</span>
      </a>
      <a href="#priceHeatmap" class="nav-item">
        <span>&#9619;</span>
        <span class="nav-tooltip">Price Heatmap</span>
      </a>
      <a href="#spreadHeatmap" class="nav-item">
        <span>&#9608;</span>
        <span class="nav-tooltip">Spread Matrix</span>
      </a>
      <a href="#movers" class="nav-item">
        <span>&#128293;</span>
        <span class="nav-tooltip">Top Movers</span>
      </a>
      <a href="#futures" class="nav-item">
        <span>&#9783;</span>
        <span class="nav-tooltip">Futures Data</span>
      </a>
      <a href="#summary" class="nav-item">
        <span>&#9881;</span>
        <span class="nav-tooltip">Summary</span>
      </a>
      <div style="width:32px;height:1px;background:var(--border-subtle);margin:8px 0"></div>
      <a href="#commodities" class="nav-item">
        <span>&#9632;</span>
        <span class="nav-tooltip">Commodities</span>
      </a>
      <a href="#commodity-curves" class="nav-item">
        <span>&#9650;</span>
        <span class="nav-tooltip">Commodity Curves</span>
      </a>
      <a href="#cross-market" class="nav-item">
        <span>&#9619;</span>
        <span class="nav-tooltip">Cross-Market</span>
      </a>
      <div style="width:32px;height:1px;background:var(--border-subtle);margin:8px 0"></div>
      <a href="#spot" class="nav-item">
        <span>&#9733;</span>
        <span class="nav-tooltip">Spot Analysis</span>
      </a>
      <a href="#spotVsFutures" class="nav-item">
        <span>&#8646;</span>
        <span class="nav-tooltip">Spot vs Futures</span>
      </a>
      <div style="width:32px;height:1px;background:var(--border-subtle);margin:8px 0"></div>
      <a href="weekly_compare.html" class="nav-item">
        <span>&#9776;</span>
        <span class="nav-tooltip">週次比較 (Weekly)</span>
      </a>
      <a href="spread_analysis.html" class="nav-item">
        <span>&#9878;</span>
        <span class="nav-tooltip">スプレッド分析 (Spread)</span>
      </a>
    </div>
  </nav>

  <!-- Main Content -->
  <div class="main-content">

    <!-- Header -->
    <header class="header">
      <div class="header-left">
        <h1>DataServer In-House</h1>
        <div class="subtitle">JPX Derivative & Commodity Analytics</div>
      </div>
      <div class="header-right">
        <div class="header-badge">
          <span class="dot"></span>
          Live
        </div>
        <div class="header-date">{latest}<br><span style="font-size:0.65rem;color:#64748B">vs {prev_label}</span></div>
      </div>
    </header>

    <div class="container">

      <!-- KPI Cards -->
      <div class="kpi-row" style="grid-template-columns:repeat(5,1fr)">
        <div class="kpi-card">
          <div class="kpi-icon">&#128202;</div>
          <div class="kpi-value" data-target="{total}">{total:,}</div>
          <div class="kpi-label">Total Records</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-icon">&#127968;</div>
          <div class="kpi-value" data-target="{underlying_count}">{underlying_count}</div>
          <div class="kpi-label">Underlying Assets</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-icon">&#9889;</div>
          <div class="kpi-value" data-target="{power_count}">{power_count}</div>
          <div class="kpi-label">Power Futures</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-icon">&#128197;</div>
          <div class="kpi-value" data-target="{len(dates)}">{len(dates)}</div>
          <div class="kpi-label">Import Days</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-icon">&#9632;</div>
          <div class="kpi-value" data-target="{commodity_count}">{commodity_count}</div>
          <div class="kpi-label">Commodities</div>
        </div>
      </div>

      <!-- Bento Grid -->
      <div class="bento-grid">

        <!-- Overview Chart (large) -->
        <div class="card card-full" id="overview">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9670;</span> Forward Curve Comparison (Monthly)</div>
            <div style="display:flex;align-items:center;gap:0.75rem">
              <select id="overviewModeSelect" style="padding:0.4rem 2rem 0.4rem 0.8rem;font-size:0.78rem;font-family:Inter,sans-serif;border:1px solid var(--border-subtle);border-radius:4px;background:var(--bg-surface);color:var(--text-primary);cursor:pointer;appearance:none;background-image:url(&quot;data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%2394A3B8' d='M6 8L1 3h10z'/%3E%3C/svg%3E&quot;);background-repeat:no-repeat;background-position:right 8px center">
                <option value="base">Base Load (East vs West)</option>
                <option value="peak">Peak Load (East vs West)</option>
                <option value="all">All 4 Areas</option>
              </select>
              <div class="card-badge" id="overviewBadge">East vs West</div>
            </div>
          </div>
          <div class="card-body">
            <div class="chart-container" id="overviewChart" style="min-height:380px"></div>
          </div>
        </div>

        <!-- Curve Selector (feature) -->
        <div class="card card-feature" id="curves">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9699;</span> Interactive Forward Curve</div>
            <div class="card-badge" id="curveCountBadge">12 Curves</div>
          </div>
          <div class="card-body">
            <div class="curve-selector">
              <label for="curveSelect">Curve:</label>
              <select id="curveSelect">
{curve_options}          </select>
              <span class="curve-info" id="curveInfo"></span>
            </div>
            <div class="chart-container" id="curveChart" style="min-height:320px"></div>
          </div>
        </div>

        <!-- Spread Chart (side) -->
        <div class="card card-side" id="spread">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#8651;</span> E-W Spread</div>
          </div>
          <div class="card-body">
            <div class="chart-container" id="spreadChart" style="min-height:320px"></div>
          </div>
        </div>

        <!-- Power Price Change Heatmap -->
        <div class="card card-full" id="priceHeatmap">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9619;</span> Power Price Changes — DoD % Change Heatmap</div>
            <div class="card-badge">24 months &times; 4 areas</div>
          </div>
          <div class="card-body">
            <div class="chart-container" id="priceHeatmapChart" style="min-height:420px"></div>
          </div>
        </div>

        <!-- E-W Spread Heatmap -->
        <div class="card card-full" id="spreadHeatmap">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#8651;</span> East-West Spread Matrix</div>
            <div class="card-badge">{latest} vs {prev_label}</div>
          </div>
          <div class="card-body">
            <div class="chart-container" id="spreadHeatmapChart" style="min-height:420px"></div>
          </div>
        </div>

        <!-- Top Movers -->
        <div class="card card-full" id="movers">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#128293;</span> Top Movers (Day-over-Day)</div>
            <div class="card-badge">{latest} vs {prev_label}</div>
          </div>
          <div class="card-body">
            <div class="table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th>Instrument</th>
                    <th>Category</th>
                    <th>Month</th>
                    <th class="num">Today</th>
                    <th class="num">Prev</th>
                    <th class="num">Change</th>
                    <th class="num">%</th>
                  </tr>
                </thead>
                <tbody>
{movers_rows}                </tbody>
              </table>
            </div>
          </div>
        </div>

        <!-- Curve Detail Table -->
        <div class="card card-full">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9783;</span> Curve Detail</div>
          </div>
          <div class="card-body">
            <div class="table-wrapper" id="curveTable"></div>
          </div>
        </div>

        <!-- Power Futures Table -->
        <div class="card card-wide" id="futures">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9889;</span> Power Futures ({latest})</div>
            <div class="card-badge">{power_count} contracts</div>
          </div>
          <div class="card-body">
            <div class="table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th>Instrument</th>
                    <th>Category</th>
                    <th>Month</th>
                    <th class="num">Settlement</th>
                    <th class="num">DoD Change</th>
                    <th class="num">Theoretical</th>
                    <th class="num">DTE</th>
                  </tr>
                </thead>
                <tbody>
{power_rows}                </tbody>
              </table>
            </div>
          </div>
        </div>

        <!-- Summary Table -->
        <div class="card card-wide" id="summary">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9881;</span> Underlying Summary ({latest})</div>
            <div class="card-badge">{underlying_count} assets</div>
          </div>
          <div class="card-body">
            <div class="table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th>Underlying</th>
                    <th class="num">Total</th>
                    <th class="num">Futures</th>
                    <th class="num">Calls</th>
                    <th class="num">Puts</th>
                  </tr>
                </thead>
                <tbody>
{summary_rows}                </tbody>
              </table>
            </div>
          </div>
        </div>

      </div><!-- /bento-grid -->

      <!-- ==================== Commodity Section ==================== -->
      <div class="section-title" id="commodities" style="margin-top:2.5rem">COMMODITY MARKET OVERVIEW</div>

      <div class="bento-grid">

        <!-- Commodity Snapshot Table -->
        <div class="card card-full">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9632;</span> Commodity Market Overview</div>
            <div class="card-badge">{commodity_count} assets</div>
          </div>
          <div class="card-body">
            <div class="table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th>Asset</th>
                    <th>Category</th>
                    <th>Contract</th>
                    <th class="num">Price</th>
                    <th class="num">Unit</th>
                    <th class="num">Change</th>
                    <th class="num">%Change</th>
                  </tr>
                </thead>
                <tbody>
{commodity_rows}                </tbody>
              </table>
            </div>
          </div>
        </div>

        <!-- Commodity Forward Curves -->
        <div class="card card-full" id="commodity-curves">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9650;</span> Commodity Forward Curves</div>
            <div class="card-badge" id="commodityCurveBadge">Select</div>
          </div>
          <div class="card-body">
            <div class="curve-selector">
              <label for="commoditySelect">Commodity:</label>
              <select id="commoditySelect" style="min-width:280px">
{commodity_curve_options}          </select>
              <span class="curve-info" id="commodityCurveInfo"></span>
            </div>
            <div class="chart-container" id="commodityCurveChart" style="min-height:380px"></div>
          </div>
        </div>

        <!-- Cross-Market Daily Changes -->
        <div class="card card-full" id="cross-market">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9619;</span> Cross-Market Daily Changes</div>
            <div class="card-badge">{latest} vs {prev_label}</div>
          </div>
          <div class="card-body">
            <div class="chart-container" id="crossMarketChart" style="min-height:400px"></div>
          </div>
        </div>

      </div><!-- /bento-grid commodity -->

      <!-- ==================== Spot Analysis Section ==================== -->
      <div class="section-title" id="spot" style="margin-top:2.5rem">SPOT PRICE ANALYSIS</div>

      <!-- Upload Zone -->
      <div class="upload-zone" id="uploadZone">
        <div class="upload-icon">&#128200;</div>
        <div class="upload-title">JEPX Spot Price CSV</div>
        <div class="upload-desc">Drop <code>spot_summary_2025.csv</code> here or click to select</div>
        <input type="file" id="spotFileInput" accept=".csv" style="display:none">
        <button class="upload-btn" id="uploadBtn">Select CSV File</button>
      </div>

      <!-- Spot Results (hidden until CSV uploaded) -->
      <div id="spotResults" style="display:none">

        <!-- Spot KPI Cards -->
        <div class="kpi-row" style="margin-top:1.5rem">
          <div class="kpi-card">
            <div class="kpi-icon" style="background:rgba(59,130,246,0.1);color:var(--accent-blue)">&#9889;</div>
            <div class="kpi-value" id="spotKpiTokyoBase">-</div>
            <div class="kpi-label">Tokyo Base Avg (JPY/kWh)</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-icon" style="background:rgba(6,182,212,0.1);color:var(--accent-cyan)">&#9889;</div>
            <div class="kpi-value" id="spotKpiTokyoPeak">-</div>
            <div class="kpi-label">Tokyo Peak Avg (JPY/kWh)</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-icon" style="background:rgba(245,158,11,0.1);color:var(--warning)">&#9889;</div>
            <div class="kpi-value" id="spotKpiKansaiBase">-</div>
            <div class="kpi-label">Kansai Base Avg (JPY/kWh)</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-icon" style="background:rgba(139,92,246,0.1);color:var(--accent-purple)">&#128197;</div>
            <div class="kpi-value" id="spotKpiPeriod" style="font-size:1rem">-</div>
            <div class="kpi-label">Data Period</div>
          </div>
        </div>

        <div class="bento-grid">

          <!-- Monthly Spot Trend -->
          <div class="card card-full">
            <div class="card-header">
              <div class="card-title"><span class="icon">&#9733;</span> Monthly Spot Price Trend</div>
              <div class="card-badge">Base & Peak</div>
            </div>
            <div class="card-body">
              <div class="chart-container" id="spotMonthlyChart" style="min-height:380px"></div>
            </div>
          </div>

          <!-- Intraday Profile -->
          <div class="card card-full">
            <div class="card-header">
              <div class="card-title"><span class="icon">&#128336;</span> Intraday Price Profile (Duck Curve)</div>
              <div class="card-badge">48 Slots</div>
            </div>
            <div class="card-body">
              <div class="chart-container" id="spotIntradayChart" style="min-height:350px"></div>
            </div>
          </div>

          <!-- Futures vs Spot Overlay -->
          <div class="card card-full" id="spotVsFutures">
            <div class="card-header">
              <div class="card-title"><span class="icon">&#8646;</span> Futures vs Spot Comparison</div>
              <div class="card-badge">Monthly Match</div>
            </div>
            <div class="card-body">
              <div class="spot-category-selector" style="margin-bottom:1rem">
                <label style="font-weight:500;color:var(--text-secondary);font-size:0.82rem">Category:</label>
                <select id="spotCategorySelect" style="padding:0.5rem 2.2rem 0.5rem 1rem;font-size:0.82rem;font-family:Inter,sans-serif;border:1px solid var(--glass-border);border-radius:var(--radius-sm);background:var(--bg-surface);color:var(--text-primary);cursor:pointer;min-width:240px;appearance:none;background-image:url(&quot;data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%2394A3B8' d='M6 8L1 3h10z'/%3E%3C/svg%3E&quot;);background-repeat:no-repeat;background-position:right 12px center">
                  <option value="east_base">East Base (Monthly)</option>
                  <option value="east_peak">East Peak (Monthly)</option>
                  <option value="west_base">West Base (Monthly)</option>
                  <option value="west_peak">West Peak (Monthly)</option>
                </select>
              </div>
              <div class="chart-container" id="spotVsFuturesChart" style="min-height:380px"></div>
            </div>
          </div>

          <!-- Premium/Discount Bar -->
          <div class="card card-full">
            <div class="card-header">
              <div class="card-title"><span class="icon">&#128200;</span> Futures Premium / Discount</div>
              <div class="card-badge" id="premiumBadge">vs Spot</div>
            </div>
            <div class="card-body">
              <div class="chart-container" id="spotPremiumChart" style="min-height:320px"></div>
            </div>
          </div>

        </div><!-- /bento-grid spot -->

      </div><!-- /spotResults -->

    </div><!-- /container -->

    <div class="footer">
      DataServer In-House — JPX Derivative & Commodity Market Analytics Platform &middot; Generated {generated}
    </div>

  </div><!-- /main-content -->
</div><!-- /app-layout -->

<script>
const chartData = {chart_data_json};
const categoryMeta = {category_meta_json};

(function() {{
    'use strict';

    // ============================================
    // Theme config
    // ============================================
    const colors = {{
      blue: '#3B82F6',
      cyan: '#06B6D4',
      purple: '#8B5CF6',
      amber: '#F59E0B',
      pink: '#EC4899',
      green: '#22C55E',
      red: '#EF4444',
      grid: 'rgba(255,255,255,0.05)',
      text: '#94A3B8',
      textLight: '#64748B'
    }};

    const baseChartOpts = {{
      chart: {{
        background: 'transparent',
        fontFamily: "'Inter', sans-serif",
        toolbar: {{ show: true, tools: {{ download: true, selection: false, zoom: true, zoomin: true, zoomout: true, pan: true, reset: true }} }},
      }},
      theme: {{ mode: 'dark', palette: 'palette1' }},
      grid: {{
        borderColor: colors.grid,
        strokeDashArray: 3,
        xaxis: {{ lines: {{ show: false }} }},
        yaxis: {{ lines: {{ show: true }} }},
        padding: {{ left: 10, right: 10 }}
      }},
      tooltip: {{
        theme: 'dark',
        style: {{ fontSize: '12px', fontFamily: "'JetBrains Mono', monospace" }},
        y: {{ formatter: v => v !== null && v !== undefined ? v.toFixed(2) + ' JPY/kWh' : 'N/A' }}
      }},
      xaxis: {{
        labels: {{ style: {{ colors: colors.text, fontSize: '11px', fontFamily: "'Inter', sans-serif" }} }},
        axisBorder: {{ color: colors.grid }},
        axisTicks: {{ color: colors.grid }}
      }},
      yaxis: {{
        labels: {{
          style: {{ colors: colors.text, fontSize: '11px', fontFamily: "'JetBrains Mono', monospace" }},
          formatter: v => v !== null ? v.toFixed(1) : ''
        }}
      }},
      legend: {{
        labels: {{ colors: colors.text }},
        fontSize: '12px',
        fontFamily: "'Inter', sans-serif",
        itemMargin: {{ horizontal: 12 }}
      }}
    }};

    // ============================================
    // 1. Overview: Forward Curve with Base/Peak selector
    // ============================================
    const overview = chartData.overview_chart;
    const prevOverview = chartData.prev_overview_chart || {{}};

    const overviewModes = {{
      base: {{
        cats: [
          {{ key: '\u6771\u30fb\u30d9\u30fc\u30b9(\u6708\u6b21)', color: colors.blue, name: 'East Base' }},
          {{ key: '\u897f\u30fb\u30d9\u30fc\u30b9(\u6708\u6b21)', color: colors.amber, name: 'West Base' }},
        ],
        badge: 'Base Load',
      }},
      peak: {{
        cats: [
          {{ key: '\u6771\u30fb\u65e5\u4e2d(\u6708\u6b21)', color: colors.cyan, name: 'East Peak' }},
          {{ key: '\u897f\u30fb\u65e5\u4e2d(\u6708\u6b21)', color: colors.pink, name: 'West Peak' }},
        ],
        badge: 'Peak Load',
      }},
      all: {{
        cats: [
          {{ key: '\u6771\u30fb\u30d9\u30fc\u30b9(\u6708\u6b21)', color: colors.blue, name: 'East Base' }},
          {{ key: '\u6771\u30fb\u65e5\u4e2d(\u6708\u6b21)', color: colors.cyan, name: 'East Peak' }},
          {{ key: '\u897f\u30fb\u30d9\u30fc\u30b9(\u6708\u6b21)', color: colors.amber, name: 'West Base' }},
          {{ key: '\u897f\u30fb\u65e5\u4e2d(\u6708\u6b21)', color: colors.pink, name: 'West Peak' }},
        ],
        badge: '4 Areas',
      }},
    }};

    // Collect all months across all curves
    const allMonths = new Set();
    Object.values(overview).forEach(items => {{
      items.forEach(item => allMonths.add(item.month));
    }});
    const months = Array.from(allMonths).sort();
    const monthLabels = months.map(m => m.substring(0, 4) + '/' + m.substring(4));

    let overviewInstance = null;

    function renderOverview(mode) {{
      const cfg = overviewModes[mode];
      document.getElementById('overviewBadge').textContent = cfg.badge;

      const series = [];
      const chartColors = [];
      const dashArr = [];
      const widths = [];

      // Current day curves
      cfg.cats.forEach(c => {{
        const items = overview[c.key] || [];
        const priceMap = {{}};
        items.forEach(item => {{ priceMap[item.month] = item.price; }});
        series.push({{ name: c.name, data: months.map(m => priceMap[m] || null) }});
        chartColors.push(c.color);
        dashArr.push(0);
        widths.push(2.5);
      }});

      // Previous day curves (dashed)
      cfg.cats.forEach(c => {{
        const prevItems = (prevOverview[c.key] || []);
        if (prevItems.length === 0) return;
        const prevMap = {{}};
        prevItems.forEach(item => {{ prevMap[item.month] = item.price; }});
        series.push({{ name: c.name + ' (prev)', data: months.map(m => prevMap[m] || null) }});
        chartColors.push(c.color + '55');
        dashArr.push(5);
        widths.push(1.2);
      }});

      if (overviewInstance) overviewInstance.destroy();

      overviewInstance = new ApexCharts(document.getElementById('overviewChart'), {{
        ...baseChartOpts,
        chart: {{ ...baseChartOpts.chart, type: 'area', height: 380, animations: {{ enabled: true, easing: 'easeinout', speed: 600 }} }},
        series: series,
        colors: chartColors,
        stroke: {{ curve: 'smooth', width: widths, dashArray: dashArr }},
        fill: {{ type: 'gradient', gradient: {{ shadeIntensity: 1, opacityFrom: 0.15, opacityTo: 0.01, stops: [0, 95, 100] }} }},
        xaxis: {{ ...baseChartOpts.xaxis, categories: monthLabels, title: {{ text: 'Contract Month', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
        yaxis: {{ ...baseChartOpts.yaxis, title: {{ text: 'JPY/kWh', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
        markers: {{ size: 3, strokeWidth: 0, hover: {{ size: 6 }} }},
        legend: {{ ...baseChartOpts.legend, position: 'top', horizontalAlign: 'right' }},
        dataLabels: {{ enabled: false }}
      }});
      overviewInstance.render();
    }}

    // Initial render: Base Load
    renderOverview('base');

    // Selector event
    document.getElementById('overviewModeSelect').addEventListener('change', function() {{
      renderOverview(this.value);
    }});

    // ============================================
    // 2. Interactive curve selector
    // ============================================
    const curves = chartData.forward_curves;
    const select = document.getElementById('curveSelect');
    let curveChartInstance = null;

    const prevCurves = chartData.prev_forward_curves || {{}};

    function renderCurve(catName) {{
      const items = curves[catName] || [];
      const prevItems = prevCurves[catName] || [];
      document.getElementById('curveInfo').textContent = items.length + ' contracts';

      const labels = items.map(i => {{
        const m = i.month;
        if (m.length === 6) return m.substring(0, 4) + '/' + m.substring(4);
        if (m.length === 8) return m.substring(0, 4) + '/' + m.substring(4, 6) + '/' + m.substring(6);
        return m;
      }});

      const series = [{{
        name: 'Settlement',
        data: items.map(i => i.settlement)
      }}];

      const chartColors = [colors.blue];
      const dashArrays = [0];
      const widths = [2.5];

      // Previous day curve
      if (prevItems.length > 0) {{
        const prevMap = {{}};
        prevItems.forEach(p => {{ prevMap[p.month] = p.settlement; }});
        series.push({{
          name: 'Prev Day',
          data: items.map(i => prevMap[i.month] || null)
        }});
        chartColors.push(colors.amber);
        dashArrays.push(5);
        widths.push(1.5);
      }}

      const hasTheo = items.some(i => i.theoretical !== null);
      if (hasTheo) {{
        series.push({{
          name: 'Theoretical',
          data: items.map(i => i.theoretical)
        }});
        chartColors.push(colors.purple);
        dashArrays.push(3);
        widths.push(1.2);
      }}

      if (curveChartInstance) curveChartInstance.destroy();

      curveChartInstance = new ApexCharts(document.getElementById('curveChart'), {{
        ...baseChartOpts,
        chart: {{ ...baseChartOpts.chart, type: 'area', height: 320, animations: {{ enabled: true, easing: 'easeinout', speed: 600 }} }},
        series: series,
        colors: chartColors,
        stroke: {{ curve: 'smooth', width: widths, dashArray: dashArrays }},
        fill: {{ type: 'gradient', gradient: {{ shadeIntensity: 1, opacityFrom: 0.2, opacityTo: 0.01, stops: [0, 95, 100] }} }},
        xaxis: {{ ...baseChartOpts.xaxis, categories: labels }},
        yaxis: {{ ...baseChartOpts.yaxis, title: {{ text: 'JPY/kWh', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
        markers: {{ size: 4, strokeWidth: 0, hover: {{ size: 7 }} }},
        title: {{ text: catName, align: 'left', style: {{ fontSize: '14px', fontWeight: 600, color: '#F1F5F9' }} }},
        legend: {{ ...baseChartOpts.legend, position: 'top', horizontalAlign: 'right' }},
        dataLabels: {{ enabled: false }},
        tooltip: {{
          ...baseChartOpts.tooltip,
          custom: function({{ series: s, seriesIndex, dataPointIndex, w }}) {{
            const item = items[dataPointIndex];
            let html = '<div style="padding:8px 12px;font-family:JetBrains Mono,monospace;font-size:12px">';
            html += '<div style="color:#F1F5F9;font-weight:600;margin-bottom:4px">' + item.name + '</div>';
            html += '<div style="color:#94A3B8">Settlement: <span style="color:#3B82F6">' + (item.settlement !== null ? item.settlement.toFixed(2) : 'N/A') + '</span></div>';
            if (item.prev_settlement !== null && item.prev_settlement !== undefined) {{
              html += '<div style="color:#94A3B8">Prev Day: <span style="color:#F59E0B">' + item.prev_settlement.toFixed(2) + '</span></div>';
            }}
            if (item.change_diff !== null && item.change_diff !== undefined) {{
              const chgColor = item.change_diff >= 0 ? '#22C55E' : '#EF4444';
              const chgSign = item.change_diff >= 0 ? '+' : '';
              html += '<div style="color:' + chgColor + ';font-weight:600">' + chgSign + item.change_diff.toFixed(2) + ' (' + chgSign + item.change_pct.toFixed(1) + '%)</div>';
            }}
            if (item.theoretical !== null) {{
              html += '<div style="color:#94A3B8">Theoretical: <span style="color:#8B5CF6">' + item.theoretical.toFixed(2) + '</span></div>';
            }}
            html += '<div style="color:#64748B;margin-top:2px">DTE: ' + (item.days || 'N/A') + '</div>';
            html += '</div>';
            return html;
          }}
        }}
      }});
      curveChartInstance.render();

      // Render detail table with change column
      let tableHtml = '<table><thead><tr><th>Instrument</th><th>Month</th><th class="num">Settlement</th><th class="num">DoD Change</th><th class="num">Theoretical</th><th class="num">DTE</th></tr></thead><tbody>';
      items.forEach(i => {{
        const s = i.settlement !== null ? i.settlement : '';
        const t = i.theoretical !== null ? i.theoretical : '';
        const d = i.days !== null ? i.days : '';
        let chg = '-';
        let chgClass = '';
        if (i.change_diff !== null && i.change_diff !== undefined) {{
          const sign = i.change_diff >= 0 ? '+' : '';
          chgClass = i.change_diff > 0 ? 'positive' : i.change_diff < 0 ? 'negative' : '';
          chg = sign + i.change_diff.toFixed(2) + ' <span class="change-pct">(' + sign + i.change_pct.toFixed(1) + '%)</span>';
        }}
        tableHtml += '<tr><td>' + i.name + '</td><td>' + i.month + '</td><td class="num">' + s + '</td><td class="num change ' + chgClass + '">' + chg + '</td><td class="num">' + t + '</td><td class="num">' + d + '</td></tr>';
      }});
      tableHtml += '</tbody></table>';
      document.getElementById('curveTable').innerHTML = tableHtml;
    }}

    select.addEventListener('change', () => renderCurve(select.value));
    if (select.options.length > 0) renderCurve(select.value);

    // ============================================
    // 3. East vs West spread chart
    // ============================================
    const eastBase = curves['\u6771\u30fb\u30d9\u30fc\u30b9(\u6708\u6b21)'] || [];
    const westBase = curves['\u897f\u30fb\u30d9\u30fc\u30b9(\u6708\u6b21)'] || [];

    if (eastBase.length && westBase.length) {{
      const westMap = {{}};
      westBase.forEach(w => {{ westMap[w.month] = w.settlement; }});

      const spreadLabels = [];
      const spreadValues = [];
      const spreadColors = [];

      eastBase.forEach(e => {{
        if (e.settlement !== null && westMap[e.month] !== undefined && westMap[e.month] !== null) {{
          const spread = Math.round((e.settlement - westMap[e.month]) * 100) / 100;
          spreadLabels.push(e.month.substring(0, 4) + '/' + e.month.substring(4));
          spreadValues.push(spread);
          spreadColors.push(spread >= 0 ? colors.blue : colors.red);
        }}
      }});

      new ApexCharts(document.getElementById('spreadChart'), {{
        ...baseChartOpts,
        chart: {{ ...baseChartOpts.chart, type: 'bar', height: 320, animations: {{ enabled: true, easing: 'easeinout', speed: 600 }} }},
        series: [{{ name: 'E-W Spread', data: spreadValues }}],
        plotOptions: {{
          bar: {{
            borderRadius: 4,
            columnWidth: '65%',
            colors: {{
              ranges: [
                {{ from: -9999, to: -0.001, color: colors.red }},
                {{ from: 0, to: 9999, color: colors.blue }}
              ]
            }}
          }}
        }},
        colors: [colors.blue],
        xaxis: {{ ...baseChartOpts.xaxis, categories: spreadLabels }},
        yaxis: {{ ...baseChartOpts.yaxis, title: {{ text: 'JPY/kWh', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
        dataLabels: {{ enabled: false }},
        tooltip: {{
          ...baseChartOpts.tooltip,
          y: {{ formatter: v => (v >= 0 ? '+' : '') + v.toFixed(2) + ' JPY/kWh' }}
        }}
      }}).render();
    }}

    // ============================================
    // 3b. Power Price Change Heatmap
    // ============================================
    const heatmapData = chartData.power_heatmap || {{}};
    if (heatmapData.price_changes && document.getElementById('priceHeatmapChart')) {{
      const typeKeys = ['\u6771\u30fb\u30d9\u30fc\u30b9', '\u6771\u30fb\u65e5\u4e2d', '\u897f\u30fb\u30d9\u30fc\u30b9', '\u897f\u30fb\u65e5\u4e2d'];
      const typeLabels = ['East Base', 'East Peak', 'West Base', 'West Peak'];

      // ApexCharts heatmap: series = rows (reversed for display), data = columns
      const priceSeries = typeLabels.map((label, idx) => {{
        const items = heatmapData.price_changes[typeKeys[idx]] || [];
        return {{
          name: label,
          data: items.map(it => ({{
            x: it.month.substring(0, 4) + '/' + it.month.substring(4),
            y: it.pct !== null && it.pct !== undefined ? it.pct : 0,
            settlement: it.settlement,
            prev: it.prev,
            diff: it.diff,
          }}))
        }};
      }}).reverse();

      new ApexCharts(document.getElementById('priceHeatmapChart'), {{
        chart: {{
          type: 'heatmap',
          height: 420,
          background: 'transparent',
          fontFamily: "'Inter', sans-serif",
          toolbar: {{ show: true, tools: {{ download: true, selection: false, zoom: false, zoomin: false, zoomout: false, pan: false, reset: false }} }},
        }},
        theme: {{ mode: 'dark' }},
        series: priceSeries,
        plotOptions: {{
          heatmap: {{
            shadeIntensity: 0,
            radius: 2,
            enableShades: false,
            colorScale: {{
              ranges: [
                {{ from: -20, to: -5, color: '#DC2626', name: '< -5%' }},
                {{ from: -5, to: -2, color: '#EF4444', name: '-5% to -2%' }},
                {{ from: -2, to: -0.5, color: '#F87171', name: '-2% to -0.5%' }},
                {{ from: -0.5, to: 0.5, color: '#374151', name: 'Flat' }},
                {{ from: 0.5, to: 2, color: '#34D399', name: '+0.5% to +2%' }},
                {{ from: 2, to: 5, color: '#10B981', name: '+2% to +5%' }},
                {{ from: 5, to: 20, color: '#059669', name: '> +5%' }},
              ]
            }}
          }}
        }},
        dataLabels: {{
          enabled: true,
          formatter: function(val) {{ return val !== 0 ? (val > 0 ? '+' : '') + val.toFixed(1) + '%' : '-'; }},
          style: {{ fontSize: '10px', fontFamily: 'JetBrains Mono, monospace', fontWeight: 500, colors: ['#E2E8F0'] }}
        }},
        grid: {{ show: false }},
        xaxis: {{
          labels: {{ style: {{ colors: colors.text, fontSize: '10px' }}, rotate: -45, rotateAlways: false }},
          position: 'top',
          axisBorder: {{ show: false }},
          axisTicks: {{ show: false }},
        }},
        yaxis: {{
          labels: {{ style: {{ colors: colors.text, fontSize: '11px' }} }}
        }},
        legend: {{
          show: true,
          position: 'bottom',
          labels: {{ colors: colors.text }},
          fontSize: '11px',
        }},
        tooltip: {{
          custom: function({{ seriesIndex, dataPointIndex, w }}) {{
            const d = w.config.series[seriesIndex].data[dataPointIndex];
            const name = w.config.series[seriesIndex].name;
            const pctColor = d.y >= 0 ? '#10B981' : '#EF4444';
            const sign = d.y >= 0 ? '+' : '';
            let html = '<div style="padding:10px 14px;font-family:JetBrains Mono,monospace;font-size:12px;background:#151B28;border:1px solid #1E293B">';
            html += '<div style="color:#F1F5F9;font-weight:600;margin-bottom:6px">' + name + ' — ' + d.x + '</div>';
            html += '<div style="color:#94A3B8">Settlement: <span style="color:#E2E8F0">' + (d.settlement !== null ? d.settlement.toFixed(2) : 'N/A') + ' JPY/kWh</span></div>';
            if (d.prev !== null && d.prev !== undefined) {{
              html += '<div style="color:#94A3B8">Previous: <span style="color:#94A3B8">' + d.prev.toFixed(2) + '</span></div>';
            }}
            if (d.diff !== null && d.diff !== undefined) {{
              html += '<div style="color:#94A3B8">Change: <span style="color:' + pctColor + '">' + (d.diff >= 0 ? '+' : '') + d.diff.toFixed(2) + '</span></div>';
            }}
            html += '<div style="color:' + pctColor + ';font-weight:600;font-size:13px;margin-top:4px">' + sign + d.y.toFixed(2) + '%</div>';
            html += '</div>';
            return html;
          }}
        }}
      }}).render();
    }}

    // ============================================
    // 3c. E-W Spread Heatmap
    // ============================================
    if (heatmapData.spreads && document.getElementById('spreadHeatmapChart')) {{
      const baseSpreads = heatmapData.spreads.base || [];
      const peakSpreads = heatmapData.spreads.peak || [];

      const spreadSeries = [
        {{
          name: 'Peak Spread Change',
          data: peakSpreads.map(s => ({{
            x: s.month.substring(0, 4) + '/' + s.month.substring(4),
            y: s.spread_change !== null ? s.spread_change : 0,
            spread: s.spread,
            prevSpread: s.prev_spread,
          }}))
        }},
        {{
          name: 'Base Spread Change',
          data: baseSpreads.map(s => ({{
            x: s.month.substring(0, 4) + '/' + s.month.substring(4),
            y: s.spread_change !== null ? s.spread_change : 0,
            spread: s.spread,
            prevSpread: s.prev_spread,
          }}))
        }},
        {{
          name: 'Peak Spread (JPY/kWh)',
          data: peakSpreads.map(s => ({{
            x: s.month.substring(0, 4) + '/' + s.month.substring(4),
            y: s.spread !== null ? s.spread : 0,
            prevSpread: s.prev_spread,
            spreadChange: s.spread_change,
          }}))
        }},
        {{
          name: 'Base Spread (JPY/kWh)',
          data: baseSpreads.map(s => ({{
            x: s.month.substring(0, 4) + '/' + s.month.substring(4),
            y: s.spread !== null ? s.spread : 0,
            prevSpread: s.prev_spread,
            spreadChange: s.spread_change,
          }}))
        }},
      ];

      new ApexCharts(document.getElementById('spreadHeatmapChart'), {{
        chart: {{
          type: 'heatmap',
          height: 420,
          background: 'transparent',
          fontFamily: "'Inter', sans-serif",
          toolbar: {{ show: true, tools: {{ download: true, selection: false, zoom: false, zoomin: false, zoomout: false, pan: false, reset: false }} }},
        }},
        theme: {{ mode: 'dark' }},
        series: spreadSeries,
        plotOptions: {{
          heatmap: {{
            shadeIntensity: 0,
            radius: 2,
            enableShades: false,
            colorScale: {{
              ranges: [
                {{ from: -3, to: -1, color: '#DC2626', name: '< -1.0' }},
                {{ from: -1, to: -0.3, color: '#F87171', name: '-1.0 to -0.3' }},
                {{ from: -0.3, to: 0.3, color: '#374151', name: 'Flat' }},
                {{ from: 0.3, to: 1, color: '#60A5FA', name: '+0.3 to +1.0' }},
                {{ from: 1, to: 3, color: '#2563EB', name: '+1.0 to +3.0' }},
                {{ from: 3, to: 10, color: '#1D4ED8', name: '> +3.0' }},
              ]
            }}
          }}
        }},
        dataLabels: {{
          enabled: true,
          formatter: function(val) {{ return val !== 0 ? (val > 0 ? '+' : '') + val.toFixed(1) : '-'; }},
          style: {{ fontSize: '10px', fontFamily: 'JetBrains Mono, monospace', fontWeight: 500, colors: ['#E2E8F0'] }}
        }},
        grid: {{ show: false }},
        xaxis: {{
          labels: {{ style: {{ colors: colors.text, fontSize: '10px' }}, rotate: -45, rotateAlways: false }},
          position: 'top',
          axisBorder: {{ show: false }},
          axisTicks: {{ show: false }},
        }},
        yaxis: {{
          labels: {{ style: {{ colors: colors.text, fontSize: '11px' }} }}
        }},
        legend: {{
          show: true,
          position: 'bottom',
          labels: {{ colors: colors.text }},
          fontSize: '11px',
        }},
        tooltip: {{
          custom: function({{ seriesIndex, dataPointIndex, w }}) {{
            const d = w.config.series[seriesIndex].data[dataPointIndex];
            const name = w.config.series[seriesIndex].name;
            let html = '<div style="padding:10px 14px;font-family:JetBrains Mono,monospace;font-size:12px;background:#151B28;border:1px solid #1E293B">';
            html += '<div style="color:#F1F5F9;font-weight:600;margin-bottom:6px">' + name + ' — ' + d.x + '</div>';
            if (name.includes('Change')) {{
              if (d.spread !== undefined && d.spread !== null) {{
                html += '<div style="color:#94A3B8">Today Spread: <span style="color:#E2E8F0">' + d.spread.toFixed(2) + ' JPY/kWh</span></div>';
              }}
              if (d.prevSpread !== undefined && d.prevSpread !== null) {{
                html += '<div style="color:#94A3B8">Prev Spread: <span style="color:#94A3B8">' + d.prevSpread.toFixed(2) + '</span></div>';
              }}
              const chgColor = d.y >= 0 ? '#10B981' : '#EF4444';
              html += '<div style="color:' + chgColor + ';font-weight:600;font-size:13px;margin-top:4px">Change: ' + (d.y >= 0 ? '+' : '') + d.y.toFixed(2) + '</div>';
            }} else {{
              html += '<div style="color:#E2E8F0;font-weight:600;font-size:14px">' + d.y.toFixed(2) + ' JPY/kWh</div>';
              if (d.prevSpread !== undefined && d.prevSpread !== null) {{
                html += '<div style="color:#94A3B8">Prev: ' + d.prevSpread.toFixed(2) + '</div>';
              }}
              if (d.spreadChange !== undefined && d.spreadChange !== null) {{
                const chgColor = d.spreadChange >= 0 ? '#10B981' : '#EF4444';
                html += '<div style="color:' + chgColor + '">DoD: ' + (d.spreadChange >= 0 ? '+' : '') + d.spreadChange.toFixed(2) + '</div>';
              }}
            }}
            html += '</div>';
            return html;
          }}
        }}
      }}).render();
    }}

    // ============================================
    // 4. Commodity Forward Curves
    // ============================================
    const commodityCurves = chartData.commodity_curves || {{}};
    const commoditySelect = document.getElementById('commoditySelect');
    let commodityCurveInstance = null;

    function renderCommodityCurve(underlyingName) {{
      const data = commodityCurves[underlyingName];
      if (!data) return;

      const curveData = data.curve || [];
      const prevCurveData = data.prev_curve || [];
      const unit = data.unit || '';

      document.getElementById('commodityCurveInfo').textContent = curveData.length + ' contracts | ' + unit;
      document.getElementById('commodityCurveBadge').textContent = data.display_en;

      const labels = curveData.map(i => {{
        const m = i.month;
        if (m.length === 6) return m.substring(0, 4) + '/' + m.substring(4);
        if (m.length === 8) return m.substring(0, 4) + '/' + m.substring(4, 6) + '/' + m.substring(6);
        return m;
      }});

      const series = [{{
        name: 'Settlement (' + data.display_en + ')',
        data: curveData.map(i => i.settlement)
      }}];

      const catColor = (categoryMeta[data.category] || {{}}).color || colors.blue;
      const chartColors = [catColor];
      const dashArrays = [0];
      const widths = [2.5];

      // Previous day curve overlay
      if (prevCurveData.length > 0) {{
        const prevMap = {{}};
        prevCurveData.forEach(p => {{ prevMap[p.month] = p.settlement; }});
        series.push({{
          name: 'Prev Day',
          data: curveData.map(i => prevMap[i.month] || null)
        }});
        chartColors.push(colors.amber);
        dashArrays.push(5);
        widths.push(1.5);
      }}

      if (commodityCurveInstance) commodityCurveInstance.destroy();

      commodityCurveInstance = new ApexCharts(document.getElementById('commodityCurveChart'), {{
        ...baseChartOpts,
        chart: {{ ...baseChartOpts.chart, type: 'area', height: 380, animations: {{ enabled: true, easing: 'easeinout', speed: 600 }} }},
        series: series,
        colors: chartColors,
        stroke: {{ curve: 'smooth', width: widths, dashArray: dashArrays }},
        fill: {{ type: 'gradient', gradient: {{ shadeIntensity: 1, opacityFrom: 0.2, opacityTo: 0.01, stops: [0, 95, 100] }} }},
        xaxis: {{ ...baseChartOpts.xaxis, categories: labels, title: {{ text: 'Contract Month', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
        yaxis: {{
          ...baseChartOpts.yaxis,
          title: {{ text: unit, style: {{ color: colors.textLight, fontSize: '11px' }} }},
          labels: {{
            style: {{ colors: colors.text, fontSize: '11px', fontFamily: "'JetBrains Mono', monospace" }},
            formatter: v => v !== null ? v.toLocaleString() : ''
          }}
        }},
        markers: {{ size: 4, strokeWidth: 0, hover: {{ size: 7 }} }},
        title: {{ text: data.display_en + ' (' + data.display_ja + ')', align: 'left', style: {{ fontSize: '14px', fontWeight: 600, color: '#F1F5F9' }} }},
        legend: {{ ...baseChartOpts.legend, position: 'top', horizontalAlign: 'right' }},
        dataLabels: {{ enabled: false }},
        tooltip: {{
          theme: 'dark',
          style: {{ fontSize: '12px', fontFamily: "'JetBrains Mono', monospace" }},
          custom: function({{ series: s, seriesIndex, dataPointIndex, w }}) {{
            const item = curveData[dataPointIndex];
            let html = '<div style="padding:8px 12px;font-family:JetBrains Mono,monospace;font-size:12px">';
            html += '<div style="color:#F1F5F9;font-weight:600;margin-bottom:4px">' + (item.name || item.month) + '</div>';
            html += '<div style="color:#94A3B8">Settlement: <span style="color:' + catColor + '">' + (item.settlement !== null ? item.settlement.toLocaleString() : 'N/A') + ' ' + unit + '</span></div>';
            if (item.theoretical !== null && item.theoretical !== undefined) {{
              html += '<div style="color:#94A3B8">Theoretical: <span style="color:#8B5CF6">' + item.theoretical.toLocaleString() + '</span></div>';
            }}
            html += '<div style="color:#64748B;margin-top:2px">DTE: ' + (item.days || 'N/A') + '</div>';
            html += '</div>';
            return html;
          }}
        }}
      }});
      commodityCurveInstance.render();
    }}

    if (commoditySelect) {{
      commoditySelect.addEventListener('change', () => renderCommodityCurve(commoditySelect.value));
      if (commoditySelect.options.length > 0) renderCommodityCurve(commoditySelect.value);
    }}

    // ============================================
    // 5. Cross-Market Daily Changes
    // ============================================
    const commoditySnapshot = chartData.commodity_snapshot || [];
    if (commoditySnapshot.length > 0 && document.getElementById('crossMarketChart')) {{
      const sorted = [...commoditySnapshot]
        .filter(a => a.change_pct !== null && a.change_pct !== undefined)
        .sort((a, b) => Math.abs(b.change_pct) - Math.abs(a.change_pct));

      if (sorted.length > 0) {{
        const cmLabels = sorted.map(s => s.display_en);
        const cmValues = sorted.map(s => s.change_pct);
        const cmColors = sorted.map(s => s.change_pct >= 0 ? colors.green : colors.red);

        new ApexCharts(document.getElementById('crossMarketChart'), {{
          ...baseChartOpts,
          chart: {{ ...baseChartOpts.chart, type: 'bar', height: Math.max(300, sorted.length * 32), animations: {{ enabled: true, easing: 'easeinout', speed: 600 }} }},
          series: [{{ name: 'Daily Change %', data: cmValues }}],
          plotOptions: {{
            bar: {{
              horizontal: true,
              borderRadius: 4,
              barHeight: '65%',
              distributed: true,
              colors: {{
                ranges: [
                  {{ from: -9999, to: -0.001, color: colors.red }},
                  {{ from: 0, to: 9999, color: colors.green }}
                ]
              }}
            }}
          }},
          colors: cmColors,
          xaxis: {{
            ...baseChartOpts.xaxis,
            categories: cmLabels,
            title: {{ text: 'Change (%)', style: {{ color: colors.textLight, fontSize: '11px' }} }},
            labels: {{
              style: {{ colors: colors.text, fontSize: '11px', fontFamily: "'JetBrains Mono', monospace" }},
              formatter: v => (v >= 0 ? '+' : '') + v.toFixed(1) + '%'
            }}
          }},
          yaxis: {{
            labels: {{
              style: {{ colors: colors.text, fontSize: '11px', fontFamily: "'Inter', sans-serif" }}
            }}
          }},
          dataLabels: {{
            enabled: true,
            formatter: function(v) {{ return (v >= 0 ? '+' : '') + v.toFixed(2) + '%'; }},
            style: {{ fontSize: '11px', fontFamily: "'JetBrains Mono', monospace", colors: ['#F1F5F9'] }},
            offsetX: 6
          }},
          legend: {{ show: false }},
          tooltip: {{
            theme: 'dark',
            style: {{ fontSize: '12px', fontFamily: "'JetBrains Mono', monospace" }},
            custom: function({{ series: s, seriesIndex, dataPointIndex }}) {{
              const item = sorted[dataPointIndex];
              const chgColor = item.change_pct >= 0 ? colors.green : colors.red;
              const sign = item.change_pct >= 0 ? '+' : '';
              const diffSign = item.change_diff >= 0 ? '+' : '';
              return '<div style="padding:8px 12px;font-family:JetBrains Mono,monospace;font-size:12px">'
                + '<div style="color:#F1F5F9;font-weight:600;margin-bottom:4px">' + item.display_en + '</div>'
                + '<div style="color:#94A3B8">' + item.display_ja + '</div>'
                + '<div style="color:#94A3B8">Price: <span style="color:#3B82F6">' + item.settlement.toLocaleString() + ' ' + item.unit + '</span></div>'
                + '<div style="color:' + chgColor + ';font-weight:600">' + diffSign + (item.change_diff || 0).toFixed(2) + ' (' + sign + item.change_pct.toFixed(2) + '%)</div>'
                + '</div>';
            }}
          }}
        }}).render();
      }}
    }}

    // ============================================
    // 6. KPI count-up animation
    // ============================================
    function animateValue(el, end) {{
      const duration = 1200;
      const start = 0;
      const range = end - start;
      const startTime = performance.now();
      function update(currentTime) {{
        const elapsed = currentTime - startTime;
        const progress = Math.min(elapsed / duration, 1);
        const eased = 1 - Math.pow(1 - progress, 3);
        const current = Math.round(start + range * eased);
        el.textContent = current.toLocaleString();
        if (progress < 1) requestAnimationFrame(update);
      }}
      requestAnimationFrame(update);
    }}

    document.querySelectorAll('.kpi-value[data-target]').forEach(el => {{
      const target = parseInt(el.dataset.target, 10);
      if (!isNaN(target)) animateValue(el, target);
    }});

    // ============================================
    // 7. Sidebar active state on scroll
    // ============================================
    const sections = document.querySelectorAll('[id]');
    const navItems = document.querySelectorAll('.nav-item');

    const observer = new IntersectionObserver(entries => {{
      entries.forEach(entry => {{
        if (entry.isIntersecting) {{
          navItems.forEach(item => item.classList.remove('active'));
          const target = document.querySelector('.nav-item[href="#' + entry.target.id + '"]');
          if (target) target.classList.add('active');
        }}
      }});
    }}, {{ threshold: 0.3, rootMargin: '-80px 0px -50% 0px' }});

    sections.forEach(sec => {{
      if (sec.id && document.querySelector('.nav-item[href="#' + sec.id + '"]')) {{
        observer.observe(sec);
      }}
    }});

    // ============================================
    // 8. Spot Analysis — CSV Upload & Processing
    // ============================================
    const uploadZone = document.getElementById('uploadZone');
    const spotFileInput = document.getElementById('spotFileInput');
    const uploadBtn = document.getElementById('uploadBtn');

    uploadBtn.addEventListener('click', () => spotFileInput.click());
    spotFileInput.addEventListener('change', e => {{
      if (e.target.files.length > 0) handleSpotFile(e.target.files[0]);
    }});

    uploadZone.addEventListener('dragover', e => {{
      e.preventDefault();
      uploadZone.classList.add('drag-over');
    }});
    uploadZone.addEventListener('dragleave', () => {{
      uploadZone.classList.remove('drag-over');
    }});
    uploadZone.addEventListener('drop', e => {{
      e.preventDefault();
      uploadZone.classList.remove('drag-over');
      if (e.dataTransfer.files.length > 0) handleSpotFile(e.dataTransfer.files[0]);
    }});

    function handleSpotFile(file) {{
      const reader = new FileReader();
      reader.onload = function(e) {{
        // Decode CP932 (Shift-JIS)
        const decoder = new TextDecoder('shift-jis');
        const text = decoder.decode(new Uint8Array(e.target.result));
        const rows = parseSpotCSV(text);
        if (rows.length === 0) {{
          alert('CSV parsing failed. Please check the file format.');
          return;
        }}
        processSpotData(rows);
      }};
      reader.readAsArrayBuffer(file);
    }}

    function parseSpotCSV(text) {{
      const lines = text.split('\\n').filter(l => l.trim());
      if (lines.length < 2) return [];
      // Skip header
      const rows = [];
      for (let i = 1; i < lines.length; i++) {{
        const cols = lines[i].split(',');
        if (cols.length < 15) continue;
        const dateStr = cols[0].trim();  // 2025/04/01
        const date = dateStr.replace(/\\//g, '-');  // 2025-04-01
        const timeCode = parseInt(cols[1], 10);
        const systemPrice = parseFloat(cols[5]) || null;
        const tokyo = parseFloat(cols[8]) || null;
        const kansai = parseFloat(cols[11]) || null;
        if (!date || isNaN(timeCode)) continue;
        rows.push({{ date, timeCode, systemPrice, tokyo, kansai }});
      }}
      return rows;
    }}

    function processSpotData(rows) {{
      // Monthly averages
      const monthly = calcMonthlyAverages(rows);
      // Intraday profile
      const intraday = calcIntradayProfile(rows);
      // Futures comparison
      const comparison = calcFuturesComparison(monthly);

      // Show results section
      document.getElementById('spotResults').style.display = 'block';
      uploadZone.style.borderColor = 'var(--positive)';
      uploadZone.querySelector('.upload-title').textContent = 'CSV Loaded (' + rows.length.toLocaleString() + ' rows)';

      // KPI cards
      const allTokyo = rows.filter(r => r.tokyo !== null).map(r => r.tokyo);
      const allKansai = rows.filter(r => r.kansai !== null).map(r => r.kansai);
      const peakTokyo = rows.filter(r => r.tokyo !== null && r.timeCode >= 17 && r.timeCode <= 40).map(r => r.tokyo);
      const avg = arr => arr.length ? (arr.reduce((a, b) => a + b, 0) / arr.length) : 0;

      document.getElementById('spotKpiTokyoBase').textContent = avg(allTokyo).toFixed(2);
      document.getElementById('spotKpiTokyoPeak').textContent = avg(peakTokyo).toFixed(2);
      document.getElementById('spotKpiKansaiBase').textContent = avg(allKansai).toFixed(2);

      const dates = [...new Set(rows.map(r => r.date))].sort();
      document.getElementById('spotKpiPeriod').textContent = dates[0] + ' ~ ' + dates[dates.length - 1];

      // Render charts
      renderSpotMonthlyChart(monthly);
      renderIntradayChart(intraday);
      renderFuturesVsSpotChart(comparison, 'east_base');
      renderPremiumChart(comparison, 'east_base');

      // Category selector
      document.getElementById('spotCategorySelect').addEventListener('change', function() {{
        renderFuturesVsSpotChart(comparison, this.value);
        renderPremiumChart(comparison, this.value);
      }});

      // Scroll to spot section
      document.getElementById('spot').scrollIntoView({{ behavior: 'smooth' }});
    }}

    function calcMonthlyAverages(rows) {{
      const months = {{}};
      rows.forEach(r => {{
        const ym = r.date.substring(0, 7);  // YYYY-MM
        if (!months[ym]) months[ym] = {{ tokyoBase: [], tokyoPeak: [], kansaiBase: [], kansaiPeak: [] }};
        if (r.tokyo !== null) {{
          months[ym].tokyoBase.push(r.tokyo);
          if (r.timeCode >= 17 && r.timeCode <= 40) months[ym].tokyoPeak.push(r.tokyo);
        }}
        if (r.kansai !== null) {{
          months[ym].kansaiBase.push(r.kansai);
          if (r.timeCode >= 17 && r.timeCode <= 40) months[ym].kansaiPeak.push(r.kansai);
        }}
      }});
      const avg = arr => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;
      const result = {{}};
      Object.keys(months).sort().forEach(ym => {{
        result[ym] = {{
          tokyoBase: avg(months[ym].tokyoBase),
          tokyoPeak: avg(months[ym].tokyoPeak),
          kansaiBase: avg(months[ym].kansaiBase),
          kansaiPeak: avg(months[ym].kansaiPeak)
        }};
      }});
      return result;
    }}

    function calcIntradayProfile(rows) {{
      const slots = {{}};
      for (let i = 1; i <= 48; i++) slots[i] = {{ tokyo: [], kansai: [] }};
      rows.forEach(r => {{
        if (r.tokyo !== null) slots[r.timeCode].tokyo.push(r.tokyo);
        if (r.kansai !== null) slots[r.timeCode].kansai.push(r.kansai);
      }});
      const avg = arr => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0;
      const result = [];
      for (let i = 1; i <= 48; i++) {{
        const h = Math.floor((i - 1) / 2);
        const m = (i - 1) % 2 === 0 ? '00' : '30';
        result.push({{
          code: i,
          label: String(h).padStart(2, '0') + ':' + m,
          tokyo: avg(slots[i].tokyo),
          kansai: avg(slots[i].kansai)
        }});
      }}
      return result;
    }}

    function calcFuturesComparison(spotMonthly) {{
      // Map futures data from CHART_DATA to spot monthly
      const futuresCats = {{
        east_base: {{ chartKey: '\u6771\u30fb\u30d9\u30fc\u30b9(\u6708\u6b21)', spotKey: 'tokyoBase', label: 'East Base' }},
        east_peak: {{ chartKey: '\u6771\u30fb\u65e5\u4e2d(\u6708\u6b21)', spotKey: 'tokyoPeak', label: 'East Peak' }},
        west_base: {{ chartKey: '\u897f\u30fb\u30d9\u30fc\u30b9(\u6708\u6b21)', spotKey: 'kansaiBase', label: 'West Base' }},
        west_peak: {{ chartKey: '\u897f\u30fb\u65e5\u4e2d(\u6708\u6b21)', spotKey: 'kansaiPeak', label: 'West Peak' }}
      }};

      const result = {{}};
      Object.entries(futuresCats).forEach(([key, cfg]) => {{
        const futuresItems = (chartData.forward_curves[cfg.chartKey] || []);
        const matches = [];
        futuresItems.forEach(f => {{
          // f.month is YYYYMM, spot key is YYYY-MM
          const ym = f.month.substring(0, 4) + '-' + f.month.substring(4, 6);
          const spotVal = spotMonthly[ym] ? spotMonthly[ym][cfg.spotKey] : null;
          if (f.settlement !== null && spotVal !== null) {{
            const premium = f.settlement - spotVal;
            const premiumPct = (premium / spotVal) * 100;
            matches.push({{
              month: f.month,
              monthLabel: f.month.substring(0, 4) + '/' + f.month.substring(4),
              futures: f.settlement,
              spot: Math.round(spotVal * 100) / 100,
              premium: Math.round(premium * 100) / 100,
              premiumPct: Math.round(premiumPct * 100) / 100
            }});
          }}
        }});
        result[key] = {{ label: cfg.label, matches }};
      }});
      return result;
    }}

    // ============================================
    // Spot Chart Renderers
    // ============================================
    let spotMonthlyInstance = null;
    let spotIntradayInstance = null;
    let spotVsFuturesInstance = null;
    let spotPremiumInstance = null;

    function renderSpotMonthlyChart(monthly) {{
      if (spotMonthlyInstance) spotMonthlyInstance.destroy();
      const months = Object.keys(monthly);
      const labels = months.map(m => m);
      spotMonthlyInstance = new ApexCharts(document.getElementById('spotMonthlyChart'), {{
        ...baseChartOpts,
        chart: {{ ...baseChartOpts.chart, type: 'area', height: 380, animations: {{ enabled: true, easing: 'easeinout', speed: 800 }} }},
        series: [
          {{ name: 'Tokyo Base', data: months.map(m => monthly[m].tokyoBase ? +monthly[m].tokyoBase.toFixed(2) : null) }},
          {{ name: 'Tokyo Peak', data: months.map(m => monthly[m].tokyoPeak ? +monthly[m].tokyoPeak.toFixed(2) : null) }},
          {{ name: 'Kansai Base', data: months.map(m => monthly[m].kansaiBase ? +monthly[m].kansaiBase.toFixed(2) : null) }},
          {{ name: 'Kansai Peak', data: months.map(m => monthly[m].kansaiPeak ? +monthly[m].kansaiPeak.toFixed(2) : null) }}
        ],
        colors: [colors.blue, colors.cyan, colors.amber, colors.pink],
        stroke: {{ curve: 'smooth', width: 2.5 }},
        fill: {{ type: 'gradient', gradient: {{ shadeIntensity: 1, opacityFrom: 0.15, opacityTo: 0.01, stops: [0, 95, 100] }} }},
        xaxis: {{ ...baseChartOpts.xaxis, categories: labels, title: {{ text: 'Month', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
        yaxis: {{ ...baseChartOpts.yaxis, title: {{ text: 'JPY/kWh', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
        markers: {{ size: 4, strokeWidth: 0, hover: {{ size: 7 }} }},
        legend: {{ ...baseChartOpts.legend, position: 'top', horizontalAlign: 'right' }},
        dataLabels: {{ enabled: false }}
      }});
      spotMonthlyInstance.render();
    }}

    function renderIntradayChart(intraday) {{
      if (spotIntradayInstance) spotIntradayInstance.destroy();
      spotIntradayInstance = new ApexCharts(document.getElementById('spotIntradayChart'), {{
        ...baseChartOpts,
        chart: {{ ...baseChartOpts.chart, type: 'area', height: 350, animations: {{ enabled: true, easing: 'easeinout', speed: 800 }} }},
        series: [
          {{ name: 'Tokyo', data: intraday.map(s => +s.tokyo.toFixed(2)) }},
          {{ name: 'Kansai', data: intraday.map(s => +s.kansai.toFixed(2)) }}
        ],
        colors: [colors.blue, colors.amber],
        stroke: {{ curve: 'smooth', width: 2.5 }},
        fill: {{ type: 'gradient', gradient: {{ shadeIntensity: 1, opacityFrom: 0.2, opacityTo: 0.01, stops: [0, 95, 100] }} }},
        xaxis: {{
          ...baseChartOpts.xaxis,
          categories: intraday.map(s => s.label),
          title: {{ text: 'Time of Day', style: {{ color: colors.textLight, fontSize: '11px' }} }},
          tickAmount: 12,
          labels: {{ ...baseChartOpts.xaxis.labels, rotate: -45, rotateAlways: false }}
        }},
        yaxis: {{ ...baseChartOpts.yaxis, title: {{ text: 'JPY/kWh', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
        annotations: {{
          xaxis: [{{
            x: '08:00', x2: '20:00',
            fillColor: 'rgba(59,130,246,0.06)',
            borderColor: 'transparent',
            label: {{ text: 'Peak Hours', style: {{ color: colors.textLight, fontSize: '10px', background: 'transparent' }} }}
          }}]
        }},
        markers: {{ size: 0 }},
        legend: {{ ...baseChartOpts.legend, position: 'top', horizontalAlign: 'right' }},
        dataLabels: {{ enabled: false }}
      }});
      spotIntradayInstance.render();
    }}

    function renderFuturesVsSpotChart(comparison, category) {{
      if (spotVsFuturesInstance) spotVsFuturesInstance.destroy();
      const data = comparison[category];
      if (!data || data.matches.length === 0) {{
        document.getElementById('spotVsFuturesChart').innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:3rem">No matching months found for ' + data.label + '</div>';
        return;
      }}
      const labels = data.matches.map(m => m.monthLabel);
      spotVsFuturesInstance = new ApexCharts(document.getElementById('spotVsFuturesChart'), {{
        ...baseChartOpts,
        chart: {{ ...baseChartOpts.chart, type: 'bar', height: 380, animations: {{ enabled: true, easing: 'easeinout', speed: 600 }} }},
        series: [
          {{ name: 'Futures', data: data.matches.map(m => m.futures) }},
          {{ name: 'Spot', data: data.matches.map(m => m.spot) }}
        ],
        colors: [colors.blue, colors.amber],
        plotOptions: {{ bar: {{ borderRadius: 4, columnWidth: '55%', dataLabels: {{ position: 'top' }} }} }},
        xaxis: {{ ...baseChartOpts.xaxis, categories: labels, title: {{ text: 'Contract Month', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
        yaxis: {{ ...baseChartOpts.yaxis, title: {{ text: 'JPY/kWh', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
        legend: {{ ...baseChartOpts.legend, position: 'top', horizontalAlign: 'right' }},
        dataLabels: {{ enabled: true, formatter: v => v.toFixed(1), style: {{ fontSize: '10px', fontFamily: 'JetBrains Mono', colors: ['#F1F5F9'] }}, offsetY: -18 }},
        title: {{ text: data.label + ' — Futures vs Spot', align: 'left', style: {{ fontSize: '14px', fontWeight: 600, color: '#F1F5F9' }} }},
        tooltip: {{
          ...baseChartOpts.tooltip,
          shared: true,
          intersect: false,
          y: {{ formatter: v => v.toFixed(2) + ' JPY/kWh' }}
        }}
      }});
      spotVsFuturesInstance.render();
    }}

    function renderPremiumChart(comparison, category) {{
      if (spotPremiumInstance) spotPremiumInstance.destroy();
      const data = comparison[category];
      if (!data || data.matches.length === 0) {{
        document.getElementById('spotPremiumChart').innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:3rem">No data</div>';
        return;
      }}
      const labels = data.matches.map(m => m.monthLabel);
      const premiums = data.matches.map(m => m.premium);
      document.getElementById('premiumBadge').textContent = data.label + ' vs Spot';

      spotPremiumInstance = new ApexCharts(document.getElementById('spotPremiumChart'), {{
        ...baseChartOpts,
        chart: {{ ...baseChartOpts.chart, type: 'bar', height: 320, animations: {{ enabled: true, easing: 'easeinout', speed: 600 }} }},
        series: [{{ name: 'Premium', data: premiums }}],
        plotOptions: {{
          bar: {{
            borderRadius: 4,
            columnWidth: '60%',
            colors: {{
              ranges: [
                {{ from: -9999, to: -0.001, color: colors.red }},
                {{ from: 0, to: 9999, color: colors.green }}
              ]
            }}
          }}
        }},
        colors: [colors.green],
        xaxis: {{ ...baseChartOpts.xaxis, categories: labels }},
        yaxis: {{ ...baseChartOpts.yaxis, title: {{ text: 'Premium (JPY/kWh)', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
        dataLabels: {{
          enabled: true,
          formatter: function(v) {{ return (v >= 0 ? '+' : '') + v.toFixed(1); }},
          style: {{ fontSize: '10px', fontFamily: 'JetBrains Mono', colors: ['#F1F5F9'] }},
          offsetY: -6
        }},
        tooltip: {{
          ...baseChartOpts.tooltip,
          custom: function({{ series: s, seriesIndex, dataPointIndex }}) {{
            const m = data.matches[dataPointIndex];
            const pColor = m.premium >= 0 ? colors.green : colors.red;
            const sign = m.premium >= 0 ? '+' : '';
            return '<div style="padding:8px 12px;font-family:JetBrains Mono,monospace;font-size:12px">'
              + '<div style="color:#F1F5F9;font-weight:600;margin-bottom:4px">' + m.monthLabel + '</div>'
              + '<div style="color:#94A3B8">Futures: <span style="color:' + colors.blue + '">' + m.futures.toFixed(2) + '</span></div>'
              + '<div style="color:#94A3B8">Spot: <span style="color:' + colors.amber + '">' + m.spot.toFixed(2) + '</span></div>'
              + '<div style="color:' + pColor + ';font-weight:600">' + sign + m.premium.toFixed(2) + ' (' + sign + m.premiumPct.toFixed(1) + '%)</div>'
              + '</div>';
          }}
        }}
      }});
      spotPremiumInstance.render();
    }}

}})();
</script>

<style>
.commodity-ja {{
  display: block;
  font-size: 0.72rem;
  color: var(--text-muted, #64748B);
  margin-top: 2px;
}}
.cat-badge {{
  display: inline-block;
  padding: 2px 8px;
  border-radius: 4px;
  font-size: 0.72rem;
  font-weight: 600;
  letter-spacing: 0.02em;
}}
</style>

</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────
# Weekly Compare Page (週次比較ページ)
# ─────────────────────────────────────────────────────────────────────────────

MAIN_POWER_CURVES = ["東・ベース(月次)", "東・日中(月次)", "西・ベース(月次)", "西・日中(月次)"]

# 電力関連の原燃料 (発電燃料・石油製品)。電力先物のドライバー
FUEL_UNDERLYINGS = [
    "LNG(プラッツJKM)",
    "ドバイ原油",
    "バージガソリン",
    "バージ灯油",
    "バージ軽油",
    "中京ガソリン",
    "中京灯油",
]
# フォワードカーブを描く対象（流動性が高く限月数が複数あるもの）
FUEL_CURVE_TARGETS = ["LNG(プラッツJKM)", "ドバイ原油"]


def generate_weekly_compare_data(repo, latest_date: str, base_date: str) -> dict:
    """Build the dict consumed by generate_weekly_compare_html.

    Focus: power futures (12 curves) + power-related fuel commodities
    (LNG, crude, petroleum products). Equity/bond/FX/metals/agriculture
    are intentionally excluded — this page is the JERA-style "電力先物
    weekly briefing" view.
    """
    # ── Power futures: today and base ──
    power_today = [r for r in get_power_futures(repo, latest_date) if not r.get("put_call")]
    power_base = [r for r in get_power_futures(repo, base_date) if not r.get("put_call")]
    prev_power_map = {
        r["instrument_name"]: r.get("settlement_price")
        for r in power_base
        if r.get("settlement_price") is not None
    }

    # All 12 power curves (today + base) — used for overlays and sub-cat summary
    all_curve_cats = list(CURVE_CATEGORIES.keys())
    power_curves_today: dict[str, list] = {c: [] for c in all_curve_cats}
    power_curves_base: dict[str, list] = {c: [] for c in all_curve_cats}
    for r in power_today:
        cat = classify_power_future(r["instrument_name"])
        if cat in power_curves_today:
            power_curves_today[cat].append({
                "month": r.get("contract_month", ""),
                "price": r.get("settlement_price"),
            })
    for r in power_base:
        cat = classify_power_future(r["instrument_name"])
        if cat in power_curves_base:
            power_curves_base[cat].append({
                "month": r.get("contract_month", ""),
                "price": r.get("settlement_price"),
            })
    for d in (power_curves_today, power_curves_base):
        for cat in d:
            d[cat].sort(key=lambda x: x["month"])

    # Power rows (per-contract) — for tables, sorting, top movers
    power_rows: list[dict] = []
    for r in power_today:
        name = r["instrument_name"]
        sub = classify_power_future(name) or "Power"
        settle = r.get("settlement_price")
        prev = prev_power_map.get(name)
        change = _calc_change(settle, prev)
        power_rows.append({
            "name": name,
            "display": name,
            "category": "energy",
            "asset_type": "power",
            "sub_category": sub,
            "contract_month": r.get("contract_month", ""),
            "price": settle,
            "prev_price": prev,
            "diff": change["diff"],
            "pct": change["pct"],
        })

    # 12 power sub-category summary (avg WoW % per curve type)
    power_subcat_summary = []
    for cat in all_curve_cats:
        pcts = [r["pct"] for r in power_rows if r["sub_category"] == cat and r["pct"] is not None]
        if not pcts:
            continue
        up = sum(1 for p in pcts if p > 0)
        down = sum(1 for p in pcts if p < 0)
        flat = sum(1 for p in pcts if p == 0)
        power_subcat_summary.append({
            "sub_category": cat,
            "avg_pct": round(sum(pcts) / len(pcts), 2),
            "count_up": up,
            "count_down": down,
            "count_flat": flat,
            "count_total": len(pcts),
        })

    # Main KPI cards — the 4 monthly curves (East/West × Base/Peak)
    main_kpi = []
    subcat_lookup = {s["sub_category"]: s for s in power_subcat_summary}
    for cat in MAIN_POWER_CURVES:
        s = subcat_lookup.get(cat)
        main_kpi.append({
            "label": cat,
            "avg_pct": s["avg_pct"] if s else None,
            "count_up": s["count_up"] if s else 0,
            "count_down": s["count_down"] if s else 0,
            "count_total": s["count_total"] if s else 0,
        })

    # ── Fuel section: LNG, crude, petroleum products ──
    fuel_rows: list[dict] = []
    for name in FUEL_UNDERLYINGS:
        info = ASSET_TAXONOMY.get(name)
        if not info:
            continue
        today_front = get_front_month_price(repo, latest_date, name)
        base_front = get_front_month_price(repo, base_date, name)
        if not today_front or today_front.get("settlement") is None:
            continue
        price = today_front["settlement"]
        prev = base_front["settlement"] if base_front else None
        change = _calc_change(price, prev)
        fuel_rows.append({
            "name": info["display_ja"],
            "display": f'{info["display_en"]} ({info["display_ja"]})',
            "display_en": info["display_en"],
            "display_ja": info["display_ja"],
            "category": "energy",
            "asset_type": "fuel",
            "sub_category": info.get("subcategory", ""),
            "contract_month": today_front.get("month", ""),
            "price": price,
            "prev_price": prev,
            "diff": change["diff"],
            "pct": change["pct"],
            "unit": info.get("unit", ""),
        })

    # Fuel forward curves (LNG JKM, Dubai Crude) — full curve today vs base
    fuel_curves_today: dict[str, list] = {}
    fuel_curves_base: dict[str, list] = {}
    for underlying in FUEL_CURVE_TARGETS:
        cur_today = get_commodity_forward_curve(repo, latest_date, underlying)
        cur_base = get_commodity_forward_curve(repo, base_date, underlying)
        if not cur_today:
            continue
        fuel_curves_today[underlying] = [
            {"month": p["month"], "price": p["settlement"]}
            for p in cur_today if p.get("settlement") is not None and p.get("month")
        ]
        fuel_curves_base[underlying] = [
            {"month": p["month"], "price": p["settlement"]}
            for p in cur_base if p.get("settlement") is not None and p.get("month")
        ]

    # ── Top movers (split into power vs fuel) ──
    valid_power = [r for r in power_rows if r["pct"] is not None]
    power_top_gainers = sorted(
        [r for r in valid_power if r["pct"] > 0], key=lambda x: x["pct"], reverse=True
    )[:15]
    power_top_losers = sorted(
        [r for r in valid_power if r["pct"] < 0], key=lambda x: x["pct"]
    )[:15]

    valid_fuel = [r for r in fuel_rows if r["pct"] is not None]
    fuel_top_gainers = sorted(
        [r for r in valid_fuel if r["pct"] > 0], key=lambda x: x["pct"], reverse=True
    )[:15]
    fuel_top_losers = sorted(
        [r for r in valid_fuel if r["pct"] < 0], key=lambda x: x["pct"]
    )[:15]

    # Combined rows (power + fuel) for the all-symbols table
    all_rows = power_rows + fuel_rows
    valid_all = [r for r in all_rows if r["pct"] is not None]

    counts = {
        "up": sum(1 for r in valid_all if r["pct"] > 0),
        "down": sum(1 for r in valid_all if r["pct"] < 0),
        "flat": sum(1 for r in valid_all if r["pct"] == 0),
        "total_valid": len(valid_all),
        "total_rows": len(all_rows),
        "power_count": len(power_rows),
        "fuel_count": len(fuel_rows),
    }

    def _parse(s: str):
        return datetime.strptime(s, "%Y-%m-%d" if "-" in s else "%Y%m%d")

    try:
        day_diff = (_parse(latest_date) - _parse(base_date)).days
    except ValueError:
        day_diff = None

    return {
        "latest_date": latest_date,
        "base_date": base_date,
        "calendar_days_back": day_diff,
        "rows": all_rows,
        "main_kpi": main_kpi,
        "power_subcat_summary": power_subcat_summary,
        "power_forward_curves": {"today": power_curves_today, "base": power_curves_base},
        "fuel_snapshot": fuel_rows,
        "fuel_forward_curves": {"today": fuel_curves_today, "base": fuel_curves_base},
        "power_top_gainers": power_top_gainers,
        "power_top_losers": power_top_losers,
        "fuel_top_gainers": fuel_top_gainers,
        "fuel_top_losers": fuel_top_losers,
        "counts": counts,
    }


def load_news_events(yaml_path: Path, start_date: str, end_date: str) -> list[dict]:
    """Load news/event entries from a YAML file, filtered to [start_date, end_date].

    Dates in YAML are ISO (YYYY-MM-DD). Comparison done after normalising.
    Missing/empty file returns an empty list. PyYAML is optional — if not
    installed, returns [] gracefully so the page still renders.
    """
    if not yaml_path.exists():
        return []
    try:
        import yaml  # type: ignore
    except ImportError:
        return []

    try:
        with open(yaml_path, encoding="utf-8") as f:
            data = yaml.safe_load(f) or []
    except Exception:
        return []
    if not isinstance(data, list):
        return []

    def _normalise(s: str) -> str:
        return str(s).replace("-", "")[:8]

    s_norm = _normalise(start_date)
    e_norm = _normalise(end_date)
    events: list[dict] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        d = _normalise(item.get("date", ""))
        if not d or not (s_norm <= d <= e_norm):
            continue
        events.append({
            "date": item.get("date", ""),
            "title": item.get("title", ""),
            "category": item.get("category", ""),
            "source": item.get("source", ""),
            "url": item.get("url", ""),
            "note": item.get("note", ""),
        })
    events.sort(key=lambda x: str(x["date"]))
    return events


def _wk_change_cell(diff, pct, *, show_pct: bool = True) -> str:
    """Compact change cell (used in tables on the weekly page)."""
    if diff is None or pct is None:
        return '<td class="num change">-</td>'
    sign = "+" if diff >= 0 else ""
    css = "positive" if diff > 0 else "negative" if diff < 0 else ""
    pct_html = (
        f'<span class="change-pct">({sign}{pct:.1f}%)</span>' if show_pct else ""
    )
    return (
        f'<td class="num change {css}">'
        f'{sign}{diff:.2f}{pct_html}'
        f'</td>'
    )


def _fmt_price(v) -> str:
    if v is None:
        return "-"
    try:
        return f"{float(v):.2f}"
    except (TypeError, ValueError):
        return str(v)


def _fmt_date_dotted(date_str: str) -> str:
    """Normalise to YYYY-MM-DD. Accepts ISO or packed 8-digit."""
    if not date_str:
        return ""
    s = str(date_str)
    if len(s) == 10 and s[4] == "-" and s[7] == "-":
        return s
    if len(s) == 8 and s.isdigit():
        return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
    return s


def _mover_rows_html(items: list[dict]) -> str:
    out = ""
    for it in items:
        sign = "+" if (it["diff"] or 0) >= 0 else ""
        css = "positive" if it["pct"] > 0 else "negative"
        meta = CATEGORY_META.get(it["category"], {})
        cat_label = meta.get("display_en", it["category"])
        cat_color = meta.get("color", "#666")
        out += (
            f'<tr>'
            f'<td>{it["display"]}'
            f'<span class="commodity-ja" style="display:block;font-size:0.7rem;opacity:0.55">{it.get("sub_category", "")}</span></td>'
            f'<td><span class="cat-badge" style="background:{cat_color}22;color:{cat_color}">{cat_label}</span></td>'
            f'<td class="num">{_fmt_price(it["prev_price"])}</td>'
            f'<td class="num">{_fmt_price(it["price"])}</td>'
            f'<td class="num {css}">{sign}{it["diff"]:.2f}</td>'
            f'<td class="num {css}">{sign}{it["pct"]:.2f}%</td>'
            f'</tr>'
        )
    if not out:
        out = '<tr><td colspan="6" style="text-align:center;opacity:0.6">該当銘柄なし</td></tr>'
    return out


def _all_rows_html(rows: list[dict]) -> str:
    out = ""
    for r in rows:
        asset_type = r.get("asset_type", "power")
        if asset_type == "power":
            type_label, type_color = "Power", "#F59E0B"
        else:
            type_label, type_color = "Fuel", "#FB7185"
        diff = r.get("diff")
        pct = r.get("pct")
        if diff is None or pct is None:
            change_cell = '<td class="num change">-</td>'
            pct_cell = '<td class="num change">-</td>'
        else:
            sign = "+" if diff >= 0 else ""
            css = "positive" if diff > 0 else "negative" if diff < 0 else ""
            change_cell = f'<td class="num {css}">{sign}{diff:.2f}</td>'
            pct_cell = f'<td class="num {css}">{sign}{pct:.2f}%</td>'
        out += (
            f'<tr data-asset-type="{asset_type}" data-pct="{pct if pct is not None else ""}">'
            f'<td>{r["display"]}</td>'
            f'<td><span class="cat-badge" style="background:{type_color}22;color:{type_color}">{type_label}</span></td>'
            f'<td>{r.get("sub_category", "")}</td>'
            f'<td>{r.get("contract_month", "")}</td>'
            f'<td class="num">{_fmt_price(r.get("prev_price"))}</td>'
            f'<td class="num">{_fmt_price(r.get("price"))}</td>'
            f'{change_cell}{pct_cell}'
            f'</tr>'
        )
    return out


def _asset_type_filter_options() -> str:
    return (
        '<option value="">All (Power + Fuel)</option>'
        '<option value="power">Power のみ</option>'
        '<option value="fuel">Fuel のみ</option>'
    )


def _fuel_snapshot_table_html(fuel_rows: list[dict]) -> str:
    """Snapshot table of fuel front-month prices with WoW change."""
    if not fuel_rows:
        return '<div style="opacity:0.6;padding:12px">原燃料データがありません。</div>'
    out = (
        '<div class="table-wrapper"><table class="wk-table-compact"><thead><tr>'
        '<th>原燃料</th><th>サブ</th><th>限月</th><th class="num">単位</th>'
        '<th class="num">1週前</th><th class="num">今日</th>'
        '<th class="num">差</th><th class="num">変化率</th>'
        '</tr></thead><tbody>'
    )
    for r in fuel_rows:
        diff = r.get("diff")
        pct = r.get("pct")
        if diff is None or pct is None:
            change_cell = '<td class="num change">-</td>'
            pct_cell = '<td class="num change">-</td>'
        else:
            sign = "+" if diff >= 0 else ""
            css = "positive" if diff > 0 else "negative" if diff < 0 else ""
            change_cell = f'<td class="num {css}">{sign}{diff:.2f}</td>'
            pct_cell = f'<td class="num {css}">{sign}{pct:.2f}%</td>'
        out += (
            f'<tr>'
            f'<td>{r.get("display_en", "")}<span class="commodity-ja">{r.get("display_ja", "")}</span></td>'
            f'<td>{r.get("sub_category", "")}</td>'
            f'<td>{r.get("contract_month", "")}</td>'
            f'<td class="num" style="opacity:0.6">{r.get("unit", "")}</td>'
            f'<td class="num">{_fmt_price(r.get("prev_price"))}</td>'
            f'<td class="num">{_fmt_price(r.get("price"))}</td>'
            f'{change_cell}{pct_cell}'
            f'</tr>'
        )
    out += '</tbody></table></div>'
    return out


def _events_html(events: list[dict]) -> str:
    if not events:
        return (
            '<div style="opacity:0.6;padding:12px;font-size:0.9rem">'
            '対象期間に登録されたイベントはありません。'
            '<br><span style="font-size:0.75rem">docs/news_events.yaml を編集してイベントを追加できます。</span>'
            '</div>'
        )
    out = '<ul style="list-style:none;padding:0;margin:0">'
    for e in events:
        date_label = _fmt_date_dotted(str(e["date"]).replace("-", ""))
        title = e["title"] or "(no title)"
        url = e.get("url", "")
        title_html = (
            f'<a href="{url}" target="_blank" rel="noopener" style="color:inherit;text-decoration:none">{title}</a>'
            if url else title
        )
        cat = e.get("category", "")
        source = e.get("source", "")
        note = e.get("note", "")
        meta_bits = " · ".join([b for b in (cat, source) if b])
        note_html = (
            f'<div style="font-size:0.78rem;opacity:0.7;margin-top:2px">{note}</div>'
            if note else ""
        )
        out += (
            f'<li style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,0.06)">'
            f'<div style="font-size:0.72rem;opacity:0.6">{date_label} · {meta_bits}</div>'
            f'<div style="font-weight:600;margin-top:2px">{title_html}</div>'
            f'{note_html}'
            f'</li>'
        )
    out += '</ul>'
    return out


def generate_weekly_compare_html(data: dict) -> str:
    """Generate the docs/weekly_compare.html page.

    Focused on power futures + power-related fuels (LNG, crude, petroleum
    products). Equity/bond/FX/metals/agriculture are out of scope here —
    this is the "電力先物 週次ブリーフィング" view.
    """
    latest = data.get("latest_date", "N/A")
    base = data.get("base_date", "N/A")
    days_back = data.get("calendar_days_back")
    counts = data.get("counts", {})
    rows = data.get("rows", [])
    main_kpi = data.get("main_kpi", [])
    power_subcat = data.get("power_subcat_summary", [])
    power_curves = data.get("power_forward_curves", {"today": {}, "base": {}})
    power_gainers = data.get("power_top_gainers", [])
    power_losers = data.get("power_top_losers", [])
    fuel_snapshot = data.get("fuel_snapshot", [])
    fuel_curves = data.get("fuel_forward_curves", {"today": {}, "base": {}})
    fuel_gainers = data.get("fuel_top_gainers", [])
    fuel_losers = data.get("fuel_top_losers", [])
    events = data.get("events", [])

    latest_label = _fmt_date_dotted(latest)
    base_label = _fmt_date_dotted(base)
    days_back_label = f"{days_back} 日前" if isinstance(days_back, int) else "—"

    power_gainers_html = _mover_rows_html(power_gainers)
    power_losers_html = _mover_rows_html(power_losers)
    fuel_gainers_html = _mover_rows_html(fuel_gainers)
    fuel_losers_html = _mover_rows_html(fuel_losers)
    all_rows_html = _all_rows_html(rows)
    asset_options_html = _asset_type_filter_options()
    fuel_snapshot_html = _fuel_snapshot_table_html(fuel_snapshot)
    events_html = _events_html(events)

    # KPI cards (4 main monthly power curves)
    kpi_cards_html = ""
    kpi_color = {
        "東・ベース(月次)": ("var(--accent-blue)", "東 Base (月次)", "East Base"),
        "東・日中(月次)":   ("var(--gold)",        "東 Peak (月次)", "East Peak"),
        "西・ベース(月次)": ("var(--accent-cyan)", "西 Base (月次)", "West Base"),
        "西・日中(月次)":   ("var(--positive)",    "西 Peak (月次)", "West Peak"),
    }
    for kpi in main_kpi:
        color, ja_label, en_label = kpi_color.get(
            kpi["label"], ("#888", kpi["label"], kpi["label"])
        )
        avg = kpi.get("avg_pct")
        if avg is None:
            avg_str = "—"
            avg_color = "#888"
        else:
            sign = "+" if avg >= 0 else ""
            avg_str = f"{sign}{avg:.2f}%"
            avg_color = "var(--positive)" if avg > 0 else "var(--negative)" if avg < 0 else "#888"
        kpi_cards_html += (
            f'<div class="kpi-card" style="border-left:3px solid {color}">'
            f'<div class="kpi-icon" style="color:{color}">&#9650;</div>'
            f'<div class="kpi-label">{ja_label}</div>'
            f'<div class="kpi-value" style="color:{avg_color}">{avg_str}</div>'
            f'<div class="kpi-sub">{en_label} · 上 {kpi["count_up"]} / 下 {kpi["count_down"]} / 計 {kpi["count_total"]} 限月</div>'
            f'</div>'
        )

    # Inline JSON for charts
    chart_payload = json.dumps({
        "power_subcat_summary": power_subcat,
        "power_forward_curves": power_curves,
        "fuel_forward_curves": fuel_curves,
    }, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>電力先物 週次比較 — DataServer In-House</title>
  <link rel="stylesheet" href="style.css">
  <script src="https://cdn.jsdelivr.net/npm/apexcharts@3"></script>
  <style>
    .wk-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; }}
    .wk-grid-4 {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; }}
    .wk-filters {{ display:flex; gap:12px; align-items:center; margin-bottom:12px; flex-wrap:wrap; }}
    .wk-filters input, .wk-filters select {{
      background:rgba(255,255,255,0.04); border:1px solid rgba(255,255,255,0.08);
      color:inherit; padding:6px 10px; border-radius:6px; font-size:0.85rem; min-width:180px;
    }}
    .wk-filters input:focus, .wk-filters select:focus {{
      outline:none; border-color:var(--accent-blue);
    }}
    th[data-sort] {{ cursor:pointer; user-select:none; }}
    th[data-sort]:hover {{ color:var(--accent-blue); }}
    .wk-table-compact td, .wk-table-compact th {{ padding:6px 10px; font-size:0.82rem; }}
    .wk-section-divider {{
      margin:32px 0 12px 0; padding:8px 14px; border-left:3px solid var(--accent-blue);
      background:linear-gradient(90deg, rgba(59,130,246,0.10), transparent);
      border-radius:4px;
      display:flex; align-items:baseline; gap:10px;
    }}
    .wk-section-divider.fuel {{ border-color:#FB7185; background:linear-gradient(90deg, rgba(251,113,133,0.10), transparent); }}
    .wk-section-divider h2 {{ font-size:1.05rem; margin:0; font-weight:700; }}
    .wk-section-divider .wk-section-tag {{ font-size:0.7rem; opacity:0.55; letter-spacing:0.04em; }}
    .wk-mini-summary {{
      display:flex; gap:18px; flex-wrap:wrap; margin-top:8px;
      font-size:0.78rem; opacity:0.85;
    }}
    .wk-mini-summary span b {{ font-weight:700; }}
    @media (max-width: 980px) {{
      .wk-grid, .wk-grid-4 {{ grid-template-columns:1fr; }}
    }}
  </style>
</head>
<body>

<div class="bg-gradient"></div>

<div class="app-layout">

  <!-- Sidebar -->
  <nav class="sidebar">
    <div class="sidebar-logo">DS</div>
    <div class="sidebar-nav">
      <a href="index.html" class="nav-item">
        <span>&#9670;</span>
        <span class="nav-tooltip">Dashboard</span>
      </a>
      <a href="weekly_compare.html" class="nav-item active">
        <span>&#9776;</span>
        <span class="nav-tooltip">Weekly Compare</span>
      </a>
      <a href="spread_analysis.html" class="nav-item">
        <span>&#9878;</span>
        <span class="nav-tooltip">Spread Analysis</span>
      </a>
      <div style="width:32px;height:1px;background:var(--border-subtle);margin:8px 0"></div>
      <a href="#wk-power-movers" class="nav-item">
        <span>&#9889;</span>
        <span class="nav-tooltip">電力 Movers</span>
      </a>
      <a href="#wk-power-subcat" class="nav-item">
        <span>&#9619;</span>
        <span class="nav-tooltip">電力 12 サブカテゴリ</span>
      </a>
      <a href="#wk-power-curves" class="nav-item">
        <span>&#9699;</span>
        <span class="nav-tooltip">電力カーブ重ね合わせ</span>
      </a>
      <a href="#wk-fuel-snap" class="nav-item">
        <span>&#9876;</span>
        <span class="nav-tooltip">原燃料スナップショット</span>
      </a>
      <a href="#wk-fuel-curves" class="nav-item">
        <span>&#9696;</span>
        <span class="nav-tooltip">原燃料カーブ</span>
      </a>
      <a href="#wk-all" class="nav-item">
        <span>&#9783;</span>
        <span class="nav-tooltip">全銘柄テーブル</span>
      </a>
      <a href="#wk-events" class="nav-item">
        <span>&#128240;</span>
        <span class="nav-tooltip">News & Events</span>
      </a>
    </div>
  </nav>

  <!-- Main Content -->
  <div class="main-content">

    <header class="header">
      <div class="header-left">
        <h1>電力先物 週次比較</h1>
        <div class="subtitle">電力先物 + 電力関連原燃料(LNG · 原油 · 石油製品) · WoW {days_back_label}</div>
      </div>
      <div class="header-right">
        <div class="header-badge"><span class="dot"></span>WoW</div>
        <div class="header-date">{latest_label}<br><span style="font-size:0.65rem;color:#64748B">vs {base_label}</span></div>
      </div>
    </header>

    <div class="container">

      <!-- KPI Cards: 4 main monthly power curves' avg WoW% -->
      <div class="kpi-row" style="grid-template-columns:repeat(4,1fr)">
        {kpi_cards_html}
      </div>
      <div class="wk-mini-summary">
        <span>&#9889; 電力 <b>{counts.get("power_count", 0)}</b> 限月</span>
        <span>&#9876; 原燃料 <b>{counts.get("fuel_count", 0)}</b> 銘柄</span>
        <span style="color:var(--positive)">&#9650; 上昇 <b>{counts.get("up", 0)}</b></span>
        <span style="color:var(--negative)">&#9660; 下落 <b>{counts.get("down", 0)}</b></span>
        <span>= 横ばい <b>{counts.get("flat", 0)}</b></span>
      </div>

      <!-- ─── 電力セクション ─── -->
      <div class="wk-section-divider"><h2>&#9889; 電力先物</h2><span class="wk-section-tag">Power Futures</span></div>

      <section id="wk-power-movers" class="card-grid">
        <div class="card card-wide">
          <div class="card-header">
            <div class="card-title"><span class="icon" style="color:var(--positive)">&#9650;</span> 電力 Top Gainers</div>
            <div class="card-badge">{len(power_gainers)}</div>
          </div>
          <div class="card-body">
            <div class="table-wrapper">
              <table class="wk-table-compact">
                <thead><tr>
                  <th>銘柄</th><th>サブ</th><th class="num">1週前</th><th class="num">今日</th><th class="num">差</th><th class="num">変化率</th>
                </tr></thead>
                <tbody>{power_gainers_html}</tbody>
              </table>
            </div>
          </div>
        </div>
        <div class="card card-wide">
          <div class="card-header">
            <div class="card-title"><span class="icon" style="color:var(--negative)">&#9660;</span> 電力 Top Losers</div>
            <div class="card-badge">{len(power_losers)}</div>
          </div>
          <div class="card-body">
            <div class="table-wrapper">
              <table class="wk-table-compact">
                <thead><tr>
                  <th>銘柄</th><th>サブ</th><th class="num">1週前</th><th class="num">今日</th><th class="num">差</th><th class="num">変化率</th>
                </tr></thead>
                <tbody>{power_losers_html}</tbody>
              </table>
            </div>
          </div>
        </div>
      </section>

      <section id="wk-power-subcat" class="card-grid" style="margin-top:16px">
        <div class="card card-full">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9619;</span> 電力 12 サブカテゴリ別 平均 WoW%</div>
            <div class="card-badge">{len(power_subcat)} curves</div>
          </div>
          <div class="card-body">
            <div id="wk-subcat-chart" class="chart-container" style="min-height:360px"></div>
          </div>
        </div>
      </section>

      <section id="wk-power-curves" class="card-grid wk-grid-4" style="margin-top:16px">
        <div class="card">
          <div class="card-header"><div class="card-title">東・ベース (月次)</div></div>
          <div class="card-body"><div id="wk-curve-EEB" class="chart-container" style="min-height:280px"></div></div>
        </div>
        <div class="card">
          <div class="card-header"><div class="card-title">東・日中 (月次)</div></div>
          <div class="card-body"><div id="wk-curve-EEP" class="chart-container" style="min-height:280px"></div></div>
        </div>
        <div class="card">
          <div class="card-header"><div class="card-title">西・ベース (月次)</div></div>
          <div class="card-body"><div id="wk-curve-EWB" class="chart-container" style="min-height:280px"></div></div>
        </div>
        <div class="card">
          <div class="card-header"><div class="card-title">西・日中 (月次)</div></div>
          <div class="card-body"><div id="wk-curve-EWP" class="chart-container" style="min-height:280px"></div></div>
        </div>
      </section>

      <!-- ─── 原燃料セクション ─── -->
      <div class="wk-section-divider fuel"><h2>&#9876; 電力関連 原燃料</h2><span class="wk-section-tag">Power-related Fuels (LNG / Crude / Petroleum)</span></div>

      <section id="wk-fuel-snap" class="card-grid">
        <div class="card card-full">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9876;</span> 原燃料スナップショット (Front-month)</div>
            <div class="card-badge">{len(fuel_snapshot)} symbols</div>
          </div>
          <div class="card-body">{fuel_snapshot_html}</div>
        </div>
      </section>

      <section class="card-grid" style="margin-top:16px">
        <div class="card card-wide">
          <div class="card-header">
            <div class="card-title"><span class="icon" style="color:var(--positive)">&#9650;</span> 原燃料 Top Gainers</div>
            <div class="card-badge">{len(fuel_gainers)}</div>
          </div>
          <div class="card-body">
            <div class="table-wrapper">
              <table class="wk-table-compact">
                <thead><tr>
                  <th>銘柄</th><th>サブ</th><th class="num">1週前</th><th class="num">今日</th><th class="num">差</th><th class="num">変化率</th>
                </tr></thead>
                <tbody>{fuel_gainers_html}</tbody>
              </table>
            </div>
          </div>
        </div>
        <div class="card card-wide">
          <div class="card-header">
            <div class="card-title"><span class="icon" style="color:var(--negative)">&#9660;</span> 原燃料 Top Losers</div>
            <div class="card-badge">{len(fuel_losers)}</div>
          </div>
          <div class="card-body">
            <div class="table-wrapper">
              <table class="wk-table-compact">
                <thead><tr>
                  <th>銘柄</th><th>サブ</th><th class="num">1週前</th><th class="num">今日</th><th class="num">差</th><th class="num">変化率</th>
                </tr></thead>
                <tbody>{fuel_losers_html}</tbody>
              </table>
            </div>
          </div>
        </div>
      </section>

      <section id="wk-fuel-curves" class="card-grid wk-grid" style="margin-top:16px">
        <div class="card">
          <div class="card-header"><div class="card-title">LNG (プラッツ JKM) フォワードカーブ</div></div>
          <div class="card-body"><div id="wk-fcurve-LNG" class="chart-container" style="min-height:280px"></div></div>
        </div>
        <div class="card">
          <div class="card-header"><div class="card-title">ドバイ原油 フォワードカーブ</div></div>
          <div class="card-body"><div id="wk-fcurve-DUBAI" class="chart-container" style="min-height:280px"></div></div>
        </div>
      </section>

      <!-- ─── 全銘柄テーブル ─── -->
      <section id="wk-all" class="card-grid" style="margin-top:24px">
        <div class="card card-full">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9783;</span> 全銘柄 (電力 + 原燃料) 週次比較</div>
            <div class="card-badge">{len(rows)} symbols</div>
          </div>
          <div class="card-body">
            <div class="wk-filters">
              <input id="wk-search" type="text" placeholder="銘柄名で絞り込み...">
              <select id="wk-asset-filter">{asset_options_html}</select>
              <span style="font-size:0.78rem;opacity:0.6" id="wk-row-count"></span>
            </div>
            <div class="table-wrapper">
              <table class="wk-table-compact" id="wk-all-table">
                <thead><tr>
                  <th data-sort="name">銘柄</th>
                  <th data-sort="type">種別</th>
                  <th data-sort="sub">サブ</th>
                  <th data-sort="month">限月</th>
                  <th class="num" data-sort="prev">1週前</th>
                  <th class="num" data-sort="price">今日</th>
                  <th class="num" data-sort="diff">差</th>
                  <th class="num" data-sort="pct">変化率</th>
                </tr></thead>
                <tbody>{all_rows_html}</tbody>
              </table>
            </div>
          </div>
        </div>
      </section>

      <!-- News & Events -->
      <section id="wk-events" class="card-grid" style="margin-top:24px">
        <div class="card card-full">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#128240;</span> 期間内のニュース・イベント</div>
            <div class="card-badge">{len(events)} items</div>
          </div>
          <div class="card-body">{events_html}</div>
        </div>
      </section>

    </div>
  </div>
</div>

<script>
const weeklyData = {chart_payload};

const apexCommonOpts = {{
  chart: {{ background:'transparent', toolbar:{{show:false}}, fontFamily:'inherit' }},
  theme: {{ mode:'dark' }},
  grid: {{ borderColor:'rgba(255,255,255,0.06)', strokeDashArray:3 }},
  tooltip: {{ theme:'dark' }},
}};

// 電力 12 サブカテゴリ平均 WoW% 棒チャート
(function renderSubcatChart() {{
  const items = weeklyData.power_subcat_summary || [];
  if (!items.length) return;
  const labels = items.map(i => i.sub_category);
  const values = items.map(i => i.avg_pct == null ? 0 : i.avg_pct);
  const colors = values.map(v => v > 0 ? '#10B981' : v < 0 ? '#EF4444' : '#64748b');
  const opts = Object.assign({{}}, apexCommonOpts, {{
    chart: Object.assign({{type:'bar', height:360}}, apexCommonOpts.chart),
    series: [{{ name:'平均 WoW (%)', data: values }}],
    xaxis: {{
      categories: labels,
      labels: {{ style:{{colors:'#94a3b8', fontSize:'0.7rem'}}, rotate:-30 }},
    }},
    yaxis: {{ labels:{{ style:{{colors:'#94a3b8'}}, formatter:(v)=> v.toFixed(2)+'%' }} }},
    plotOptions: {{ bar:{{ distributed:true, borderRadius:4, columnWidth:'60%' }} }},
    colors: colors,
    legend: {{ show:false }},
    dataLabels: {{
      enabled:true,
      formatter:(v)=> (v>=0?'+':'') + v.toFixed(2) + '%',
      style: {{ colors:['#e2e8f0'], fontSize:'0.72rem' }},
      offsetY: -16,
    }},
    tooltip: {{
      theme:'dark',
      custom: function(o) {{
        const i = items[o.dataPointIndex];
        if (!i) return '';
        const sign = i.avg_pct >= 0 ? '+' : '';
        return '<div style="padding:8px 10px">'
          + '<div style="font-weight:600">' + i.sub_category + '</div>'
          + '<div>平均 ' + sign + i.avg_pct.toFixed(2) + '%</div>'
          + '<div style="font-size:0.75rem;opacity:0.75">上 ' + i.count_up + ' / 下 ' + i.count_down + ' / 計 ' + i.count_total + '</div>'
          + '</div>';
      }},
    }},
  }});
  new ApexCharts(document.querySelector('#wk-subcat-chart'), opts).render();
}})();

// フォワードカーブ重ね合わせ (汎用)
function renderOverlayCurve(id, today, base) {{
  const monthSet = new Set();
  today.forEach(p => monthSet.add(p.month));
  base.forEach(p => monthSet.add(p.month));
  const months = Array.from(monthSet).sort();
  const el = document.querySelector(id);
  if (!months.length) {{
    if (el) el.innerHTML = '<div style="padding:24px;opacity:0.5">データなし</div>';
    return;
  }}
  const tMap = Object.fromEntries(today.map(p => [p.month, p.price]));
  const bMap = Object.fromEntries(base.map(p => [p.month, p.price]));
  const tSeries = months.map(m => tMap[m] == null ? null : tMap[m]);
  const bSeries = months.map(m => bMap[m] == null ? null : bMap[m]);
  const opts = Object.assign({{}}, apexCommonOpts, {{
    chart: Object.assign({{type:'line', height:280}}, apexCommonOpts.chart),
    series: [
      {{ name:'今日', data: tSeries }},
      {{ name:'1週間前', data: bSeries }},
    ],
    xaxis: {{ categories: months, labels:{{ style:{{colors:'#94a3b8'}}, rotate:-45 }} }},
    yaxis: {{ labels:{{ style:{{colors:'#94a3b8'}}, formatter:(v)=> v==null?'-':v.toFixed(2) }} }},
    colors: ['#3B82F6', '#F59E0B'],
    stroke: {{ width:[2.5, 1.8], dashArray:[0, 6], curve:'straight' }},
    markers: {{ size:[4, 0] }},
    legend: {{ labels:{{ colors:'#cbd5e1' }} }},
    dataLabels: {{ enabled:false }},
  }});
  if (el) new ApexCharts(el, opts).render();
}}

function renderPowerCurve(id, cat) {{
  const today = (weeklyData.power_forward_curves.today || {{}})[cat] || [];
  const base = (weeklyData.power_forward_curves.base || {{}})[cat] || [];
  renderOverlayCurve(id, today, base);
}}
function renderFuelCurve(id, underlying) {{
  const today = (weeklyData.fuel_forward_curves.today || {{}})[underlying] || [];
  const base = (weeklyData.fuel_forward_curves.base || {{}})[underlying] || [];
  renderOverlayCurve(id, today, base);
}}

renderPowerCurve('#wk-curve-EEB', '東・ベース(月次)');
renderPowerCurve('#wk-curve-EEP', '東・日中(月次)');
renderPowerCurve('#wk-curve-EWB', '西・ベース(月次)');
renderPowerCurve('#wk-curve-EWP', '西・日中(月次)');
renderFuelCurve('#wk-fcurve-LNG', 'LNG(プラッツJKM)');
renderFuelCurve('#wk-fcurve-DUBAI', 'ドバイ原油');

// 全銘柄テーブルのフィルタ + ソート
(function setupAllTable() {{
  const table = document.querySelector('#wk-all-table');
  if (!table) return;
  const tbody = table.querySelector('tbody');
  const allRows = Array.from(tbody.querySelectorAll('tr'));
  const search = document.querySelector('#wk-search');
  const assetFilter = document.querySelector('#wk-asset-filter');
  const rowCount = document.querySelector('#wk-row-count');

  function applyFilter() {{
    const q = (search.value || '').toLowerCase();
    const a = assetFilter.value;
    let visible = 0;
    allRows.forEach(tr => {{
      const text = tr.textContent.toLowerCase();
      const at = tr.getAttribute('data-asset-type') || '';
      const ok = (!q || text.includes(q)) && (!a || at === a);
      tr.style.display = ok ? '' : 'none';
      if (ok) visible++;
    }});
    rowCount.textContent = visible + ' 件表示';
  }}
  search.addEventListener('input', applyFilter);
  assetFilter.addEventListener('change', applyFilter);
  applyFilter();

  // Sort
  const headers = table.querySelectorAll('th[data-sort]');
  let sortState = {{ key:null, asc:true }};
  headers.forEach((th, idx) => {{
    th.addEventListener('click', () => {{
      const key = th.getAttribute('data-sort');
      sortState.asc = (sortState.key === key) ? !sortState.asc : true;
      sortState.key = key;
      const sign = sortState.asc ? 1 : -1;
      const rowsArr = Array.from(tbody.querySelectorAll('tr'));
      rowsArr.sort((a, b) => {{
        const av = a.children[idx]?.textContent.trim() || '';
        const bv = b.children[idx]?.textContent.trim() || '';
        const an = parseFloat(av.replace(/[+%,]/g, ''));
        const bn = parseFloat(bv.replace(/[+%,]/g, ''));
        if (!isNaN(an) && !isNaN(bn)) return (an - bn) * sign;
        return av.localeCompare(bv) * sign;
      }});
      rowsArr.forEach(r => tbody.appendChild(r));
    }});
  }});
}})();
</script>

</body>
</html>"""


def write_weekly_compare(repo, site_dir: Path, latest_date: str) -> Path | None:
    """Generate and write `docs/weekly_compare.html`.

    Picks the comparison base date as the nearest business day at-or-before
    (latest_date − 7 calendar days). Returns the output path, or None if
    there is no usable history before `latest_date`.

    Safe to call from any site-regeneration entry point (generate_site.main
    or fetch_jpx.main) — keeps the weekly page in sync with the daily one.
    """
    base_date = find_nearest_business_day_before(repo, latest_date, days_back=7)
    if not base_date:
        return None
    wk_data = generate_weekly_compare_data(repo, latest_date, base_date)
    wk_data["events"] = load_news_events(
        site_dir / "news_events.yaml", base_date, latest_date
    )
    wk_html = generate_weekly_compare_html(wk_data)
    wk_html_path = site_dir / "weekly_compare.html"
    with open(wk_html_path, "w", encoding="utf-8") as f:
        f.write(wk_html)
    return wk_html_path


# ─────────────────────────────────────────────────────────────────────────────
# Spread Analysis Page (スプレッド分析) — recreates Spread Calculator_v2.xlsx
# ─────────────────────────────────────────────────────────────────────────────

def generate_spread_data(repo, latest_date: str, prev_date: str | None = None) -> dict:
    """Thin wrapper around the self-contained spread computation module."""
    return compute_spread_analysis(repo, latest_date, prev_date)


# Standalone HTML built via placeholder replacement (no f-string brace doubling
# so the embedded JS stays readable). Placeholders: __PAYLOAD__ __LATEST__
# __HORIZON__ __PARAMS_ROWS__.
_SPREAD_HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>スプレッド分析 — DataServer In-House</title>
  <link rel="stylesheet" href="style.css">
  <script src="https://cdn.jsdelivr.net/npm/apexcharts@3"></script>
  <style>
    .sp-sliders { display:grid; grid-template-columns:repeat(3,1fr); gap:18px; }
    .sp-slider label { display:block; font-size:0.8rem; opacity:0.8; margin-bottom:6px; }
    .sp-slider input[type=range] { width:100%; accent-color:var(--accent-blue); }
    .sp-slider .sp-val { font-weight:700; color:var(--accent-cyan); font-variant-numeric:tabular-nums; }
    .sp-reset { margin-top:12px; background:rgba(255,255,255,0.05);
      border:1px solid rgba(255,255,255,0.12); color:inherit; padding:6px 14px;
      border-radius:6px; font-size:0.8rem; cursor:pointer; }
    .sp-reset:hover { border-color:var(--accent-blue); }
    .sp-params-table td { padding:4px 12px; font-size:0.82rem; }
    .sp-params-table td:first-child { opacity:0.7; }
    .sp-note { font-size:0.78rem; opacity:0.6; margin-top:10px; }
    @media (max-width:980px) { .sp-sliders { grid-template-columns:1fr; } }
  </style>
</head>
<body>

<div class="bg-gradient"></div>

<div class="app-layout">

  <!-- Sidebar -->
  <nav class="sidebar">
    <div class="sidebar-logo">DS</div>
    <div class="sidebar-nav">
      <a href="index.html" class="nav-item">
        <span>&#9670;</span>
        <span class="nav-tooltip">Dashboard</span>
      </a>
      <a href="weekly_compare.html" class="nav-item">
        <span>&#9776;</span>
        <span class="nav-tooltip">Weekly Compare</span>
      </a>
      <a href="spread_analysis.html" class="nav-item active">
        <span>&#9878;</span>
        <span class="nav-tooltip">Spread Analysis</span>
      </a>
      <div style="width:32px;height:1px;background:var(--border-subtle);margin:8px 0"></div>
      <a href="#sp-forward" class="nav-item">
        <span>&#9699;</span>
        <span class="nav-tooltip">フォワードカーブ</span>
      </a>
      <a href="#sp-spark" class="nav-item">
        <span>&#9889;</span>
        <span class="nav-tooltip">Spark Spread</span>
      </a>
      <a href="#sp-params" class="nav-item">
        <span>&#9881;</span>
        <span class="nav-tooltip">パラメータ</span>
      </a>
    </div>
  </nav>

  <!-- Main Content -->
  <div class="main-content">

    <header class="header">
      <div class="header-left">
        <h1>スプレッド分析</h1>
        <div class="subtitle">電力フォワードカーブ + ガス火力 Spark / Clean Spark スプレッド · Spread Calculator_v2 を日次再現</div>
      </div>
      <div class="header-right">
        <div class="header-badge"><span class="dot"></span>Spark</div>
        <div class="header-date">__LATEST__</div>
      </div>
    </header>

    <div class="container">

      <!-- Forward curves -->
      <section id="sp-forward" class="card-grid">
        <div class="card card-full">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9699;</span> 電力フォワードカーブ (6 系列 · ¥/kWh)</div>
            <div class="card-badge">東/西/中部 × ベース/日中</div>
          </div>
          <div class="card-body">
            <div id="sp-fwd-chart" class="chart-container" style="min-height:420px"></div>
          </div>
        </div>
      </section>

      <!-- Parameter sliders -->
      <section id="sp-params" class="card-grid" style="margin-top:16px">
        <div class="card card-full">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9881;</span> 発電パラメータ (スライダーで Spark を再計算)</div>
          </div>
          <div class="card-body">
            <div class="sp-sliders">
              <div class="sp-slider">
                <label>ガス火力 熱効率 <span class="sp-val" id="sp-eff-val"></span></label>
                <input type="range" id="sp-eff" min="0.30" max="0.65" step="0.01">
              </div>
              <div class="sp-slider">
                <label>ガス CO2 排出係数 (kg-CO2/kWh) <span class="sp-val" id="sp-co2f-val"></span></label>
                <input type="range" id="sp-co2f" min="0.0" max="0.8" step="0.05">
              </div>
              <div class="sp-slider">
                <label>CO2 価格 (円/t) <span class="sp-val" id="sp-co2p-val"></span></label>
                <input type="range" id="sp-co2p" min="0" max="20000" step="500">
              </div>
            </div>
            <button class="sp-reset" id="sp-reset">既定値に戻す</button>
            <table class="sp-params-table" style="margin-top:14px">
              __PARAMS_ROWS__
            </table>
            <div class="sp-note">
              Spark[m] = 電力価格[m] − ガス発電コスト[m] ／ ガス発電コスト[m] = (JKM[m] ÷ 熱効率) ÷ 293.07 ／
              Clean Spark[m] = Spark[m] − (CO2係数 × CO2価格 ÷ 1000) ／ 対象限月: __HORIZON__
            </div>
          </div>
        </div>
      </section>

      <!-- Spark spread -->
      <section id="sp-spark" class="card-grid" style="margin-top:16px">
        <div class="card card-full">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9889;</span> Spark Spread (ガス火力採算 · ¥/kWh)</div>
          </div>
          <div class="card-body">
            <div id="sp-spark-chart" class="chart-container" style="min-height:380px"></div>
          </div>
        </div>
      </section>

      <section class="card-grid" style="margin-top:16px">
        <div class="card card-full">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#127807;</span> Clean Spark Spread (CO2コスト控除後 · ¥/kWh)</div>
          </div>
          <div class="card-body">
            <div id="sp-clean-chart" class="chart-container" style="min-height:380px"></div>
          </div>
        </div>
      </section>

    </div>
  </div>
</div>

<script>
const spreadData = __PAYLOAD__;

const apexCommonOpts = {
  chart: { background:'transparent', toolbar:{show:false}, fontFamily:'inherit' },
  theme: { mode:'dark' },
  grid: { borderColor:'rgba(255,255,255,0.06)', strokeDashArray:3 },
  tooltip: { theme:'dark' },
};

const SP_COLORS = ['#3B82F6','#F59E0B','#0EA5E9','#10B981','#A855F7','#EF4444'];
const POWER_LABELS = Object.keys(spreadData.forward_curves || {});
const GAS_CONV = spreadData.params.gas_conv_kwh_mmbtu;
const jkmMap = Object.fromEntries((spreadData.jkm_curve || []).map(p => [p.month, p.price]));
const SPARK_MONTHS = (spreadData.jkm_curve || []).map(p => p.month);

function fmtMonth(ym) {
  return (ym && ym.length >= 6) ? ym.slice(0,4) + '-' + ym.slice(4,6) : ym;
}

// ── Forward curve chart (static) ──
(function renderForward() {
  const el = document.querySelector('#sp-fwd-chart');
  if (!el) return;
  const monthSet = new Set();
  POWER_LABELS.forEach(l => (spreadData.forward_curves[l] || []).forEach(p => monthSet.add(p.month)));
  const months = Array.from(monthSet).sort();
  if (!months.length) { el.innerHTML = '<div style="padding:24px;opacity:0.5">データなし</div>'; return; }
  const series = POWER_LABELS.map(l => {
    const m = Object.fromEntries((spreadData.forward_curves[l] || []).map(p => [p.month, p.price]));
    return { name: l, data: months.map(x => m[x] == null ? null : m[x]) };
  });
  const opts = Object.assign({}, apexCommonOpts, {
    chart: Object.assign({type:'line', height:420}, apexCommonOpts.chart),
    series: series,
    xaxis: { categories: months.map(fmtMonth), labels:{ style:{colors:'#94a3b8'}, rotate:-45 } },
    yaxis: { labels:{ style:{colors:'#94a3b8'}, formatter:(v)=> v==null?'-':v.toFixed(1) }, title:{ text:'¥/kWh', style:{color:'#94a3b8'} } },
    colors: SP_COLORS,
    stroke: { width:2.2, curve:'straight' },
    markers: { size:0, hover:{size:4} },
    legend: { labels:{ colors:'#cbd5e1' }, position:'top' },
    dataLabels: { enabled:false },
  });
  new ApexCharts(el, opts).render();
})();

// ── Spark / Clean spark (interactive) ──
function computeSpark(gasEff, gasCo2, co2Price) {
  const co2 = gasCo2 * co2Price / 1000;
  const sparkSeries = [], cleanSeries = [];
  POWER_LABELS.forEach(label => {
    const pmap = Object.fromEntries((spreadData.forward_curves[label] || []).map(p => [p.month, p.price]));
    const sp = [], cl = [];
    SPARK_MONTHS.forEach(m => {
      if (pmap[m] == null || jkmMap[m] == null) { sp.push(null); cl.push(null); return; }
      const gas = (jkmMap[m] / gasEff) / GAS_CONV;
      const s = pmap[m] - gas;
      sp.push(+s.toFixed(4));
      cl.push(+(s - co2).toFixed(4));
    });
    sparkSeries.push({ name: label, data: sp });
    cleanSeries.push({ name: label, data: cl });
  });
  return { sparkSeries, cleanSeries };
}

function sparkChartOpts(series, height) {
  return Object.assign({}, apexCommonOpts, {
    chart: Object.assign({type:'line', height:height}, apexCommonOpts.chart),
    series: series,
    xaxis: { categories: SPARK_MONTHS.map(fmtMonth), labels:{ style:{colors:'#94a3b8'}, rotate:-45 } },
    yaxis: { labels:{ style:{colors:'#94a3b8'}, formatter:(v)=> v==null?'-':v.toFixed(2) }, title:{ text:'¥/kWh', style:{color:'#94a3b8'} } },
    colors: SP_COLORS,
    stroke: { width:2.2, curve:'straight' },
    markers: { size:0, hover:{size:4} },
    legend: { labels:{ colors:'#cbd5e1' }, position:'top' },
    dataLabels: { enabled:false },
    annotations: { yaxis: [{ y:0, borderColor:'#64748b', strokeDashArray:4,
      label:{ text:'損益分岐 0', style:{ color:'#94a3b8', background:'transparent' } } }] },
  });
}

const _init = computeSpark(
  spreadData.params.gas_thermal_eff,
  spreadData.params.gas_co2_kg_per_kwh,
  spreadData.params.co2_price_yen_per_t
);
let sparkChart = null, cleanChart = null;
if (SPARK_MONTHS.length) {
  sparkChart = new ApexCharts(document.querySelector('#sp-spark-chart'), sparkChartOpts(_init.sparkSeries, 380));
  cleanChart = new ApexCharts(document.querySelector('#sp-clean-chart'), sparkChartOpts(_init.cleanSeries, 380));
  sparkChart.render();
  cleanChart.render();
} else {
  ['#sp-spark-chart','#sp-clean-chart'].forEach(id => {
    const e = document.querySelector(id);
    if (e) e.innerHTML = '<div style="padding:24px;opacity:0.5">電力∩JKM の重なる限月がありません</div>';
  });
}

// ── Sliders ──
const effEl = document.querySelector('#sp-eff');
const co2fEl = document.querySelector('#sp-co2f');
const co2pEl = document.querySelector('#sp-co2p');
const DEFAULTS = {
  eff: spreadData.params.gas_thermal_eff,
  co2f: spreadData.params.gas_co2_kg_per_kwh,
  co2p: spreadData.params.co2_price_yen_per_t,
};

function setLabels() {
  document.querySelector('#sp-eff-val').textContent = (+effEl.value).toFixed(2);
  document.querySelector('#sp-co2f-val').textContent = (+co2fEl.value).toFixed(2);
  document.querySelector('#sp-co2p-val').textContent = (+co2pEl.value).toLocaleString();
}
function recompute() {
  setLabels();
  if (!sparkChart) return;
  const r = computeSpark(+effEl.value, +co2fEl.value, +co2pEl.value);
  sparkChart.updateSeries(r.sparkSeries, true);
  cleanChart.updateSeries(r.cleanSeries, true);
}
function resetDefaults() {
  effEl.value = DEFAULTS.eff;
  co2fEl.value = DEFAULTS.co2f;
  co2pEl.value = DEFAULTS.co2p;
  recompute();
}
if (effEl) {
  resetDefaults();
  [effEl, co2fEl, co2pEl].forEach(el => el.addEventListener('input', recompute));
  document.querySelector('#sp-reset').addEventListener('click', resetDefaults);
}
</script>

</body>
</html>"""


def generate_spread_html(data: dict) -> str:
    """Render docs/spread_analysis.html from the spread payload."""
    p = data.get("params", {})
    dark = data.get("dark_spread") or {}
    params_rows = "".join([
        f"<tr><td>ガス火力 熱効率</td><td>{p.get('gas_thermal_eff')}</td></tr>",
        f"<tr><td>ガス単位換算</td><td>{p.get('gas_conv_kwh_mmbtu')} kWh/MMBtu</td></tr>",
        f"<tr><td>ガス CO2 排出係数</td><td>{p.get('gas_co2_kg_per_kwh')} kg-CO2/kWh</td></tr>",
        f"<tr><td>CO2 価格</td><td>{p.get('co2_price_yen_per_t')} 円/t</td></tr>",
        f"<tr><td>JKM 単位の仮定</td><td>{p.get('jkm_unit_assumption')}</td></tr>",
        f"<tr><td>JKM 系列</td><td>{p.get('jkm_underlying')}</td></tr>",
        f"<tr><td>FX 系列</td><td>{p.get('fx_underlying')}</td></tr>",
        f"<tr><td>Dark Spread (石炭)</td><td>{dark.get('note', '—')}</td></tr>",
    ])
    payload = json.dumps(data, ensure_ascii=False)
    return (
        _SPREAD_HTML_TEMPLATE
        .replace("__PAYLOAD__", payload)
        .replace("__LATEST__", _fmt_date_dotted(data.get("latest_date", "N/A")))
        .replace("__HORIZON__", str(data.get("horizon") or "—"))
        .replace("__PARAMS_ROWS__", params_rows)
    )


def write_spread_analysis(repo, site_dir: Path, latest_date: str) -> Path | None:
    """Generate and write docs/spread_analysis.html.

    Mirrors write_weekly_compare's lifecycle so the page stays in sync with the
    daily dashboard. Returns the path, or None if no power forward curves exist.
    """
    prev_date = find_nearest_business_day_before(repo, latest_date, days_back=7)
    data = generate_spread_data(repo, latest_date, prev_date)
    if not any(data.get("forward_curves", {}).values()):
        return None
    html = generate_spread_html(data)
    out_path = site_dir / "spread_analysis.html"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


def main():
    repo = get_repository()
    try:
        print("Generating site data...")
        data = generate_data_json(repo)

        if not data:
            print("ERROR: No data found. Run import_csv.py first.")
            return

        # Write data.json
        SITE_DIR.mkdir(parents=True, exist_ok=True)
        data_path = SITE_DIR / "data.json"
        with open(data_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"  Written: {data_path}")

        # Write index.html
        html = generate_html(data)
        html_path = SITE_DIR / "index.html"
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"  Written: {html_path}")

        # Weekly compare page (1 週間前と比較)
        latest_date = data["latest_date"]
        wk_html_path = write_weekly_compare(repo, SITE_DIR, latest_date)
        if wk_html_path:
            print(f"  Written: {wk_html_path}")
        else:
            print("  Skipped weekly_compare.html (insufficient history before latest_date).")

        # Spread analysis page (スプレッド分析 — Spark/Clean Spark + forward curves)
        sp_html_path = write_spread_analysis(repo, SITE_DIR, latest_date)
        if sp_html_path:
            print(f"  Written: {sp_html_path}")
        else:
            print("  Skipped spread_analysis.html (no power forward curves).")

        # Report curve stats
        curves = data.get("forward_curves", {})
        commodity_count = data.get("commodity_count", 0)
        print(f"\nSite generated successfully!")
        print(f"  Total records: {data['total_records']:,}")
        print(f"  Power futures: {data['power_futures_count']}")
        print(f"  Forward curves: {len(curves)}")
        for cat, items in sorted(curves.items()):
            print(f"    {cat}: {len(items)} 限月")
        print(f"  Commodities: {commodity_count}")
        print(f"\n  Open in browser: {html_path}")
    finally:
        repo.close()


if __name__ == "__main__":
    main()
