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
    get_cross_commodity_snapshot,
    get_all_commodity_forward_curves,
)
from src.asset_taxonomy import CATEGORY_META, ASSET_TAXONOMY

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
