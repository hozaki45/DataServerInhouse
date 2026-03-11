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

    # Inline JSON for chart data (avoids fetch/CORS issues on file://)
    chart_data = {
        "overview_chart": data.get("overview_chart", {}),
        "prev_overview_chart": data.get("prev_overview_chart", {}),
        "forward_curves": data.get("forward_curves", {}),
        "prev_forward_curves": data.get("prev_forward_curves", {}),
    }
    chart_data_json = json.dumps(chart_data, ensure_ascii=False)
    prev_label = prev_date or "N/A"

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>DataServer In-House — Energy Futures Dashboard</title>
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
        <div class="subtitle">JPX Derivative Theoretical Price — Energy Futures Analytics</div>
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
      <div class="kpi-row">
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
      </div>

      <!-- Bento Grid -->
      <div class="bento-grid">

        <!-- Overview Chart (large) -->
        <div class="card card-full" id="overview">
          <div class="card-header">
            <div class="card-title"><span class="icon">&#9670;</span> Forward Curve Comparison (Monthly)</div>
            <div class="card-badge">4 Areas</div>
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
      DataServer In-House — JPX Derivative Data Analytics Platform &middot; Generated {generated}
    </div>

  </div><!-- /main-content -->
</div><!-- /app-layout -->

<script>
const chartData = {chart_data_json};

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
    // 1. Overview: 4 main monthly curves
    // ============================================
    const overview = chartData.overview_chart;
    const lineConfigs = {{
      '\u6771\u30fb\u30d9\u30fc\u30b9(\u6708\u6b21)': {{ color: colors.blue, name: 'East Base' }},
      '\u6771\u30fb\u65e5\u4e2d(\u6708\u6b21)': {{ color: colors.cyan, name: 'East Peak' }},
      '\u897f\u30fb\u30d9\u30fc\u30b9(\u6708\u6b21)': {{ color: colors.amber, name: 'West Base' }},
      '\u897f\u30fb\u65e5\u4e2d(\u6708\u6b21)': {{ color: colors.pink, name: 'West Peak' }}
    }};

    const allMonths = new Set();
    Object.values(overview).forEach(items => {{
      items.forEach(item => allMonths.add(item.month));
    }});
    const months = Array.from(allMonths).sort();
    const monthLabels = months.map(m => m.substring(0, 4) + '/' + m.substring(4));

    const prevOverview = chartData.prev_overview_chart || {{}};

    const overviewSeries = [];
    const overviewColors = [];
    const overviewDash = [];
    const overviewWidths = [];

    Object.entries(overview).forEach(([cat, items]) => {{
      const priceMap = {{}};
      items.forEach(item => {{ priceMap[item.month] = item.price; }});
      const cfg = lineConfigs[cat] || {{ color: '#999', name: cat }};
      overviewSeries.push({{
        name: cfg.name,
        data: months.map(m => priceMap[m] || null)
      }});
      overviewColors.push(cfg.color);
      overviewDash.push(0);
      overviewWidths.push(2.5);
    }});

    // Add previous day lines (dashed, dimmed)
    Object.entries(overview).forEach(([cat, items]) => {{
      const prevItems = prevOverview[cat] || [];
      if (prevItems.length === 0) return;
      const prevMap = {{}};
      prevItems.forEach(item => {{ prevMap[item.month] = item.price; }});
      const cfg = lineConfigs[cat] || {{ color: '#999', name: cat }};
      overviewSeries.push({{
        name: cfg.name + ' (prev)',
        data: months.map(m => prevMap[m] || null)
      }});
      overviewColors.push(cfg.color + '55');
      overviewDash.push(5);
      overviewWidths.push(1.2);
    }});

    new ApexCharts(document.getElementById('overviewChart'), {{
      ...baseChartOpts,
      chart: {{ ...baseChartOpts.chart, type: 'area', height: 380, animations: {{ enabled: true, easing: 'easeinout', speed: 800 }} }},
      series: overviewSeries,
      colors: overviewColors,
      stroke: {{ curve: 'smooth', width: overviewWidths, dashArray: overviewDash }},
      fill: {{ type: 'gradient', gradient: {{ shadeIntensity: 1, opacityFrom: 0.15, opacityTo: 0.01, stops: [0, 95, 100] }} }},
      xaxis: {{ ...baseChartOpts.xaxis, categories: monthLabels, title: {{ text: 'Contract Month', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
      yaxis: {{ ...baseChartOpts.yaxis, title: {{ text: 'JPY/kWh', style: {{ color: colors.textLight, fontSize: '11px' }} }} }},
      markers: {{ size: 3, strokeWidth: 0, hover: {{ size: 6 }} }},
      legend: {{ ...baseChartOpts.legend, position: 'top', horizontalAlign: 'right' }},
      dataLabels: {{ enabled: false }}
    }}).render();

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
    // 4. KPI count-up animation
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
    // 5. Sidebar active state on scroll
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
    // 6. Spot Analysis — CSV Upload & Processing
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
        print(f"\nSite generated successfully!")
        print(f"  Total records: {data['total_records']:,}")
        print(f"  Power futures: {data['power_futures_count']}")
        print(f"  Forward curves: {len(curves)}")
        for cat, items in sorted(curves.items()):
            print(f"    {cat}: {len(items)} 限月")
        print(f"\n  Open in browser: {html_path}")
    finally:
        repo.close()


if __name__ == "__main__":
    main()
