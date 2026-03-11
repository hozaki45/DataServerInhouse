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


def generate_data_json(repo) -> dict:
    """Generate data.json from database for Chart.js consumption."""
    log = repo.get_import_log()
    if not log:
        return {}

    latest_date = log[0]["trade_date"]

    # Power futures
    power = get_power_futures(repo, latest_date)
    power_fut = [r for r in power if not r.get("put_call")]

    # Build forward curves per category
    forward_curves = {}
    for r in power_fut:
        cat = classify_power_future(r["instrument_name"])
        if cat is None:
            continue
        if cat not in forward_curves:
            forward_curves[cat] = []

        month = r.get("contract_month", "")
        forward_curves[cat].append({
            "month": month,
            "settlement": r.get("settlement_price"),
            "theoretical": r.get("theoretical_price"),
            "days": r.get("days_to_expiry"),
            "name": r["instrument_name"],
        })

    # Sort each curve by month/date
    for cat in forward_curves:
        forward_curves[cat].sort(key=lambda x: x["month"])

    # Overview bar chart: monthly curves for the 4 main types
    overview_chart = {}
    for cat_name in ["東・ベース(月次)", "東・日中(月次)", "西・ベース(月次)", "西・日中(月次)"]:
        if cat_name in forward_curves:
            overview_chart[cat_name] = [
                {"month": p["month"], "price": p["settlement"]}
                for p in forward_curves[cat_name]
                if p["settlement"] is not None
            ]

    # Summary by underlying
    summary = summary_by_underlying(repo, latest_date)

    # Import log
    import_dates = [entry["trade_date"] for entry in log if entry["status"] == "success"]

    return {
        "latest_date": latest_date,
        "generated_at": datetime.now().isoformat(),
        "total_records": sum(s["total"] for s in summary),
        "underlying_count": len(summary),
        "import_dates": import_dates,
        "power_futures_count": len(power_fut),
        "overview_chart": overview_chart,
        "forward_curves": forward_curves,
        "power_futures": [
            {
                "name": r["instrument_name"],
                "underlying": r.get("underlying_name", ""),
                "month": r.get("contract_month", ""),
                "settlement": r.get("settlement_price"),
                "theoretical": r.get("theoretical_price"),
                "days": r.get("days_to_expiry"),
                "category": classify_power_future(r["instrument_name"]),
            }
            for r in power_fut
        ],
        "summary": summary,
    }


def generate_html(data: dict) -> str:
    """Generate the HTML dashboard page."""
    latest = data.get("latest_date", "N/A")
    generated = data.get("generated_at", "")[:19].replace("T", " ")
    total = data.get("total_records", 0)
    underlying_count = data.get("underlying_count", 0)
    dates = data.get("import_dates", [])
    power_count = data.get("power_futures_count", 0)

    # Power futures table rows
    power_rows = ""
    for pf in data.get("power_futures", []):
        settle = pf["settlement"] if pf["settlement"] is not None else ""
        theo = pf["theoretical"] if pf["theoretical"] is not None else ""
        days = pf["days"] if pf["days"] is not None else ""
        cat = pf.get("category", "") or ""
        power_rows += f"""        <tr>
          <td>{pf['name']}</td>
          <td>{cat}</td>
          <td>{pf['month']}</td>
          <td class="num">{settle}</td>
          <td class="num">{theo}</td>
          <td class="num">{days}</td>
        </tr>\n"""

    # Summary table rows
    summary_rows = ""
    for s in data.get("summary", []):
        summary_rows += f"""        <tr>
          <td>{s['underlying_name']}</td>
          <td class="num">{s['total']}</td>
          <td class="num">{s['fut']}</td>
          <td class="num">{s['cal']}</td>
          <td class="num">{s['put']}</td>
        </tr>\n"""

    # Build curve selector options
    curve_options = ""
    for cat_name in sorted(data.get("forward_curves", {}).keys()):
        curve_options += f'    <option value="{cat_name}">{cat_name}</option>\n'

    # Inline JSON for chart data (avoids fetch/CORS issues on file://)
    chart_data = {
        "overview_chart": data.get("overview_chart", {}),
        "forward_curves": data.get("forward_curves", {}),
    }
    chart_data_json = json.dumps(chart_data, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>DataServer In-House — 電力先物データダッシュボード</title>
  <link rel="stylesheet" href="style.css">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
</head>
<body>

<div class="header">
  <div>
    <h1>DataServer In-House</h1>
    <div class="subtitle">電力先物データ分析基盤 — JPXデリバティブ理論価格</div>
  </div>
  <div class="updated">
    最新データ: {latest}<br>
    生成日時: {generated}
  </div>
</div>

<div class="container">

  <!-- KPI -->
  <div class="kpi-row">
    <div class="kpi-box">
      <div class="value">{total:,}</div>
      <div class="label">総レコード数</div>
    </div>
    <div class="kpi-box">
      <div class="value">{underlying_count}</div>
      <div class="label">原資産数</div>
    </div>
    <div class="kpi-box">
      <div class="value">{power_count}</div>
      <div class="label">電力先物契約数</div>
    </div>
    <div class="kpi-box">
      <div class="value">{len(dates)}</div>
      <div class="label">取り込み済み日数</div>
    </div>
  </div>

  <!-- Overview: 4 main curves comparison -->
  <div class="card">
    <div class="card-header">電力先物フォワードカーブ 比較（月次 4エリア）</div>
    <div class="card-body">
      <div class="chart-container" style="height:400px">
        <canvas id="overviewChart"></canvas>
      </div>
    </div>
  </div>

  <!-- Interactive Forward Curve Selector -->
  <div class="card">
    <div class="card-header">フォワードカーブ（プルダウン選択）</div>
    <div class="card-body">
      <div class="curve-selector">
        <label for="curveSelect">カーブ選択:</label>
        <select id="curveSelect">
{curve_options}        </select>
        <span class="curve-info" id="curveInfo"></span>
      </div>
      <div class="chart-container" style="height:350px">
        <canvas id="curveChart"></canvas>
      </div>
      <div id="curveTable"></div>
    </div>
  </div>

  <!-- East vs West Spread -->
  <div class="card">
    <div class="card-header">東西スプレッド（東ベース − 西ベース、月次）</div>
    <div class="card-body">
      <div class="chart-container" style="height:300px">
        <canvas id="spreadChart"></canvas>
      </div>
    </div>
  </div>

  <!-- Power Futures Table -->
  <div class="card">
    <div class="card-header">電力先物 全データ一覧（{latest}）</div>
    <div class="card-body">
      <table>
        <thead>
          <tr>
            <th>銘柄名称</th>
            <th>カテゴリ</th>
            <th>限月</th>
            <th class="num">清算価格</th>
            <th class="num">理論価格</th>
            <th class="num">残日数</th>
          </tr>
        </thead>
        <tbody>
{power_rows}        </tbody>
      </table>
    </div>
  </div>

  <!-- All Underlying Summary -->
  <div class="card">
    <div class="card-header">原資産別データ件数（{latest}）</div>
    <div class="card-body">
      <table>
        <thead>
          <tr>
            <th>原資産名称</th>
            <th class="num">合計</th>
            <th class="num">先物</th>
            <th class="num">コール</th>
            <th class="num">プット</th>
          </tr>
        </thead>
        <tbody>
{summary_rows}        </tbody>
      </table>
    </div>
  </div>

</div>

<div class="footer">
  DataServer In-House — Generated automatically from JPX derivative data
</div>

<script>
let curveChartInstance = null;

const data = {chart_data_json};

(function() {{
    // ============================================
    // 1. Overview: 4 main monthly curves (line chart)
    // ============================================
    const overview = data.overview_chart;
    const lineColors = {{
      '東・ベース(月次)': {{ border: '#2E75B6', bg: 'rgba(46,117,182,0.1)' }},
      '東・日中(月次)': {{ border: '#5DADE2', bg: 'rgba(93,173,226,0.1)' }},
      '西・ベース(月次)': {{ border: '#E67E22', bg: 'rgba(230,126,34,0.1)' }},
      '西・日中(月次)': {{ border: '#F5B041', bg: 'rgba(245,176,65,0.1)' }}
    }};

    const allMonths = new Set();
    Object.values(overview).forEach(items => {{
      items.forEach(item => allMonths.add(item.month));
    }});
    const months = Array.from(allMonths).sort();

    const overviewDatasets = Object.entries(overview).map(([cat, items]) => {{
      const priceMap = {{}};
      items.forEach(item => {{ priceMap[item.month] = item.price; }});
      const c = lineColors[cat] || {{ border: '#999', bg: 'rgba(153,153,153,0.1)' }};
      return {{
        label: cat,
        data: months.map(m => priceMap[m] || null),
        borderColor: c.border,
        backgroundColor: c.bg,
        borderWidth: 2.5,
        pointRadius: 4,
        pointHoverRadius: 6,
        tension: 0.3,
        fill: false,
        spanGaps: true
      }};
    }});

    new Chart(document.getElementById('overviewChart'), {{
      type: 'line',
      data: {{
        labels: months.map(m => m.substring(0, 4) + '/' + m.substring(4)),
        datasets: overviewDatasets
      }},
      options: {{
        responsive: true,
        maintainAspectRatio: false,
        interaction: {{ mode: 'index', intersect: false }},
        plugins: {{
          legend: {{ position: 'top' }},
          tooltip: {{
            callbacks: {{
              label: ctx => ctx.dataset.label + ': ' + (ctx.parsed.y !== null ? ctx.parsed.y.toFixed(2) + ' 円/kWh' : 'N/A')
            }}
          }}
        }},
        scales: {{
          y: {{
            title: {{ display: true, text: '円/kWh' }},
            grid: {{ color: 'rgba(0,0,0,0.06)' }}
          }},
          x: {{
            title: {{ display: true, text: '限月' }},
            grid: {{ display: false }}
          }}
        }}
      }}
    }});

    // ============================================
    // 2. Interactive curve selector
    // ============================================
    const curves = data.forward_curves;
    const select = document.getElementById('curveSelect');

    function renderCurve(catName) {{
      const items = curves[catName] || [];
      const info = document.getElementById('curveInfo');
      info.textContent = items.length + ' 限月';

      // Destroy previous chart
      if (curveChartInstance) curveChartInstance.destroy();

      const labels = items.map(i => {{
        const m = i.month;
        if (m.length === 6) return m.substring(0, 4) + '/' + m.substring(4);
        if (m.length === 8) return m.substring(0, 4) + '/' + m.substring(4, 6) + '/' + m.substring(6);
        return m;
      }});

      const datasets = [
        {{
          label: '清算価格',
          data: items.map(i => i.settlement),
          borderColor: '#2E75B6',
          backgroundColor: 'rgba(46,117,182,0.15)',
          borderWidth: 2.5,
          pointRadius: 5,
          pointHoverRadius: 8,
          pointBackgroundColor: '#2E75B6',
          tension: 0.3,
          fill: true
        }}
      ];

      // Add theoretical price if available
      const hasTheo = items.some(i => i.theoretical !== null);
      if (hasTheo) {{
        datasets.push({{
          label: '理論価格',
          data: items.map(i => i.theoretical),
          borderColor: '#E74C3C',
          backgroundColor: 'rgba(231,76,60,0.05)',
          borderWidth: 1.5,
          borderDash: [5, 3],
          pointRadius: 3,
          pointHoverRadius: 6,
          tension: 0.3,
          fill: false
        }});
      }}

      curveChartInstance = new Chart(document.getElementById('curveChart'), {{
        type: 'line',
        data: {{ labels, datasets }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          interaction: {{ mode: 'index', intersect: false }},
          plugins: {{
            title: {{ display: true, text: catName + ' フォワードカーブ', font: {{ size: 16, weight: 'bold' }}, color: '#0C2140' }},
            tooltip: {{
              callbacks: {{
                afterLabel: ctx => {{
                  const item = items[ctx.dataIndex];
                  return '残日数: ' + (item.days || 'N/A');
                }}
              }}
            }}
          }},
          scales: {{
            y: {{
              title: {{ display: true, text: '円/kWh' }},
              grid: {{ color: 'rgba(0,0,0,0.06)' }}
            }},
            x: {{
              title: {{ display: true, text: '限月' }},
              grid: {{ display: false }}
            }}
          }}
        }}
      }});

      // Render detail table
      let tableHtml = '<table><thead><tr><th>銘柄</th><th>限月</th><th class="num">清算価格</th><th class="num">理論価格</th><th class="num">残日数</th></tr></thead><tbody>';
      items.forEach(i => {{
        const s = i.settlement !== null ? i.settlement : '';
        const t = i.theoretical !== null ? i.theoretical : '';
        const d = i.days !== null ? i.days : '';
        tableHtml += '<tr><td>' + i.name + '</td><td>' + i.month + '</td><td class="num">' + s + '</td><td class="num">' + t + '</td><td class="num">' + d + '</td></tr>';
      }});
      tableHtml += '</tbody></table>';
      document.getElementById('curveTable').innerHTML = tableHtml;
    }}

    select.addEventListener('change', () => renderCurve(select.value));
    if (select.options.length > 0) renderCurve(select.value);

    // ============================================
    // 3. East vs West spread chart
    // ============================================
    const eastBase = curves['東・ベース(月次)'] || [];
    const westBase = curves['西・ベース(月次)'] || [];

    if (eastBase.length && westBase.length) {{
      const westMap = {{}};
      westBase.forEach(w => {{ westMap[w.month] = w.settlement; }});

      const spreadMonths = [];
      const spreadValues = [];
      const spreadColors = [];

      eastBase.forEach(e => {{
        if (e.settlement !== null && westMap[e.month] !== undefined && westMap[e.month] !== null) {{
          const spread = e.settlement - westMap[e.month];
          const m = e.month;
          spreadMonths.push(m.substring(0, 4) + '/' + m.substring(4));
          spreadValues.push(Math.round(spread * 100) / 100);
          spreadColors.push(spread >= 0 ? 'rgba(46,117,182,0.7)' : 'rgba(231,76,60,0.7)');
        }}
      }});

      new Chart(document.getElementById('spreadChart'), {{
        type: 'bar',
        data: {{
          labels: spreadMonths,
          datasets: [{{
            label: '東西スプレッド（東−西）',
            data: spreadValues,
            backgroundColor: spreadColors,
            borderColor: spreadColors.map(c => c.replace('0.7', '1')),
            borderWidth: 1
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            tooltip: {{
              callbacks: {{
                label: ctx => '東−西: ' + ctx.parsed.y.toFixed(2) + ' 円/kWh'
              }}
            }}
          }},
          scales: {{
            y: {{
              title: {{ display: true, text: '円/kWh' }},
              grid: {{ color: 'rgba(0,0,0,0.06)' }}
            }},
            x: {{
              title: {{ display: true, text: '限月' }},
              grid: {{ display: false }}
            }}
          }}
        }}
      }});
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
