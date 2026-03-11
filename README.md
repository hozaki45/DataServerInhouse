# DataServer In-House

JPX デリバティブ清算価格の自動収集・分析基盤。電力先物を中心とした市場データを日次で蓄積し、インタラクティブなダッシュボードで可視化します。

**Dashboard**: https://hozaki45.github.io/DataServerInhouse/

---

## What's New

### 2026-03-12: Day-over-Day Change Tracking

- **Top Movers**: 前日比が最も大きい15銘柄をハイライト表示（緑=上昇/赤=下落）
- **Forward Curve Overlay**: 前日カーブをアンバー色破線でオーバーレイ比較
- **DoD Change列**: 全テーブルに前日比（差額+変化率%）を色分け表示
- **カスタムツールチップ**: ホバー時に前日価格・変化額・変化率を表示

### 2026-03-11: Daily Auto-Fetch (GitHub Actions)

- 毎営業日 17:00 JST に JPX から最新清算価格CSVを自動取得
- データベース更新 → ダッシュボード再生成 → 自動コミット&デプロイ
- 手動トリガーも対応（Actions > "Daily JPX Data Update" > Run workflow）

### 2026-03-11: 2025 Dark Mode Dashboard

- ダークモード + グラスモーフィズムの最新UIデザイン
- ApexCharts によるグラデーションエリアチャート
- Bento Grid レイアウト + サイドバーナビゲーション
- KPI カウントアップアニメーション

---

## Architecture

```
Data Source          Pipeline              Storage            Frontend
┌──────────┐     ┌──────────────┐     ┌─────────────┐    ┌──────────────┐
│ JPX Web  │────>│ fetch_jpx.py │────>│ SQLite DB   │───>│ GitHub Pages │
│ (CSV)    │     │ csv_parser   │     │ (market_data│    │ (HTML+JS)    │
└──────────┘     │ data_loader  │     │  .db)       │    └──────────────┘
                 └──────────────┘     └─────────────┘
                       │                                        ^
                 GitHub Actions                                 │
                 (daily cron)───────── commit & push ───────────┘
```

## Data Coverage

| Category | Assets | Description |
|----------|--------|-------------|
| Power Futures | 12 curves, 124 contracts | East/West, Base/Peak, Monthly/Weekly/Annual |
| Commodities | Gold, Platinum, Rubber, etc. | TOCOM derivatives |
| Equity Index | Nikkei 225, TOPIX, etc. | Index futures & options |
| Volatility | VIX-related | Volatility futures |
| Bonds | JGB futures | Government bond derivatives |

## Project Structure

```
DataServerInhouse/
├── .github/workflows/
│   └── daily-update.yml       # GitHub Actions: daily auto-fetch
├── Data/                      # Raw CSV files (rb{YYYYMMDD}.csv)
├── db/
│   └── market_data.db         # SQLite database
├── docs/                      # GitHub Pages dashboard
│   ├── index.html
│   ├── style.css
│   └── data.json
├── scripts/
│   ├── fetch_jpx.py           # JPX scraper + import + site regen
│   ├── generate_site.py       # Dashboard HTML generator
│   ├── import_csv.py          # Manual CSV import CLI
│   ├── check_data.py          # Data verification
│   └── generate_presentation.py  # PowerPoint generator
├── src/
│   ├── config.py              # Environment-based configuration
│   ├── csv_parser.py          # CP932 CSV parser
│   ├── data_loader.py         # Import orchestrator
│   ├── db_schema.py           # SQLite schema
│   ├── query.py               # Query utilities
│   ├── repository.py          # Data access abstraction (SQLite/DynamoDB)
│   └── storage.py             # Storage abstraction (Local/S3)
└── pyproject.toml
```

## Quick Start

```bash
# Install dependencies
uv sync

# Manual import of a CSV file
uv run python scripts/import_csv.py

# Fetch latest data from JPX
uv run python scripts/fetch_jpx.py

# Regenerate dashboard
uv run python scripts/generate_site.py
```

## Technology

- **Python 3.12** + uv package manager
- **SQLite** (WAL mode) — designed for future DynamoDB migration
- **ApexCharts** — interactive dark-mode charts (CDN)
- **GitHub Actions** — daily automated data pipeline
- **GitHub Pages** — static dashboard hosting

## Future Roadmap

- AWS serverless migration (DynamoDB + Lambda + S3)
- Multi-day time series charts (price history over weeks/months)
- Options analytics (volatility surface, Greeks)
- Alert system for significant price movements
