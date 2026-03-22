"""Fetch latest JPX derivative settlement price CSV.

Scrapes the JPX settlement price page to find the CSV download link,
downloads the file, saves it to Data/, imports into DB, and regenerates the site.

Usage:
    uv run python scripts/fetch_jpx.py
"""

import re
import sys
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import requests
from bs4 import BeautifulSoup

from src.config import DATA_DIR
from src.data_loader import DuplicateDataError, import_csv_file
from src.repository import get_repository
from src.storage import get_storage

JPX_BASE_URL = "https://www.jpx.co.jp"
JPX_PAGE_URL = (
    f"{JPX_BASE_URL}/markets/derivatives/settlement-price/index.html"
)
CSV_PATTERN = re.compile(r"rb\d{8}\.csv")


def find_csv_url() -> tuple[str, str]:
    """Scrape JPX page to find the latest CSV download URL.

    Returns (full_url, filename) tuple.
    """
    print(f"Fetching JPX page: {JPX_PAGE_URL}")
    resp = requests.get(JPX_PAGE_URL, timeout=30)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")

    # Look for CSV download link
    for a_tag in soup.find_all("a", href=True):
        href = a_tag["href"]
        match = CSV_PATTERN.search(href)
        if match:
            filename = match.group(0)
            full_url = href if href.startswith("http") else JPX_BASE_URL + href
            print(f"  Found CSV link: {filename}")
            return full_url, filename

    raise RuntimeError(
        "Could not find CSV download link on JPX page. "
        "The page structure may have changed."
    )


def download_csv(url: str, filename: str) -> Path:
    """Download CSV file and save to Data/ directory."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    filepath = DATA_DIR / filename

    if filepath.exists():
        print(f"  File already exists: {filepath}")
        return filepath

    print(f"  Downloading: {url}")
    resp = requests.get(url, timeout=60)
    resp.raise_for_status()

    filepath.write_bytes(resp.content)
    print(f"  Saved: {filepath} ({len(resp.content):,} bytes)")
    return filepath


def main():
    print(f"=== JPX Settlement Price Fetcher ===")
    print(f"    {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()

    # Step 1: Find CSV URL
    csv_url, filename = find_csv_url()

    # Step 2: Download (always download to allow hash comparison)
    download_csv(csv_url, filename)

    # Step 3: Import into database (with hash-based duplicate detection)
    storage = get_storage()
    repo = get_repository()

    try:
        imported_files = set(repo.get_imported_files())
        if filename in imported_files:
            print(f"\n  Already imported: {filename} (skipping)")
        else:
            print(f"\n  Importing: {filename}")
            try:
                trade_date, count = import_csv_file(filename, storage, repo)
                print(f"  Imported {count:,} records for {trade_date}")
            except DuplicateDataError as e:
                print(f"\n  SKIPPED (holiday duplicate): {e}")
                return

        # Step 4: Regenerate site
        print("\n  Regenerating dashboard...")
        # Import here to avoid circular imports at module level
        from scripts.generate_site import generate_data_json, generate_html

        site_dir = Path(__file__).resolve().parent.parent / "docs"
        data = generate_data_json(repo)

        if data:
            import json

            site_dir.mkdir(parents=True, exist_ok=True)

            data_path = site_dir / "data.json"
            with open(data_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            html = generate_html(data)
            html_path = site_dir / "index.html"
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html)

            print(f"  Site regenerated: {html_path}")
        else:
            print("  WARNING: No data for site generation")

    finally:
        repo.close()

    print("\nDone!")


if __name__ == "__main__":
    main()
