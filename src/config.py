"""Configuration module for DataServer In-House.

Supports environment variable-based switching between local and AWS backends.
"""

import os
from pathlib import Path

# Project root directory
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Storage backend: "local" or "s3"
DATA_STORAGE = os.environ.get("DATA_STORAGE", "local")

# Database backend: "sqlite" or "dynamodb"
DB_BACKEND = os.environ.get("DB_BACKEND", "sqlite")

# Local paths
DATA_DIR = Path(os.environ.get("DATA_DIR", str(PROJECT_ROOT / "Data")))
DB_DIR = Path(os.environ.get("DB_DIR", str(PROJECT_ROOT / "db")))
SQLITE_DB_PATH = DB_DIR / "market_data.db"

# AWS settings (used when DATA_STORAGE=s3 or DB_BACKEND=dynamodb)
AWS_REGION = os.environ.get("AWS_REGION", "ap-northeast-1")
S3_BUCKET = os.environ.get("S3_BUCKET", "")
DYNAMODB_TABLE = os.environ.get("DYNAMODB_TABLE", "derivative_prices")

# CSV parsing settings
CSV_ENCODING = "cp932"
CSV_HEADER_ROWS = 3  # Number of header/comment rows to skip
CSV_COLUMNS = [
    "instrument_code",
    "instrument_name",
    "put_call",
    "contract_month",
    "strike_price",
    "settlement_price",
    "theoretical_price",
    "underlying_price",
    "volatility",
    "interest_rate",
    "days_to_expiry",
    "underlying_name",
]
