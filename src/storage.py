"""Storage abstraction layer.

Provides a unified interface for file access, supporting local filesystem
and S3 (future). Switch via DATA_STORAGE environment variable.
"""

from __future__ import annotations

import re
from abc import ABC, abstractmethod
from pathlib import Path

from src.config import DATA_DIR, DATA_STORAGE


class StorageBackend(ABC):
    """Abstract base class for storage backends."""

    @abstractmethod
    def list_csv_files(self) -> list[str]:
        """Return sorted list of CSV filenames (e.g., ['rb20260310.csv'])."""

    @abstractmethod
    def read_file(self, filename: str) -> bytes:
        """Read file contents as bytes."""

    @abstractmethod
    def file_exists(self, filename: str) -> bool:
        """Check if a file exists."""


class LocalStorage(StorageBackend):
    """Local filesystem storage backend."""

    def __init__(self, data_dir: Path = DATA_DIR):
        self.data_dir = data_dir

    def list_csv_files(self) -> list[str]:
        pattern = re.compile(r"^rb\d{8}\.csv$")
        files = [
            f.name
            for f in self.data_dir.iterdir()
            if f.is_file() and pattern.match(f.name)
        ]
        return sorted(files)

    def read_file(self, filename: str) -> bytes:
        file_path = self.data_dir / filename
        return file_path.read_bytes()

    def file_exists(self, filename: str) -> bool:
        return (self.data_dir / filename).is_file()


def get_storage() -> StorageBackend:
    """Factory function to get the configured storage backend."""
    if DATA_STORAGE == "local":
        return LocalStorage()
    elif DATA_STORAGE == "s3":
        raise NotImplementedError("S3 storage backend not yet implemented")
    else:
        raise ValueError(f"Unknown DATA_STORAGE: {DATA_STORAGE}")
