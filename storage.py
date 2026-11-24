from __future__ import annotations

import hashlib
import os
import re
import tempfile
from pathlib import Path

from analyzer import format_month


HISTORY_PATTERN = re.compile(r"^(20\d{2})-(0[1-9]|1[0-2])\.xlsx$")


def file_digest(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()


def history_files(history_directory: Path) -> list[Path]:
    history_directory.mkdir(parents=True, exist_ok=True)
    return sorted(
        (
            path
            for path in history_directory.iterdir()
            if path.is_file() and HISTORY_PATTERN.fullmatch(path.name)
        ),
        reverse=True,
    )


def month_path(history_directory: Path, year: int, month: int) -> Path:
    return history_directory / f"{format_month(year, month)}.xlsx"


def save_monthly_workbook(
    content: bytes,
    year: int,
    month: int,
    history_directory: Path,
    replace_existing: bool = False,
) -> str:
    history_directory.mkdir(parents=True, exist_ok=True)
    destination = month_path(history_directory, year, month)
    if destination.exists():
        if file_digest(destination.read_bytes()) == file_digest(content):
            return "duplicate"
        if not replace_existing:
            return "rejected"

    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb", suffix=".xlsx.tmp", dir=history_directory, delete=False
        ) as temporary_file:
            temporary_file.write(content)
            temporary_file.flush()
            os.fsync(temporary_file.fileno())
            temporary_path = Path(temporary_file.name)
        os.replace(temporary_path, destination)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()
    return "saved"
