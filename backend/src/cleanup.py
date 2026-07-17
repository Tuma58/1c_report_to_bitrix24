"""Очистка сгенерированных Excel-файлов."""
from __future__ import annotations

import os
import time
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv


BACKEND_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = BACKEND_DIR / ".env"
OUTPUT_DIR = BACKEND_DIR / "output"
load_dotenv(ENV_PATH)


def truthy_value(value: Optional[str], default: bool = False) -> bool:
    if value is None or value == "":
        return default
    return value.strip().lower() in ("1", "true", "yes", "on")


def _value(values: Optional[dict[str, str]], key: str, default: str) -> str:
    if values is not None:
        return values.get(key, default)
    return os.getenv(key, default)


def cleanup_output_files(
    values: Optional[dict[str, str]] = None,
    output_dir: Path = OUTPUT_DIR,
) -> int:
    if not truthy_value(_value(values, "FILE_CLEANUP_ENABLED", "0")):
        return 0
    if not output_dir.exists():
        return 0

    try:
        keep_days = max(1, int(_value(values, "FILE_CLEANUP_DAYS", "30") or "30"))
    except ValueError:
        keep_days = 30
    try:
        max_files = max(1, int(_value(values, "FILE_CLEANUP_MAX_FILES", "100") or "100"))
    except ValueError:
        max_files = 100

    files = sorted(
        [p for p in output_dir.glob("*.xlsx") if p.is_file()],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    cutoff = time.time() - keep_days * 86400
    to_delete: set[Path] = {p for p in files if p.stat().st_mtime < cutoff}
    to_delete.update(files[max_files:])

    deleted = 0
    for path in to_delete:
        try:
            path.unlink()
            deleted += 1
        except OSError:
            continue
    return deleted
