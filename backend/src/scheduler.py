"""Установка cron-расписания автоматической генерации отчётов."""
from __future__ import annotations

import os
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path


BACKEND_DIR = Path(__file__).resolve().parent.parent
PROJECT_DIR = BACKEND_DIR.parent
PYTHON_BIN = BACKEND_DIR / "venv" / "bin" / "python"
SRC_DIR = BACKEND_DIR / "src"
LOG_DIR = BACKEND_DIR / "logs"
MARKER = "# 1c_report generate_reports"


class ScheduleError(RuntimeError):
    """Ошибка настройки cron-расписания."""


@dataclass(frozen=True)
class CronEntry:
    name: str
    line: str


def _truthy(value: str, default: bool = False) -> bool:
    raw = value.strip().lower()
    if not raw:
        return default
    return raw in ("1", "true", "yes", "on")


def _env_file_values(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip()
    return values


def _get(values: dict[str, str], key: str, default: str = "") -> str:
    return values.get(key, os.getenv(key, default)).strip()


def _parse_time(value: str, field: str) -> tuple[int, int]:
    if not re.fullmatch(r"\d{1,2}:\d{2}", value):
        raise ScheduleError(f"{field}: время должно быть в формате HH:MM")
    hour_raw, minute_raw = value.split(":", 1)
    hour = int(hour_raw)
    minute = int(minute_raw)
    if hour > 23 or minute > 59:
        raise ScheduleError(f"{field}: время вне диапазона 00:00..23:59")
    return hour, minute


def _parse_weekday(value: str) -> int:
    try:
        day = int(value)
    except ValueError as exc:
        raise ScheduleError("SCHEDULE_WEEKLY_DAY должен быть числом 0..7") from exc
    if day < 0 or day > 7:
        raise ScheduleError("SCHEDULE_WEEKLY_DAY должен быть числом 0..7")
    return day


def build_entries(values: dict[str, str]) -> list[CronEntry]:
    if not _truthy(_get(values, "SCHEDULE_ENABLED", "0")):
        return []

    send_flags = []
    if _truthy(_get(values, "SCHEDULE_SEND_BITRIX", "1"), default=True):
        if not _get(values, "BITRIX_WEBHOOK_URL") or not _get(values, "BITRIX_CHAT_ID"):
            raise ScheduleError(
                "SCHEDULE_SEND_BITRIX=1, но BITRIX_WEBHOOK_URL или BITRIX_CHAT_ID не заданы"
            )
        send_flags.append("--send-bitrix")
    if _truthy(_get(values, "SCHEDULE_SEND_EMAIL", "0")):
        if not _get(values, "EMAIL_SMTP_HOST") or not _get(values, "EMAIL_FROM") or not _get(values, "EMAIL_TO"):
            raise ScheduleError(
                "SCHEDULE_SEND_EMAIL=1, но EMAIL_SMTP_HOST, EMAIL_FROM или EMAIL_TO не заданы"
            )
        send_flags.append("--send-email")
    send_args = " ".join(send_flags)

    entries: list[CronEntry] = []
    if _truthy(_get(values, "SCHEDULE_DAILY_ENABLED", "1"), default=True):
        hour, minute = _parse_time(_get(values, "SCHEDULE_DAILY_TIME", "11:00"), "SCHEDULE_DAILY_TIME")
        command = (
            f"cd {SRC_DIR} && {PYTHON_BIN} generate_reports.py --mode daily {send_args} "
            f">> {LOG_DIR}/daily.log 2>&1 {MARKER}"
        )
        entries.append(CronEntry("daily", f"{minute} {hour} * * * {command}"))

    if _truthy(_get(values, "SCHEDULE_WEEKLY_ENABLED", "1"), default=True):
        hour, minute = _parse_time(
            _get(values, "SCHEDULE_WEEKLY_TIME", "11:00"),
            "SCHEDULE_WEEKLY_TIME",
        )
        weekday = _parse_weekday(_get(values, "SCHEDULE_WEEKLY_DAY", "5"))
        command = (
            f"cd {SRC_DIR} && {PYTHON_BIN} generate_reports.py --mode weekly {send_args} "
            f">> {LOG_DIR}/weekly.log 2>&1 {MARKER}"
        )
        entries.append(CronEntry("weekly", f"{minute} {hour} * * {weekday} {command}"))

    return entries


def install_from_values(values: dict[str, str]) -> list[CronEntry]:
    if not PYTHON_BIN.exists():
        raise ScheduleError(f"Python venv не найден: {PYTHON_BIN}")
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    entries = build_entries(values)

    existing = subprocess.run(
        ["crontab", "-l"],
        text=True,
        capture_output=True,
        check=False,
    )
    if existing.returncode not in (0, 1):
        raise ScheduleError(existing.stderr.strip() or "не удалось прочитать crontab")

    kept = [line for line in existing.stdout.splitlines() if MARKER not in line]
    lines = kept + [entry.line for entry in entries]
    payload = "\n".join(lines).rstrip() + ("\n" if lines else "")

    updated = subprocess.run(
        ["crontab", "-"],
        input=payload,
        text=True,
        capture_output=True,
        check=False,
    )
    if updated.returncode != 0:
        raise ScheduleError(updated.stderr.strip() or "не удалось записать crontab")
    return entries


def install_from_env_file(path: Path = BACKEND_DIR / ".env") -> list[CronEntry]:
    return install_from_values(_env_file_values(path))


def main() -> int:
    try:
        entries = install_from_env_file()
    except ScheduleError as exc:
        print(f"Schedule error: {exc}", file=sys.stderr)
        return 2
    if entries:
        for entry in entries:
            print(f"[OK] cron {entry.name}: {entry.line}")
    else:
        print("[OK] cron disabled; project cron entries removed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
