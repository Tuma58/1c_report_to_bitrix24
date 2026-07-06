"""Загрузка конфигурации из .env. Без хардкода секретов."""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

# .env лежит в backend/ (на уровень выше src/)
_ENV_PATH = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(_ENV_PATH)


def _get_int(name: str, default: int) -> int:
    raw = os.getenv(name)
    if raw is None or raw.strip() == "":
        return default
    try:
        return int(raw)
    except ValueError:
        return default


@dataclass(frozen=True)
class Settings:
    odata_base_url: str
    odata_login: str
    odata_password: str
    timeout: int
    retries: int

    @classmethod
    def from_env(cls) -> "Settings":
        base_url = os.getenv("ODATA_BASE_URL", "").strip()
        if base_url and not base_url.endswith("/"):
            base_url += "/"
        return cls(
            odata_base_url=base_url,
            odata_login=os.getenv("ODATA_LOGIN", ""),
            odata_password=os.getenv("ODATA_PASSWORD", ""),
            timeout=_get_int("ODATA_TIMEOUT", 30),
            retries=_get_int("ODATA_RETRIES", 3),
        )


# Готовый объект настроек для импорта
settings = Settings.from_env()
