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
    bitrix_webhook_url: str
    bitrix_chat_id: str
    email_smtp_host: str
    email_smtp_port: int
    email_smtp_login: str
    email_smtp_password: str
    email_from: str
    email_to: str
    email_use_tls: bool
    email_use_ssl: bool

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
            bitrix_webhook_url=os.getenv("BITRIX_WEBHOOK_URL", "").strip(),
            bitrix_chat_id=os.getenv("BITRIX_CHAT_ID", "").strip(),
            email_smtp_host=os.getenv("EMAIL_SMTP_HOST", "").strip(),
            email_smtp_port=_get_int("EMAIL_SMTP_PORT", 587),
            email_smtp_login=os.getenv("EMAIL_SMTP_LOGIN", "").strip(),
            email_smtp_password=os.getenv("EMAIL_SMTP_PASSWORD", ""),
            email_from=os.getenv("EMAIL_FROM", "").strip(),
            email_to=os.getenv("EMAIL_TO", "").strip(),
            email_use_tls=os.getenv("EMAIL_USE_TLS", "1").strip().lower() in ("1", "true", "yes", "on"),
            email_use_ssl=os.getenv("EMAIL_USE_SSL", "0").strip().lower() in ("1", "true", "yes", "on"),
        )


# Готовый объект настроек для импорта
settings = Settings.from_env()
