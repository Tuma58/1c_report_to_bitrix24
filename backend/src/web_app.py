"""Минимальный web GUI для настройки и запуска отчётов.

Работает без внешнего web-фреймворка: стандартный ThreadingHTTPServer.
По умолчанию слушает 127.0.0.1:8080. Для VPS можно задать WEB_UI_HOST,
WEB_UI_PORT и WEB_UI_PASSWORD в backend/.env.
"""
from __future__ import annotations

import base64
import argparse
import html
import os
import secrets
import sys
import tempfile
import time
from datetime import date, datetime, timedelta
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Optional
from urllib.parse import parse_qs, quote, urlencode, unquote, urlparse

try:
    from .bitrix_sender import BitrixError
    from .email_sender import EmailError
    from .generate_reports import (
        OUTPUT_DIR,
        GeneratedReport,
        generate_workbook,
        send_to_email,
        send_to_bitrix,
    )
    from .excel_reporter import ExcelReporter
    from .metrics import MetricsService
    from .odata_client import ODataError, ODataUnavailableError
    from .scheduler import ScheduleError, install_from_values as install_schedule
except ImportError:  # запуск как скрипта из backend/src
    from bitrix_sender import BitrixError
    from email_sender import EmailError
    from generate_reports import (
        OUTPUT_DIR,
        GeneratedReport,
        generate_workbook,
        send_to_email,
        send_to_bitrix,
    )
    from excel_reporter import ExcelReporter
    from metrics import MetricsService
    from odata_client import ODataError, ODataUnavailableError
    from scheduler import ScheduleError, install_from_values as install_schedule


APP_TITLE = "Отчёты АТЦ"
BACKEND_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = BACKEND_DIR / ".env"
ENV_EXAMPLE_PATH = BACKEND_DIR / ".env.example"
MODES = {
    "all": "Все листы",
    "daily": "День",
    "weekly": "Неделя",
}
SETTINGS_FIELDS = [
    {
        "key": "ODATA_BASE_URL",
        "label": "URL 1С OData",
        "section": "1С OData",
        "help": "Например: http://host/base/odata/standard.odata/",
    },
    {
        "key": "ODATA_LOGIN",
        "label": "Логин 1С",
        "section": "1С OData",
        "autocomplete": "username",
    },
    {
        "key": "ODATA_PASSWORD",
        "label": "Пароль 1С",
        "section": "1С OData",
        "secret": True,
        "autocomplete": "current-password",
    },
    {
        "key": "ODATA_TIMEOUT",
        "label": "Таймаут OData, сек.",
        "section": "1С OData",
        "kind": "int",
        "min": 1,
        "max": 600,
    },
    {
        "key": "ODATA_RETRIES",
        "label": "Повторы OData",
        "section": "1С OData",
        "kind": "int",
        "min": 1,
        "max": 20,
    },
    {
        "key": "BITRIX_WEBHOOK_URL",
        "label": "Webhook Bitrix24",
        "section": "Bitrix24",
        "secret": True,
        "help": "URL вида https://portal.bitrix24.ru/rest/<user>/<token>/",
        "autocomplete": "off",
    },
    {
        "key": "BITRIX_CHAT_ID",
        "label": "ID чата Bitrix24",
        "section": "Bitrix24",
        "help": "Например: chat123",
    },
    {
        "key": "BITRIX_DISK_FOLDER_ID",
        "label": "ID папки Диска Bitrix24",
        "section": "Bitrix24",
    },
    {
        "key": "EMAIL_SMTP_HOST",
        "label": "SMTP сервер",
        "section": "Email",
        "help": "Например: smtp.yandex.ru или smtp.gmail.com",
    },
    {
        "key": "EMAIL_SMTP_PORT",
        "label": "SMTP порт",
        "section": "Email",
        "kind": "int",
        "min": 1,
        "max": 65535,
    },
    {
        "key": "EMAIL_SMTP_LOGIN",
        "label": "SMTP логин",
        "section": "Email",
        "autocomplete": "username",
    },
    {
        "key": "EMAIL_SMTP_PASSWORD",
        "label": "SMTP пароль",
        "section": "Email",
        "secret": True,
        "autocomplete": "new-password",
    },
    {
        "key": "EMAIL_FROM",
        "label": "Email отправителя",
        "section": "Email",
    },
    {
        "key": "EMAIL_TO",
        "label": "Получатели",
        "section": "Email",
        "help": "Несколько адресов можно разделить запятыми или точками с запятой.",
    },
    {
        "key": "EMAIL_USE_TLS",
        "label": "Использовать STARTTLS",
        "section": "Email",
        "kind": "bool",
    },
    {
        "key": "EMAIL_USE_SSL",
        "label": "Использовать SMTP SSL",
        "section": "Email",
        "kind": "bool",
    },
    {
        "key": "SCHEDULE_ENABLED",
        "label": "Включить автоматическую генерацию",
        "section": "Расписание",
        "kind": "bool",
    },
    {
        "key": "SCHEDULE_DAILY_ENABLED",
        "label": "Дневной отчёт",
        "section": "Расписание",
        "kind": "bool",
    },
    {
        "key": "SCHEDULE_DAILY_TIME",
        "label": "Время дневного отчёта",
        "section": "Расписание",
        "kind": "time",
    },
    {
        "key": "SCHEDULE_WEEKLY_ENABLED",
        "label": "Недельный отчёт",
        "section": "Расписание",
        "kind": "bool",
    },
    {
        "key": "SCHEDULE_WEEKLY_DAY",
        "label": "День недельного отчёта",
        "section": "Расписание",
        "kind": "select",
        "options": [
            ("1", "Понедельник"),
            ("2", "Вторник"),
            ("3", "Среда"),
            ("4", "Четверг"),
            ("5", "Пятница"),
            ("6", "Суббота"),
            ("0", "Воскресенье"),
        ],
    },
    {
        "key": "SCHEDULE_WEEKLY_TIME",
        "label": "Время недельного отчёта",
        "section": "Расписание",
        "kind": "time",
    },
    {
        "key": "SCHEDULE_SEND_BITRIX",
        "label": "Рассылать в Bitrix24",
        "section": "Расписание",
        "kind": "bool",
    },
    {
        "key": "SCHEDULE_SEND_EMAIL",
        "label": "Рассылать по email",
        "section": "Расписание",
        "kind": "bool",
    },
    {
        "key": "WEB_UI_HOST",
        "label": "Адрес Web UI",
        "section": "Web UI",
        "help": "0.0.0.0 для доступа извне, 127.0.0.1 для reverse proxy.",
    },
    {
        "key": "WEB_UI_PORT",
        "label": "Порт Web UI",
        "section": "Web UI",
        "kind": "port",
        "help": "Порты 80 и 443 запрещены для этого приложения.",
    },
    {
        "key": "WEB_UI_USER",
        "label": "Логин Web UI",
        "section": "Web UI",
        "autocomplete": "username",
    },
    {
        "key": "WEB_UI_PASSWORD",
        "label": "Пароль Web UI",
        "section": "Web UI",
        "secret": True,
        "autocomplete": "new-password",
    },
]
SETTINGS_KEYS = [field["key"] for field in SETTINGS_FIELDS]


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name, "").strip()
    if not raw:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def _parse_report_date(raw: str) -> date:
    return datetime.strptime(raw, "%Y-%m-%d").date()


def _default_date() -> date:
    return date.today() - timedelta(days=1)


def _file_url(path: Path) -> str:
    return f"/download?file={quote(path.name)}"


def _recent_reports(limit: int = 10) -> list[Path]:
    if not OUTPUT_DIR.exists():
        return []
    files = [p for p in OUTPUT_DIR.glob("*.xlsx") if p.is_file()]
    return sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)[:limit]


def _fmt_mtime(path: Path) -> str:
    return datetime.fromtimestamp(path.stat().st_mtime).strftime("%d.%m.%Y %H:%M")


def _read_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value
    return values


def _current_settings() -> dict[str, str]:
    values = _read_env_file(ENV_EXAMPLE_PATH)
    values.update(_read_env_file(ENV_PATH))
    for key in SETTINGS_KEYS:
        if key not in values:
            values[key] = os.getenv(key, "")
    return values


def _setting_status(value: str) -> str:
    return "задано" if value else "не задано"


def _write_env_values(values: dict[str, str]) -> None:
    source = ENV_PATH if ENV_PATH.exists() else ENV_EXAMPLE_PATH
    lines = source.read_text(encoding="utf-8").splitlines() if source.exists() else []
    updated: list[str] = []
    seen: set[str] = set()

    for line in lines:
        stripped = line.strip()
        if stripped and not stripped.startswith("#") and "=" in line:
            key, _value = line.split("=", 1)
            key = key.strip()
            if key in values:
                updated.append(f"{key}={values[key]}")
                seen.add(key)
                continue
        updated.append(line)

    missing = [key for key in SETTINGS_KEYS if key not in seen]
    if missing:
        if updated and updated[-1].strip():
            updated.append("")
        for key in missing:
            updated.append(f"{key}={values.get(key, '')}")

    ENV_PATH.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix=".env.", dir=str(ENV_PATH.parent), text=True)
    tmp_path = Path(tmp_name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write("\n".join(updated).rstrip() + "\n")
        os.chmod(tmp_path, 0o600)
        tmp_path.replace(ENV_PATH)
        os.chmod(ENV_PATH, 0o600)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def _reload_runtime_settings(values: dict[str, str]) -> None:
    for key in SETTINGS_KEYS:
        os.environ[key] = values.get(key, "")

    for module_name in ("config", "odata_client", "bitrix_sender", "email_sender"):
        module = sys.modules.get(module_name) or sys.modules.get(f"{__package__}.{module_name}")
        if module is None:
            continue
        if module_name == "config" and hasattr(module, "Settings"):
            module.settings = module.Settings.from_env()
        elif hasattr(module, "default_settings"):
            config_module = sys.modules.get("config") or sys.modules.get(f"{__package__}.config")
            if config_module is not None and hasattr(config_module, "settings"):
                module.default_settings = config_module.settings


def _validated_settings(form: dict[str, list[str]]) -> tuple[dict[str, str], Optional[str]]:
    current = _current_settings()
    values = dict(current)

    for field in SETTINGS_FIELDS:
        key = field["key"]
        kind = field.get("kind")
        if kind == "bool":
            values[key] = "1" if form.get(key, [""])[0] == "1" else "0"
            continue

        raw = form.get(key, [""])[0]
        if "\n" in raw or "\r" in raw:
            return values, f"{field['label']}: переносы строк недопустимы."
        raw = raw.strip()

        if field.get("secret"):
            if form.get(f"{key}__clear", [""])[0] == "1":
                values[key] = ""
            elif raw:
                values[key] = raw
            else:
                values[key] = current.get(key, "")
        else:
            values[key] = raw

    for field in SETTINGS_FIELDS:
        key = field["key"]
        value = values.get(key, "")
        kind = field.get("kind")
        if kind in ("int", "port"):
            try:
                numeric = int(value)
            except ValueError:
                return values, f"{field['label']}: укажите целое число."
            min_value = int(field.get("min", 1))
            max_value = int(field.get("max", 65535))
            if numeric < min_value or numeric > max_value:
                return values, f"{field['label']}: допустимо от {min_value} до {max_value}."
            if kind == "port" and numeric in (80, 443):
                return values, "Web UI не должен занимать порты 80 и 443."
            values[key] = str(numeric)
        elif kind == "time":
            if not value:
                return values, f"{field['label']}: укажите время."
            parts = value.split(":", 1)
            if len(parts) != 2 or not all(part.isdigit() for part in parts):
                return values, f"{field['label']}: время должно быть в формате HH:MM."
            hour, minute = int(parts[0]), int(parts[1])
            if hour > 23 or minute > 59:
                return values, f"{field['label']}: время вне диапазона 00:00..23:59."
            values[key] = f"{hour:02d}:{minute:02d}"
        elif kind == "select":
            allowed = {option[0] for option in field.get("options", [])}
            if value not in allowed:
                return values, f"{field['label']}: выберите значение из списка."

    return values, None


def _settings_form(values: dict[str, str]) -> str:
    sections: list[str] = []
    for section in dict.fromkeys(field["section"] for field in SETTINGS_FIELDS):
        controls = []
        for field in SETTINGS_FIELDS:
            if field["section"] != section:
                continue
            key = field["key"]
            value = values.get(key, "")
            label = html.escape(field["label"])
            help_text = field.get("help", "")
            autocomplete = html.escape(field.get("autocomplete", "off"))
            kind = field.get("kind")
            if kind == "bool":
                checked = " checked" if value.strip().lower() in ("1", "true", "yes", "on") else ""
                controls.append(
                    f"""
                    <label class="check">
                      <input type="checkbox" name="{key}" value="1"{checked}>
                      <span>{label}</span>
                    </label>
                    {f'<p class="hint">{html.escape(help_text)}</p>' if help_text else ''}
                    """
                )
            elif kind == "select":
                options = []
                for option_value, option_label in field.get("options", []):
                    selected = " selected" if option_value == value else ""
                    options.append(
                        f'<option value="{html.escape(option_value, quote=True)}"{selected}>'
                        f'{html.escape(option_label)}</option>'
                    )
                controls.append(
                    f"""
                    <div class="field">
                      <label class="title" for="{key}">{label}</label>
                      <select id="{key}" name="{key}">{''.join(options)}</select>
                      {f'<p class="hint">{html.escape(help_text)}</p>' if help_text else ''}
                    </div>
                    """
                )
            elif field.get("secret"):
                status = html.escape(_setting_status(value))
                controls.append(
                    f"""
                    <div class="field">
                      <div class="label-row">
                        <label class="title" for="{key}">{label}</label>
                        <span class="secret-state">{status}</span>
                      </div>
                      <input id="{key}" name="{key}" type="password"
                             autocomplete="{autocomplete}"
                             placeholder="Оставьте пустым, чтобы не менять">
                      <label class="check small">
                        <input type="checkbox" name="{key}__clear" value="1">
                        <span>Очистить значение</span>
                      </label>
                      {f'<p class="hint">{html.escape(help_text)}</p>' if help_text else ''}
                    </div>
                    """
                )
            else:
                input_type = "number" if kind in ("int", "port") else "time" if kind == "time" else "text"
                extra = ""
                if kind in ("int", "port"):
                    extra = (
                        f' min="{int(field.get("min", 1))}"'
                        f' max="{int(field.get("max", 65535))}"'
                        ' step="1"'
                    )
                controls.append(
                    f"""
                    <div class="field">
                      <label class="title" for="{key}">{label}</label>
                      <input id="{key}" name="{key}" type="{input_type}"
                             value="{html.escape(value, quote=True)}"
                             autocomplete="{autocomplete}"{extra}>
                      {f'<p class="hint">{html.escape(help_text)}</p>' if help_text else ''}
                    </div>
                    """
                )
        sections.append(
            f"""
            <fieldset>
              <legend>{html.escape(section)}</legend>
              {''.join(controls)}
            </fieldset>
            """
        )

    return f"""
    <section class="panel">
      <h2>Секреты и подключения</h2>
      <form method="post" action="/settings" autocomplete="off">
        {''.join(sections)}
        <div class="actions">
          <button type="submit">Сохранить настройки</button>
          <a class="button secondary" href="/">Вернуться к отчёту</a>
        </div>
      </form>
    </section>
    """


def _page(
    *,
    selected_mode: str = "all",
    selected_date: Optional[date] = None,
    send_bitrix: bool = False,
    send_email: bool = False,
    message: str = "",
    error: str = "",
    generated: Optional[list[GeneratedReport]] = None,
) -> str:
    selected_date = selected_date or _default_date()
    mode_options = []
    for value, label in MODES.items():
        checked = " checked" if value == selected_mode else ""
        mode_options.append(
            f"""
            <label class="seg-item">
              <input type="radio" name="mode" value="{value}"{checked}>
              <span>{html.escape(label)}</span>
            </label>
            """
        )

    generated_html = ""
    if generated:
        items = []
        for report in generated:
            name = html.escape(report.path.name)
            items.append(
                f"""
                <li>
                  <span>{html.escape(report.kind)}</span>
                  <a href="{_file_url(report.path)}">{name}</a>
                </li>
                """
            )
        generated_html = f"""
        <section class="panel">
          <h2>Готовый файл</h2>
          <ul class="file-list">{''.join(items)}</ul>
        </section>
        """

    recent_items = []
    for path in _recent_reports():
        recent_items.append(
            f"""
            <li>
              <span>{html.escape(_fmt_mtime(path))}</span>
              <a href="{_file_url(path)}">{html.escape(path.name)}</a>
            </li>
            """
        )
    recent_html = "".join(recent_items) or "<li><span></span><em>Нет файлов</em></li>"

    bitrix_checked = " checked" if send_bitrix else ""
    email_checked = " checked" if send_email else ""
    message_html = f'<div class="notice ok">{html.escape(message)}</div>' if message else ""
    error_html = f'<div class="notice err">{html.escape(error)}</div>' if error else ""

    return f"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{APP_TITLE}</title>
  <style>
    :root {{
      --bg: #f6f7f8;
      --panel: #ffffff;
      --text: #202428;
      --muted: #687079;
      --line: #d9dee3;
      --accent: #167c80;
      --accent-dark: #0f5e61;
      --danger: #b42318;
      --ok: #1f7a4d;
      --shadow: 0 10px 28px rgba(32, 36, 40, .08);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      min-height: 100vh;
      background: var(--bg);
      color: var(--text);
      font: 15px/1.45 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }}
    header {{
      border-bottom: 1px solid var(--line);
      background: #ffffff;
    }}
    .wrap {{
      width: min(1040px, calc(100% - 32px));
      margin: 0 auto;
    }}
    .topbar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      min-height: 64px;
      gap: 16px;
    }}
    h1 {{
      margin: 0;
      font-size: 22px;
      font-weight: 650;
      letter-spacing: 0;
    }}
    .status {{
      color: var(--muted);
      font-size: 13px;
      white-space: nowrap;
    }}
    main {{
      padding: 28px 0 40px;
    }}
    .grid {{
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(320px, .8fr);
      gap: 20px;
      align-items: start;
    }}
    .panel {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 20px;
    }}
    h2 {{
      margin: 0 0 16px;
      font-size: 16px;
      font-weight: 650;
      letter-spacing: 0;
    }}
    form {{
      display: grid;
      gap: 18px;
    }}
    .field {{
      display: grid;
      gap: 8px;
    }}
    label.title {{
      font-weight: 600;
      font-size: 13px;
      color: #30363c;
    }}
    input[type="date"], input[type="text"], input[type="password"], input[type="number"], input[type="time"], select {{
      width: 100%;
      min-height: 42px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 8px 10px;
      font: inherit;
      color: var(--text);
      background: #fff;
    }}
    input::placeholder {{ color: #9aa2aa; }}
    .label-row {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
    }}
    .secret-state {{
      color: var(--muted);
      font-size: 12px;
      white-space: nowrap;
    }}
    .hint {{
      margin: -2px 0 0;
      color: var(--muted);
      font-size: 12px;
    }}
    fieldset {{
      display: grid;
      gap: 14px;
      margin: 0;
      padding: 0 0 18px;
      border: 0;
      border-bottom: 1px solid var(--line);
    }}
    fieldset:last-of-type {{
      border-bottom: 0;
      padding-bottom: 0;
    }}
    legend {{
      padding: 0;
      margin: 0 0 2px;
      font-size: 14px;
      font-weight: 700;
    }}
    .seg {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      background: #eef1f3;
    }}
    .seg-item {{
      min-width: 0;
      cursor: pointer;
    }}
    .seg-item input {{
      position: absolute;
      opacity: 0;
      pointer-events: none;
    }}
    .seg-item span {{
      display: grid;
      place-items: center;
      min-height: 42px;
      padding: 8px 10px;
      color: var(--muted);
      border-right: 1px solid var(--line);
      text-align: center;
      white-space: nowrap;
    }}
    .seg-item:last-child span {{ border-right: 0; }}
    .seg-item input:checked + span {{
      background: #ffffff;
      color: var(--text);
      font-weight: 650;
    }}
    .check {{
      display: flex;
      align-items: center;
      gap: 10px;
      color: var(--text);
      font-weight: 500;
    }}
    .check input {{
      width: 18px;
      height: 18px;
      accent-color: var(--accent);
    }}
    .check.small {{
      gap: 8px;
      color: var(--muted);
      font-size: 13px;
      font-weight: 400;
    }}
    .check.small input {{
      width: 16px;
      height: 16px;
    }}
    .actions {{
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
      align-items: center;
    }}
    button, .button {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 42px;
      padding: 9px 16px;
      border-radius: 6px;
      border: 1px solid var(--accent);
      background: var(--accent);
      color: #fff;
      font: inherit;
      font-weight: 650;
      text-decoration: none;
      cursor: pointer;
    }}
    button:hover, .button:hover {{ background: var(--accent-dark); }}
    .button.secondary {{
      background: #fff;
      color: var(--accent-dark);
      border-color: var(--line);
    }}
    .button.secondary:hover {{
      background: #eef3f3;
      border-color: #b8c5c6;
    }}
    nav {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      align-items: center;
      justify-content: flex-end;
    }}
    nav a {{
      display: inline-flex;
      align-items: center;
      min-height: 34px;
      padding: 6px 10px;
      border-radius: 6px;
      color: var(--accent-dark);
      text-decoration: none;
      font-weight: 600;
    }}
    nav a:hover {{ background: #eef3f3; }}
    .notice {{
      border-radius: 8px;
      padding: 12px 14px;
      margin-bottom: 16px;
      border: 1px solid;
      background: #fff;
    }}
    .notice.ok {{
      color: var(--ok);
      border-color: rgba(31, 122, 77, .35);
    }}
    .notice.err {{
      color: var(--danger);
      border-color: rgba(180, 35, 24, .35);
    }}
    .file-list {{
      list-style: none;
      padding: 0;
      margin: 0;
      display: grid;
      gap: 10px;
    }}
    .file-list li {{
      display: grid;
      grid-template-columns: 140px minmax(0, 1fr);
      gap: 12px;
      align-items: baseline;
      border-bottom: 1px solid var(--line);
      padding-bottom: 10px;
    }}
    .file-list li:last-child {{
      border-bottom: 0;
      padding-bottom: 0;
    }}
    .file-list span, .file-list em {{
      color: var(--muted);
      font-size: 13px;
      font-style: normal;
    }}
    a {{
      color: var(--accent-dark);
      overflow-wrap: anywhere;
      text-decoration-thickness: 1px;
      text-underline-offset: 3px;
    }}
    @media (max-width: 760px) {{
      .grid {{ grid-template-columns: 1fr; }}
      .topbar {{ align-items: flex-start; flex-direction: column; padding: 14px 0; }}
      .status {{ white-space: normal; }}
      .file-list li {{ grid-template-columns: 1fr; gap: 4px; }}
      .seg {{ grid-template-columns: 1fr; }}
      .seg-item span {{ border-right: 0; border-bottom: 1px solid var(--line); }}
      .seg-item:last-child span {{ border-bottom: 0; }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="wrap topbar">
      <h1>{APP_TITLE}</h1>
      <nav aria-label="Разделы">
        <a href="/">Отчёт</a>
        <a href="/settings">Настройки</a>
      </nav>
    </div>
  </header>
  <main class="wrap">
    {message_html}
    {error_html}
    <div class="grid">
      <section class="panel">
        <h2>Настройка отчёта</h2>
        <form method="post" action="/generate">
          <div class="field">
            <label class="title">Режим</label>
            <div class="seg">{''.join(mode_options)}</div>
          </div>
          <div class="field">
            <label class="title" for="date">Дата</label>
            <input id="date" name="date" type="date" value="{selected_date.isoformat()}" required>
          </div>
          <label class="check">
            <input type="checkbox" name="send_bitrix" value="1"{bitrix_checked}>
            <span>Отправить в Bitrix24</span>
          </label>
          <label class="check">
            <input type="checkbox" name="send_email" value="1"{email_checked}>
            <span>Отправить по email</span>
          </label>
          <div class="actions">
            <button type="submit">Сформировать</button>
          </div>
        </form>
      </section>
      <section class="panel">
        <h2>Последние файлы</h2>
        <ul class="file-list">{recent_html}</ul>
      </section>
      {generated_html}
    </div>
  </main>
</body>
</html>"""


def _settings_page(message: str = "", error: str = "") -> str:
    message_html = f'<div class="notice ok">{html.escape(message)}</div>' if message else ""
    error_html = f'<div class="notice err">{html.escape(error)}</div>' if error else ""
    return f"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{APP_TITLE} · Настройки</title>
  <style>
    :root {{
      --bg: #f6f7f8;
      --panel: #ffffff;
      --text: #202428;
      --muted: #687079;
      --line: #d9dee3;
      --accent: #167c80;
      --accent-dark: #0f5e61;
      --danger: #b42318;
      --ok: #1f7a4d;
      --shadow: 0 10px 28px rgba(32, 36, 40, .08);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      min-height: 100vh;
      background: var(--bg);
      color: var(--text);
      font: 15px/1.45 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }}
    header {{
      border-bottom: 1px solid var(--line);
      background: #ffffff;
    }}
    .wrap {{
      width: min(1040px, calc(100% - 32px));
      margin: 0 auto;
    }}
    .topbar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      min-height: 64px;
      gap: 16px;
    }}
    h1 {{
      margin: 0;
      font-size: 22px;
      font-weight: 650;
      letter-spacing: 0;
    }}
    main {{
      padding: 28px 0 40px;
    }}
    .panel {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 20px;
    }}
    h2 {{
      margin: 0 0 16px;
      font-size: 16px;
      font-weight: 650;
      letter-spacing: 0;
    }}
    form {{
      display: grid;
      gap: 18px;
    }}
    .field {{
      display: grid;
      gap: 8px;
    }}
    label.title {{
      font-weight: 600;
      font-size: 13px;
      color: #30363c;
    }}
    input[type="date"], input[type="text"], input[type="password"], input[type="number"], input[type="time"], select {{
      width: 100%;
      min-height: 42px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 8px 10px;
      font: inherit;
      color: var(--text);
      background: #fff;
    }}
    input::placeholder {{ color: #9aa2aa; }}
    .label-row {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
    }}
    .secret-state {{
      color: var(--muted);
      font-size: 12px;
      white-space: nowrap;
    }}
    .hint {{
      margin: -2px 0 0;
      color: var(--muted);
      font-size: 12px;
    }}
    fieldset {{
      display: grid;
      gap: 14px;
      margin: 0;
      padding: 0 0 18px;
      border: 0;
      border-bottom: 1px solid var(--line);
    }}
    fieldset:last-of-type {{
      border-bottom: 0;
      padding-bottom: 0;
    }}
    legend {{
      padding: 0;
      margin: 0 0 2px;
      font-size: 14px;
      font-weight: 700;
    }}
    .check {{
      display: flex;
      align-items: center;
      gap: 10px;
      color: var(--text);
      font-weight: 500;
    }}
    .check input {{
      width: 18px;
      height: 18px;
      accent-color: var(--accent);
    }}
    .check.small {{
      gap: 8px;
      color: var(--muted);
      font-size: 13px;
      font-weight: 400;
    }}
    .check.small input {{
      width: 16px;
      height: 16px;
    }}
    .actions {{
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
      align-items: center;
    }}
    button, .button {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 42px;
      padding: 9px 16px;
      border-radius: 6px;
      border: 1px solid var(--accent);
      background: var(--accent);
      color: #fff;
      font: inherit;
      font-weight: 650;
      text-decoration: none;
      cursor: pointer;
    }}
    button:hover, .button:hover {{ background: var(--accent-dark); }}
    .button.secondary {{
      background: #fff;
      color: var(--accent-dark);
      border-color: var(--line);
    }}
    .button.secondary:hover {{
      background: #eef3f3;
      border-color: #b8c5c6;
    }}
    nav {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      align-items: center;
      justify-content: flex-end;
    }}
    nav a {{
      display: inline-flex;
      align-items: center;
      min-height: 34px;
      padding: 6px 10px;
      border-radius: 6px;
      color: var(--accent-dark);
      text-decoration: none;
      font-weight: 600;
    }}
    nav a:hover {{ background: #eef3f3; }}
    .notice {{
      border-radius: 8px;
      padding: 12px 14px;
      margin-bottom: 16px;
      border: 1px solid;
      background: #fff;
    }}
    .notice.ok {{
      color: var(--ok);
      border-color: rgba(31, 122, 77, .35);
    }}
    .notice.err {{
      color: var(--danger);
      border-color: rgba(180, 35, 24, .35);
    }}
    @media (max-width: 760px) {{
      .topbar {{ align-items: flex-start; flex-direction: column; padding: 14px 0; }}
      nav {{ justify-content: flex-start; }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="wrap topbar">
      <h1>{APP_TITLE}</h1>
      <nav aria-label="Разделы">
        <a href="/">Отчёт</a>
        <a href="/settings">Настройки</a>
      </nav>
    </div>
  </header>
  <main class="wrap">
    {message_html}
    {error_html}
    {_settings_form(_current_settings())}
  </main>
</body>
</html>"""


def _apply_schedule(values: dict[str, str]) -> str:
    try:
        entries = install_schedule(values)
    except (ScheduleError, OSError) as exc:
        return f"Расписание не обновлено: {exc}"
    if entries:
        return f"Расписание обновлено: {len(entries)} задач."
    return "Расписание выключено, задачи cron удалены."


def _reports_from_query(params: dict[str, list[str]]) -> list[GeneratedReport]:
    reports: list[GeneratedReport] = []
    for filename in params.get("file", []):
        if not filename or "/" in filename or "\\" in filename or not filename.endswith(".xlsx"):
            continue
        path = (OUTPUT_DIR / filename).resolve()
        if OUTPUT_DIR.resolve() in path.parents and path.exists():
            reports.append(GeneratedReport("Файл", filename, path))
    return reports


class ReportWebHandler(BaseHTTPRequestHandler):
    server_version = "ATCReportWeb/1.0"

    def _auth_required(self) -> bool:
        return bool(os.getenv("WEB_UI_PASSWORD", "").strip())

    def _is_authorized(self) -> bool:
        password = os.getenv("WEB_UI_PASSWORD", "").strip()
        if not password:
            return True
        expected_user = os.getenv("WEB_UI_USER", "admin").strip() or "admin"
        header = self.headers.get("Authorization", "")
        if not header.startswith("Basic "):
            return False
        try:
            raw = base64.b64decode(header[6:]).decode("utf-8")
        except (ValueError, UnicodeDecodeError):
            return False
        user, sep, supplied_password = raw.partition(":")
        if not sep:
            return False
        return secrets.compare_digest(user, expected_user) and secrets.compare_digest(
            supplied_password, password
        )

    def _require_auth(self) -> bool:
        if self._is_authorized():
            return False
        self.send_response(HTTPStatus.UNAUTHORIZED)
        self.send_header("WWW-Authenticate", 'Basic realm="ATC Reports"')
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write("Требуется авторизация".encode("utf-8"))
        return True

    def _send_html(self, content: str, status: HTTPStatus = HTTPStatus.OK) -> None:
        data = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _redirect(self, location: str) -> None:
        self.send_response(HTTPStatus.SEE_OTHER)
        self.send_header("Location", location)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def _report_location(
        self,
        *,
        mode: str = "all",
        report_date: Optional[date] = None,
        send_bitrix: bool = False,
        send_email: bool = False,
        message: str = "",
        error: str = "",
        generated: Optional[list[GeneratedReport]] = None,
    ) -> str:
        query: dict[str, object] = {
            "mode": mode,
            "send_bitrix": "1" if send_bitrix else "0",
            "send_email": "1" if send_email else "0",
        }
        if report_date is not None:
            query["date"] = report_date.isoformat()
        if message:
            query["message"] = message
        if error:
            query["error"] = error
        if generated:
            query["file"] = [report.path.name for report in generated]
        return "/?" + urlencode(query, doseq=True)

    def _settings_location(self, *, message: str = "", error: str = "") -> str:
        query = {}
        if message:
            query["message"] = message
        if error:
            query["error"] = error
        return "/settings" + (("?" + urlencode(query)) if query else "")

    def _send_error_page(
        self,
        error: str,
        *,
        selected_mode: str = "all",
        selected_date: Optional[date] = None,
        send_bitrix: bool = False,
    ) -> None:
        self._send_html(
            _page(
                selected_mode=selected_mode,
                selected_date=selected_date,
                send_bitrix=send_bitrix,
                send_email=False,
                error=error,
            ),
            HTTPStatus.BAD_REQUEST,
        )

    def do_GET(self) -> None:  # noqa: N802
        if self._require_auth():
            return
        parsed = urlparse(self.path)
        if parsed.path == "/":
            params = parse_qs(parsed.query)
            mode = params.get("mode", ["all"])[0]
            if mode not in MODES:
                mode = "all"
            try:
                selected_date = _parse_report_date(params.get("date", [""])[0])
            except ValueError:
                selected_date = _default_date()
            self._send_html(
                _page(
                    selected_mode=mode,
                    selected_date=selected_date,
                    send_bitrix=params.get("send_bitrix", [""])[0] == "1",
                    send_email=params.get("send_email", [""])[0] == "1",
                    message=params.get("message", [""])[0],
                    error=params.get("error", [""])[0],
                    generated=_reports_from_query(params),
                )
            )
            return
        if parsed.path == "/settings":
            params = parse_qs(parsed.query)
            self._send_html(
                _settings_page(
                    message=params.get("message", [""])[0],
                    error=params.get("error", [""])[0],
                )
            )
            return
        if parsed.path == "/download":
            self._download(parsed.query)
            return
        if parsed.path == "/health":
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write("ok".encode("utf-8"))
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_HEAD(self) -> None:  # noqa: N802
        if self._require_auth():
            return
        parsed = urlparse(self.path)
        if parsed.path == "/":
            data = _page().encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            return
        if parsed.path == "/settings":
            data = _settings_page().encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            return
        if parsed.path == "/health":
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", "2")
            self.end_headers()
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:  # noqa: N802
        if self._require_auth():
            return
        parsed = urlparse(self.path)
        if parsed.path == "/settings":
            self._save_settings()
            return
        if parsed.path == "/generate":
            self._generate_report()
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def _save_settings(self) -> None:
        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length).decode("utf-8")
        form = parse_qs(raw)
        values, error = _validated_settings(form)
        if error:
            self._redirect(self._settings_location(error=error))
            return
        try:
            _write_env_values(values)
            _reload_runtime_settings(values)
        except OSError as exc:
            self._redirect(self._settings_location(error=f"Не удалось сохранить backend/.env: {exc}"))
            return
        schedule_message = _apply_schedule(values)
        self._redirect(self._settings_location(message=f"Настройки сохранены. {schedule_message}"))

    def _generate_report(self) -> None:
        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length).decode("utf-8")
        form = parse_qs(raw)
        mode = form.get("mode", ["all"])[0]
        date_raw = form.get("date", [""])[0]
        send_bitrix = form.get("send_bitrix", [""])[0] == "1"
        send_email = form.get("send_email", [""])[0] == "1"

        if mode not in MODES:
            self._redirect(
                self._report_location(
                    mode="all",
                    send_bitrix=send_bitrix,
                    send_email=send_email,
                    error="Неизвестный режим отчёта.",
                )
            )
            return
        try:
            report_date = _parse_report_date(date_raw)
        except ValueError:
            self._redirect(
                self._report_location(
                    mode=mode,
                    send_bitrix=send_bitrix,
                    send_email=send_email,
                    error="Дата должна быть в формате YYYY-MM-DD.",
                )
            )
            return

        started = time.monotonic()
        try:
            service = MetricsService()
            reporter = ExcelReporter()
            daily_day = report_date
            weekly_day = report_date
            if mode == "weekly":
                weekly_day = report_date
            generated = generate_workbook(service, reporter, mode, daily_day, weekly_day)
            if send_bitrix:
                send_to_bitrix(generated, "Отчёты АТЦ: Excel-файл с листами отчётов")
            if send_email:
                send_to_email(generated, "Отчёты АТЦ: Excel-файл с листами отчётов")
        except ODataUnavailableError as exc:
            self._redirect(
                self._report_location(
                    mode=mode,
                    report_date=report_date,
                    send_bitrix=send_bitrix,
                    send_email=send_email,
                    error=f"OData недоступен: {exc}",
                )
            )
            return
        except (ODataError, BitrixError, EmailError, RuntimeError) as exc:
            self._redirect(
                self._report_location(
                    mode=mode,
                    report_date=report_date,
                    send_bitrix=send_bitrix,
                    send_email=send_email,
                    error=str(exc),
                )
            )
            return

        elapsed = time.monotonic() - started
        channels = []
        if send_bitrix:
            channels.append("Bitrix24")
        if send_email:
            channels.append("email")
        suffix = f" Отправлено: {', '.join(channels)}." if channels else ""
        self._redirect(
            self._report_location(
                mode=mode,
                report_date=report_date,
                send_bitrix=send_bitrix,
                send_email=send_email,
                message=f"Отчёт сформирован за {elapsed:.1f} сек.{suffix}",
                generated=generated,
            )
        )

    def _download(self, query: str) -> None:
        params = parse_qs(query)
        filename = unquote(params.get("file", [""])[0])
        if not filename or "/" in filename or "\\" in filename or not filename.endswith(".xlsx"):
            self.send_error(HTTPStatus.BAD_REQUEST)
            return
        path = (OUTPUT_DIR / filename).resolve()
        if OUTPUT_DIR.resolve() not in path.parents or not path.exists():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Length", str(len(data)))
        self.send_header(
            "Content-Disposition",
            f"attachment; filename*=UTF-8''{quote(path.name)}",
        )
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt: str, *args) -> None:
        sys.stderr.write("[%s] %s\n" % (self.log_date_time_string(), fmt % args))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Web UI настройки отчётов АТЦ")
    parser.add_argument(
        "--host",
        default=os.getenv("WEB_UI_HOST", "127.0.0.1").strip() or "127.0.0.1",
        help="адрес для web-интерфейса",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=_env_int("WEB_UI_PORT", 8080),
        help="порт для web-интерфейса",
    )
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    host = args.host
    port = args.port
    if port in (80, 443):
        print("WEB UI must not use ports 80 or 443. Set WEB_UI_PORT to 8080 or another high port.", file=sys.stderr)
        return 2
    password_set = bool(os.getenv("WEB_UI_PASSWORD", "").strip())
    if host not in ("127.0.0.1", "localhost") and not password_set:
        print("Warning: WEB_UI_PASSWORD is empty while WEB_UI_HOST is not local.", file=sys.stderr)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((host, port), ReportWebHandler)
    print(f"Web UI: http://{host}:{port}/")
    if password_set:
        print(f"Basic Auth user: {os.getenv('WEB_UI_USER', 'admin') or 'admin'}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
