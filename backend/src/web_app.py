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
import time
from datetime import date, datetime, timedelta
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Optional
from urllib.parse import parse_qs, quote, unquote, urlparse

try:
    from .bitrix_sender import BitrixError
    from .generate_reports import (
        OUTPUT_DIR,
        GeneratedReport,
        generate_workbook,
        send_to_bitrix,
    )
    from .excel_reporter import ExcelReporter
    from .metrics import MetricsService
    from .odata_client import ODataError, ODataUnavailableError
except ImportError:  # запуск как скрипта из backend/src
    from bitrix_sender import BitrixError
    from generate_reports import (
        OUTPUT_DIR,
        GeneratedReport,
        generate_workbook,
        send_to_bitrix,
    )
    from excel_reporter import ExcelReporter
    from metrics import MetricsService
    from odata_client import ODataError, ODataUnavailableError


APP_TITLE = "Отчёты АТЦ"
MODES = {
    "all": "Все листы",
    "daily": "День",
    "weekly": "Неделя",
}


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


def _page(
    *,
    selected_mode: str = "all",
    selected_date: Optional[date] = None,
    send_bitrix: bool = False,
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
    input[type="date"] {{
      width: 100%;
      min-height: 42px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 8px 10px;
      font: inherit;
      color: var(--text);
      background: #fff;
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
      <div class="status">1С OData · Excel · Bitrix24</div>
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
                error=error,
            ),
            HTTPStatus.BAD_REQUEST,
        )

    def do_GET(self) -> None:  # noqa: N802
        if self._require_auth():
            return
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._send_html(_page())
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
        if parsed.path != "/generate":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length).decode("utf-8")
        form = parse_qs(raw)
        mode = form.get("mode", ["all"])[0]
        date_raw = form.get("date", [""])[0]
        send_bitrix = form.get("send_bitrix", [""])[0] == "1"

        if mode not in MODES:
            self._send_error_page("Неизвестный режим отчёта.", send_bitrix=send_bitrix)
            return
        try:
            report_date = _parse_report_date(date_raw)
        except ValueError:
            self._send_error_page(
                "Дата должна быть в формате YYYY-MM-DD.",
                selected_mode=mode,
                send_bitrix=send_bitrix,
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
        except ODataUnavailableError as exc:
            self._send_error_page(
                f"OData недоступен: {exc}",
                selected_mode=mode,
                selected_date=report_date,
                send_bitrix=send_bitrix,
            )
            return
        except (ODataError, BitrixError, RuntimeError) as exc:
            self._send_error_page(
                str(exc),
                selected_mode=mode,
                selected_date=report_date,
                send_bitrix=send_bitrix,
            )
            return

        elapsed = time.monotonic() - started
        suffix = " Отправлено в Bitrix24." if send_bitrix else ""
        self._send_html(
            _page(
                selected_mode=mode,
                selected_date=report_date,
                send_bitrix=send_bitrix,
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
