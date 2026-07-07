"""CLI генерации Excel-отчётов по ТЗ.

Создаёт один Excel-файл на базе шаблона: каждый отчёт заполняется на своём
листе, как в `templates/report_template.xlsx`.
Опционально отправляет созданный файл в чат Bitrix24.
"""
from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    from .bitrix_sender import BitrixSender, BitrixError
    from .excel_reporter import ExcelReporter
    from .metrics import MetricsService
    from .odata_client import ODataError, ODataUnavailableError
except ImportError:  # запуск как скрипта из backend/src
    from bitrix_sender import BitrixSender, BitrixError
    from excel_reporter import ExcelReporter
    from metrics import MetricsService
    from odata_client import ODataError, ODataUnavailableError


OUTPUT_DIR = Path(__file__).resolve().parent.parent / "output"
SERVICE_REPORTS = ("Арсенал", "Реф. Сервис")
SHOP_REPORTS = ("Шаркер", "ЦКР")
WASH_STROIT = "Мойка на ул. Строителей"
WASH_ULYAN = "Мойка на ул. Ульяновская"


@dataclass
class GeneratedReport:
    kind: str
    name: str
    path: Path


def _parse_date(value: str) -> date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise argparse.ArgumentTypeError("дата должна быть в формате YYYY-MM-DD") from exc


def previous_week(ref: date) -> date:
    """Понедельник предыдущей ISO-недели."""
    return ref - timedelta(days=ref.weekday() + 7)


def fill_daily_sheets(service: MetricsService, reporter: ExcelReporter, wb, day: date) -> None:
    for name in SERVICE_REPORTS:
        metrics = service.daily(name, day)
        reporter.fill_daily_in_workbook(wb, name, day, metrics)

    stroit = service.daily_wash(WASH_STROIT, day)
    ulyan = service.daily_wash(WASH_ULYAN, day)
    reporter.fill_daily_wash_in_workbook(wb, day, stroit, ulyan)

    for name in SHOP_REPORTS:
        metrics = service.daily_shop(name, day)
        reporter.fill_daily_shop_in_workbook(wb, name, metrics)


def fill_weekly_sheets(
    service: MetricsService,
    reporter: ExcelReporter,
    wb,
    any_date: date,
) -> None:
    for name in SERVICE_REPORTS:
        weekly = service.weekly(name, any_date)
        days = service.week_daily_breakdown(name, any_date)
        reporter.fill_weekly_in_workbook(wb, name, weekly, days)

    stroit = service.weekly_wash(WASH_STROIT, any_date)
    ulyan = service.weekly_wash(WASH_ULYAN, any_date)
    days_stroit = service.week_wash_daily_breakdown(WASH_STROIT, any_date)
    days_ulyan = service.week_wash_daily_breakdown(WASH_ULYAN, any_date)
    reporter.fill_weekly_wash_in_workbook(wb, WASH_STROIT, stroit, days_stroit)
    reporter.fill_weekly_wash_in_workbook(wb, WASH_ULYAN, ulyan, days_ulyan)

    for name in SHOP_REPORTS:
        weekly = service.weekly_shop(name, any_date)
        days = service.week_shop_daily_breakdown(name, any_date)
        reporter.fill_weekly_shop_in_workbook(wb, name, weekly, days)


def _weekly_start(service: MetricsService, any_date: date) -> date:
    start, _ = service._week_bounds(any_date)
    return start


def _bundle_path(mode: str, service: MetricsService, daily_day: date, weekly_day: date) -> Path:
    if mode == "daily":
        return OUTPUT_DIR / f"reports_daily_{daily_day.isoformat()}.xlsx"
    if mode == "weekly":
        return OUTPUT_DIR / f"reports_weekly_{_weekly_start(service, weekly_day).isoformat()}.xlsx"
    return OUTPUT_DIR / (
        f"reports_all_daily_{daily_day.isoformat()}_"
        f"weekly_{_weekly_start(service, weekly_day).isoformat()}.xlsx"
    )


def generate_workbook(
    service: MetricsService,
    reporter: ExcelReporter,
    mode: str,
    daily_day: date,
    weekly_day: date,
) -> list[GeneratedReport]:
    wb = reporter.new_workbook()
    if mode in ("daily", "all"):
        fill_daily_sheets(service, reporter, wb, daily_day)
    if mode in ("weekly", "all"):
        fill_weekly_sheets(service, reporter, wb, weekly_day)

    out = reporter.save_workbook(wb, _bundle_path(mode, service, daily_day, weekly_day))
    return [GeneratedReport(mode, "Все отчёты", out)]


def send_to_bitrix(reports: list[GeneratedReport], title: str) -> None:
    sender = BitrixSender()
    sender.send_files([r.path for r in reports], title)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Генерация Excel-отчётов АТЦ")
    parser.add_argument(
        "--mode",
        choices=("daily", "weekly", "all"),
        default="all",
        help="какие формы создавать",
    )
    parser.add_argument(
        "--date",
        type=_parse_date,
        default=None,
        help="дата отчёта YYYY-MM-DD; для weekly берётся неделя этой даты",
    )
    parser.add_argument(
        "--send-bitrix",
        action="store_true",
        help="отправить созданный файл в чат Bitrix24",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    today = date.today()
    daily_day = args.date or (today - timedelta(days=1))
    weekly_day = args.date or previous_week(today)

    try:
        service = MetricsService()
        reporter = ExcelReporter()
        generated = generate_workbook(service, reporter, args.mode, daily_day, weekly_day)
    except ODataUnavailableError as exc:
        print(f"OData недоступен: {exc}", file=sys.stderr)
        return 2
    except ODataError as exc:
        print(f"Ошибка OData: {exc}", file=sys.stderr)
        return 3

    for report in generated:
        print(f"[OK] {report.kind:6} {report.name:24} {report.path}")

    if args.send_bitrix:
        try:
            send_to_bitrix(generated, "Отчёты АТЦ: Excel-файл с листами отчётов")
            print("[OK] Отправлено в Bitrix24")
        except BitrixError as exc:
            print(f"Ошибка Bitrix24: {exc}", file=sys.stderr)
            return 4

    return 0


if __name__ == "__main__":
    sys.exit(main())
