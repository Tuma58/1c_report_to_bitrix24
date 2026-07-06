"""CLI генерации Excel-отчётов по ТЗ.

Создаёт все 10 форм: 5 направлений × дневная/недельная форма.
Опционально отправляет созданные файлы в чат Bitrix24.
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
    from .metrics import Metric, MetricsService, WashMetrics
    from .odata_client import ODataError, ODataUnavailableError
except ImportError:  # запуск как скрипта из backend/src
    from bitrix_sender import BitrixSender, BitrixError
    from excel_reporter import ExcelReporter
    from metrics import Metric, MetricsService, WashMetrics
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


def _safe_name(name: str) -> str:
    return name.replace(" ", "_").replace(".", "").replace("/", "_")


def _parse_date(value: str) -> date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise argparse.ArgumentTypeError("дата должна быть в формате YYYY-MM-DD") from exc


def previous_week(ref: date) -> date:
    """Понедельник предыдущей ISO-недели."""
    return ref - timedelta(days=ref.weekday() + 7)


def _sum_plan(a, b):
    if a is None and b is None:
        return None
    return (a or 0.0) + (b or 0.0)


def combine_wash(name: str, a: WashMetrics, b: WashMetrics) -> WashMetrics:
    result = WashMetrics(
        report_name=name,
        period_start=a.period_start,
        period_end=a.period_end,
        division_key=None,
    )
    result.cars = Metric(
        "Машин обслужено",
        (a.cars.fact or 0.0) + (b.cars.fact or 0.0),
        _sum_plan(a.cars.plan, b.cars.plan),
    )
    result.revenue = Metric(
        "Выручка",
        (a.revenue.fact or 0.0) + (b.revenue.fact or 0.0),
        _sum_plan(a.revenue.plan, b.revenue.plan),
    )
    result.executors_count = (a.executors_count or 0) + (b.executors_count or 0)
    return result


def generate_daily(service: MetricsService, reporter: ExcelReporter, day: date) -> list[GeneratedReport]:
    reports: list[GeneratedReport] = []

    for name in SERVICE_REPORTS:
        metrics = service.daily(name, day)
        out = OUTPUT_DIR / f"{_safe_name(name)}_{day.isoformat()}.xlsx"
        reports.append(GeneratedReport("daily", name, reporter.fill_daily(name, day, metrics, out)))

    stroit = service.daily_wash(WASH_STROIT, day)
    ulyan = service.daily_wash(WASH_ULYAN, day)
    out = OUTPUT_DIR / f"wash_daily_{day.isoformat()}.xlsx"
    reports.append(GeneratedReport("daily", "Мойки", reporter.fill_daily_wash(day, stroit, ulyan, out)))

    for name in SHOP_REPORTS:
        metrics = service.daily_shop(name, day)
        out = OUTPUT_DIR / f"shop_{name}_{day.isoformat()}.xlsx"
        reports.append(GeneratedReport("daily", name, reporter.fill_daily_shop(name, metrics, out)))

    return reports


def generate_weekly(
    service: MetricsService,
    reporter: ExcelReporter,
    any_date: date,
) -> list[GeneratedReport]:
    reports: list[GeneratedReport] = []

    for name in SERVICE_REPORTS:
        weekly = service.weekly(name, any_date)
        days = service.week_daily_breakdown(name, any_date)
        out = OUTPUT_DIR / f"weekly_{name}_{weekly.week_start.isoformat()}.xlsx"
        reports.append(GeneratedReport("weekly", name, reporter.fill_weekly(name, weekly, days, out)))

    stroit = service.weekly_wash(WASH_STROIT, any_date)
    ulyan = service.weekly_wash(WASH_ULYAN, any_date)
    combined = combine_wash("Мойки", stroit, ulyan)
    days_stroit = service.week_wash_daily_breakdown(WASH_STROIT, any_date)
    days_ulyan = service.week_wash_daily_breakdown(WASH_ULYAN, any_date)
    day_combined = [
        combine_wash("Мойки", a, b) for a, b in zip(days_stroit, days_ulyan)
    ]
    out = OUTPUT_DIR / f"wash_weekly_{combined.period_start.isoformat()}.xlsx"
    reports.append(
        GeneratedReport(
            "weekly",
            "Мойки",
            reporter.fill_weekly_wash(any_date, combined, day_combined, out),
        )
    )

    for name in SHOP_REPORTS:
        weekly = service.weekly_shop(name, any_date)
        days = service.week_shop_daily_breakdown(name, any_date)
        out = OUTPUT_DIR / f"shop_weekly_{name}_{weekly.week_start.isoformat()}.xlsx"
        reports.append(
            GeneratedReport("weekly", name, reporter.fill_weekly_shop(name, weekly, days, out))
        )

    return reports


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
        help="отправить созданные файлы в чат Bitrix24",
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
        generated: list[GeneratedReport] = []
        if args.mode in ("daily", "all"):
            generated.extend(generate_daily(service, reporter, daily_day))
        if args.mode in ("weekly", "all"):
            generated.extend(generate_weekly(service, reporter, weekly_day))
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
            send_to_bitrix(generated, "Отчёты АТЦ: созданные Excel-файлы")
            print("[OK] Отправлено в Bitrix24")
        except BitrixError as exc:
            print(f"Ошибка Bitrix24: {exc}", file=sys.stderr)
            return 4

    return 0


if __name__ == "__main__":
    sys.exit(main())
