"""Smoke-тест дневных показателей (Инкремент 2).

Прогоняет MetricsService.daily() для «Арсенал» и «Реф. Сервис» за будний
день июня 2026, печатает разбор ФАКТ/ПЛАН и проверяет, что ФАКТ ненулевой.

Только GET-запросы к живой базе. Если порт 8888 недоступен — фиксируем
это отдельным сообщением (не вина кода) и выходим с кодом 2.
"""
from __future__ import annotations

import os
import sys
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from odata_client import ODataUnavailableError  # noqa: E402
from metrics import MetricsService, DailyMetrics  # noqa: E402


TEST_DAY = date(2026, 6, 10)  # будний день (среда) июня 2026
REPORTS = ["Арсенал", "Реф. Сервис"]


def _fmt(v):
    if v is None:
        return "—"
    return f"{v:,.2f}".replace(",", " ")


def print_metrics(m: DailyMetrics) -> None:
    print(f"\n=== {m.report_name} — {m.day.isoformat()} ===")
    print(f"  Подразделение Ref_Key: {m.division_key}")
    print(f"  Закрыто ЗН (шт): {m.closed_orders_count}")
    print(f"  Уникальных мастеров: {m.executors_count}")
    print(f"  {'Показатель':<38} {'ФАКТ':>16} {'ПЛАН/день':>16}")
    for met in m.all_metrics():
        print(f"  {met.name:<38} {_fmt(met.fact):>16} {_fmt(met.plan):>16}")


def main() -> int:
    try:
        service = MetricsService()
        results = {rn: service.daily(rn, TEST_DAY) for rn in REPORTS}
    except ODataUnavailableError as exc:
        print("OData НЕДОСТУПЕН (порт 8888 / сеть) — не вина кода:")
        print(f"  {exc}")
        return 2

    failures: list[str] = []
    for rn in REPORTS:
        m = results[rn]
        print_metrics(m)
        # ФАКТ ненулевой: считаем по ключевым ФАКТ-показателям
        fact_sum = (
            m.normhours.fact
            + m.revenue.fact
            + m.cost.fact
            + m.closed_orders_count
        )
        if fact_sum <= 0:
            failures.append(f"{rn}: ФАКТ нулевой (нормочасы+выручка+себест+ЗН = {fact_sum})")

    print("\n" + "=" * 60)
    if failures:
        print("ПРОВАЛЕНО:")
        for f in failures:
            print("  -", f)
        return 1
    print("ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ: ФАКТ ненулевой для обоих подразделений.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
