"""Тест Инкремента 4: наценка parts-only, сводная, недельные показатели.

Проверяет:
- наценка ФАКТ Арсенал 01.07.2026 ≈ 49.2%, Реф ≈ 63.0% (parts-only);
- регресс дневного расчёта Арсенал 2026-06-10 (нормочасы 170.36,
  выручка 581 397.50, себест 186 026.96, закрыто 28);
- ConsolidatedReporter.daily(2026-07-01) — печать по 6 орг. + сохранение xlsx;
- weekly('Арсенал', дата июля 2026) — недельный ФАКТ и ПЛАН заполнены.

Только GET из 1С; запись только в xlsx. Если порт недоступен — код 2.
"""
from __future__ import annotations

import os
import sys
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from odata_client import ODataUnavailableError  # noqa: E402
from metrics import MetricsService  # noqa: E402
from consolidated import ConsolidatedReporter  # noqa: E402


MARKUP_DAY = date(2026, 7, 1)
REGRESS_DAY = date(2026, 6, 10)
CONSOL_DAY = date(2026, 7, 1)
WEEK_DAY = date(2026, 7, 1)  # среда — ISU-неделя 29.06–05.07.2026


def _fmt(v):
    if v is None:
        return "—"
    return f"{v:,.2f}".replace(",", " ")


def _approx(a, b, tol) -> bool:
    return a is not None and abs(a - b) <= tol


def main() -> int:
    try:
        service = MetricsService()
    except ODataUnavailableError as exc:
        print("OData НЕДОСТУПЕН (порт / сеть) — не вина кода:")
        print(f"  {exc}")
        return 2

    failures: list[str] = []
    passed: list[str] = []

    try:
        # --- 1. Наценка parts-only ---
        print("=" * 70)
        print("ЗАДАЧА 1. Наценка на ЗЧ (parts-only)")
        print("=" * 70)
        ars = service.daily("Арсенал", MARKUP_DAY)
        ref = service.daily("Реф. Сервис", MARKUP_DAY)
        ars_mk = ars.markup_pct.fact
        ref_mk = ref.markup_pct.fact
        print(f"  Арсенал {MARKUP_DAY.isoformat()}: наценка ФАКТ = {_fmt(ars_mk)} %")
        print(f"  Реф.Сервис {MARKUP_DAY.isoformat()}: наценка ФАКТ = {_fmt(ref_mk)} %")

        if _approx(ars_mk, 49.2, 1.0):
            passed.append("наценка Арсенал ≈ 49.2%")
        else:
            failures.append(f"наценка Арсенал = {_fmt(ars_mk)} (ожид. ≈49.2)")
        if _approx(ref_mk, 63.0, 1.5):
            passed.append("наценка Реф ≈ 63.0%")
        else:
            failures.append(f"наценка Реф = {_fmt(ref_mk)} (ожид. ≈63.0)")

        # --- 2. Регресс дневного Арсенал 2026-06-10 ---
        print("\n" + "=" * 70)
        print("РЕГРЕСС. Дневной Арсенал 2026-06-10")
        print("=" * 70)
        reg = service.daily("Арсенал", REGRESS_DAY)
        print(f"  Нормочасы:  {_fmt(reg.normhours.fact)} (ожид. 170.36)")
        print(f"  Выручка:    {_fmt(reg.revenue.fact)} (ожид. 581 397.50)")
        print(f"  Себест.:    {_fmt(reg.cost.fact)} (ожид. 186 026.96)")
        print(f"  Закрыто ЗН: {reg.closed_orders_count} (ожид. 28)")

        checks = [
            ("нормочасы 170.36", _approx(reg.normhours.fact, 170.36, 0.05)),
            ("выручка 581 397.50", _approx(reg.revenue.fact, 581397.50, 0.5)),
            ("себест 186 026.96", _approx(reg.cost.fact, 186026.96, 0.5)),
            ("закрыто 28", reg.closed_orders_count == 28),
        ]
        for name, ok in checks:
            (passed if ok else failures).append(f"регресс {name}")

        # --- 3. Сводная за день ---
        print("\n" + "=" * 70)
        print("ЗАДАЧА 2. Сводная таблица 2026-07-01")
        print("=" * 70)
        reporter = ConsolidatedReporter(service=service)
        consol = reporter.daily(CONSOL_DAY)
        if len(consol) == 6:
            passed.append("сводная: 6 организаций")
        else:
            failures.append(f"сводная: {len(consol)} орг. (ожид. 6)")

        # --- 4. Недельные показатели ---
        print("\n" + "=" * 70)
        print("ЗАДАЧА 3. Недельные показатели — Арсенал")
        print("=" * 70)
        wk = service.weekly("Арсенал", WEEK_DAY)
        print(f"  Неделя: {wk.week_start.isoformat()} .. {wk.week_end.isoformat()}")
        print(f"  Закрыто ЗН (шт): {wk.closed_orders_count}, мастеров: {wk.executors_count}")
        print(f"  {'Показатель':<40} {'ФАКТ':>16} {'ПЛАН/нед':>16}")
        for met in wk.all_metrics():
            print(f"  {met.name:<40} {_fmt(met.fact):>16} {_fmt(met.plan):>16}")

        wk_fact_ok = wk.revenue.fact > 0 or wk.closed_orders_count > 0
        wk_plan_ok = wk.revenue.plan is not None and wk.normhours.plan is not None
        (passed if wk_fact_ok else failures).append("недельный ФАКТ заполнен")
        (passed if wk_plan_ok else failures).append("недельный ПЛАН заполнен (месяц/дней×7)")

    except ODataUnavailableError as exc:
        print("OData стал недоступен во время теста — не вина кода:")
        print(f"  {exc}")
        return 2

    # --- Итог ---
    print("\n" + "=" * 70)
    print("ИТОГ КРИТЕРИЕВ")
    print("=" * 70)
    for p in passed:
        print(f"  [PASS] {p}")
    for f in failures:
        print(f"  [FAIL] {f}")

    if failures:
        print(f"\nНЕ ПРОШЛО: {len(failures)} из {len(passed) + len(failures)}")
        return 1
    print(f"\nВСЕ КРИТЕРИИ ПРОШЛИ: {len(passed)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
