"""Расчёт показателей ПЛАН/ФАКТ (Арсенал, Реф. Сервис и др.).

MetricsService.daily(report_name, day)   — дневной расчёт (период [day, day+1)).
MetricsService.weekly(report_name, date) — недельный расчёт (ISO-неделя Пн–Вс).

ФАКТ собирается из репозиториев, ПЛАН — из документа плана:
- дневной ПЛАН   = месячное значение ÷ число дней месяца;
- недельный ПЛАН = месячное значение ÷ число дней месяца × 7;
- средний чек НЕ масштабируется (это чек, а не поток).

ПЛАН по строке = None, если Code в plan_mapping не задан или отсутствует в
документе плана — отчёт при этом не падает.
"""
from __future__ import annotations

import calendar
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Optional

try:
    from .odata_client import ODataClient
    from .references import References
    from .repositories import (
        OrderRepository,
        IncomeExpenseRepository,
        PlanRepository,
        PaymentRepository,
    )
    from .plan_mapping import plan_code
except ImportError:  # запуск как скрипта
    from odata_client import ODataClient
    from references import References
    from repositories import (
        OrderRepository,
        IncomeExpenseRepository,
        PlanRepository,
        PaymentRepository,
    )
    from plan_mapping import plan_code


@dataclass
class Metric:
    """Одна строка отчёта: ФАКТ и (опционально) ПЛАН."""
    name: str
    fact: float
    plan: Optional[float] = None


@dataclass
class DailyMetrics:
    """Результат дневного расчёта по подразделению."""
    report_name: str
    day: date
    division_key: Optional[str]

    # ФАКТ вспомогательные
    closed_orders_count: int = 0
    executors_count: int = 0

    normhours: Metric = field(default_factory=lambda: Metric("Нормочасы за день (общая выработка)", 0.0))
    output_per_master: Metric = field(default_factory=lambda: Metric("Выработка на 1 мастера", 0.0))
    closed_orders: Metric = field(default_factory=lambda: Metric("Закрыто ЗН за день", 0.0))
    revenue: Metric = field(default_factory=lambda: Metric("Выручка за день (без НДС)", 0.0))
    cost: Metric = field(default_factory=lambda: Metric("Себестоимость ЗЧ (без НДС)", 0.0))
    markup: Metric = field(default_factory=lambda: Metric("Наценка", 0.0))
    markup_pct: Metric = field(default_factory=lambda: Metric("Наценка на ЗЧ, %", 0.0))
    unclosed_over_1day: Metric = field(default_factory=lambda: Metric("Незакрытые ЗН старше 1 дня", 0.0))

    def all_metrics(self) -> list[Metric]:
        return [
            self.normhours,
            self.output_per_master,
            self.closed_orders,
            self.revenue,
            self.cost,
            self.markup,
            self.markup_pct,
            self.unclosed_over_1day,
        ]


@dataclass
class WeeklyMetrics:
    """Результат недельного расчёта (ISO-неделя Пн–Вс) по подразделению."""
    report_name: str
    week_start: date          # понедельник
    week_end: date            # воскресенье (включительно)
    division_key: Optional[str]

    closed_orders_count: int = 0
    executors_count: int = 0

    normhours: Metric = field(default_factory=lambda: Metric("Нормочасы за неделю (общая выработка)", 0.0))
    output_per_master: Metric = field(default_factory=lambda: Metric("Выработка на 1 мастера", 0.0))
    closed_orders: Metric = field(default_factory=lambda: Metric("Закрыто ЗН за неделю", 0.0))
    revenue: Metric = field(default_factory=lambda: Metric("Выручка за неделю (без НДС)", 0.0))
    cost: Metric = field(default_factory=lambda: Metric("Себестоимость ЗЧ (без НДС)", 0.0))
    markup: Metric = field(default_factory=lambda: Metric("Наценка", 0.0))
    markup_pct: Metric = field(default_factory=lambda: Metric("Наценка на ЗЧ, %", 0.0))
    unclosed_over_1day: Metric = field(default_factory=lambda: Metric("Незакрытые ЗН старше 1 дня", 0.0))

    def all_metrics(self) -> list[Metric]:
        return [
            self.normhours,
            self.output_per_master,
            self.closed_orders,
            self.revenue,
            self.cost,
            self.markup,
            self.markup_pct,
            self.unclosed_over_1day,
        ]


class MetricsService:
    def __init__(self, client: Optional[ODataClient] = None) -> None:
        self.client = client or ODataClient()
        self.references = References(self.client)
        self.orders = OrderRepository(self.client, self.references)
        self.income = IncomeExpenseRepository(self.client, self.references)
        self.plan = PlanRepository(self.client, self.references)
        self.payments = PaymentRepository(self.client, self.references)

    @staticmethod
    def _days_in_month(day: date) -> int:
        return calendar.monthrange(day.year, day.month)[1]

    @staticmethod
    def _week_bounds(any_date: date) -> tuple[date, date]:
        """ISO-неделя (Пн–Вс), содержащая any_date.

        Возвращает (понедельник, следующий понедельник) — интервал [start, end).
        """
        start = any_date - timedelta(days=any_date.weekday())  # понедельник
        end = start + timedelta(days=7)                        # следующий понедельник
        return start, end

    def _plan_day(self, monthly: dict[str, float], row: str, day: date) -> Optional[float]:
        """Дневное плановое значение для строки отчёта или None."""
        code = plan_code(row)
        if not code:
            return None
        month_value = monthly.get(code)
        if month_value is None:
            return None
        return month_value / self._days_in_month(day)

    def _plan_week(self, monthly: dict[str, float], row: str, ref_day: date) -> Optional[float]:
        """Недельное плановое значение = месяц ÷ дней в месяце × 7 или None."""
        daily = self._plan_day(monthly, row, ref_day)
        if daily is None:
            return None
        return daily * 7.0

    @staticmethod
    def _markup_pct_fact(parts: float, cost: float):
        """ФАКТ наценки на ЗЧ, % — формула «только запчасти» (parts-only).

        РЕШЕНО(наценка-на-ЗЧ): по решению владельца — parts-only.
        Наценка на ЗЧ, % = (Выручка по заказ-нарядам − Себестоимость материалов)
                           / Себестоимость материалов × 100.
        Используется ТОЛЬКО статья «Выручка по заказ-нарядам» (запчасти),
        БЕЗ «Работы по заказ-нарядам» (труд). Себестоимость — «Себестоимость
        материалов». При cost = 0 → None (прочерк).
        Проверенные значения (01.07.2026): Арсенал 49.2% (parts=187 717,
        cost=125 807), Реф. Сервис 63.0%.

        Прим.: требуем parts > 0 — иначе (напр. мойки: продаж ЗЧ нет, но есть
        списание материалов) формула давала бы бессмысленные −100%. Наценка на
        ЗЧ определена только при наличии продаж запчастей.
        """
        if cost > 0 and parts > 0:
            return (parts - cost) / cost * 100.0
        return None

    # ------------------------------------------------------------------ ПЛАН
    def _plan_block(
        self,
        monthly: dict[str, float],
        ref_day: date,
        scale_weekly: bool,
    ) -> dict:
        """Собирает плановые значения (дневные или недельные) в общий словарь."""
        pf = self._plan_week if scale_weekly else self._plan_day

        plan_revenue = pf(monthly, "выручка", ref_day)
        plan_cost = pf(monthly, "себестоимость_зч", ref_day)
        plan_output_master = pf(monthly, "выработка_на_мастера", ref_day)
        plan_normhours = pf(monthly, "нормочасы", ref_day)

        # средний чек — месячный, НЕ масштабируем (это чек, а не поток)
        avg_check_code = plan_code("средний_чек")
        avg_check = monthly.get(avg_check_code) if avg_check_code else None

        # закрыто ЗН план = выручка_план / средний_чек
        plan_closed = None
        if plan_revenue is not None and avg_check:
            plan_closed = plan_revenue / avg_check

        # наценка (сумма) план = выручка_план − себест_план
        plan_markup = None
        if plan_revenue is not None and plan_cost is not None:
            plan_markup = plan_revenue - plan_cost

        # наценка% план — код 000000024 (проценты), НЕ масштабируем
        markup_pct_code = plan_code("наценка")
        plan_markup_pct = monthly.get(markup_pct_code) if markup_pct_code else None

        return {
            "revenue": plan_revenue,
            "cost": plan_cost,
            "output_master": plan_output_master,
            "normhours": plan_normhours,
            "closed": plan_closed,
            "markup": plan_markup,
            "markup_pct": plan_markup_pct,
        }

    # ------------------------------------------------------------------ ФАКТ
    def _fact_block(self, division_key: str, start: date, end: date) -> dict:
        """Собирает ФАКТ за интервал [start, end): closed/normhours/executors/revenue/parts/cost."""
        closed = self.orders.closed_orders_period(division_key, start, end)
        order_keys = [o["Ref_Key"] for o in closed]

        works = self.orders.works(order_keys)
        executors = self.orders.executors(order_keys)
        normhours = self.orders.normhours(works)
        executors_count = self.orders.unique_executors(executors)

        breakdown = self.income.revenue_breakdown_period(division_key, start, end)
        parts = breakdown["parts"]
        cost = breakdown["cost"]
        revenue = parts + breakdown["works"]  # полная выручка = запчасти + работы

        return {
            "order_keys": order_keys,
            "closed_count": len(order_keys),
            "normhours": normhours,
            "executors_count": executors_count,
            "revenue": revenue,      # полная (parts+works)
            "parts": parts,          # только запчасти (для наценки)
            "cost": cost,
        }

    def daily(self, report_name: str, day: date) -> DailyMetrics:
        division_key = self.references.division_key(report_name)
        result = DailyMetrics(report_name=report_name, day=day, division_key=division_key)
        if not division_key:
            # подразделение не сопоставлено — вернуть пустой результат (ФАКТ=0), не падать
            return result

        start, end = day, day + timedelta(days=1)  # период [day, day+1)
        fact = self._fact_block(division_key, start, end)
        result.closed_orders_count = fact["closed_count"]
        result.executors_count = fact["executors_count"]

        monthly = self.plan.monthly_values(division_key, day)
        plan = self._plan_block(monthly, day, scale_weekly=False)

        normhours = fact["normhours"]
        executors_count = fact["executors_count"]

        result.normhours = Metric(result.normhours.name, normhours, plan["normhours"])
        result.output_per_master = Metric(
            result.output_per_master.name,
            (normhours / executors_count) if executors_count else 0.0,
            plan["output_master"],
        )
        result.closed_orders = Metric(
            result.closed_orders.name, float(fact["closed_count"]), plan["closed"]
        )
        result.revenue = Metric(result.revenue.name, fact["revenue"], plan["revenue"])
        result.cost = Metric(result.cost.name, fact["cost"], plan["cost"])
        result.markup = Metric(result.markup.name, fact["revenue"] - fact["cost"], plan["markup"])

        # Наценка на ЗЧ, % — parts-only (только запчасти).
        markup_pct_fact = self._markup_pct_fact(fact["parts"], fact["cost"])
        result.markup_pct = Metric(result.markup_pct.name, markup_pct_fact, plan["markup_pct"])

        # Незакрытые ЗН старше 1 дня (ПЛАН = None).
        unclosed = self.orders.open_orders_older_than_1day_asof(division_key, day)
        result.unclosed_over_1day = Metric(
            result.unclosed_over_1day.name, float(unclosed), None
        )

        return result

    def weekly(self, report_name: str, any_date: date) -> WeeklyMetrics:
        """Недельный расчёт по ISO-неделе (Пн–Вс), содержащей any_date."""
        start, end = self._week_bounds(any_date)  # [понедельник, след. понедельник)
        week_end_incl = end - timedelta(days=1)   # воскресенье (включительно)

        division_key = self.references.division_key(report_name)
        result = WeeklyMetrics(
            report_name=report_name,
            week_start=start,
            week_end=week_end_incl,
            division_key=division_key,
        )
        if not division_key:
            return result

        fact = self._fact_block(division_key, start, end)
        result.closed_orders_count = fact["closed_count"]
        result.executors_count = fact["executors_count"]

        # ПЛАН по месяцу опорной даты any_date (неделя может пересекать месяцы —
        # берём месяц запрошенной даты как опорный: масштаб «месяц ÷ дней × 7»).
        monthly = self.plan.monthly_values(division_key, any_date)
        plan = self._plan_block(monthly, any_date, scale_weekly=True)

        normhours = fact["normhours"]
        executors_count = fact["executors_count"]

        result.normhours = Metric(result.normhours.name, normhours, plan["normhours"])
        result.output_per_master = Metric(
            result.output_per_master.name,
            (normhours / executors_count) if executors_count else 0.0,
            plan["output_master"],
        )
        result.closed_orders = Metric(
            result.closed_orders.name, float(fact["closed_count"]), plan["closed"]
        )
        result.revenue = Metric(result.revenue.name, fact["revenue"], plan["revenue"])
        result.cost = Metric(result.cost.name, fact["cost"], plan["cost"])
        result.markup = Metric(result.markup.name, fact["revenue"] - fact["cost"], plan["markup"])

        markup_pct_fact = self._markup_pct_fact(fact["parts"], fact["cost"])
        result.markup_pct = Metric(result.markup_pct.name, markup_pct_fact, plan["markup_pct"])

        # Незакрытые ЗН старше 1 дня — на конец недели (воскресенье).
        unclosed = self.orders.open_orders_older_than_1day_asof(division_key, week_end_incl)
        result.unclosed_over_1day = Metric(
            result.unclosed_over_1day.name, float(unclosed), None
        )

        return result

    def week_daily_breakdown(self, report_name: str, any_date: date) -> list[DailyMetrics]:
        """Разбивка по дням ISO-недели (Пн..Вс), содержащей any_date.

        Возвращает ровно 7 DailyMetrics: индекс 0 = понедельник (week_start),
        индекс 6 = воскресенье. Переиспользует daily() для каждого дня.
        """
        start, _ = self._week_bounds(any_date)  # понедельник
        return [self.daily(report_name, start + timedelta(days=i)) for i in range(7)]


# ============================================================ ИНКРЕМЕНТ 6: МОЙКИ
@dataclass
class WashMetrics:
    """Результат wash-расчёта (мойка) за период — дневной или недельный.

    Показатели простые: машин обслужено (cars) и выручка (revenue), каждый
    ФАКТ/ПЛАН. Средний чек в коде НЕ считается (формула Excel).
    """
    report_name: str
    period_start: date        # начало периода (день или понедельник недели)
    period_end: date          # конец периода включительно (тот же день / воскресенье)
    division_key: Optional[str]

    cars: Metric = field(default_factory=lambda: Metric("Машин обслужено", 0.0))
    revenue: Metric = field(default_factory=lambda: Metric("Выручка", 0.0))

    # вспомогательное (для динамики операторов)
    executors_count: int = 0

    def all_metrics(self) -> list[Metric]:
        return [self.cars, self.revenue]


# --- расширение MetricsService wash-методами (не трогаем daily()/weekly()) ---
def _wash_fact(self, division_key: str, start: date, end: date) -> dict:
    """ФАКТ мойки за [start, end): машин = число закрытых ЗН; выручка = все статьи."""
    closed = self.orders.closed_orders_period(division_key, start, end)
    order_keys = [o["Ref_Key"] for o in closed]
    executors = self.orders.executors(order_keys)
    executors_count = self.orders.unique_executors(executors)

    breakdown = self.income.revenue_breakdown_period(division_key, start, end)
    revenue = breakdown["parts"] + breakdown["works"]  # все статьи выручки

    return {
        "cars": len(order_keys),
        "revenue": revenue,
        "executors_count": executors_count,
    }


def daily_wash(self, report_name: str, day: date):
    """Дневной wash-расчёт (период [day, day+1))."""
    division_key = self.references.division_key(report_name)
    result = WashMetrics(
        report_name=report_name, period_start=day, period_end=day,
        division_key=division_key,
    )
    if not division_key:
        return result

    start, end = day, day + timedelta(days=1)
    fact = self._wash_fact(division_key, start, end)
    result.executors_count = fact["executors_count"]

    monthly = self.plan.monthly_values(division_key, day)
    plan_cars = self._plan_day(monthly, "машин_обслужено", day)   # код 000000074
    plan_revenue = self._plan_day(monthly, "выручка", day)         # код 000000067

    result.cars = Metric(result.cars.name, float(fact["cars"]), plan_cars)
    result.revenue = Metric(result.revenue.name, fact["revenue"], plan_revenue)
    return result


def weekly_wash(self, report_name: str, any_date: date):
    """Недельный wash-расчёт по ISO-неделе (Пн–Вс), содержащей any_date."""
    start, end = self._week_bounds(any_date)      # [Пн, след. Пн)
    week_end_incl = end - timedelta(days=1)        # воскресенье включительно

    division_key = self.references.division_key(report_name)
    result = WashMetrics(
        report_name=report_name, period_start=start, period_end=week_end_incl,
        division_key=division_key,
    )
    if not division_key:
        return result

    fact = self._wash_fact(division_key, start, end)
    result.executors_count = fact["executors_count"]

    monthly = self.plan.monthly_values(division_key, any_date)
    plan_cars = self._plan_week(monthly, "машин_обслужено", any_date)   # 074 ×7
    plan_revenue = self._plan_week(monthly, "выручка", any_date)         # 067 ×7

    result.cars = Metric(result.cars.name, float(fact["cars"]), plan_cars)
    result.revenue = Metric(result.revenue.name, fact["revenue"], plan_revenue)
    return result


def week_wash_daily_breakdown(self, report_name: str, any_date: date) -> list:
    """Разбивка wash по дням ISO-недели (Пн..Вс) — ровно 7 WashMetrics."""
    start, _ = self._week_bounds(any_date)
    return [self.daily_wash(report_name, start + timedelta(days=i)) for i in range(7)]


MetricsService._wash_fact = _wash_fact
MetricsService.daily_wash = daily_wash
MetricsService.weekly_wash = weekly_wash
MetricsService.week_wash_daily_breakdown = week_wash_daily_breakdown


# ============================================================ ИНКРЕМЕНТ 7: ШАРКЕР/ЦКР
@dataclass
class ShopMetrics:
    """Дневные показатели формы «длинного ремонта» (Шаркер / ЦКР).

    Блокированные показатели (нет данных в OData / формула не определена) — None
    с пометкой TODO; в отчёте выводятся прочерком «—».
    """
    report_name: str
    day: date
    division_key: Optional[str]

    # ВЫРАБОТКА
    normhours: Metric = field(default_factory=lambda: Metric("Нормочасы за день", 0.0))
    output_per_master: Metric = field(default_factory=lambda: Metric("Выработка на 1 мастера", 0.0))
    # только ЦКР; ПЛАН/ФАКТ BLOCKED
    output_cost: Metric = field(default_factory=lambda: Metric("Стоимость выработки", None))

    # ПАЙПЛАЙН (значения 1С на дату; ПЛАН отсутствует)
    active: int = 0
    ready_today: int = 0
    awaiting_parts: int = 0
    overdue: int = 0
    planned_close_week: int = 0

    # ФИНАНСЫ
    closed_orders: Metric = field(default_factory=lambda: Metric("Закрыто ЗН за день", 0.0))
    revenue_closed: Metric = field(default_factory=lambda: Metric("Выручка закрытых ЗН", 0.0))
    payments: Metric = field(default_factory=lambda: Metric("Поступления оплат", 0.0))
    insurance_count: Metric = field(default_factory=lambda: Metric("Кол-во страховых ЗН прошлых мес.", None))
    insurance_sum: Metric = field(default_factory=lambda: Metric("Сумма страховых ЗН прошлых мес.", None))


def daily_shop(self, report_name: str, day: date) -> ShopMetrics:
    """Дневной расчёт формы длинного ремонта (Шаркер / ЦКР) за день day.

    Доступное считаем, заблокированное = None (TODO).
    """
    division_key = self.references.division_key(report_name)
    result = ShopMetrics(report_name=report_name, day=day, division_key=division_key)
    if not division_key:
        return result

    start, end = day, day + timedelta(days=1)

    # ---- ПЛАН (месячные показатели ÷ дней месяца) ----
    monthly = self.plan.monthly_values(division_key, day)
    plan_normhours = self._plan_day(monthly, "нормочасы", day)               # 000000029
    plan_output_master = self._plan_day(monthly, "выработка_на_мастера", day)  # 000000068
    plan_revenue = self._plan_day(monthly, "выручка", day)                   # 000000067
    avg_check_code = plan_code("средний_чек")                                # 000000069
    avg_check = monthly.get(avg_check_code) if avg_check_code else None
    plan_closed = (plan_revenue / avg_check) if (plan_revenue is not None and avg_check) else None

    # ---- ВЫРАБОТКА ----
    # ФАКТ нормочасов/выработки — BLOCKED: «мастер вводит вручную», в базе нет.
    # TODO(BLOCKED normhours/output ФАКТ): источник данных отсутствует в OData.
    result.normhours = Metric(result.normhours.name, None, plan_normhours)
    result.output_per_master = Metric(result.output_per_master.name, None, plan_output_master)
    # Стоимость выработки (только ЦКР) — BLOCKED: формула не определена.
    # TODO(BLOCKED output_cost): согласовать формулу стоимости выработки.
    result.output_cost = Metric(result.output_cost.name, None, None)

    # ---- ПАЙПЛАЙН (значения 1С на дату day) ----
    result.active = self.orders.pipeline_active(division_key, day)
    result.ready_today = self.orders.pipeline_ready_today(division_key, day)
    result.awaiting_parts = self.orders.pipeline_awaiting_parts(division_key, day)
    result.overdue = self.orders.pipeline_overdue(division_key, day)
    result.planned_close_week = self.orders.pipeline_planned_close_week(division_key, day)

    # ---- ФИНАНСЫ ----
    closed = self.orders.closed_orders_period(division_key, start, end)
    closed_count = len(closed)
    result.closed_orders = Metric(result.closed_orders.name, float(closed_count), plan_closed)

    breakdown = self.income.revenue_breakdown_period(division_key, start, end)
    revenue = breakdown["parts"] + breakdown["works"]  # без НДС; у длинных ремонтов дневной ФАКТ часто 0
    result.revenue_closed = Metric(result.revenue_closed.name, revenue, plan_revenue)

    payments = self.payments.payments_period(division_key, start, end)
    result.payments = Metric(result.payments.name, payments, None)

    # Страховые ЗН прошлых мес. — BLOCKED: «Дата вручения КА» отсутствует в OData счёт-фактуры.
    # TODO(BLOCKED insurance): нет поля даты вручения контрагенту.
    result.insurance_count = Metric(result.insurance_count.name, None, None)
    result.insurance_sum = Metric(result.insurance_sum.name, None, None)

    return result


MetricsService.daily_shop = daily_shop

