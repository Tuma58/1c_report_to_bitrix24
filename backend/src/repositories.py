"""Репозитории чтения данных 1С (только GET).

- OrderRepository       — закрытые ЗН за день/период, их Работы (нормочасы) и Исполнители.
- IncomeExpenseRepository — выручка/себестоимость из регистра ДоходыИРасходы_RecordType.
- PlanRepository        — последний документ плана за месяц + строки показателей.

Общие правила:
- Даты в фильтрах OData — литерал datetime'YYYY-MM-DDTHH:MM:SS' (без кавычек-строк).
- Интервал [start, end): field ge datetime'start' and field lt datetime'end'.
- Дневной интервал — частный случай периода [day, day+1).
- guid-фильтры группируются батчами (~20) через 'Ref_Key eq guid'..' or ...'.
"""
from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Iterable, Optional

try:
    from .odata_client import ODataClient, build_params
    from .references import References
except ImportError:  # запуск как скрипта
    from odata_client import ODataClient, build_params
    from references import References


GUID_BATCH = 20


def _dt(d: date) -> str:
    """OData-литерал datetime для начала дня d."""
    return f"datetime'{d.strftime('%Y-%m-%d')}T00:00:00'"


def _range_interval(field: str, start: date, end: date) -> str:
    """field ge datetime'start' and field lt datetime'end' — интервал [start, end)."""
    return f"{field} ge {_dt(start)} and {field} lt {_dt(end)}"


def _day_interval(field: str, d: date) -> str:
    """Дневной интервал [d, d+1) — частный случай _range_interval."""
    return _range_interval(field, d, d + timedelta(days=1))


def _batched(items: list[str], size: int = GUID_BATCH) -> Iterable[list[str]]:
    for i in range(0, len(items), size):
        yield items[i : i + size]


def _guid_or(field: str, keys: list[str]) -> str:
    """Собирает 'field eq guid'k1' or field eq guid'k2' ...'."""
    return " or ".join(f"{field} eq guid'{k}'" for k in keys)


def _num(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _parse_odata_datetime(value) -> Optional[datetime]:
    if not value or not isinstance(value, str):
        return None
    raw = value.strip().replace("Z", "+00:00")
    try:
        return datetime.fromisoformat(raw)
    except ValueError:
        try:
            return datetime.strptime(raw[:19], "%Y-%m-%dT%H:%M:%S")
        except ValueError:
            return None


def _has_filled_date(value) -> bool:
    parsed = _parse_odata_datetime(value)
    return parsed is not None and parsed.year > 1901


class OrderRepository:
    """Заказ-наряды: шапки, работы (нормочасы), исполнители."""

    def __init__(self, client: ODataClient, references: References) -> None:
        self.client = client
        self.references = references

    def closed_orders_period(self, division_key: str, start: date, end: date) -> list[dict]:
        """Шапки закрытых ЗН по подразделению за период [start, end).

        Закрытый = Состояние 'Закрыт' и ДатаЗакрытия в интервале [start, end).
        """
        closed_key = self.references.statuses().get("Закрыт")
        if not closed_key:
            raise RuntimeError("Не найден статус 'Закрыт' в References.statuses()")
        f = (
            f"Состояние_Key eq guid'{closed_key}' "
            f"and ПодразделениеКомпании_Key eq guid'{division_key}' "
            f"and {_range_interval('ДатаЗакрытия', start, end)}"
        )
        return self.client.get(
            "Document_ЗаказНаряд",
            params=build_params(
                filter=f,
                select=(
                    "Ref_Key,Number,Date,ДатаНачала,ДатаЗакрытия,СуммаДокумента,"
                    "Состояние_Key,ПодразделениеКомпании_Key"
                ),
            ),
        )

    def closed_orders(self, division_key: str, day: date) -> list[dict]:
        """Шапки закрытых за день ЗН — период [day, day+1)."""
        return self.closed_orders_period(division_key, day, day + timedelta(days=1))

    def works(self, order_keys: list[str]) -> list[dict]:
        """Строки Работы для набора ЗН (батчами по Ref_Key)."""
        if not order_keys:
            return []
        rows: list[dict] = []
        for batch in _batched(order_keys):
            f = _guid_or("Ref_Key", batch)
            rows.extend(
                self.client.get(
                    "Document_ЗаказНаряд_Работы",
                    params=build_params(
                        filter=f,
                        select="Ref_Key,Количество,Коэффициент",
                    ),
                )
            )
        return rows

    def executors(self, order_keys: list[str]) -> list[dict]:
        """Строки Исполнители для набора ЗН (батчами по Ref_Key)."""
        if not order_keys:
            return []
        rows: list[dict] = []
        for batch in _batched(order_keys):
            f = _guid_or("Ref_Key", batch)
            rows.extend(
                self.client.get(
                    "Document_ЗаказНаряд_Исполнители",
                    params=build_params(
                        filter=f,
                        select="Ref_Key,Исполнитель_Key",
                    ),
                )
            )
        return rows

    # Терминальные статусы: ЗН считается «завершённым» и НЕ висит в работе.
    # Исключаются из «незакрытых старше 1 дня». «Выполнен» намеренно НЕ включён
    # (выполнен, но не закрыт/не выдан — легитимный флаг). Подтвердить у владельца.
    TERMINAL_STATUSES = ("Закрыт", "Отказ", "Архив")

    def open_orders_older_than_1day_asof(self, division_key: str, as_of: date) -> int:
        """Количество незакрытых ЗН старше 1 дня по подразделению на дату as_of.

        Незакрытый = состояние НЕ в TERMINAL_STATUSES (Закрыт/Отказ/Архив).
        Старше 1 дня = Date < начало (as_of - 1) 00:00.
        """
        statuses = self.references.statuses()
        term_keys = [statuses[s] for s in self.TERMINAL_STATUSES if s in statuses]
        if not term_keys:
            raise RuntimeError("Не найдены терминальные статусы в References.statuses()")
        cutoff = as_of - timedelta(days=1)  # ЗН старше 1 дня: Date < (as_of-1) 00:00
        ne_clause = " and ".join(f"Состояние_Key ne guid'{k}'" for k in term_keys)
        f = (
            f"ПодразделениеКомпании_Key eq guid'{division_key}' "
            f"and {ne_clause} "
            f"and Date lt {_dt(cutoff)}"
        )
        rows = self.client.get(
            "Document_ЗаказНаряд",
            params=build_params(filter=f, select="Ref_Key"),
        )
        return len(rows)

    def open_orders_older_than_1day(self, division_key: str, day: date) -> int:
        """Незакрытые ЗН старше 1 дня на день day (обёртка over asof)."""
        return self.open_orders_older_than_1day_asof(division_key, day)

    # ---------------------------------------------------------------- ПАЙПЛАЙН
    # Терминальные статусы для пайплайна длинных ремонтов (Шаркер/ЦКР).
    # Для просроченных ЗН ТЗ явно требует исключить «Закрыт» и «Отказ»;
    # «Архив» тоже оставляем исключённым как терминальное состояние.
    PIPELINE_TERMINAL = ("Закрыт", "Отказ", "Архив")
    AWAITING_PARTS_STATUSES = ("Ожидание автозапчастей клиента", "Ожидание деталей")

    def _terminal_status_keys(self) -> list[str]:
        statuses = self.references.statuses()
        keys = [statuses[s] for s in self.PIPELINE_TERMINAL if s in statuses]
        if not keys:
            raise RuntimeError("Не найдены терминальные статусы в References.statuses()")
        return keys

    def _not_terminal_clause(self) -> str:
        """'Состояние_Key ne guid'..' and ...' по терминальным статусам."""
        return " and ".join(
            f"Состояние_Key ne guid'{k}'" for k in self._terminal_status_keys()
        )

    def _count(self, filter_str: str) -> int:
        rows = self.client.get(
            "Document_ЗаказНаряд",
            params=build_params(filter=filter_str, select="Ref_Key"),
        )
        return len(rows)

    def pipeline_active(self, division_key: str, day: date) -> int:
        """Активных ЗН = НЕ терминальные И НЕ «Заявка» (значение 1С на дату day)."""
        statuses = self.references.statuses()
        zayavka = statuses.get("Заявка")
        f = (
            f"ПодразделениеКомпании_Key eq guid'{division_key}' "
            f"and {self._not_terminal_clause()}"
        )
        if zayavka:
            f += f" and Состояние_Key ne guid'{zayavka}'"
        return self._count(f)

    def pipeline_ready_today(self, division_key: str, day: date) -> int:
        """Готовы к выдаче сегодня = НЕ терминальные И ПлановаяДатаВыдачи in [day, day+1)."""
        f = (
            f"ПодразделениеКомпании_Key eq guid'{division_key}' "
            f"and {self._not_terminal_clause()} "
            f"and {_day_interval('ПлановаяДатаВыдачи', day)}"
        )
        return self._count(f)

    def pipeline_awaiting_parts(self, division_key: str, day: date) -> int:
        """Ожидают ЗЧ = статус in {Ожидание автозапчастей клиента, Ожидание деталей}."""
        statuses = self.references.statuses()
        keys = [statuses[s] for s in self.AWAITING_PARTS_STATUSES if s in statuses]
        if not keys:
            return 0
        or_clause = " or ".join(f"Состояние_Key eq guid'{k}'" for k in keys)
        f = (
            f"ПодразделениеКомпании_Key eq guid'{division_key}' "
            f"and ({or_clause})"
        )
        return self._count(f)

    def pipeline_overdue(self, division_key: str, day: date) -> int:
        """Просрочены = НЕ Закрыт/Отказ/Архив И ПлановаяДатаВыдачи > 1901 И < day.

        1С OData НЕ поддерживает 'ne null' — пустые даты отсекаем нижней
        границей 'gt datetime'1901-01-01T00:00:00''.
        """
        f = (
            f"ПодразделениеКомпании_Key eq guid'{division_key}' "
            f"and {self._not_terminal_clause()} "
            f"and ПлановаяДатаВыдачи gt datetime'1901-01-01T00:00:00' "
            f"and ПлановаяДатаВыдачи lt {_dt(day)}"
        )
        return self._count(f)

    def pipeline_planned_close_week(self, division_key: str, day: date) -> int:
        """Планируется закрыть на неделе = НЕ терминальные И ПлановаяДатаВыдачи
        in [понедельник, следующий понедельник) недели day."""
        week_start = day - timedelta(days=day.weekday())
        week_end = week_start + timedelta(days=7)
        f = (
            f"ПодразделениеКомпании_Key eq guid'{division_key}' "
            f"and {self._not_terminal_clause()} "
            f"and {_range_interval('ПлановаяДатаВыдачи', week_start, week_end)}"
        )
        return self._count(f)

    @staticmethod
    def normhours(work_rows: list[dict]) -> float:
        """Нормочасы = Σ Количество × Коэффициент по строкам Работы."""
        return sum(_num(r.get("Количество")) * _num(r.get("Коэффициент")) for r in work_rows)

    @staticmethod
    def unique_executors(executor_rows: list[dict]) -> int:
        """Кол-во уникальных Исполнитель_Key."""
        keys = {r.get("Исполнитель_Key") for r in executor_rows if r.get("Исполнитель_Key")}
        return len(keys)


class IncomeExpenseRepository:
    """Регистр AccumulationRegister_ДоходыИРасходы_RecordType."""

    # статьи выручки (ДоходБезНДС) и себестоимости (РасходБезНДС)
    PARTS_ARTICLE = "Выручка по заказ-нарядам"   # запчасти
    WORKS_ARTICLE = "Работы по заказ-нарядам"    # труд/услуги
    COST_ARTICLE = "Себестоимость материалов"

    REVENUE_ARTICLES = (PARTS_ARTICLE, WORKS_ARTICLE)
    COST_ARTICLES = (COST_ARTICLE,)

    def __init__(self, client: ODataClient, references: References) -> None:
        self.client = client
        self.references = references

    def _rows_period(self, division_key: str, start: date, end: date) -> list[dict]:
        # Active eq true — исключаем сторнированные/неактивные движения регистра.
        f = (
            f"ПодразделениеКомпании_Key eq guid'{division_key}' "
            f"and {_range_interval('Period', start, end)} "
            f"and Active eq true"
        )
        return self.client.get(
            "AccumulationRegister_ДоходыИРасходы_RecordType",
            params=build_params(
                filter=f,
                select="Period,ПодразделениеКомпании_Key,СтатьяДоходовИРасходов_Key,ДоходБезНДС,РасходБезНДС,Active",
            ),
        )

    def _article_keys(self, names: Iterable[str]) -> set[str]:
        articles = self.references.income_expense_articles()
        keys = set()
        for n in names:
            k = articles.get(n)
            if k:
                keys.add(k)
        return keys

    def revenue_breakdown_period(
        self, division_key: str, start: date, end: date
    ) -> dict[str, float]:
        """Раздельная выручка/себестоимость за период [start, end).

        Возвращает {"parts": .., "works": .., "cost": ..}:
        - parts = Σ ДоходБезНДС по «Выручка по заказ-нарядам» (запчасти);
        - works = Σ ДоходБезНДС по «Работы по заказ-нарядам» (труд/услуги);
        - cost  = Σ РасходБезНДС по «Себестоимость материалов».
        """
        rows = self._rows_period(division_key, start, end)
        parts_keys = self._article_keys((self.PARTS_ARTICLE,))
        works_keys = self._article_keys((self.WORKS_ARTICLE,))
        cost_keys = self._article_keys(self.COST_ARTICLES)
        parts = works = cost = 0.0
        for r in rows:
            art = r.get("СтатьяДоходовИРасходов_Key")
            if art in parts_keys:
                parts += _num(r.get("ДоходБезНДС"))
            if art in works_keys:
                works += _num(r.get("ДоходБезНДС"))
            if art in cost_keys:
                cost += _num(r.get("РасходБезНДС"))
        return {"parts": parts, "works": works, "cost": cost}

    def revenue_breakdown(self, division_key: str, day: date) -> dict[str, float]:
        """Раздельная выручка/себестоимость за день — период [day, day+1)."""
        return self.revenue_breakdown_period(division_key, day, day + timedelta(days=1))

    def revenue_and_cost_period(
        self, division_key: str, start: date, end: date
    ) -> tuple[float, float]:
        """(полная выручка = parts+works, себестоимость) за период [start, end)."""
        b = self.revenue_breakdown_period(division_key, start, end)
        return b["parts"] + b["works"], b["cost"]

    def revenue_and_cost(self, division_key: str, day: date) -> tuple[float, float]:
        """(полная выручка = parts+works, себестоимость) за день."""
        return self.revenue_and_cost_period(division_key, day, day + timedelta(days=1))


class InsuranceRepository:
    """Неврученные к оплате страховые ЗН прошлых месяцев."""

    SELECTION_START = date(2025, 12, 30)
    REPAIR_TYPES = (
        "Страховой ремонт термофургонов",
        "Кузовной страховой ремонт",
    )
    INVOICE_DOC = "Document_СчетФактураВыданный"
    INVOICE_ORDER_TYPE = "StandardODATA.Document_ЗаказНаряд"
    ORIGINALS_REGISTER = "InformationRegister_олОригиналДокументаПолучен"
    INVOICE_ORIGINAL_TYPE = "StandardODATA.Document_СчетФактураВыданный"
    DELIVERY_DATE_FIELD = "Дата2"
    PAYMENT_REGISTER = "AccumulationRegister_ВзаиморасчетыКомпании_RecordType"
    PAYMENT_ORDER_TYPE = "StandardODATA.Document_ЗаказНаряд"
    PAYMENT_RECORDER_TYPE = "StandardODATA.Document_Выписка"
    PAYMENT_OPERATION = "ПогашениеДебиторскойЗадолженности"
    PAYMENT_TOLERANCE = 1.0

    def __init__(self, client: ODataClient, references: References) -> None:
        self.client = client
        self.references = references
        self._repair_keys: Optional[set[str]] = None
        self._invoice_cache: dict[tuple[str, date], dict[str, list[dict]]] = {}
        self._delivered_cache: dict[date, set[str]] = {}
        self._payment_cache: dict[date, dict[str, float]] = {}

    def _repair_type_keys(self) -> set[str]:
        if self._repair_keys is not None:
            return self._repair_keys
        rows = self.client.get(
            "Catalog_ВидыРемонта",
            params=build_params(select="Ref_Key,Description"),
        )
        wanted = {name.casefold() for name in self.REPAIR_TYPES}
        self._repair_keys = {
            r["Ref_Key"]
            for r in rows
            if r.get("Ref_Key") and (r.get("Description") or "").casefold() in wanted
        }
        return self._repair_keys

    @staticmethod
    def _selection_end(day: date) -> date:
        return day + timedelta(days=1)

    def _closed_insurance_orders(self, division_key: str, as_of: date) -> list[dict]:
        closed_key = self.references.statuses().get("Закрыт")
        if not closed_key:
            raise RuntimeError("Не найден статус 'Закрыт' в References.statuses()")
        repair_keys = self._repair_type_keys()
        if not repair_keys:
            return []
        repair_clause = " or ".join(f"ВидРемонта_Key eq guid'{k}'" for k in repair_keys)
        f = (
            f"Состояние_Key eq guid'{closed_key}' "
            f"and ПодразделениеКомпании_Key eq guid'{division_key}' "
            f"and ДатаЗакрытия ge {_dt(self.SELECTION_START)} "
            f"and ДатаЗакрытия lt {_dt(self._selection_end(as_of))} "
            f"and ({repair_clause})"
        )
        return self.client.get(
            "Document_ЗаказНаряд",
            params=build_params(
                filter=f,
                select="Ref_Key,Number,ДатаЗакрытия,ВидРемонта_Key,СуммаДокумента",
            ),
        )

    def _invoices_by_order(self, division_key: str, as_of: date) -> dict[str, list[dict]]:
        cache_key = (division_key, as_of)
        if cache_key in self._invoice_cache:
            return self._invoice_cache[cache_key]
        # Поле ДокументОснование в 1С OData имеет строковый тип неограниченной длины:
        # фильтр по нему либо не работает, либо не возвращает строки. Поэтому читаем
        # счет-фактуры подразделения до даты отчета и сопоставляем с ЗН в Python.
        f = (
            f"ПодразделениеКомпании_Key eq guid'{division_key}' "
            f"and Date lt {_dt(as_of + timedelta(days=1))} "
            f"and ДокументОснование_Type eq '{self.INVOICE_ORDER_TYPE}'"
        )
        rows = self.client.get(
            self.INVOICE_DOC,
            params=build_params(
                filter=f,
                select=(
                    "Ref_Key,Number,Date,ДокументОснование,ДокументОснование_Type,"
                    "СуммаДокумента,Выставлен"
                ),
            ),
        )
        result: dict[str, list[dict]] = {}
        for row in rows:
            order_key = row.get("ДокументОснование")
            if order_key:
                result.setdefault(order_key, []).append(row)
        self._invoice_cache[cache_key] = result
        return result

    def _delivered_invoice_keys(self, as_of: date) -> set[str]:
        if as_of in self._delivered_cache:
            return self._delivered_cache[as_of]
        # Регистр соответствует форме «Оригинал документа получен». Дата2 — дата
        # строки «Получен КА», показанная в форме на скриншоте.
        f = (
            f"ДокументСсылка_Type eq '{self.INVOICE_ORIGINAL_TYPE}' "
            "and ОригиналПолучен eq true "
            "and Дата2 gt datetime'1901-01-01T00:00:00' "
            f"and Дата2 lt {_dt(as_of + timedelta(days=1))}"
        )
        rows = self.client.get(
            self.ORIGINALS_REGISTER,
            params=build_params(
                filter=f,
                select=(
                    "ДокументСсылка,ДокументСсылка_Type,ОригиналПолучен,"
                    f"Дата,{self.DELIVERY_DATE_FIELD}"
                ),
            ),
        )
        delivered = {
            row["ДокументСсылка"]
            for row in rows
            if row.get("ДокументСсылка")
            and row.get("ОригиналПолучен") is True
            and _has_filled_date(row.get(self.DELIVERY_DATE_FIELD))
        }
        self._delivered_cache[as_of] = delivered
        return delivered

    @staticmethod
    def _payable_amount(order: dict, invoices: list[dict]) -> float:
        order_sum = _num(order.get("СуммаДокумента"))
        # У одного ЗН могут быть повторные/переоформленные счет-фактуры на ту же
        # сумму. Суммировать их нельзя: это задваивает ожидаемую оплату.
        invoice_sum = max((_num(inv.get("СуммаДокумента")) for inv in invoices), default=0.0)
        return max(order_sum, invoice_sum)

    def _paid_by_order(self, as_of: date) -> dict[str, float]:
        if as_of in self._payment_cache:
            return self._payment_cache[as_of]
        # Оплаты страховых ЗН проходят движениями банковской выписки по регистру
        # взаиморасчетов. Сделка хранит ссылку на ЗН, но поле строковое
        # неограниченной длины, поэтому сопоставление с нужными ЗН делаем в Python.
        f = (
            f"Recorder_Type eq '{self.PAYMENT_RECORDER_TYPE}' "
            f"and Сделка_Type eq '{self.PAYMENT_ORDER_TYPE}' "
            f"and ВидОперации eq '{self.PAYMENT_OPERATION}' "
            "and RecordType eq 'Expense' "
            f"and Period lt {_dt(as_of + timedelta(days=1))}"
        )
        rows = self.client.get(
            self.PAYMENT_REGISTER,
            params=build_params(
                filter=f,
                select=(
                    "Period,Recorder,Recorder_Type,RecordType,Сделка,Сделка_Type,"
                    "Сумма,ВидОперации,ДоговорВзаиморасчетов_Key"
                ),
            ),
        )
        paid: dict[str, float] = {}
        for row in rows:
            order_key = row.get("Сделка")
            if not order_key:
                continue
            paid[order_key] = paid.get(order_key, 0.0) + _num(row.get("Сумма"))
        self._payment_cache[as_of] = paid
        return paid

    def undelivered_previous_months(self, division_key: str, as_of: date) -> dict:
        """Возвращает count/sum и отладку по страховым ЗН прошлых месяцев.

        Поле «Дата вручения КА» хранится в регистре «Оригинал документа
        получен» как Дата2 по ссылке на счет-фактуру. Пустая дата 0001-01-01
        считается незаполненной. Полностью оплаченные ЗН исключаются, частично
        оплаченные остаются с суммой неоплаченного остатка.
        """
        orders = self._closed_insurance_orders(division_key, as_of)
        invoices = self._invoices_by_order(division_key, as_of)
        delivered_invoice_keys = self._delivered_invoice_keys(as_of)
        paid_by_order = self._paid_by_order(as_of)
        undelivered: list[dict] = []
        delivered = 0
        with_invoice = 0
        without_invoice = 0
        without_delivery_date = 0
        fully_paid = 0
        partially_paid = 0
        unpaid_sum = 0.0

        for order in orders:
            order_key = order.get("Ref_Key")
            related = invoices.get(order_key, [])
            payable = self._payable_amount(order, related)
            paid = paid_by_order.get(order_key, 0.0)
            unpaid = max(payable - paid, 0.0)
            if not related:
                without_invoice += 1
            else:
                with_invoice += 1
                if any(inv.get("Ref_Key") in delivered_invoice_keys for inv in related):
                    delivered += 1
                    continue
                without_delivery_date += 1

            if unpaid <= self.PAYMENT_TOLERANCE:
                fully_paid += 1
                continue
            if paid > self.PAYMENT_TOLERANCE:
                partially_paid += 1
            enriched = dict(order)
            enriched["_payable_sum"] = payable
            enriched["_paid_sum"] = paid
            enriched["_unpaid_sum"] = unpaid
            undelivered.append(enriched)
            unpaid_sum += unpaid

        return {
            "count": len(undelivered),
            "sum": unpaid_sum,
            "debug": {
                "as_of": as_of.isoformat(),
                "selection_start": self.SELECTION_START.isoformat(),
                "selection_end_exclusive": self._selection_end(as_of).isoformat(),
                "repair_type_keys": sorted(self._repair_type_keys()),
                "candidate_orders": len(orders),
                "invoice_order_links": len(invoices),
                "delivered_invoice_links": len(delivered_invoice_keys),
                "payment_order_links": len(paid_by_order),
                "with_invoice": with_invoice,
                "delivered": delivered,
                "without_invoice": without_invoice,
                "without_delivery_date": without_delivery_date,
                "fully_paid_excluded": fully_paid,
                "partially_paid": partially_paid,
                "delivery_date_field": self.DELIVERY_DATE_FIELD,
                "sample": [
                    {
                        "number": order.get("Number"),
                        "ref": order.get("Ref_Key"),
                        "closed_at": order.get("ДатаЗакрытия"),
                        "sum": order.get("СуммаДокумента"),
                        "payable": order.get("_payable_sum"),
                        "paid": order.get("_paid_sum"),
                        "unpaid": order.get("_unpaid_sum"),
                        "invoice_count": len(invoices.get(order.get("Ref_Key"), [])),
                    }
                    for order in undelivered[:10]
                ],
            },
        }


class PlanRepository:
    """Плановые показатели: последний документ месяца + строки Показатели."""

    DOC = "Document_ПрибылиУбытки_УстановкаЗначенийПлановыхПоказателей"
    ROWS = "Document_ПрибылиУбытки_УстановкаЗначенийПлановыхПоказателей_Показатели"

    def __init__(self, client: ODataClient, references: References) -> None:
        self.client = client
        self.references = references
        self._month_docs_cache: dict[tuple[str, date], list[dict]] = {}
        self._day_docs_cache: dict[tuple[str, date], list[dict]] = {}
        self._indicator_values_cache: dict[str, dict[str, float]] = {}

    def _month_bounds(self, day: date) -> tuple[date, date]:
        start = day.replace(day=1)
        if start.month == 12:
            nxt = start.replace(year=start.year + 1, month=1)
        else:
            nxt = start.replace(month=start.month + 1)
        return start, nxt

    def documents_month(self, division_key: str, day: date) -> list[dict]:
        """Документы подразделения с ДатаФиксации в месяце day, новые первыми."""
        cache_key = (division_key, day.replace(day=1))
        if cache_key in self._month_docs_cache:
            return self._month_docs_cache[cache_key]
        start, nxt = self._month_bounds(day)
        f = (
            f"ПодразделениеКомпании_Key eq guid'{division_key}' "
            f"and ДатаФиксации ge {_dt(start)} and ДатаФиксации lt {_dt(nxt)}"
        )
        rows = self.client.get(
            self.DOC,
            params=build_params(
                filter=f,
                select="Ref_Key,Date,ДатаФиксации,ПодразделениеКомпании_Key",
                orderby="Date desc",
            ),
        )
        self._month_docs_cache[cache_key] = rows
        return rows

    def documents_day(self, division_key: str, day: date) -> list[dict]:
        """Документы подразделения с ДатаФиксации ровно в день day, новые первыми."""
        cache_key = (division_key, day)
        if cache_key in self._day_docs_cache:
            return self._day_docs_cache[cache_key]
        f = (
            f"ПодразделениеКомпании_Key eq guid'{division_key}' "
            f"and ДатаФиксации ge {_dt(day)} and ДатаФиксации lt {_dt(day + timedelta(days=1))}"
        )
        rows = self.client.get(
            self.DOC,
            params=build_params(
                filter=f,
                select="Ref_Key,Date,ДатаФиксации,ПодразделениеКомпании_Key",
                orderby="Date desc",
            ),
        )
        self._day_docs_cache[cache_key] = rows
        return rows

    def latest_document(self, division_key: str, day: date) -> Optional[dict]:
        """Последний по Date документ подразделения, чей ДатаФиксации в месяце дня."""
        rows = self.documents_month(division_key, day)
        return rows[0] if rows else None

    def indicator_values(self, doc_key: str) -> dict[str, float]:
        """Code показателя -> ЗначениеПоказателя для документа плана."""
        if doc_key in self._indicator_values_cache:
            return self._indicator_values_cache[doc_key]
        rows = self.client.get(
            self.ROWS,
            params=build_params(
                filter=f"Ref_Key eq guid'{doc_key}'",
                select="Ref_Key,Показатель,ЗначениеПоказателя",
            ),
        )
        code_by_ref = {ref: code for code, ref in self.references.plan_indicators().items()}
        result: dict[str, float] = {}
        for r in rows:
            ref = r.get("Показатель")
            code = code_by_ref.get(ref)
            if code:
                result[code] = _num(r.get("ЗначениеПоказателя"))
        self._indicator_values_cache[doc_key] = result
        return result

    def monthly_values(self, division_key: str, day: date) -> dict[str, float]:
        """Code -> последнее значение за месяц по каждому показателю.

        В 1С могут быть отдельные дневные документы факта с тем же месяцем
        фиксации, но без строк месячного плана. Поэтому собираем значения
        по всем документам месяца, новыми первыми: для каждого Code берётся
        первая найденная строка.
        """
        result: dict[str, float] = {}
        for doc in self.documents_month(division_key, day):
            for code, value in self.indicator_values(doc["Ref_Key"]).items():
                result.setdefault(code, value)
        return result

    def daily_values(self, division_key: str, day: date) -> dict[str, float]:
        """Code -> последнее значение за конкретную ДатаФиксации day."""
        result: dict[str, float] = {}
        for doc in self.documents_day(division_key, day):
            for code, value in self.indicator_values(doc["Ref_Key"]).items():
                result.setdefault(code, value)
        return result


class PaymentRepository:
    """Поступления оплат покупателей через договор -> подразделение.

    Оплата = Document_Выписка со статьёй ДДС «Оплата от покупателя», чей
    ДоговорВзаиморасчетов_Key принадлежит подразделению. НЕ включает
    «Предоплата от покупателя» (отдельная статья).
    """

    PAYMENT_ARTICLE = "Оплата от покупателя"

    def __init__(self, client: ODataClient, references: References) -> None:
        self.client = client
        self.references = references

    def payments_period(self, division_key: str, start: date, end: date) -> float:
        """Сумма СуммаДокументаПриход оплат подразделения за [start, end).

        Выписки фильтруются в 1С по статье и дате; принадлежность договору
        подразделению проверяется в Python по множеству договоров (НЕ строим
        огромный OR).
        """
        article_key = self.references.dds_article(self.PAYMENT_ARTICLE)
        if not article_key:
            raise RuntimeError(
                f"Не найдена статья ДДС '{self.PAYMENT_ARTICLE}' в References.dds_article()"
            )
        contracts = set(self.references.contracts_by_division(division_key))
        if not contracts:
            return 0.0
        f = (
            f"СтатьяДДС_Key eq guid'{article_key}' "
            f"and {_range_interval('Date', start, end)}"
        )
        rows = self.client.get(
            "Document_Выписка",
            params=build_params(
                filter=f,
                select="Ref_Key,Date,СтатьяДДС_Key,ДоговорВзаиморасчетов_Key,СуммаДокументаПриход",
            ),
        )
        total = 0.0
        for r in rows:
            if r.get("ДоговорВзаиморасчетов_Key") in contracts:
                total += _num(r.get("СуммаДокументаПриход"))
        return total

    def payments_day(self, division_key: str, day: date) -> float:
        """Сумма оплат за день — период [day, day+1)."""
        return self.payments_period(division_key, day, day + timedelta(days=1))
