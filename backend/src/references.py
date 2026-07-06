"""Справочники 1С: подразделения и статусы ЗН.

Кэширует map Description -> Ref_Key. Содержит маппинг названий отчёта на
Description подразделений в 1С (см. ARCHITECTURE.md).
"""
from __future__ import annotations

from typing import Optional

try:
    from .odata_client import ODataClient, build_params
except ImportError:  # запуск как скрипта
    from odata_client import ODataClient, build_params


# Маппинг: Название в отчёте -> Description в Catalog_ПодразделенияКомпании
# (по ARCHITECTURE.md). Сопоставление регистронезависимо: в реальной базе
# встречается расхождение регистра (например, 'ТО и Ремонт' vs 'ТО и ремонт').
REPORT_TO_DIVISION_DESCRIPTION: dict[str, str] = {
    "Арсенал": "ТО и Ремонт",  # точное написание из базы; НЕ одноимённое пустое подразд. «Арсенал»
    "Реф. Сервис": "Рефика Сервис",
    "Шаркер": "Шаркер Трейлер",
    "ЦКР": "Кузовной",
    "Мойка на ул. Строителей": "Автомойка",
    "Мойка на ул. Ульяновская": "Мойка Ульяновская",
}


class References:
    def __init__(self, client: ODataClient) -> None:
        self.client = client
        self._divisions: Optional[dict[str, str]] = None
        self._statuses: Optional[dict[str, str]] = None
        self._articles: Optional[dict[str, str]] = None
        self._plan_indicators: Optional[dict[str, str]] = None
        self._dds_articles: Optional[dict[str, str]] = None
        self._contracts_by_division: dict[str, list[str]] = {}

    def divisions(self, refresh: bool = False) -> dict[str, str]:
        """Description -> Ref_Key для Catalog_ПодразделенияКомпании."""
        if self._divisions is None or refresh:
            rows = self.client.get(
                "Catalog_ПодразделенияКомпании",
                params=build_params(select="Ref_Key,Description"),
            )
            self._divisions = {
                r["Description"]: r["Ref_Key"]
                for r in rows
                if r.get("Description")
            }
        return self._divisions

    def division_key(self, report_name: str) -> Optional[str]:
        """Ref_Key подразделения по названию из отчёта. None, если не найдено.

        Сопоставление Description выполняется без учёта регистра, чтобы
        пережить расхождения регистра в данных 1С.
        """
        description = REPORT_TO_DIVISION_DESCRIPTION.get(report_name)
        if description is None:
            return None
        divisions = self.divisions()
        # точное совпадение
        if description in divisions:
            return divisions[description]
        # регистронезависимое совпадение
        target = description.casefold()
        for descr, key in divisions.items():
            if descr.casefold() == target:
                return key
        return None

    def statuses(self, refresh: bool = False) -> dict[str, str]:
        """Description -> Ref_Key для Catalog_ВидыСостоянийЗаказНарядов."""
        if self._statuses is None or refresh:
            rows = self.client.get(
                "Catalog_ВидыСостоянийЗаказНарядов",
                params=build_params(select="Ref_Key,Description"),
            )
            self._statuses = {
                r["Description"]: r["Ref_Key"]
                for r in rows
                if r.get("Description")
            }
        return self._statuses

    def income_expense_articles(self, refresh: bool = False) -> dict[str, str]:
        """Description -> Ref_Key для Catalog_СтатьиДоходовИРасходов.

        Используется для резолва статей выручки/себестоимости в регистре
        AccumulationRegister_ДоходыИРасходы_RecordType.
        """
        if getattr(self, "_articles", None) is None or refresh:
            rows = self.client.get(
                "Catalog_СтатьиДоходовИРасходов",
                params=build_params(select="Ref_Key,Description"),
            )
            self._articles = {
                r["Description"]: r["Ref_Key"]
                for r in rows
                if r.get("Description")
            }
        return self._articles

    def plan_indicators(self, refresh: bool = False) -> dict[str, str]:
        """Code -> Ref_Key для Catalog_ПрибылиУбытки_ПлановыеПоказатели.

        Code хранится строкой с ведущими нулями (например '000000067').
        """
        if getattr(self, "_plan_indicators", None) is None or refresh:
            rows = self.client.get(
                "Catalog_ПрибылиУбытки_ПлановыеПоказатели",
                params=build_params(select="Ref_Key,Code,Description"),
            )
            self._plan_indicators = {
                r["Code"]: r["Ref_Key"]
                for r in rows
                if r.get("Code")
            }
        return self._plan_indicators

    def dds_articles(self, refresh: bool = False) -> dict[str, str]:
        """Description -> Ref_Key для Catalog_СтатьиДДС (статьи движения ДС)."""
        if getattr(self, "_dds_articles", None) is None or refresh:
            rows = self.client.get(
                "Catalog_СтатьиДДС",
                params=build_params(select="Ref_Key,Description"),
            )
            self._dds_articles = {
                r["Description"]: r["Ref_Key"]
                for r in rows
                if r.get("Description")
            }
        return self._dds_articles

    def dds_article(self, name: str) -> Optional[str]:
        """Ref_Key статьи ДДС по Description (напр. 'Оплата от покупателя').

        Точное совпадение, при промахе — регистронезависимое.
        """
        articles = self.dds_articles()
        if name in articles:
            return articles[name]
        target = name.casefold()
        for descr, key in articles.items():
            if descr.casefold() == target:
                return key
        return None

    def contracts_by_division(self, division_key: str) -> list[str]:
        """Список Ref_Key договоров взаиморасчётов подразделения (кэш по div).

        Catalog_ДоговорыВзаиморасчетов where Подразделение_Key eq guid'div'.
        """
        if division_key in self._contracts_by_division:
            return self._contracts_by_division[division_key]
        rows = self.client.get(
            "Catalog_ДоговорыВзаиморасчетов",
            params=build_params(
                filter=f"Подразделение_Key eq guid'{division_key}'",
                select="Ref_Key,Подразделение_Key",
            ),
        )
        keys = [r["Ref_Key"] for r in rows if r.get("Ref_Key")]
        self._contracts_by_division[division_key] = keys
        return keys
