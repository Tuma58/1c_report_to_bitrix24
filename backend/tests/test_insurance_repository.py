"""Regression tests for undelivered insurance orders payment filtering."""
from __future__ import annotations

import os
import sys
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from repositories import InsuranceRepository  # noqa: E402


class FilterClient:
    def __init__(self) -> None:
        self.filter = ""

    def get(self, entity: str, params: dict | None = None) -> list[dict]:
        assert entity == "Document_ЗаказНаряд"
        self.filter = (params or {}).get("$filter", "")
        return []


class BalanceClient:
    def __init__(self) -> None:
        self.entity = ""
        self.filter = ""

    def get(self, entity: str, params: dict | None = None) -> list[dict]:
        self.entity = entity
        self.filter = (params or {}).get("$filter", "")
        return [
            {"Сделка": "order-a", "СуммаBalance": 1200.0},
            {"Сделка": "order-a", "СуммаBalance": 300.0},
            {"Сделка": "order-b", "СуммаBalance": 0.5},
        ]


class FilterReferences:
    def statuses(self) -> dict[str, str]:
        return {"Закрыт": "closed-status"}


def main() -> int:
    repo = InsuranceRepository(client=None, references=None)  # type: ignore[arg-type]
    as_of = date(2026, 7, 10)
    orders = [
        {"Ref_Key": "no-invoice-open", "Number": "ZN-1", "СуммаДокумента": 1000},
        {"Ref_Key": "invoice-paid", "Number": "ZN-2", "СуммаДокумента": 1990},
        {"Ref_Key": "invoice-partial", "Number": "ZN-3", "СуммаДокумента": 3000},
        {"Ref_Key": "invoice-delivered", "Number": "ZN-4", "СуммаДокумента": 4000},
        {"Ref_Key": "no-invoice-paid", "Number": "ZN-5", "СуммаДокумента": 5000},
        {"Ref_Key": "duplicate-invoices-paid", "Number": "ZN-6", "СуммаДокумента": 6000},
    ]
    invoices = {
        "invoice-paid": [{"Ref_Key": "sf-2", "СуммаДокумента": 1990}],
        "invoice-partial": [{"Ref_Key": "sf-3", "СуммаДокумента": 3000}],
        "invoice-delivered": [{"Ref_Key": "sf-4", "СуммаДокумента": 4000}],
        "duplicate-invoices-paid": [
            {"Ref_Key": "sf-6-a", "СуммаДокумента": 6000},
            {"Ref_Key": "sf-6-b", "СуммаДокумента": 6000},
        ],
    }
    balances = {
        "no-invoice-open": 1000,
        "invoice-partial": 2000,
        "invoice-delivered": 4000,
    }

    repo._closed_insurance_orders = lambda _division, _as_of: orders  # type: ignore[method-assign]
    repo._invoices_by_order = lambda _division, _as_of: invoices  # type: ignore[method-assign]
    repo._delivered_invoice_keys = lambda _as_of: {"sf-4"}  # type: ignore[method-assign]
    repo._balance_by_order = lambda _as_of: balances  # type: ignore[method-assign]
    repo._repair_type_keys = lambda: {"repair-type"}  # type: ignore[method-assign]

    result = repo.undelivered_previous_months("division", as_of)

    assert result["count"] == 2, result
    assert result["sum"] == 3000, result
    assert result["debug"]["no_positive_balance_excluded"] == 3, result["debug"]

    sample = {item["number"]: item for item in result["debug"]["sample"]}
    assert sample["ZN-1"]["unpaid"] == 1000
    assert sample["ZN-3"]["unpaid"] == 2000

    filter_client = FilterClient()
    filter_repo = InsuranceRepository(filter_client, FilterReferences())  # type: ignore[arg-type]
    filter_repo._repair_type_keys = lambda: {"repair-type"}  # type: ignore[method-assign]
    filter_repo._closed_insurance_orders("division-main", as_of)
    assert "ДатаЗакрытия ge datetime'2025-12-30T00:00:00'" in filter_client.filter
    assert "ДатаЗакрытия lt datetime'2026-07-11T00:00:00'" in filter_client.filter

    balance_client = BalanceClient()
    balance_repo = InsuranceRepository(balance_client, FilterReferences())  # type: ignore[arg-type]
    balances = balance_repo._balance_by_order(as_of)
    assert balance_client.entity == (
        "AccumulationRegister_ВзаиморасчетыКомпании/Balance"
        "(Period=datetime'2026-07-11T00:00:00')"
    )
    assert "Сделка_Type eq 'StandardODATA.Document_ЗаказНаряд'" in balance_client.filter
    assert "СуммаBalance gt 1" in balance_client.filter
    assert balances == {"order-a": 1500.0}

    print("insurance repository payment filtering ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
