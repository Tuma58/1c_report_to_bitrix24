"""Regression tests for undelivered insurance orders payment filtering."""
from __future__ import annotations

import os
import sys
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from repositories import InsuranceRepository  # noqa: E402


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
    paid = {
        "invoice-paid": 1990,
        "invoice-partial": 1000,
        "invoice-delivered": 0,
        "no-invoice-paid": 5000,
        "duplicate-invoices-paid": 6000,
    }

    repo._closed_insurance_orders = lambda _division, _as_of: orders  # type: ignore[method-assign]
    repo._invoices_by_order = lambda _division, _as_of: invoices  # type: ignore[method-assign]
    repo._delivered_invoice_keys = lambda _as_of: {"sf-4"}  # type: ignore[method-assign]
    repo._paid_by_order = lambda _as_of: paid  # type: ignore[method-assign]
    repo._repair_type_keys = lambda: {"repair-type"}  # type: ignore[method-assign]

    result = repo.undelivered_previous_months("division", as_of)

    assert result["count"] == 2, result
    assert result["sum"] == 3000, result
    assert result["debug"]["fully_paid_excluded"] == 3, result["debug"]
    assert result["debug"]["partially_paid"] == 1, result["debug"]

    sample = {item["number"]: item for item in result["debug"]["sample"]}
    assert sample["ZN-1"]["unpaid"] == 1000
    assert sample["ZN-3"]["paid"] == 1000
    assert sample["ZN-3"]["unpaid"] == 2000

    print("insurance repository payment filtering ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
