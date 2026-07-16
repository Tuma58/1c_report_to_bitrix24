"""Regression test for payment sums from Document_Выписка.

No network calls: the OData client and references are in-memory fakes.
"""
from __future__ import annotations

import os
import sys
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from repositories import PaymentRepository  # noqa: E402


class FakeClient:
    def __init__(self) -> None:
        self.calls: list[tuple[str, dict]] = []

    def get(self, entity: str, params: dict | None = None) -> list[dict]:
        self.calls.append((entity, params or {}))
        assert entity == "Document_Выписка"
        select = (params or {}).get("$select", "")
        assert "СуммаДокументаПриход" in select
        assert "СуммаДокумента," not in select
        return [
            {
                "ДоговорВзаиморасчетов_Key": "contract-main",
                "СуммаДокументаПриход": 1250.75,
            },
            {
                "ДоговорВзаиморасчетов_Key": "contract-main",
                "СуммаДокументаПриход": "249.25",
            },
            {
                "ДоговорВзаиморасчетов_Key": "contract-other",
                "СуммаДокументаПриход": 999999,
            },
        ]


class FakeReferences:
    def dds_article(self, name: str) -> str:
        assert name == "Оплата от покупателя"
        return "article-payment"

    def contracts_by_division(self, division_key: str) -> list[str]:
        assert division_key == "division-main"
        return ["contract-main"]


def main() -> int:
    client = FakeClient()
    repo = PaymentRepository(client, FakeReferences())
    total = repo.payments_period("division-main", date(2026, 7, 1), date(2026, 7, 8))

    assert total == 1500.0, total
    entity, params = client.calls[0]
    assert entity == "Document_Выписка"
    assert "СтатьяДДС_Key eq guid'article-payment'" in params["$filter"]
    assert "Date ge datetime'2026-07-01T00:00:00'" in params["$filter"]
    assert "Date lt datetime'2026-07-08T00:00:00'" in params["$filter"]

    print("payment repository uses СуммаДокументаПриход ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
