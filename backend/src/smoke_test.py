"""Smoke-тест Инкремента 1: 5 критериев приёмки из ARCHITECTURE.md.

Возвращает код 0 при успехе, ненулевой — при недоступности OData.
"""
from __future__ import annotations

import sys
from datetime import datetime, timedelta

try:
    from .config import settings
    from .odata_client import (
        ODataClient,
        ODataError,
        ODataUnavailableError,
        build_params,
    )
    from .references import References, REPORT_TO_DIVISION_DESCRIPTION
except ImportError:  # запуск как скрипта: python src/smoke_test.py
    from config import settings
    from odata_client import (
        ODataClient,
        ODataError,
        ODataUnavailableError,
        build_params,
    )
    from references import References, REPORT_TO_DIVISION_DESCRIPTION


LINE = "=" * 68


def _hdr(text: str) -> None:
    print(f"\n{text}")
    print("-" * len(text))


def main() -> int:
    print(LINE)
    print("SMOKE-ТЕСТ «Отчёты АТЦ» — Инкремент 1 (фундамент)")
    print(LINE)
    print(f"Base URL : {settings.odata_base_url}")
    print(f"Login    : {settings.odata_login} (кодировка UTF-8)")
    print(f"Timeout  : {settings.timeout}s, Retries: {settings.retries}")

    passed: list[str] = []
    failed: list[str] = []

    try:
        client = ODataClient(settings)
        refs = References(client)

        # --- Критерий 1: подключение + число EntitySet ---
        _hdr("1) Подключение и число EntitySet ($metadata)")
        count = client.metadata_entityset_count()
        print(f"   EntitySet в составе: {count}")
        if count > 0:
            passed.append("1. Подключение / число EntitySet")
        else:
            failed.append("1. Подключение / число EntitySet (0 сущностей)")

        # --- Критерий 2: подразделения, проверка 6 названий ---
        _hdr("2) Catalog_ПодразделенияКомпании — проверка маппинга (6 названий)")
        divisions = refs.divisions()
        print(f"   Всего подразделений: {len(divisions)}")
        missing: list[str] = []
        for report_name, descr in REPORT_TO_DIVISION_DESCRIPTION.items():
            # division_key сопоставляет Description без учёта регистра
            key = refs.division_key(report_name)
            mark = "OK" if key else "НЕ НАЙДЕНО"
            print(f"   [{mark:>10}] {report_name!r} -> '{descr}'"
                  + (f" -> {key}" if key else ""))
            if not key:
                missing.append(f"{report_name} ('{descr}')")
        if not missing:
            passed.append("2. Все 6 подразделений найдены")
        else:
            failed.append("2. Не найдены подразделения: " + ", ".join(missing))

        # --- Критерий 3: статусы ЗН ---
        _hdr("3) Catalog_ВидыСостоянийЗаказНарядов — список статусов")
        statuses = refs.statuses()
        print(f"   Всего статусов: {len(statuses)}")
        for name in sorted(statuses):
            print(f"   - {name}")
        if statuses:
            passed.append("3. Статусы прочитаны")
        else:
            failed.append("3. Список статусов пуст")

        # --- Критерий 4: ЗН за последние 30 дней ---
        _hdr("4) Document_ЗаказНаряд — последние 30 дней")
        since = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%dT00:00:00")
        # OData 1С: литерал datetime, без строковых кавычек
        date_filter = f"Date ge datetime'{since}'"
        print(f"   Фильтр: {date_filter}")
        orders = client.get(
            "Document_ЗаказНаряд",
            params=build_params(
                filter=date_filter,
                orderby="Date desc",
                select=(
                    "Ref_Key,Number,Date,СуммаДокумента,"
                    "ПодразделениеКомпании_Key,Состояние_Key"
                ),
            ),
        )
        print(f"   Найдено заказ-нарядов: {len(orders)}")
        if orders:
            div_by_key = {v: k for k, v in divisions.items()}
            st_by_key = {v: k for k, v in statuses.items()}
            ex = orders[0]
            div_name = div_by_key.get(ex.get("ПодразделениеКомпании_Key"), "?")
            st_name = st_by_key.get(ex.get("Состояние_Key"), "?")
            print("   Пример:")
            print(f"     Номер       : {ex.get('Number')}")
            print(f"     Дата        : {ex.get('Date')}")
            print(f"     Подразделение: {div_name}")
            print(f"     Статус      : {st_name}")
            print(f"     Сумма       : {ex.get('СуммаДокумента')}")
            passed.append("4. ЗН за 30 дней прочитаны (есть пример)")
        else:
            # запрос прошёл, но данных нет — не считаем провалом чтения
            print("   За последние 30 дней заказ-нарядов не найдено "
                  "(запрос выполнен успешно).")
            passed.append("4. Запрос ЗН выполнен (0 записей за период)")

    except ODataUnavailableError as exc:
        # --- Критерий 5: недоступность порта ---
        print("\n" + LINE)
        print("ODATA НЕДОСТУПЕН")
        print(LINE)
        print(f"Причина: {exc}")
        print("Вероятно, порт 8888 закрыт из-за IP-allowlist. "
              "Клиент корректно обработал ошибку и вернул ненулевой код.")
        return 2
    except ODataError as exc:
        print("\n" + LINE)
        print("ОШИБКА ODATA (не сетевая)")
        print(LINE)
        print(f"Причина: {exc}")
        return 3

    # --- Итог ---
    print("\n" + LINE)
    print("ИТОГ ПО КРИТЕРИЯМ ПРИЁМКИ")
    print(LINE)
    for p in passed:
        print(f"  [PASS] {p}")
    for f in failed:
        print(f"  [FAIL] {f}")
    print(LINE)

    if failed:
        print("Результат: часть проверок не пройдена.")
        return 1
    print("Результат: все проверки пройдены.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
