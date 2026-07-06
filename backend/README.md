# Отчёты АТЦ — бэкенд

Python-бэкенд для отчётов АТЦ из 1С OData. Проект читает данные из 1С
только через `GET`, считает показатели ПЛАН/ФАКТ и заполняет Excel-шаблон
`templates/report_template.xlsx`.

## Что уже реализовано

- Подключение к 1С OData с Basic Auth, UTF-8 логином, `$format=json`,
  таймаутами и ретраями.
- Кэшируемые справочники 1С: подразделения, статусы, статьи доходов/расходов,
  статьи ДДС, плановые показатели.
- Дневные и недельные показатели для `Арсенал` и `Реф. Сервис`.
- Дневные и недельные формы моек.
- Дневные формы `Шаркер` и `ЦКР` с помеченными заблокированными показателями,
  для которых в OData нет подтверждённого источника.
- Сводная дневная/недельная таблица по организациям.
- Проверки заполнения Excel-файлов и сохранности формульных ячеек.

## Требования

- Python 3.11+
- Доступ к тестовой базе 1С OData.
- Сетевой доступ к порту 8888 может зависеть от IP-allowlist.

## Установка

```bash
cd backend
python3 -m venv venv
./venv/bin/pip install -r requirements.txt
```

## Настройка

Секреты берутся из `backend/.env`; пример лежит в `.env.example`.
Минимально нужны:

```dotenv
ODATA_BASE_URL=http://91.144.178.239:8888/test_moskva/odata/standard.odata/
ODATA_LOGIN=...
ODATA_PASSWORD=...
ODATA_TIMEOUT=30
ODATA_RETRIES=3
```

`backend/.env` не должен попадать в git. Корневой `.gitignore` и
`backend/.gitignore` дополнительно исключают локальные секреты, `venv`,
`__pycache__` и сгенерированные отчёты.

## Основные команды

Smoke-тест подключения и базовых справочников:

```bash
cd backend
./venv/bin/python src/smoke_test.py
```

Быстрая проверка синтаксиса без обращения к 1С:

```bash
cd ..
backend/venv/bin/python -m compileall -q backend/src backend/tests
```

Интеграционные проверки отчётов:

```bash
cd backend
./venv/bin/python tests/test_metrics_smoke.py
./venv/bin/python tests/test_excel_smoke.py
./venv/bin/python tests/test_increment4.py
./venv/bin/python tests/test_weekly_excel.py
./venv/bin/python tests/test_wash_excel.py
./venv/bin/python tests/test_shop_excel.py
```

Эти проверки ходят в живую 1С и могут завершиться кодом `2`, если OData
недоступен из-за сети, таймаута или IP-allowlist.

## Структура

```text
backend/
  .env.example
  requirements.txt
  ARCHITECTURE.md
  INCREMENT*.md
  src/
    config.py          загрузка .env
    odata_client.py    низкоуровневый OData-клиент
    references.py      справочники и маппинг отчётов на подразделения
    repositories.py    чтение заказ-нарядов, регистров, планов, оплат
    metrics.py         расчёт дневных/недельных/спец. метрик
    excel_reporter.py  заполнение Excel-шаблона
    consolidated.py    сводные таблицы
    smoke_test.py      проверка подключения
  templates/
    report_template.xlsx
  tests/
    test_*.py          интеграционные smoke-проверки
```

## Важные ограничения

- Записи в 1С нет: только чтение через OData.
- Формулы в Excel-шаблоне должны оставаться формулами; код пишет только ячейки
  данных.
- Некоторые показатели `Шаркер`/`ЦКР` намеренно выводятся прочерком, пока не
  подтверждён источник данных или формула.
- Сгенерированные `.xlsx` лежат в `backend/output/` и не коммитятся.
