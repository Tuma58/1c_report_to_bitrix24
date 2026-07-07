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
- Web-интерфейс настройки и запуска отчётов.
- Проверки заполнения Excel-файлов и сохранности формульных ячеек.

## Требования

- Python 3.11+
- Доступ к тестовой базе 1С OData.
- Сетевой доступ к порту 8888 может зависеть от IP-allowlist.

## Установка

На чистом Debian 12 после клонирования репозитория:

```bash
./setup.sh
```

Скрипт установит системные зависимости через `apt`, создаст `backend/venv`,
установит Python-пакеты, подготовит `backend/.env` из `.env.example`, проверит
компиляцию модулей и развернёт web-интерфейс как systemd-сервис
`1c-report-web`.

Web-интерфейс по умолчанию слушает порт `8080`. Приложение не занимает порты
`80` и `443`; если в `.env` случайно указан `WEB_UI_PORT=80` или `443`,
`setup.sh` заменит порт на `8080`, а `web_app.py` откажется стартовать на этих
портах.

Чтобы пропустить установку web-сервиса:

```bash
INSTALL_WEB=0 ./setup.sh
```

Для установки cron-задач:

```bash
INSTALL_CRON=1 ./setup.sh
```

Задачи cron:
- ежедневно в 11:00 — дневные отчёты за вчера;
- по пятницам в 11:00 — недельные отчёты за предыдущую неделю;
- обе задачи отправляют созданные файлы в Bitrix24, если заполнены `BITRIX_*`.

Ручная установка для разработки:

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
BITRIX_WEBHOOK_URL=
BITRIX_CHAT_ID=
BITRIX_DISK_FOLDER_ID=
WEB_UI_HOST=0.0.0.0
WEB_UI_PORT=8080
WEB_UI_USER=admin
WEB_UI_PASSWORD=
```

Если `WEB_UI_PASSWORD` пустой, `setup.sh` автоматически сгенерирует пароль и
сохранит его в `backend/.env`. Посмотреть пароль на сервере:

```bash
grep '^WEB_UI_PASSWORD=' backend/.env
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

Создать один Excel-файл со всеми листами отчётов из ТЗ:

```bash
cd backend/src
../venv/bin/python generate_reports.py --mode all --date 2026-07-01
```

Создать один файл отчётов и отправить его в Bitrix24:

```bash
cd backend/src
../venv/bin/python generate_reports.py --mode all --send-bitrix
```

Запустить web-интерфейс настройки отчёта вручную:

```bash
cd backend
./venv/bin/python src/web_app.py
```

## Web-интерфейс на VPS

После `./setup.sh` web-интерфейс развёрнут как systemd-сервис:

```bash
systemctl status 1c-report-web
systemctl restart 1c-report-web
journalctl -u 1c-report-web -f
```

Стандартная первичная настройка:

```dotenv
WEB_UI_HOST=0.0.0.0
WEB_UI_PORT=8080
WEB_UI_USER=admin
WEB_UI_PASSWORD=<сгенерировано setup.sh>
```

Адрес для прямого доступа:

```text
http://<ip-vps>:8080/
```

Порт `8080` нужно открыть в firewall VPS или панели облачного провайдера, если
доступ должен быть извне. Порты `80` и `443` приложение не использует.

Для закрытой схемы через reverse proxy можно поменять:

```dotenv
WEB_UI_HOST=127.0.0.1
WEB_UI_PORT=8080
```

После изменения `backend/.env` перезапустите сервис:

```bash
systemctl restart 1c-report-web
```

В интерфейсе доступны:
- выбор режима отчёта: `Все листы`, `День`, `Неделя`;
- выбор даты;
- опциональная отправка файла в Bitrix24;
- список последних `.xlsx`;
- скачивание сформированного файла.

## Bitrix24

Отправка использует входящий REST webhook:

- `BITRIX_WEBHOOK_URL` — URL вида `https://portal.bitrix24.ru/rest/<user>/<token>/`;
- `BITRIX_CHAT_ID` — диалог/чат, например `chat123`;
- `BITRIX_DISK_FOLDER_ID` — ID папки Диска Bitrix24 для загрузки `.xlsx`.

Файл загружается в указанную папку Диска, затем в чат отправляется сообщение
со ссылкой на созданную таблицу. Без `BITRIX_DISK_FOLDER_ID` отправка файла
завершится ошибкой, чтобы не имитировать успешную доставку.

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
    bitrix_sender.py   отправка ссылок на xlsx в чат Bitrix24
    generate_reports.py единый CLI создания всех форм
    web_app.py          web-интерфейс настройки отчёта
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
