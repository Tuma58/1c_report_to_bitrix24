# 1C Report to Bitrix24

Автоматическое формирование книги Excel с отчётами АТЦ из 1С OData и отправкой в чат Bitrix24. В проекте есть CLI-генератор, web-интерфейс настройки отчёта и скрипт установки для чистого VPS на Debian 12.

## Установка одной командой

Выполните на чистом VPS под `root`:

```bash
bash -lc 'set -e; apt-get update && apt-get install -y git ca-certificates; if [ -d /opt/1c_report/.git ]; then cd /opt/1c_report && git pull --ff-only; else git clone https://github.com/Tuma58/1c_report_to_bitrix24.git /opt/1c_report && cd /opt/1c_report; fi; ./setup.sh'
```

Команда подходит и для первой установки, и для повторного запуска после обновления репозитория. Скрипт установит системные пакеты, создаст Python-окружение, поставит зависимости, создаст `backend/.env`, проверит Python-модули и развернёт web-интерфейс как systemd-сервис `1c-report-web`.

Web-интерфейс по умолчанию работает на порту `8080` и не занимает порты `80` и `443`.

## Обновление VPS без потери данных

Для обновления уже установленного проекта выполните на VPS под `root`:

```bash
bash -lc 'set -euo pipefail; APP=/opt/1c_report; cd "$APP"; ts=$(date +%Y%m%d_%H%M%S); [ -f backend/.env ] && cp -a backend/.env "backend/.env.backup.$ts"; [ -f backend/users.json ] && cp -a backend/users.json "backend/users.json.backup.$ts"; git fetch origin main; git pull --ff-only origin main; ./setup.sh; systemctl restart 1c-report-web || true'
```

Команда сохраняет резервные копии локальных настроек перед обновлением. `backend/.env`, `backend/users.json`, `backend/output/` и сгенерированные отчёты не хранятся в Git и не затираются при `git pull --ff-only`. Если на VPS есть локальные изменения в tracked-файлах проекта, обновление остановится вместо перезаписи этих файлов.

## Первичная настройка

После установки заполните секреты и параметры подключения через web-интерфейс:

```text
http://<ip-vps>:8080/settings
```

Или напрямую в файле:

```bash
nano /opt/1c_report/backend/.env
```

Обязательные значения:

- `ODATA_BASE_URL` — URL опубликованной базы 1С OData.
- `ODATA_LOGIN` — логин read-only пользователя 1С.
- `ODATA_PASSWORD` — пароль пользователя 1С.
- `BITRIX_WEBHOOK_URL` — входящий REST webhook Bitrix24 с правами `im` и `disk`.
- `BITRIX_CHAT_ID` — ID чата Bitrix24, например `chat226489`.
- `EMAIL_SMTP_HOST`, `EMAIL_SMTP_PORT`, `EMAIL_SMTP_LOGIN`, `EMAIL_SMTP_PASSWORD` — SMTP-доступ для email-рассылки.
- `EMAIL_FROM`, `EMAIL_TO` — отправитель и получатели email.
- `SCHEDULE_*` — расписание автоматической генерации и каналы рассылки.
- `WEB_UI_USER` и `WEB_UI_PASSWORD` — доступ к web-интерфейсу.

Пароль web-интерфейса генерируется автоматически, если `WEB_UI_PASSWORD` пустой. Посмотреть его можно так:

```bash
grep '^WEB_UI_PASSWORD=' /opt/1c_report/backend/.env
```

Пользователями web-интерфейса управляет только роль `admin`:

```text
http://<ip-vps>:8080/users
```

На этой странице можно создать пользователя, сменить пароль, заблокировать
доступ или удалить учётную запись. Данные хранятся в
`/opt/1c_report/backend/users.json`; файл не коммитится в Git.

## Проверка

```bash
cd /opt/1c_report/backend && ./venv/bin/python src/smoke_test.py
```

Сформировать одну книгу со всеми листами:

```bash
cd /opt/1c_report/backend/src && ../venv/bin/python generate_reports.py --mode all --date YYYY-MM-DD
```

В web-интерфейсе созданные книги можно открыть как HTML-страницы отчёта:

```text
http://<ip-vps>:8080/view?file=<имя-файла.xlsx>
```

Сформировать книгу и отправить в Bitrix24:

```bash
cd /opt/1c_report/backend/src && ../venv/bin/python generate_reports.py --mode all --send-bitrix
```

Сформировать книгу и отправить по email:

```bash
cd /opt/1c_report/backend/src && ../venv/bin/python generate_reports.py --mode all --send-email
```

## Web-интерфейс

Адрес после установки:

```text
http://<ip-vps>:8080/
```

Команды управления:

```bash
systemctl status 1c-report-web
systemctl restart 1c-report-web
journalctl -u 1c-report-web -f
```

## Секреты

Секреты не хранятся в Git. Файлы `.env`, локальные заметки, сгенерированные отчёты и временные офисные файлы игнорируются через `.gitignore`.
