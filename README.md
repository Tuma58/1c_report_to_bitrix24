# 1C Report to Bitrix24

Автоматическое формирование книги Excel с отчётами АТЦ из 1С OData и отправкой в чат Bitrix24. В проекте есть CLI-генератор, web-интерфейс настройки отчёта и скрипт установки для чистого VPS на Debian 12.

## Установка одной командой

Выполните на чистом VPS под `root`:

```bash
bash -lc 'set -e; apt-get update && apt-get install -y git ca-certificates; if [ -d /opt/1c_report/.git ]; then cd /opt/1c_report && git pull --ff-only; else git clone https://github.com/Tuma58/1c_report_to_bitrix24.git /opt/1c_report && cd /opt/1c_report; fi; ./setup.sh'
```

Команда подходит и для первой установки, и для повторного запуска после обновления репозитория. Скрипт установит системные пакеты, создаст Python-окружение, поставит зависимости, создаст `backend/.env`, проверит Python-модули и развернёт web-интерфейс как systemd-сервис `1c-report-web`.

Web-интерфейс по умолчанию работает на порту `8080` и не занимает порты `80` и `443`.

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
- `BITRIX_WEBHOOK_URL` — входящий REST webhook Bitrix24.
- `BITRIX_CHAT_ID` — ID чата Bitrix24.
- `BITRIX_DISK_FOLDER_ID` — ID папки Диска для загрузки книги.
- `WEB_UI_USER` и `WEB_UI_PASSWORD` — доступ к web-интерфейсу.

Пароль web-интерфейса генерируется автоматически, если `WEB_UI_PASSWORD` пустой. Посмотреть его можно так:

```bash
grep '^WEB_UI_PASSWORD=' /opt/1c_report/backend/.env
```

## Проверка

```bash
cd /opt/1c_report/backend && ./venv/bin/python src/smoke_test.py
```

Сформировать одну книгу со всеми листами:

```bash
cd /opt/1c_report/backend/src && ../venv/bin/python generate_reports.py --mode all --date YYYY-MM-DD
```

Сформировать книгу и отправить в Bitrix24:

```bash
cd /opt/1c_report/backend/src && ../venv/bin/python generate_reports.py --mode all --send-bitrix
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
