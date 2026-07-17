#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BACKEND_DIR="$PROJECT_DIR/backend"
VENV_DIR="$BACKEND_DIR/venv"
PYTHON_BIN="$VENV_DIR/bin/python"
PIP_BIN="$VENV_DIR/bin/pip"
ENV_FILE="$BACKEND_DIR/.env"
WEB_SERVICE_NAME="1c-report-web"
WEB_SERVICE_FILE="/etc/systemd/system/${WEB_SERVICE_NAME}.service"
WEB_PASSWORD_GENERATED=0

SUDO=""
if [[ "${EUID}" -ne 0 ]]; then
  if command -v sudo >/dev/null 2>&1; then
    SUDO="sudo"
  else
    echo "Run as root or install sudo first." >&2
    exit 1
  fi
fi

if [[ -r /etc/os-release ]]; then
  . /etc/os-release
  if [[ "${ID:-}" != "debian" || "${VERSION_ID:-}" != "12" ]]; then
    echo "Warning: expected Debian 12, got ${PRETTY_NAME:-unknown OS}." >&2
  fi
fi

get_env_var() {
  local key="$1"
  if [[ ! -f "$ENV_FILE" ]]; then
    return 0
  fi
  grep -E "^${key}=" "$ENV_FILE" | tail -n 1 | cut -d= -f2- || true
}

set_env_var() {
  local key="$1"
  local value="$2"
  local tmp
  tmp="$(mktemp)"
  if [[ -f "$ENV_FILE" ]] && grep -qE "^${key}=" "$ENV_FILE"; then
    awk -v k="$key" -v v="$value" '
      $0 ~ "^" k "=" { print k "=" v; next }
      { print }
    ' "$ENV_FILE" > "$tmp"
    mv "$tmp" "$ENV_FILE"
  else
    printf '%s=%s\n' "$key" "$value" >> "$ENV_FILE"
    rm -f "$tmp"
  fi
}

ensure_env_var() {
  local key="$1"
  local value="$2"
  local current
  current="$(get_env_var "$key")"
  if [[ -z "$current" ]]; then
    set_env_var "$key" "$value"
  fi
}

generate_web_password() {
  if command -v openssl >/dev/null 2>&1; then
    openssl rand -hex 16
  else
    "$PYTHON_BIN" -c 'import secrets; print(secrets.token_hex(16))'
  fi
}

configure_web_env() {
  ensure_env_var "WEB_UI_HOST" "0.0.0.0"
  ensure_env_var "WEB_UI_PORT" "8080"
  ensure_env_var "WEB_UI_USER" "admin"

  local port
  port="$(get_env_var WEB_UI_PORT)"
  if [[ "$port" == "80" || "$port" == "443" ]]; then
    echo "WEB_UI_PORT=$port is not allowed; switching Web UI to 8080."
    set_env_var "WEB_UI_PORT" "8080"
  fi

  local password
  password="$(get_env_var WEB_UI_PASSWORD)"
  if [[ -z "$password" ]]; then
    set_env_var "WEB_UI_PASSWORD" "$(generate_web_password)"
    WEB_PASSWORD_GENERATED=1
  fi
  chmod 600 "$ENV_FILE"
}

configure_schedule_env() {
  ensure_env_var "EMAIL_SMTP_HOST" ""
  ensure_env_var "EMAIL_SMTP_PORT" "587"
  ensure_env_var "EMAIL_SMTP_LOGIN" ""
  ensure_env_var "EMAIL_SMTP_PASSWORD" ""
  ensure_env_var "EMAIL_FROM" ""
  ensure_env_var "EMAIL_TO" ""
  ensure_env_var "EMAIL_USE_TLS" "1"
  ensure_env_var "EMAIL_USE_SSL" "0"

  ensure_env_var "SCHEDULE_ENABLED" "0"
  ensure_env_var "SCHEDULE_DAILY_ENABLED" "1"
  ensure_env_var "SCHEDULE_DAILY_TIME" "11:00"
  ensure_env_var "SCHEDULE_WEEKLY_ENABLED" "1"
  ensure_env_var "SCHEDULE_WEEKLY_DAY" "5"
  ensure_env_var "SCHEDULE_WEEKLY_TIME" "11:00"
  ensure_env_var "SCHEDULE_SEND_BITRIX" "1"
  ensure_env_var "SCHEDULE_SEND_EMAIL" "0"

  if [[ "${INSTALL_CRON:-0}" == "1" ]]; then
    set_env_var "SCHEDULE_ENABLED" "1"
  fi
  chmod 600 "$ENV_FILE"
}

install_web_service() {
  if [[ "${INSTALL_WEB:-1}" != "1" ]]; then
    echo "[7/8] Web UI service skipped. Re-run with INSTALL_WEB=1 ./setup.sh to install it."
    return 0
  fi
  if ! command -v systemctl >/dev/null 2>&1; then
    echo "[7/8] systemctl not found; Web UI service was not installed."
    return 0
  fi

  echo "[7/8] Installing Web UI systemd service"
  local service_user
  service_user="$(id -un)"
  $SUDO tee "$WEB_SERVICE_FILE" >/dev/null <<EOF_SERVICE
[Unit]
Description=1C Report Web UI
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
User=${service_user}
WorkingDirectory=${BACKEND_DIR}
Environment=PYTHONUNBUFFERED=1
ExecStart=${PYTHON_BIN} ${BACKEND_DIR}/src/web_app.py
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF_SERVICE

  $SUDO systemctl daemon-reload
  $SUDO systemctl enable "$WEB_SERVICE_NAME" >/dev/null
  $SUDO systemctl restart "$WEB_SERVICE_NAME"
}

echo "[1/8] Installing system packages"
$SUDO apt-get update
$SUDO env DEBIAN_FRONTEND=noninteractive apt-get install -y \
  ca-certificates \
  cron \
  curl \
  git \
  openssl \
  python3 \
  python3-pip \
  python3-venv \
  tzdata

echo "[2/8] Creating Python virtual environment"
python3 -m venv "$VENV_DIR"

echo "[3/8] Installing Python dependencies"
"$PIP_BIN" install --upgrade pip
"$PIP_BIN" install -r "$BACKEND_DIR/requirements.txt"

echo "[4/8] Preparing local files"
mkdir -p "$BACKEND_DIR/output" "$BACKEND_DIR/logs"
if [[ ! -f "$ENV_FILE" ]]; then
  cp "$BACKEND_DIR/.env.example" "$ENV_FILE"
  chmod 600 "$ENV_FILE"
  echo "Created backend/.env from example. Fill ODATA_* and BITRIX_* values before production run."
else
  chmod 600 "$ENV_FILE"
fi

echo "[5/8] Configuring Web UI"
configure_web_env
configure_schedule_env

echo "[6/8] Smoke-checking Python modules"
"$PYTHON_BIN" -m compileall -q "$BACKEND_DIR/src" "$BACKEND_DIR/tests"

install_web_service

if command -v systemctl >/dev/null 2>&1; then
  $SUDO systemctl enable --now cron >/dev/null 2>&1 || true
fi

if [[ "${INSTALL_CRON:-0}" == "1" ]]; then
  echo "[8/8] Installing cron jobs from backend/.env schedule settings"
  "$PYTHON_BIN" "$BACKEND_DIR/src/scheduler.py"
else
  echo "[8/8] Cron skipped. Configure schedule in Web UI or re-run with INSTALL_CRON=1 ./setup.sh."
fi

WEB_UI_HOST_VALUE="$(get_env_var WEB_UI_HOST)"
WEB_UI_PORT_VALUE="$(get_env_var WEB_UI_PORT)"
WEB_UI_USER_VALUE="$(get_env_var WEB_UI_USER)"
WEB_URL_HOST="$WEB_UI_HOST_VALUE"
if [[ "$WEB_URL_HOST" == "0.0.0.0" ]]; then
  WEB_URL_HOST="<server-ip>"
fi
WEB_PASSWORD_NOTE="stored in $ENV_FILE as WEB_UI_PASSWORD"
if [[ "$WEB_PASSWORD_GENERATED" == "1" ]]; then
  WEB_PASSWORD_NOTE="generated and stored in $ENV_FILE as WEB_UI_PASSWORD"
fi

cat <<EOF

Setup complete.

Web UI:
  URL:      http://$WEB_URL_HOST:$WEB_UI_PORT_VALUE/
  Login:    $WEB_UI_USER_VALUE
  Password: $WEB_PASSWORD_NOTE
  Service:  systemctl status $WEB_SERVICE_NAME
  Ports:    80/443 are not used by this app; Web UI uses $WEB_UI_PORT_VALUE.

Next steps:
1. Edit $BACKEND_DIR/.env
2. Test OData:
   cd $BACKEND_DIR && ./venv/bin/python src/smoke_test.py
3. Generate one workbook with all report sheets:
   cd $BACKEND_DIR/src && ../venv/bin/python generate_reports.py --mode all --date YYYY-MM-DD
4. Generate one workbook and send it to Bitrix24:
   cd $BACKEND_DIR/src && ../venv/bin/python generate_reports.py --mode all --send-bitrix
5. Generate one workbook and send it by email:
   cd $BACKEND_DIR/src && ../venv/bin/python generate_reports.py --mode all --send-email
6. Open Web UI:
   http://$WEB_URL_HOST:$WEB_UI_PORT_VALUE/
7. Open Web UI settings:
   http://$WEB_URL_HOST:$WEB_UI_PORT_VALUE/settings
EOF
