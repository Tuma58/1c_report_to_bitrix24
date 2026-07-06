#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BACKEND_DIR="$PROJECT_DIR/backend"
VENV_DIR="$BACKEND_DIR/venv"
PYTHON_BIN="$VENV_DIR/bin/python"
PIP_BIN="$VENV_DIR/bin/pip"

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

echo "[1/6] Installing system packages"
$SUDO apt-get update
$SUDO DEBIAN_FRONTEND=noninteractive apt-get install -y \
  ca-certificates \
  cron \
  curl \
  git \
  python3 \
  python3-pip \
  python3-venv \
  tzdata

echo "[2/6] Creating Python virtual environment"
python3 -m venv "$VENV_DIR"

echo "[3/6] Installing Python dependencies"
"$PIP_BIN" install --upgrade pip
"$PIP_BIN" install -r "$BACKEND_DIR/requirements.txt"

echo "[4/6] Preparing local files"
mkdir -p "$BACKEND_DIR/output" "$BACKEND_DIR/logs"
if [[ ! -f "$BACKEND_DIR/.env" ]]; then
  cp "$BACKEND_DIR/.env.example" "$BACKEND_DIR/.env"
  chmod 600 "$BACKEND_DIR/.env"
  echo "Created backend/.env from example. Fill ODATA_* and BITRIX_* values before production run."
else
  chmod 600 "$BACKEND_DIR/.env"
fi

echo "[5/6] Smoke-checking Python modules"
"$PYTHON_BIN" -m compileall -q "$BACKEND_DIR/src" "$BACKEND_DIR/tests"

if [[ "${INSTALL_CRON:-0}" == "1" ]]; then
  echo "[6/6] Installing cron jobs"
  CRON_FILE="$(mktemp)"
  crontab -l 2>/dev/null > "$CRON_FILE" || true
  sed -i '/1c_report generate_reports/d' "$CRON_FILE"
  {
    echo "0 11 * * * cd $BACKEND_DIR/src && $PYTHON_BIN generate_reports.py --mode daily --send-bitrix >> $BACKEND_DIR/logs/daily.log 2>&1 # 1c_report generate_reports"
    echo "0 11 * * 5 cd $BACKEND_DIR/src && $PYTHON_BIN generate_reports.py --mode weekly --send-bitrix >> $BACKEND_DIR/logs/weekly.log 2>&1 # 1c_report generate_reports"
  } >> "$CRON_FILE"
  crontab "$CRON_FILE"
  rm -f "$CRON_FILE"
  $SUDO systemctl enable --now cron >/dev/null 2>&1 || true
else
  echo "[6/6] Cron skipped. Re-run with INSTALL_CRON=1 ./setup.sh to install jobs."
fi

cat <<EOF

Setup complete.

Next steps:
1. Edit $BACKEND_DIR/.env
2. Test OData:
   cd $BACKEND_DIR && ./venv/bin/python src/smoke_test.py
3. Generate all reports:
   cd $BACKEND_DIR/src && ../venv/bin/python generate_reports.py --mode all --date YYYY-MM-DD
4. Generate and send to Bitrix24:
   cd $BACKEND_DIR/src && ../venv/bin/python generate_reports.py --mode all --send-bitrix
EOF
