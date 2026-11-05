#!/bin/bash
# start-all.sh — dipanggil oleh "npm start" (package.json) di server panel.
# Tugas: (1) pasang Hermes Agent resmi bila belum ada, (2) konfigurasikan dari
# hermes-secrets.env bila user sudah mengisinya, (3) nyalakan gateway Hermes di
# latar, (4) jalankan Bot Nasabah di depan (proses utama panel).
cd /home/container || exit 1
export HOME=/home/container
export PATH="$HOME/.local/bin:/usr/local/bin:/usr/bin:/bin:$PATH"
export HERMES_HOME="$HOME/.hermes"
LOG="$HOME/hermes-boot.log"

# Buat templat secrets bila belum ada (user mengisinya lewat panel Files)
if [ ! -f "$HOME/hermes-secrets.env" ]; then
  cat > "$HOME/hermes-secrets.env" << 'EOF'
# ============================================================
#  HERMES — isi 3 nilai di bawah ini lewat panel (Files > edit)
#  Jangan beri file ini ke siapa pun. Setelah disimpan,
#  kabari VanX supaya server direstart dan gateway dinyalakan.
# ============================================================
# 1) API key Inception Labs (yang dipakai di Mika kemarin).
#    Lupa/hilang? Buat baru di https://platform.inceptionlabs.ai
INCEPTION_API_KEY=

# 2) Token bot Telegram BARU khusus Hermes.
#    Buat di Telegram: chat @BotFather > /newbot > ikuti langkahnya.
TELEGRAM_BOT_TOKEN=

# 3) ID numerik Telegram kamu (angka saja).
#    Ambil di Telegram: chat @userinfobot, dia membalas ID kamu.
TELEGRAM_USER_ID=
EOF
fi

hermes_bg() {
  echo "[$(date '+%F %T')] === hermes background start ===" >> "$LOG"
  # 1) INSTALL bila belum ada biner hermes
  if ! command -v hermes >/dev/null 2>&1; then
    echo "[$(date '+%F %T')] memasang hermes-agent (unduh installer resmi)..." >> "$LOG"
    curl -fsSL https://hermes-agent.nousresearch.com/install.sh -o "$HOME/hermes-install.sh" >> "$LOG" 2>&1
    bash "$HOME/hermes-install.sh" </dev/null >> "$LOG" 2>&1
    echo "[$(date '+%F %T')] installer selesai, rc=$?" >> "$LOG"
  fi
  if ! command -v hermes >/dev/null 2>&1; then
    echo "[$(date '+%F %T')] hermes MASIH tidak ditemukan di PATH — cek log di atas." >> "$LOG"
    return 0
  fi
  echo "[$(date '+%F %T')] hermes terdeteksi: $(hermes --version 2>/dev/null | head -1)" >> "$LOG"

  # 2) KONFIGURASI dari hermes-secrets.env (diisi user lewat panel Files)
  SEC="$HOME/hermes-secrets.env"
  if [ -f "$SEC" ]; then
    INC_KEY=$(grep -E '^INCEPTION_API_KEY=' "$SEC" | head -1 | cut -d= -f2- | tr -d '[:space:]')
    TG_TOKEN=$(grep -E '^TELEGRAM_BOT_TOKEN=' "$SEC" | head -1 | cut -d= -f2- | tr -d '[:space:]')
    TG_UID=$(grep -E '^TELEGRAM_USER_ID=' "$SEC" | head -1 | cut -d= -f2- | tr -d '[:space:]')
  fi
  if [ -n "$INC_KEY" ] && [ ! -f "$HERMES_HOME/.configured_ok" ]; then
    echo "[$(date '+%F %T')] mengkonfigurasi model Inception mercury-3..." >> "$LOG"
    hermes config set model.provider custom >> "$LOG" 2>&1
    hermes config set model.base_url https://api.inceptionlabs.ai/v1 >> "$LOG" 2>&1
    hermes config set model.api_key "$INC_KEY" >> "$LOG" 2>&1
    hermes config set model.name mercury-3 >> "$LOG" 2>&1
    hermes config set model.default custom/mercury-3 >> "$LOG" 2>&1
    mkdir -p "$HERMES_HOME"
    touch "$HERMES_HOME/.configured_ok"
    echo "[$(date '+%F %T')] konfigurasi model selesai." >> "$LOG"
  fi

  # 3) GATEWAY Telegram — hanya bila model terkonfigurasi dan token ada
  if [ -f "$HERMES_HOME/.configured_ok" ] && [ -n "$TG_TOKEN" ]; then
    # tulis token + allowlist ke .env Hermes bila belum
    touch "$HERMES_HOME/.env"
    if ! grep -q '^TELEGRAM_BOT_TOKEN=' "$HERMES_HOME/.env"; then
      echo "TELEGRAM_BOT_TOKEN=$TG_TOKEN" >> "$HERMES_HOME/.env"
    fi
    if [ -n "$TG_UID" ] && ! grep -q 'allowed_user_ids' "$HERMES_HOME/config.yaml" 2>/dev/null; then
      {
        echo 'gateway:'
        echo '  telegram:'
        echo '    allowed_user_ids:'
        echo "      - $TG_UID"
      } >> "$HERMES_HOME/config.yaml"
    fi
    if ! pgrep -f 'hermes gateway' >/dev/null 2>&1; then
      echo "[$(date '+%F %T')] menyalakan hermes gateway..." >> "$LOG"
      mkdir -p "$HERMES_HOME/logs"
      nohup hermes gateway run >> "$HERMES_HOME/logs/gateway.log" 2>&1 &
    else
      echo "[$(date '+%F %T')] gateway sudah jalan." >> "$LOG"
    fi
  else
    echo "[$(date '+%F %T')] gateway belum dinyalakan (secrets belum lengkap)." >> "$LOG"
  fi
  echo "[$(date '+%F %T')] === hermes background selesai ===" >> "$LOG"
}

# Hermes dikerjakan di latar; bot TIDAK menunggu.
( hermes_bg ) &
exec node index.js
