#!/bin/bash
# start-hermes.sh — boot KHUSUS server Hermes terpisah (tanpa Bot Nasabah).
# Dipanggil oleh "npm start" (package.json) di server panel 'Hermes Test'.
# Tugas: (1) pasang Hermes Agent resmi bila belum ada, (2) konfigurasikan dari
# hermes-secrets.env, (3) setujui pairing Telegram dari pairing-code.txt bila
# ada, (4) jalankan gateway Hermes sebagai proses utama (foreground).
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

echo "[$(date '+%F %T')] === hermes boot (server terpisah) ===" >> "$LOG"

# 1) INSTALL bila belum ada biner hermes
if ! command -v hermes >/dev/null 2>&1; then
  echo "[$(date '+%F %T')] memasang hermes-agent (unduh installer resmi)..." >> "$LOG"
  curl -fsSL https://hermes-agent.nousresearch.com/install.sh -o "$HOME/hermes-install.sh" >> "$LOG" 2>&1
  bash "$HOME/hermes-install.sh" </dev/null >> "$LOG" 2>&1
  echo "[$(date '+%F %T')] installer selesai, rc=$?" >> "$LOG"
fi
if ! command -v hermes >/dev/null 2>&1; then
  echo "[$(date '+%F %T')] hermes MASIH tidak ditemukan di PATH — bot tidak dinyalakan." >> "$LOG"
  echo "Hermes gagal terpasang; cek hermes-boot.log"
  sleep 60
  exit 1
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

# Token Telegram untuk gateway — tulis ke .env Hermes bila belum ada.
# CATATAN: JANGAN menambahkan blok 'gateway:' ke config.yaml — config bawaan
# installer sudah punya kunci gateway (kosong); blok kedua = duplikat yang
# membuat seluruh config dianggap tidak valid (kejadian 5 Nov di alvanda).
# Otorisasi pengguna diatur lewat pairing, bukan allowlist config.
if [ -n "$TG_TOKEN" ]; then
  touch "$HERMES_HOME/.env"
  if ! grep -q '^TELEGRAM_BOT_TOKEN=' "$HERMES_HOME/.env"; then
    echo "TELEGRAM_BOT_TOKEN=$TG_TOKEN" >> "$HERMES_HOME/.env"
  fi
fi

# 3) PERSETUJUAN PAIRING Telegram — kode ditulis user ke pairing-code.txt
#    lewat panel Files (satu baris, 8 karakter). Sekali berhasil, penanda
#    dibuat dan berkas kode dihapus.
if [ -f "$HOME/pairing-code.txt" ] && [ ! -f "$HERMES_HOME/.pairing_ok" ]; then
  CODE=$(tr -d '[:space:]' < "$HOME/pairing-code.txt")
  if [ -n "$CODE" ]; then
    echo "[$(date '+%F %T')] menyetujui pairing telegram dari pairing-code.txt..." >> "$LOG"
    if hermes pairing approve telegram "$CODE" >> "$LOG" 2>&1; then
      touch "$HERMES_HOME/.pairing_ok"
      rm -f "$HOME/pairing-code.txt"
      echo "[$(date '+%F %T')] pairing telegram DISETUJUI." >> "$LOG"
    else
      echo "[$(date '+%F %T')] persetujuan pairing GAGAL — kode mungkin kedaluwarsa; minta kode baru ke bot." >> "$LOG"
    fi
  fi
fi

# 4) GATEWAY sebagai proses utama (foreground)
if [ -f "$HERMES_HOME/.configured_ok" ] && [ -n "$TG_TOKEN" ]; then
  echo "[$(date '+%F %T')] menyalakan hermes gateway (foreground)..." >> "$LOG"
  exec hermes gateway run
else
  echo "[$(date '+%F %T')] gateway TIDAK dinyalakan (secrets belum lengkap). Isi hermes-secrets.env lalu restart." >> "$LOG"
  echo "Secrets belum lengkap — isi hermes-secrets.env lewat Files, lalu restart server."
  sleep 60
  exit 1
fi
