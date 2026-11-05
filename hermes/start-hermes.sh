#!/bin/bash
# start-hermes.sh v3 — boot KHUSUS server Hermes terpisah (ANTI-WIZARD + MODEL).
# v3 menambahkan langkah 2b: model tambahan dari hermes-models.txt (berkas
# kecil yang diedit user lewat panel) diterjemahkan ke config Hermes oleh
# hermes-models-apply.sh (diambil dari repo, terkunci commit d32f402).
# Pelajaran boot pertama (5 Nov 2026): installer Hermes diakhiri Setup Wizard
# interaktif yang menunggu ketikan di /dev/tty dan membekukan seluruh boot.
# Karena itu di v2: (1) installer jalan di latar, skrip TIDAK menunggu dia
# selesai — cukup biner hermes muncul; (2) SEMUA perintah hermes diberi
# stdin /dev/null + timeout; (3) gateway jalan di latar dengan loop
# pengawasan, proses utama hanya menunggu (wait) — pola yang terbukti
# lolos wizard di server alvanda.
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

echo "[$(date '+%F %T')] === hermes boot v3 (anti-wizard + model) ===" >> "$LOG"

# Baca secrets lebih dulu (dipakai beberapa langkah)
SEC="$HOME/hermes-secrets.env"
if [ -f "$SEC" ]; then
  INC_KEY=$(grep -E '^INCEPTION_API_KEY=' "$SEC" | head -1 | cut -d= -f2- | tr -d '[:space:]')
  TG_TOKEN=$(grep -E '^TELEGRAM_BOT_TOKEN=' "$SEC" | head -1 | cut -d= -f2- | tr -d '[:space:]')
  TG_UID=$(grep -E '^TELEGRAM_USER_ID=' "$SEC" | head -1 | cut -d= -f2- | tr -d '[:space:]')
fi

# 1) INSTALL bila biner hermes belum ada — di LATAR, tanpa menunggu selesai.
if ! command -v hermes >/dev/null 2>&1; then
  echo "[$(date '+%F %T')] memasang hermes-agent di latar (installer resmi)..." >> "$LOG"
  curl -fsSL https://hermes-agent.nousresearch.com/install.sh -o "$HOME/hermes-install.sh" >> "$LOG" 2>&1
  ( bash "$HOME/hermes-install.sh" </dev/null >> "$LOG" 2>&1 ) &
  # Tunggu sampai biner hermes muncul (maks ±12 menit), lalu lanjut.
  for i in $(seq 1 144); do
    command -v hermes >/dev/null 2>&1 && break
    sleep 5
  done
fi
if ! command -v hermes >/dev/null 2>&1; then
  echo "[$(date '+%F %T')] hermes TIDAK muncul juga — boot dihentikan." >> "$LOG"
  echo "Hermes gagal terpasang; cek hermes-boot.log"
  sleep 60
  exit 1
fi
echo "[$(date '+%F %T')] hermes terdeteksi: $(timeout 15 hermes --version </dev/null 2>/dev/null | head -1)" >> "$LOG"

# 2) KONFIGURASI model dari secrets (sekali saja; timeout + stdin tertutup)
if [ -n "$INC_KEY" ] && [ ! -f "$HERMES_HOME/.configured_ok" ]; then
  echo "[$(date '+%F %T')] mengkonfigurasi model Inception mercury-3..." >> "$LOG"
  timeout 60 hermes config set model.provider custom </dev/null >> "$LOG" 2>&1
  timeout 60 hermes config set model.base_url https://api.inceptionlabs.ai/v1 </dev/null >> "$LOG" 2>&1
  timeout 60 hermes config set model.api_key "$INC_KEY" </dev/null >> "$LOG" 2>&1
  timeout 60 hermes config set model.name mercury-3 </dev/null >> "$LOG" 2>&1
  timeout 60 hermes config set model.default custom/mercury-3 </dev/null >> "$LOG" 2>&1
  mkdir -p "$HERMES_HOME"
  touch "$HERMES_HOME/.configured_ok"
  echo "[$(date '+%F %T')] konfigurasi model selesai." >> "$LOG"
fi

# Token Telegram untuk gateway — tulis ke .env Hermes bila belum ada.
# JANGAN menambahkan blok 'gateway:' ke config.yaml (config bawaan sudah
# punya kunci itu; blok kedua = duplikat yang membatalkan seluruh config).
# Otorisasi pengguna diatur lewat pairing, bukan allowlist config.
if [ -n "$TG_TOKEN" ]; then
  mkdir -p "$HERMES_HOME"
  touch "$HERMES_HOME/.env"
  if ! grep -q '^TELEGRAM_BOT_TOKEN=' "$HERMES_HOME/.env"; then
    echo "TELEGRAM_BOT_TOKEN=$TG_TOKEN" >> "$HERMES_HOME/.env"
  fi
fi

# 2b) MODEL TAMBAHAN dari hermes-models.txt (diisi user lewat panel Files).
# Penerjemahnya diambil dari repo webirvan2 terkunci commit; bila gagal
# diambil, langkah ini dilewati dengan aman (model bawaan tetap jalan).
APPLY="$HOME/hermes-models-apply.sh"
if curl -fsSL "https://raw.githubusercontent.com/irvan1406/webirvan2/d32f402/hermes/hermes-models-apply.sh" -o "$APPLY" >> "$LOG" 2>&1 && [ -s "$APPLY" ]; then
  echo "[$(date '+%F %T')] menerapkan hermes-models.txt..." >> "$LOG"
  bash "$APPLY" >> "$LOG" 2>&1
else
  echo "[$(date '+%F %T')] hermes-models-apply.sh gagal diambil — model tambahan dilewati." >> "$LOG"
fi

# 3) PERSETUJUAN PAIRING Telegram dari pairing-code.txt (sekali jalan)
if [ -f "$HOME/pairing-code.txt" ] && [ ! -f "$HERMES_HOME/.pairing_ok" ]; then
  CODE=$(tr -d '[:space:]' < "$HOME/pairing-code.txt")
  if [ -n "$CODE" ]; then
    echo "[$(date '+%F %T')] menyetujui pairing telegram dari pairing-code.txt..." >> "$LOG"
    if timeout 60 hermes pairing approve telegram "$CODE" </dev/null >> "$LOG" 2>&1; then
      touch "$HERMES_HOME/.pairing_ok"
      rm -f "$HOME/pairing-code.txt"
      echo "[$(date '+%F %T')] pairing telegram DISETUJUI." >> "$LOG"
    else
      echo "[$(date '+%F %T')] persetujuan pairing GAGAL — kode mungkin kedaluwarsa; minta kode baru ke bot." >> "$LOG"
    fi
  fi
fi

# 4) GATEWAY di latar dengan loop pengawasan; proses utama menunggu.
if [ -f "$HERMES_HOME/.configured_ok" ] && [ -n "$TG_TOKEN" ]; then
  echo "[$(date '+%F %T')] menyalakan hermes gateway (latar, diawasi)..." >> "$LOG"
  (
    while true; do
      echo "[$(date '+%F %T')] gateway mulai" >> "$LOG"
      hermes gateway run </dev/null >> "$HERMES_HOME/logs/gateway-console.log" 2>&1
      echo "[$(date '+%F %T')] gateway berhenti (rc=$?) — mulai lagi 5 detik lagi" >> "$LOG"
      sleep 5
    done
  ) &
  wait
else
  echo "[$(date '+%F %T')] gateway TIDAK dinyalakan (secrets belum lengkap). Isi hermes-secrets.env lalu restart." >> "$LOG"
  echo "Secrets belum lengkap — isi hermes-secrets.env lewat Files, lalu restart server."
  sleep 60
  exit 1
fi
