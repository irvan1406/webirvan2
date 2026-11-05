#!/bin/bash
# hermes-models-apply.sh — terjemahkan hermes-models.txt menjadi config Hermes.
# Dipanggil start-hermes.sh saat boot (hasilnya masuk hermes-boot.log).
#
# Format hermes-models.txt (satu blok per provider):
#   [nama]
#   base_url = https://.../v1
#   model    = nama-model
#   api_key  = kunci
#   default  = false        (opsional; true = jadikan model bawaan)
#
# Skema tujuan TERVERIFIKASI dari config.yaml Hermes terpasang (5 Nov 2026):
# providers.<nama>.base_url / .api_mode / .model / .key_env, dengan nilai key
# disimpan di .hermes/.env di bawah nama variabel yang ditunjuk key_env.
# Blok yang base_url/model/api_key-nya belum lengkap dilewati dengan aman.
MODELS_FILE="${1:-$HOME/hermes-models.txt}"
HERMES_HOME="${HERMES_HOME:-$HOME/.hermes}"
[ -f "$MODELS_FILE" ] || exit 0
mkdir -p "$HERMES_HOME"
touch "$HERMES_HOME/.env"

name=""; base_url=""; model=""; api_key=""; is_default="false"

apply_block() {
  [ -n "$name" ] || return 0
  if [ -z "$base_url" ] || [ -z "$model" ] || [ -z "$api_key" ]; then
    echo "model [$name] dilewati (base_url/model/api_key belum lengkap)"
    return 0
  fi
  envname=$(printf '%s' "$name" | tr 'a-z' 'A-Z' | tr -c 'A-Z0-9' '_')
  envname="${envname}_API_KEY"
  # Simpan / perbarui key di .hermes/.env (nilai key tidak pernah dicetak)
  tmp=$(mktemp)
  grep -v "^${envname}=" "$HERMES_HOME/.env" > "$tmp" 2>/dev/null
  echo "${envname}=${api_key}" >> "$tmp"
  mv "$tmp" "$HERMES_HOME/.env"
  # Daftarkan provider bernama di config Hermes
  timeout 60 hermes config set "providers.${name}.base_url" "$base_url" </dev/null >/dev/null 2>&1
  timeout 60 hermes config set "providers.${name}.api_mode" "chat_completions" </dev/null >/dev/null 2>&1
  timeout 60 hermes config set "providers.${name}.model" "$model" </dev/null >/dev/null 2>&1
  timeout 60 hermes config set "providers.${name}.key_env" "$envname" </dev/null >/dev/null 2>&1
  echo "model [$name] terdaftar (model: $model, key_env: $envname)"
  if [ "$is_default" = "true" ]; then
    timeout 60 hermes config set model.provider custom </dev/null >/dev/null 2>&1
    timeout 60 hermes config set model.base_url "$base_url" </dev/null >/dev/null 2>&1
    timeout 60 hermes config set model.api_key "$api_key" </dev/null >/dev/null 2>&1
    timeout 60 hermes config set model.default "$model" </dev/null >/dev/null 2>&1
    echo "model [$name] dijadikan bawaan"
  fi
}

while IFS= read -r line || [ -n "$line" ]; do
  line=$(echo "$line" | sed 's/^[[:space:]]*//; s/[[:space:]]*$//')
  case "$line" in
    ""|\#*) continue ;;
    \[*\])
      apply_block
      name=$(echo "$line" | tr -d '[]' | tr -d '[:space:]' | tr 'A-Z' 'a-z' | sed 's/[^a-z0-9_-]//g')
      base_url=""; model=""; api_key=""; is_default="false"
      ;;
    *=*)
      k=$(echo "${line%%=*}" | tr -d '[:space:]')
      v=$(echo "${line#*=}" | sed 's/^[[:space:]]*//; s/[[:space:]]*$//')
      case "$k" in
        base_url) base_url="$v" ;;
        model) model="$v" ;;
        api_key) api_key="$v" ;;
        default) is_default="$v" ;;
      esac
      ;;
  esac
done < "$MODELS_FILE"
apply_block
