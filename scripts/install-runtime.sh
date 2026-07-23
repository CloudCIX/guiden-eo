#!/usr/bin/env bash
set -euo pipefail

CONTAINER="${1:-euro-office}"
BASE_URL="${SDKJS_PLUGINS_V1_URL:-https://onlyoffice.github.io/sdkjs-plugins/v1}"
TMP_DIR="$(mktemp -d)"

cleanup() {
  rm -rf "$TMP_DIR"
}
trap cleanup EXIT

PLUGIN_ROOT="$(docker exec "$CONTAINER" bash -lc '
for root in /var/www/euro-office/documentserver/sdkjs-plugins /var/www/onlyoffice/documentserver/sdkjs-plugins; do
  if [ -d "$root" ]; then
    printf "%s" "$root"
    exit 0
  fi
done
exit 1
')"

mkdir -p "$TMP_DIR/v1"
for file in plugins.js plugins-ui.js plugins.css; do
  curl -fsSL "$BASE_URL/$file" -o "$TMP_DIR/v1/$file"
done

docker exec "$CONTAINER" mkdir -p "$PLUGIN_ROOT/v1"
for file in plugins.js plugins-ui.js plugins.css; do
  docker cp "$TMP_DIR/v1/$file" "$CONTAINER:$PLUGIN_ROOT/v1/$file"
done

echo "Installed ONLYOFFICE plugin runtime to $CONTAINER:$PLUGIN_ROOT/v1"

