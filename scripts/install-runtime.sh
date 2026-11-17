#!/usr/bin/env bash
set -euo pipefail

# Copies the ONLYOFFICE plugin runtime into a running container. By default it uses the pinned copies in runtime/v1
# (runtime/README.md says which upstream commit they are), the same files the image is built with. Set
# SDKJS_PLUGINS_V1_URL to download them from somewhere else instead, such as
# https://onlyoffice.github.io/sdkjs-plugins/v1 for upstream's latest (untested with Guiden).

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CONTAINER="${1:-euro-office}"
BASE_URL="${SDKJS_PLUGINS_V1_URL:-}"
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
  if [ -n "$BASE_URL" ]; then
    curl -fsSL "$BASE_URL/$file" -o "$TMP_DIR/v1/$file"
  else
    cp "$ROOT/runtime/v1/$file" "$TMP_DIR/v1/$file"
  fi
done

docker exec "$CONTAINER" mkdir -p "$PLUGIN_ROOT/v1"
for file in plugins.js plugins-ui.js plugins.css; do
  docker cp "$TMP_DIR/v1/$file" "$CONTAINER:$PLUGIN_ROOT/v1/$file"
  # The web server prefers a .gz copy when there is one, which would still be the old file.
  docker exec "$CONTAINER" rm -f "$PLUGIN_ROOT/v1/$file.gz"
done

echo "Installed ONLYOFFICE plugin runtime to $CONTAINER:$PLUGIN_ROOT/v1"

