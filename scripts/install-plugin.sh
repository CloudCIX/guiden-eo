#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CONTAINER="${1:-euro-office}"
PLUGIN_NAME="cloudcix-guiden"
PLUGIN_SRC="$ROOT/plugin/$PLUGIN_NAME"

if [ ! -d "$PLUGIN_SRC" ]; then
  echo "Plugin source not found: $PLUGIN_SRC" >&2
  exit 1
fi

"$ROOT/scripts/preflight-runtime.sh" "$CONTAINER"

PLUGIN_DEST="$(docker exec "$CONTAINER" bash -lc '
for root in /var/www/euro-office/documentserver/sdkjs-plugins /var/www/onlyoffice/documentserver/sdkjs-plugins; do
  if [ -d "$root" ]; then
    printf "%s" "$root"
    exit 0
  fi
done
exit 1
')"

docker exec "$CONTAINER" rm -rf "$PLUGIN_DEST/$PLUGIN_NAME"
docker cp "$PLUGIN_SRC" "$CONTAINER:$PLUGIN_DEST/$PLUGIN_NAME"

echo "Installed $PLUGIN_NAME to $CONTAINER:$PLUGIN_DEST/$PLUGIN_NAME"
echo "Restarting $CONTAINER so the editor reloads plugin metadata..."
docker restart "$CONTAINER" >/dev/null
echo "Done."

