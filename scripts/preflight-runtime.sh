#!/usr/bin/env bash
set -euo pipefail

CONTAINER="${1:-euro-office}"

docker exec "$CONTAINER" bash -lc '
set -euo pipefail
for root in /var/www/euro-office/documentserver/sdkjs-plugins /var/www/onlyoffice/documentserver/sdkjs-plugins; do
  if [ -d "$root" ]; then
    missing=0
    for file in v1/plugins.js v1/plugins-ui.js v1/plugins.css; do
      if [ ! -f "$root/$file" ]; then
        echo "missing: $root/$file"
        missing=1
      fi
    done
    if [ "$missing" -eq 0 ]; then
      echo "ok: ONLYOFFICE plugin runtime found under $root"
      exit 0
    fi
    exit 1
  fi
done
echo "missing: no sdkjs-plugins directory found"
exit 1
'

