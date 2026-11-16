#!/usr/bin/env bash
# Starts the mock IdP and the Grid cloud app (serves ../dist, built with npm run build:cloud). Ports: GRID_APP_PORT (8080), GRID_IDP_PORT (8090).
# Env: GRID_DB=sqlite|mysql  GRID_EVERYONE_PRO=1|0  GRID_SHARE_RPM=60
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
RUN="${GRID_RUN_DIR:-$ROOT/dev/.run}"
DIST="$(dirname "$ROOT")/dist"
mkdir -p "$RUN" "$DIST"

if [ ! -f "$DIST/index.html" ]; then
  echo "note: dist/index.html missing; writing a placeholder (run npm run build:cloud for the real app)"
  printf '<!doctype html><meta charset="utf-8"><title>Grid (dev placeholder)</title><p>Grid API dev server. Build the SPA with <code>npm run build</code>.</p>' > "$DIST/index.html"
fi

stop() {
  for f in "$RUN"/*.pid; do
    if [ -f "$f" ]; then kill "$(cat "$f")" 2>/dev/null || true; rm -f "$f"; fi
  done
}
if [ "${1:-}" = "stop" ]; then stop; echo "stopped"; exit 0; fi
stop

export GRID_IDP_STATE="${GRID_IDP_STATE:-$RUN/grid-idp.json}"
export GRID_CONFIG="${GRID_CONFIG:-$ROOT/dev/config.dev.php}"
export GRID_DB="${GRID_DB:-sqlite}"
export GRID_SQLITE="${GRID_SQLITE:-$RUN/grid-dev.sqlite}"
export GRID_APP_PORT="${GRID_APP_PORT:-8080}"
export GRID_IDP_PORT="${GRID_IDP_PORT:-8090}"

PHP_CLI_SERVER_WORKERS=2 php -S "127.0.0.1:$GRID_IDP_PORT" "$ROOT/dev/idp/index.php" > "$RUN/idp.log" 2>&1 &
echo $! > "$RUN/idp.pid"
PHP_CLI_SERVER_WORKERS="${PHP_CLI_SERVER_WORKERS:-6}" php -S "127.0.0.1:$GRID_APP_PORT" "$ROOT/dev/router.php" > "$RUN/app.log" 2>&1 &
echo $! > "$RUN/app.pid"

for i in $(seq 1 50); do
  if curl -fs -o /dev/null http://127.0.0.1:$GRID_IDP_PORT/.well-known/openid-configuration && curl -s -o /dev/null http://127.0.0.1:$GRID_APP_PORT/api/me; then
    echo "idp  → http://127.0.0.1:$GRID_IDP_PORT  (log: $RUN/idp.log)"
    echo "app  → http://127.0.0.1:$GRID_APP_PORT  (log: $RUN/app.log, db: $GRID_DB)"
    exit 0
  fi
  sleep 0.2
done
echo "servers did not come up; see $RUN/*.log" >&2
exit 1
