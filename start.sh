#!/bin/sh
set -eu
mkdir -p backend/data
if [ -n "${LITESTREAM_ACCESS_KEY_ID:-}" ] && [ -n "${LITESTREAM_SECRET_ACCESS_KEY:-}" ] && [ -n "${LITESTREAM_BUCKET:-}" ]; then
  ./litestream restore -if-replica-exists -config litestream.yml backend/data/articles.db || true
  ./litestream replicate -config litestream.yml &
  litestream_pid=$!
  trap 'kill "$litestream_pid" 2>/dev/null || true' TERM INT
  uvicorn server:app --app-dir backend --host 0.0.0.0 --port "${PORT:-8756}" &
  app_pid=$!
  wait "$app_pid"
  status=$?
  kill "$litestream_pid" 2>/dev/null || true
  exit "$status"
fi
exec uvicorn server:app --app-dir backend --host 0.0.0.0 --port "${PORT:-8756}"
