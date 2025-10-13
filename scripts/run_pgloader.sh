#!/usr/bin/env bash
set -euo pipefail

# Usage: ./scripts/run_pgloader.sh /absolute/path/to/db.sqlite3 postgresql://user:pass@host:5432/dbname
SQLITE_PATH="$1"
PG_URL="$2"

docker run --rm --name pgloader \
  --network host \
  -v "${SQLITE_PATH}:/data/db.sqlite3:ro" \
  dimitri/pgloader pgloader \
  /data/db.sqlite3 "${PG_URL}"
