#!/usr/bin/env bash
set -euo pipefail

PID_FILE="$(cd "$(dirname "$0")" && pwd)/.app.pid"

if [ ! -f "$PID_FILE" ]; then
  echo "Not running (no PID file)"
  exit 0
fi

PID=$(cat "$PID_FILE")
if kill -0 "$PID" 2>/dev/null; then
  kill "$PID"
  rm "$PID_FILE"
  echo "Stopped (PID: $PID)"
else
  echo "Process $PID not found, removing stale PID file"
  rm "$PID_FILE"
fi
