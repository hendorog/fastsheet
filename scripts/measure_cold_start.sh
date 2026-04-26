#!/usr/bin/env bash
# Cold-start measurement: launch the release binary and time until the
# fastsheet window is mapped on X11 (i.e. visible).
#
# Usage: scripts/measure_cold_start.sh [iterations]
set -uo pipefail

ITERS="${1:-5}"
BIN="$(cd "$(dirname "$0")/.." && pwd)/src-tauri/target/release/fastsheet"

[[ -x "$BIN" ]] || { echo "build first: cargo build --release --manifest-path src-tauri/Cargo.toml" >&2; exit 1; }
command -v xwininfo >/dev/null || { echo "xwininfo missing — install x11-utils" >&2; exit 1; }

# Warm run — populate page cache so we measure runtime startup, not disk I/O.
"$BIN" >/dev/null 2>&1 &
warm=$!
sleep 4
kill "$warm" 2>/dev/null
pkill -9 -f 'release/fastsheet' 2>/dev/null
sleep 1

echo "iter,wall_ms"
total=0
for i in $(seq 1 "$ITERS"); do
  start=$(date +%s%N)
  "$BIN" >/dev/null 2>&1 &
  pid=$!
  # Poll every 25ms for the fastsheet 800x600 window to appear.
  for _ in $(seq 1 400); do
    if xwininfo -root -tree 2>/dev/null | grep -q '"fastsheet".*800x600'; then
      break
    fi
    sleep 0.025
  done
  end=$(date +%s%N)
  ms=$(( (end - start) / 1000000 ))
  echo "$i,$ms"
  total=$(( total + ms ))
  kill "$pid" 2>/dev/null
  pkill -9 -f 'release/fastsheet' 2>/dev/null
  wait "$pid" 2>/dev/null
  sleep 1
done

echo "avg,$(( total / ITERS ))"
