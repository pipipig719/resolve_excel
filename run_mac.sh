#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

mode="${1:-huaining}"
if [[ "$mode" == "huaining" || "$mode" == "feixi" ]]; then
  if [[ $# -gt 0 ]]; then
    shift
  fi
elif [[ "$mode" == --* ]]; then
  mode="huaining"
else
  echo "[ERROR] Unknown mode: $mode"
  echo "Usage: ./run_mac.sh [huaining|feixi] [processor args]"
  exit 1
fi

if [[ "$mode" == "huaining" ]]; then
  target_script="huaining/process_huaining.py"
else
  target_script="feixi/process_feixi.py"
fi

if ! command -v uv >/dev/null 2>&1; then
  echo "[ERROR] uv is not installed or not in PATH."
  echo "Install: https://docs.astral.sh/uv/getting-started/installation/"
  exit 1
fi

uv sync
uv run python "$target_script" "$@"

echo "[OK] ${mode} pipeline finished."
