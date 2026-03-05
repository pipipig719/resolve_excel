#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")/.."

if ! command -v uv >/dev/null 2>&1; then
  echo "[ERROR] uv is not installed or not in PATH."
  echo "Install: https://docs.astral.sh/uv/getting-started/installation/"
  exit 1
fi

uv sync
uv run python huaining/process_huaining.py "$@"
