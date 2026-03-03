#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

if [[ ! -d "source" ]]; then
  echo "[ERROR] source 目录不存在。请在项目目录下创建 source 并放入配方间饮片数据。"
  exit 1
fi

if ! command -v uv >/dev/null 2>&1; then
  echo "[ERROR] 未检测到 uv。请先安装 uv 后再执行。"
  echo "安装参考: https://docs.astral.sh/uv/getting-started/installation/"
  exit 1
fi

uv sync
uv run python run_pipeline.py "$@"

echo "[OK] 已生成:"
echo "  1) 根目录: 饮片货位导入最终文件.xlsx"
echo "  2) source目录: 饮片货位导入备份文件.xlsx"
