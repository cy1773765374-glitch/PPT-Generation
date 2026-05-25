#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

mkdir -p /data/share/yaq/ppt/catalog /data/share/yaq/ppt/general

if [ ! -d .venv ]; then
  python3 -m venv .venv
fi

source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt

echo "安装完成：$ROOT_DIR"
echo "输出目录：/data/share/yaq/ppt"
