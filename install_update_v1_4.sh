#!/usr/bin/env bash
set -euo pipefail

WORKDIR="${1:-$HOME/.openclaw/workspace-PPT-Generation}"
mkdir -p "$WORKDIR"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
rsync -a --delete \
  --exclude '.venv' \
  --exclude '.env' \
  "$SCRIPT_DIR/" "$WORKDIR/"

cd "$WORKDIR"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

if [ ! -f .env ]; then
  cp .env.example .env
fi

echo "已安装/更新到：$WORKDIR"
echo "请检查 .env：PPT_TIMEZONE=Asia/Shanghai、PPT_IMAGE_MODE=minimax、MINIMAX_IMAGE_API_URL=https://api.minimax.com/v1/image_generation、MINIMAX_API_KEY。"
