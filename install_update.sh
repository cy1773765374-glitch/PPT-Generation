#!/usr/bin/env bash
set -euo pipefail

WORKDIR="${1:-$HOME/.openclaw/workspace-PPT-Generation}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

mkdir -p "$WORKDIR"
mkdir -p "$WORKDIR/PPT-master"

# 清理旧版本入口，避免飞书继续调用 v12/v13/v14/v15 的旧逻辑。
rm -f "$WORKDIR"/scripts/generate_catalog_ppt_v12.py \
      "$WORKDIR"/scripts/generate_catalog_ppt_v13.py \
      "$WORKDIR"/scripts/generate_catalog_ppt_v14.py \
      "$WORKDIR"/scripts/generate_catalog_ppt_v15.py || true

# 不使用 --delete，且显式排除 PPT-master，确保已有 PPT-master 不被删除或覆盖。
rsync -a \
  --exclude '.venv/' \
  --exclude '.env' \
  --exclude 'PPT-master/' \
  "$SCRIPT_DIR/" "$WORKDIR/"

cd "$WORKDIR"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

if [ ! -f .env ]; then
  cp .env.example .env
fi

chmod +x scripts/*.py scripts/*.sh 2>/dev/null || true

echo "已安装/更新到：$WORKDIR"
echo "PPT-master 已保留：$WORKDIR/PPT-master"
echo "主入口：python scripts/generate_catalog_ppt.py --prompt '...' --json"
