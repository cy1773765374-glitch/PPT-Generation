#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")/.."
python scripts/generate_catalog_ppt_v11.py \
  --prompt "生成一个厨房餐具相关的3页的商品目录册PPT" \
  --sender-name "陈玉" \
  --sender-open-id "ou_6e056040c28331827575c0061644569c" \
  --root "${PPT_CATALOG_ROOT:-/tmp/ppt_catalog_test}" \
  --json
