#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")/.."
# 离线测试用 placeholder；生产飞书链路必须使用默认 PPT_IMAGE_MODE=minimax。
PPT_IMAGE_MODE=placeholder python scripts/generate_catalog_ppt_v12.py \
  --prompt "生成一个厨房餐具相关的3页的商品目录册PPT" \
  --sender-name "陈玉" \
  --sender-open-id "ou_6e056040c28331827575c0061644569c" \
  --root "${PPT_CATALOG_ROOT:-/tmp/ppt_catalog_test}" \
  --json
