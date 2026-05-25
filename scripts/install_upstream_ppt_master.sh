#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"
mkdir -p vendor

if [ -d vendor/ppt-master/.git ]; then
  echo "vendor/ppt-master already exists, pulling latest..."
  git -C vendor/ppt-master pull --ff-only
else
  rm -rf vendor/ppt-master
  git clone https://github.com/hugohe3/ppt-master.git vendor/ppt-master
fi

echo "PPT-master installed at: $ROOT_DIR/vendor/ppt-master"
