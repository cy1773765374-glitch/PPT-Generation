#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Smoke test for workspace-PPT-Generation."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main() -> int:
    cmd = [
        sys.executable,
        str(ROOT / "scripts" / "run_ppt_generation.py"),
        "--user",
        "smoke-test",
        "--text",
        "根据这个产品描述生成一个商品目录册PPT：冰蓝水花包装盒，适合夏季饮品包装，清爽、年轻、可定制印刷。",
        "--json",
    ]
    proc = subprocess.run(cmd, cwd=str(ROOT), text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if proc.returncode != 0:
        print(proc.stderr)
        return proc.returncode
    data = json.loads(proc.stdout)
    pptx = Path(data["pptx_path"])
    print(json.dumps(data, ensure_ascii=False, indent=2))
    if not pptx.exists():
        print(f"PPTX not found: {pptx}", file=sys.stderr)
        return 3
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
