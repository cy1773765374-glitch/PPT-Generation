#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Create deterministic run directory for PPT generation."""

from __future__ import annotations

import os
import re
from datetime import datetime
from pathlib import Path

DEFAULT_OUTPUT_ROOT = "/data/share/yaq/ppt"


def safe_name(value: str | None, default: str = "feishu-user", max_len: int = 40) -> str:
    value = value or default
    value = value.strip() or default
    value = re.sub(r"[\\/:*?\"<>|\s]+", "-", value)
    value = re.sub(r"-+", "-", value).strip("-")
    return value[:max_len] or default


def make_run_dir(mode: str, user: str | None, title: str | None = None, output_root: str | None = None) -> Path:
    root = Path(output_root or os.getenv("PPT_OUTPUT_ROOT", DEFAULT_OUTPUT_ROOT))
    mode = safe_name(mode, "catalog")
    user_part = safe_name(user, "feishu-user")
    title_part = safe_name(title, "PPT生成", max_len=30)
    now = datetime.now().strftime("%Y-%m-%d-%H%M%S")
    run_dir = root / mode / f"{now}-{user_part}-{title_part}"
    for sub in ["input/images", "project/exports", "project/svg_output", "project/svg_final", "logs"]:
        (run_dir / sub).mkdir(parents=True, exist_ok=True)
    return run_dir


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", default="catalog")
    parser.add_argument("--user", default="feishu-user")
    parser.add_argument("--title", default="商品目录册")
    parser.add_argument("--output-root", default=None)
    args = parser.parse_args()
    print(make_run_dir(args.mode, args.user, args.title, args.output_root))
