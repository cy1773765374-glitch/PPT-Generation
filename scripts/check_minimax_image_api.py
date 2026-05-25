#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MiniMax image_generation 预检：检查密钥、端点归一化、DNS、可选地发起一次极小生图请求。"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None

# 复用主脚本里的配置读取与 URL 修正逻辑。
SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
from generate_catalog_ppt_v13 import (  # noqa: E402
    MiniMaxImageClient,
    assert_hostname_resolves,
    get_minimax_api_key,
    normalize_minimax_image_api_url,
)


def main() -> None:
    parser = argparse.ArgumentParser(description="Check MiniMax image_generation connectivity")
    parser.add_argument("--live", action="store_true", help="执行一次真实 MiniMax 生图请求并保存 check_minimax.png")
    parser.add_argument("--out", default="check_minimax.png", help="--live 输出图片路径")
    args = parser.parse_args()

    if load_dotenv:
        load_dotenv()

    configured_url = (
        os.getenv("MINIMAX_IMAGE_API_URL", "").strip()
        or os.getenv("MINIMAX_IMAGE_BASE_URL", "").strip()
        or os.getenv("MINIMAX_BASE_URL", "").strip()
        or "https://api.minimax.io/v1/image_generation"
    )
    normalized_url = normalize_minimax_image_api_url(configured_url)
    api_key = get_minimax_api_key()

    result = {
        "ok": False,
        "configured_url": configured_url,
        "normalized_url": normalized_url,
        "has_api_key": bool(api_key),
        "dns_ok": False,
        "live_ok": False,
        "output": "",
        "error": "",
    }

    try:
        assert_hostname_resolves(normalized_url)
        result["dns_ok"] = True
        if args.live:
            client = MiniMaxImageClient()
            out = Path(args.out).expanduser().resolve()
            client.generate_to_file(
                "A clean commercial product catalog photo of stainless steel kitchen tableware on a white background, no text, no logo",
                out,
                "1:1",
            )
            result["live_ok"] = True
            result["output"] = str(out)
        result["ok"] = bool(api_key) and result["dns_ok"] and ((not args.live) or result["live_ok"])
        if not api_key:
            result["error"] = "MINIMAX_API_KEY 未配置，也未能从 ~/.openclaw/openclaw.json 读取。"
    except Exception as exc:
        result["error"] = str(exc)

    print(json.dumps(result, ensure_ascii=False, indent=2))
    sys.exit(0 if result["ok"] else 1)


if __name__ == "__main__":
    main()
