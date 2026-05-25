#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Task classifier for workspace-PPT-Generation v1.0."""

from __future__ import annotations

CATALOG_KEYWORDS = [
    "目录册",
    "商品目录",
    "产品目录",
    "商品册",
    "产品册",
    "选品册",
    "产品画册",
    "catalog",
    "产品介绍册",
    "实物",
    "商品介绍",
    "产品介绍",
    "根据图片做产品",
    "根据图片生成商品",
]

GENERAL_PPT_KEYWORDS = [
    "ppt",
    "PPT",
    "演示文稿",
    "汇报",
    "课件",
    "培训",
    "方案书",
    "方案PPT",
    "presentation",
]


def classify_ppt_task(text: str | None, image_count: int = 0, forced_mode: str = "auto") -> str:
    """Return catalog/general/unknown.

    forced_mode can be auto/catalog/general.
    """
    forced_mode = (forced_mode or "auto").strip().lower()
    if forced_mode in {"catalog", "general"}:
        return forced_mode

    raw = text or ""
    lower = raw.lower()

    if any(k.lower() in lower for k in CATALOG_KEYWORDS):
        return "catalog"

    # 图片 + 明确 PPT/产品倾向，优先目录册。
    if image_count > 0 and any(k.lower() in lower for k in ["产品", "商品", "目录", "介绍", "画册"]):
        return "catalog"

    if any(k.lower() in lower for k in GENERAL_PPT_KEYWORDS):
        return "general"

    # 只有图片时，第一版不自动生成，避免误触发；由 Agent 上层决定是否调用。
    return "unknown"


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--text", default="")
    parser.add_argument("--image-count", type=int, default=0)
    parser.add_argument("--mode", default="auto")
    args = parser.parse_args()
    print(classify_ppt_task(args.text, args.image_count, args.mode))
