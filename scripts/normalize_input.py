#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Normalize Feishu text/images into product_input.json."""

from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import Iterable


def copy_images(image_paths: Iterable[str], dest_dir: Path) -> list[str]:
    dest_dir.mkdir(parents=True, exist_ok=True)
    copied: list[str] = []
    for idx, src in enumerate(image_paths, start=1):
        if not src:
            continue
        src_path = Path(src).expanduser().resolve()
        if not src_path.exists():
            continue
        suffix = src_path.suffix.lower() or ".jpg"
        dst = dest_dir / f"image_{idx:02d}{suffix}"
        shutil.copy2(src_path, dst)
        copied.append(str(dst))
    return copied


def build_product_input(run_dir: Path, mode: str, user_text: str, image_paths: list[str]) -> dict:
    image_count = len(image_paths)
    source_type = "text_image" if user_text and image_count else "image" if image_count else "text"

    # 第一版只做保守结构化，真正的内容扩展由 OpenClaw + MiniMax 根据 prompt_bundle 动态完成。
    product = {
        "name": "待命名商品",
        "category": "待识别",
        "description": user_text.strip() or "用户提供了商品图片，需根据图片生成商品目录册初稿。",
        "selling_points": ["待根据输入提炼", "待根据输入提炼", "待根据输入提炼"],
        "specs": {
            "material": "待确认",
            "size": "待确认",
            "color": "待确认",
            "usage": "待确认",
            "price": "待确认",
            "moq": "待确认",
        },
        "image_paths": image_paths,
        "remarks": "AI识别结果仅供初稿使用；未明确提供的信息必须标记为待确认。",
    }

    return {
        "version": "1.0",
        "task_type": mode,
        "source_type": source_type,
        "title": "商品目录册" if mode == "catalog" else "PPT生成",
        "user_message": user_text,
        "products": [product] if mode == "catalog" else [],
        "images": image_paths,
        "output_root": "/data/share/yaq/ppt",
        "run_dir": str(run_dir),
    }


def write_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--run-dir", required=True)
    parser.add_argument("--mode", default="catalog")
    parser.add_argument("--text", default="")
    parser.add_argument("--image", action="append", default=[])
    args = parser.parse_args()

    run_dir = Path(args.run_dir)
    image_paths = copy_images(args.image, run_dir / "input" / "images")
    data = build_product_input(run_dir, args.mode, args.text, image_paths)
    write_json(run_dir / "product_input.json", data)
    print(run_dir / "product_input.json")
