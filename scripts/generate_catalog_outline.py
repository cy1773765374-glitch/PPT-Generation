#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate a conservative outline from normalized input.

The outline is not a fixed template. It records a dynamic first-pass plan and
leaves room for OpenClaw + MiniMax to expand it.
"""

from __future__ import annotations

import json
from pathlib import Path


def generate_outline(product_json_path: Path) -> str:
    data = json.loads(product_json_path.read_text(encoding="utf-8"))
    mode = data.get("task_type", "catalog")
    images = data.get("images") or []
    user_message = data.get("user_message") or ""

    if mode == "catalog":
        suggested = []
        suggested.append("封面或主题页：建立商品目录册主题。")
        if images:
            suggested.append("商品主视觉页：优先使用用户上传图片。")
        if user_message:
            suggested.append("商品定位与卖点页：根据用户文字提炼。")
        suggested.append("规格/待确认页：明确列出未确认事实参数。")
        suggested.append("结尾/沟通页：提示后续补充资料。")

        body = [
            "# 商品目录册动态规划初稿",
            "",
            "本文件不是固定模板，只是根据当前输入生成的初步规划。OpenClaw + MiniMax 可根据实际内容继续调整页面数量、顺序和视觉风格。",
            "",
            "## 输入摘要",
            f"- 文本长度：{len(user_message)}",
            f"- 图片数量：{len(images)}",
            "",
            "## 建议页面组件",
        ]
        body.extend([f"- {x}" for x in suggested])
        body.extend([
            "",
            "## 事实边界",
            "- 未提供价格：待确认。",
            "- 未提供材质：待确认。",
            "- 未提供尺寸：待确认。",
            "- 未提供认证：待确认。",
        ])
        return "\n".join(body) + "\n"

    return "# 通用 PPT 规划初稿\n\n请结合 PPT-master 原始 SKILL.md 动态规划。\n"


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--product-json", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    outline = generate_outline(Path(args.product_json))
    Path(args.output).write_text(outline, encoding="utf-8")
    print(args.output)
