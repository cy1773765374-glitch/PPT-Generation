#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将 generate_catalog_ppt_v11.py 的 JSON 输出转换为飞书回复文本。

用法：
python scripts/generate_catalog_ppt_v11.py --prompt "..." --json | python scripts/feishu_reply_formatter.py
"""

from __future__ import annotations

import json
import sys


def main() -> None:
    raw = sys.stdin.read().strip()
    if not raw:
        print("PPT 生成失败：未收到生成脚本输出。")
        return
    try:
        data = json.loads(raw)
    except Exception as exc:
        print(f"PPT 生成失败：生成脚本输出不是合法 JSON：{exc}")
        return

    if not data.get("ok"):
        print(f"PPT 生成失败：{data.get('error', '未知错误')}")
        return

    print(
        "已完成，PPT 文件：\n"
        f"{data.get('windows_path')}\n\n"
        "服务器路径：\n"
        f"{data.get('pptx_path')}\n\n"
        f"已完成 · {data.get('page_count')}页 · 耗时 {data.get('elapsed_seconds')}s"
    )


if __name__ == "__main__":
    main()
