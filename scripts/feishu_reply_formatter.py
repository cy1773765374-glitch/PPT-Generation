#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将 generate_catalog_ppt_v13.py 的 JSON 输出转换为飞书回复文本。

v1.3 规则：飞书只回复 Windows/Samba 本地路径，不回复服务器路径、页数、耗时或“已完成”。
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

    # 只回本地映射盘路径。不要加“已完成”、不要加服务器路径、不要加耗时，避免飞书里重复输出。
    print(data.get("reply_text") or data.get("windows_path") or "")


if __name__ == "__main__":
    main()
