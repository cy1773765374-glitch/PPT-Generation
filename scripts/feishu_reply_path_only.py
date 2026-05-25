#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Format final reply for Feishu."""

from __future__ import annotations

import argparse
from pathlib import Path


def format_success(path: str) -> str:
    return f"已完成，保存路径：\n{path}"


def format_failure(run_dir: str, error: str) -> str:
    return f"生成失败，运行目录：\n{run_dir}\n错误摘要：{error}"


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--pptx-path")
    parser.add_argument("--run-dir")
    parser.add_argument("--error")
    args = parser.parse_args()
    if args.pptx_path:
        print(format_success(str(Path(args.pptx_path))))
    else:
        print(format_failure(args.run_dir or "未知", args.error or "未知错误"))
