#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Main entry for OpenClaw PPT Generation v1.0.

This script is intentionally simple and stable for Feishu usage:
- Input: text and/or images.
- Output: files saved under /data/share/yaq/ppt.
- Reply: only final path.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import traceback
from pathlib import Path

# Local imports without requiring package installation.
SCRIPT_DIR = Path(__file__).resolve().parent
WORKSPACE_ROOT = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))

from classify_ppt_task import classify_ppt_task  # noqa: E402
from prepare_run_dir import make_run_dir  # noqa: E402
from normalize_input import build_product_input, copy_images, write_json  # noqa: E402
from build_prompt_bundle import build_bundle  # noqa: E402
from generate_catalog_outline import generate_outline  # noqa: E402
from render_lite_pptx import render_pptx  # noqa: E402
from ppt_master_adapter import detect_ppt_master  # noqa: E402


def _write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text or "", encoding="utf-8")


def _find_latest_pptx(exports_dir: Path) -> Path | None:
    files = sorted(exports_dir.glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def _title_for_mode(mode: str) -> str:
    return "商品目录册" if mode == "catalog" else "通用PPT"


def run(args: argparse.Namespace) -> dict:
    forced_mode = args.mode or os.getenv("PPT_GENERATION_MODE", "auto")
    mode = classify_ppt_task(args.text, len(args.image or []), forced_mode)
    if mode == "unknown":
        # 第一版保持保守，但既然脚本已经被调用，就默认进入 catalog，避免飞书侧无响应。
        mode = "catalog"

    run_dir = make_run_dir(mode=mode, user=args.user, title=_title_for_mode(mode), output_root=args.output_root)
    _write(run_dir / "input" / "message.txt", args.text or "")

    copied_images = copy_images(args.image or [], run_dir / "input" / "images")
    product_input = build_product_input(run_dir, mode, args.text or "", copied_images)
    write_json(run_dir / "product_input.json", product_input)

    outline = generate_outline(run_dir / "product_input.json")
    _write(run_dir / "catalog_outline.md", outline)

    prompt_bundle = build_bundle(WORKSPACE_ROOT, run_dir, mode)
    _write(run_dir / "prompt_bundle.md", prompt_bundle)

    ppt_master_info = detect_ppt_master(WORKSPACE_ROOT)
    _write(run_dir / "logs" / "ppt_master_detect.json", json.dumps(ppt_master_info, ensure_ascii=False, indent=2))

    # v1.0 strategy:
    # If upstream PPT-master is installed and an Agent has already generated a PPT-master project, the adapter can be used.
    # In CLI smoke/local mode, use fallback to guarantee a result.
    allow_fallback = os.getenv("PPT_ALLOW_FALLBACK", "true").lower() not in {"0", "false", "no"}
    exports_dir = run_dir / "project" / "exports"
    pptx_path = None

    # If a pptx already exists, prefer it. This allows upstream PPT-master process to be inserted before this step.
    pptx_path = _find_latest_pptx(exports_dir)

    if pptx_path is None:
        if not allow_fallback:
            raise RuntimeError("未找到 PPTX，且 PPT_ALLOW_FALLBACK=false。")
        out_name = "catalog_lite_v1.pptx" if mode == "catalog" else "general_lite_v1.pptx"
        pptx_path = render_pptx(run_dir / "product_input.json", exports_dir / out_name)

    readme = f"""# PPT Generation Run\n\n- mode: {mode}\n- user: {args.user}\n- run_dir: {run_dir}\n- pptx_path: {pptx_path}\n- output_policy: 飞书只回复保存路径\n\n## 说明\n\n本次任务由 workspace-PPT-Generation v1.0 生成。\n\n如果完整 PPT-master 已接入，可由 OpenClaw + MiniMax 根据 prompt_bundle.md 生成更丰富的 PPT-master/SVG 项目；本脚本的 lite fallback 用于确保第一版流程可跑通。\n"""
    _write(run_dir / "README.md", readme)

    reply_text = f"已完成，保存路径：\n{pptx_path}"
    return {
        "ok": True,
        "mode": mode,
        "run_dir": str(run_dir),
        "pptx_path": str(pptx_path),
        "reply_text": reply_text,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="OpenClaw PPT Generation v1.0")
    parser.add_argument("--user", default="feishu-user")
    parser.add_argument("--text", default="")
    parser.add_argument("--image", action="append", default=[])
    parser.add_argument("--mode", default="auto", choices=["auto", "catalog", "general"])
    parser.add_argument("--output-root", default=None)
    parser.add_argument("--json", action="store_true", help="Print JSON only")
    args = parser.parse_args()

    try:
        result = run(args)
        if args.json:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            print(f"RUN_DIR={result['run_dir']}")
            print(f"PPTX_PATH={result['pptx_path']}")
            print(f"REPLY_TEXT={result['reply_text']}")
        return 0
    except Exception as exc:
        err = "".join(traceback.format_exception_only(type(exc), exc)).strip()
        print(f"ERROR={err}", file=sys.stderr)
        if os.getenv("DEBUG", "").lower() in {"1", "true", "yes"}:
            traceback.print_exc()
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
