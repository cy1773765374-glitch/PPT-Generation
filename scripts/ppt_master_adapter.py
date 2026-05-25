#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Adapter for upstream PPT-master.

v1.0 keeps upstream PPT-master intact. This adapter only detects whether the
expected upstream scripts exist and can optionally run the final export scripts
when a PPT-master project has already been generated.
"""

from __future__ import annotations

import os
import subprocess
from pathlib import Path


def get_ppt_master_dir(workspace_root: Path) -> Path:
    return Path(os.getenv("PPT_MASTER_DIR", str(workspace_root / "vendor" / "ppt-master"))).resolve()


def detect_ppt_master(workspace_root: Path) -> dict:
    base = get_ppt_master_dir(workspace_root)
    script_base = base / "skills" / "ppt-master" / "scripts"
    return {
        "base": str(base),
        "skill_md": str(base / "skills" / "ppt-master" / "SKILL.md"),
        "finalize_svg": str(script_base / "finalize_svg.py"),
        "svg_to_pptx": str(script_base / "svg_to_pptx.py"),
        "available": (script_base / "finalize_svg.py").exists() and (script_base / "svg_to_pptx.py").exists(),
    }


def run_export_scripts(workspace_root: Path, project_path: Path) -> None:
    info = detect_ppt_master(workspace_root)
    if not info["available"]:
        raise RuntimeError("未检测到完整 PPT-master scripts，无法运行 upstream export。")
    subprocess.run(["python3", info["finalize_svg"], str(project_path)], check=True)
    subprocess.run(["python3", info["svg_to_pptx"], str(project_path)], check=True)


if __name__ == "__main__":
    import argparse
    import json

    parser = argparse.ArgumentParser()
    parser.add_argument("--workspace-root", default=".")
    parser.add_argument("--project-path", default=None)
    parser.add_argument("--run-export", action="store_true")
    args = parser.parse_args()
    root = Path(args.workspace_root).resolve()
    if args.run_export:
        if not args.project_path:
            raise SystemExit("--project-path is required when --run-export is set")
        run_export_scripts(root, Path(args.project_path).resolve())
    else:
        print(json.dumps(detect_ppt_master(root), ensure_ascii=False, indent=2))
