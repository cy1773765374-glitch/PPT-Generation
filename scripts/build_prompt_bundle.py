#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build prompt bundle from profiles and normalized input."""

from __future__ import annotations

import json
from pathlib import Path


def read_if_exists(path: Path) -> str:
    return path.read_text(encoding="utf-8") if path.exists() else ""


def build_bundle(workspace_root: Path, run_dir: Path, mode: str) -> str:
    common_files = [
        workspace_root / "profiles/common/feishu_input.md",
        workspace_root / "profiles/common/server_output.md",
        workspace_root / "profiles/common/pptx_quality.md",
    ]
    if mode == "catalog":
        mode_files = [
            workspace_root / "profiles/catalog/prompt.md",
            workspace_root / "profiles/catalog/content_planning.md",
            workspace_root / "profiles/catalog/quality_rules.md",
            workspace_root / "profiles/catalog/output_policy.md",
            workspace_root / "templates/product_catalog_v1/design_principles.md",
            workspace_root / "templates/product_catalog_v1/page_components.md",
        ]
    else:
        mode_files = [workspace_root / "profiles/general/prompt.md"]

    product_json = read_if_exists(run_dir / "product_input.json")
    parts: list[str] = ["# Prompt Bundle\n"]
    for file in common_files + mode_files:
        parts.append(f"\n---\n\n# Source: {file.relative_to(workspace_root)}\n\n")
        parts.append(read_if_exists(file))
    parts.append("\n---\n\n# Normalized Input\n\n```json\n")
    try:
        parsed = json.loads(product_json)
        product_json = json.dumps(parsed, ensure_ascii=False, indent=2)
    except Exception:
        pass
    parts.append(product_json)
    parts.append("\n```\n")
    return "".join(parts)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--workspace-root", default=".")
    parser.add_argument("--run-dir", required=True)
    parser.add_argument("--mode", default="catalog")
    args = parser.parse_args()

    bundle = build_bundle(Path(args.workspace_root).resolve(), Path(args.run_dir).resolve(), args.mode)
    out = Path(args.run_dir) / "prompt_bundle.md"
    out.write_text(bundle, encoding="utf-8")
    print(out)
