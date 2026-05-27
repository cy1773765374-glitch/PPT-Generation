#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OpenClaw ↔ PPT-master 适配层。

目标：
- OpenClaw 的 DeckPlan / SlidePlan 仍然负责业务理解与商品目录册规划。
- PPT-master 负责 SVG 质量检查、后处理与 SVG → native editable PPTX 导出。
- 当服务器没有完整 vendor/ppt-master 时，主程序可以按配置降级到旧的 python-pptx 渲染器。
"""

from __future__ import annotations

import html
import json
import os
import re
import shutil
import subprocess
import sys
from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple


class PPTMasterUnavailable(RuntimeError):
    """没有可用的 PPT-master vendor 或关键脚本缺失。"""


class PPTMasterPipelineError(RuntimeError):
    """PPT-master 管线执行失败。"""


STYLE_HEX: Dict[str, Dict[str, str]] = {
    "commercial": {
        "bg": "#F8FAFC",
        "panel": "#FFFFFF",
        "accent": "#2563EB",
        "accent2": "#0EA5E9",
        "accent3": "#10B981",
        "text": "#0F172A",
        "muted": "#475569",
        "line": "#E2E8F0",
    },
    "minimal": {
        "bg": "#FAFAF9",
        "panel": "#FFFFFF",
        "accent": "#1E293B",
        "accent2": "#64748B",
        "accent3": "#94A3B8",
        "text": "#111827",
        "muted": "#52525B",
        "line": "#E7E5E4",
    },
    "luxury": {
        "bg": "#FAF7F2",
        "panel": "#FFFCF7",
        "accent": "#92400E",
        "accent2": "#B45309",
        "accent3": "#D97706",
        "text": "#1C1917",
        "muted": "#57534E",
        "line": "#E7D7C6",
    },
    "playful": {
        "bg": "#FFF7ED",
        "panel": "#FFFFFF",
        "accent": "#EA580C",
        "accent2": "#D946EF",
        "accent3": "#0EA5E9",
        "text": "#1E293B",
        "muted": "#475569",
        "line": "#FED7AA",
    },
    "festive": {
        "bg": "#FFFBEB",
        "panel": "#FFFFFF",
        "accent": "#DC2626",
        "accent2": "#F59E0B",
        "accent3": "#16A34A",
        "text": "#27272A",
        "muted": "#52525B",
        "line": "#FDE68A",
    },
    "nature": {
        "bg": "#F6F8F0",
        "panel": "#FFFFFB",
        "accent": "#4C7856",
        "accent2": "#846F4A",
        "accent3": "#65A30D",
        "text": "#1F2937",
        "muted": "#4B5563",
        "line": "#D8E2C2",
    },
    "tech": {
        "bg": "#EFF6FF",
        "panel": "#FFFFFF",
        "accent": "#1D4ED8",
        "accent2": "#06B6D4",
        "accent3": "#6366F1",
        "text": "#0F172A",
        "muted": "#334155",
        "line": "#BFDBFE",
    },
}

PAGE_W = 1280
PAGE_H = 720


def _obj_to_dict(obj: Any) -> Dict[str, Any]:
    if obj is None:
        return {}
    if isinstance(obj, dict):
        return obj
    if is_dataclass(obj):
        return asdict(obj)
    data: Dict[str, Any] = {}
    for name in dir(obj):
        if name.startswith("_"):
            continue
        value = getattr(obj, name)
        if callable(value):
            continue
        try:
            json.dumps(value, ensure_ascii=False, default=str)
            data[name] = value
        except Exception:
            data[name] = str(value)
    return data


def _get(obj: Any, name: str, default: Any = None) -> Any:
    if isinstance(obj, dict):
        return obj.get(name, default)
    return getattr(obj, name, default)


def _safe_text(value: Any) -> str:
    return html.escape(str(value or ""), quote=True)


def _wrap_text(text: str, max_chars: int, max_lines: int) -> List[str]:
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    if not text:
        return []
    has_cjk = bool(re.search(r"[\u3400-\u4dbf\u4e00-\u9fff]", text))
    lines: List[str] = []
    if has_cjk:
        buf = ""
        for ch in text:
            buf += ch
            if len(buf) >= max_chars:
                lines.append(buf)
                buf = ""
                if len(lines) >= max_lines:
                    break
        if buf and len(lines) < max_lines:
            lines.append(buf)
    else:
        words = text.split(" ")
        buf = ""
        for word in words:
            candidate = word if not buf else f"{buf} {word}"
            if len(candidate) <= max_chars:
                buf = candidate
            else:
                if buf:
                    lines.append(buf)
                buf = word
                if len(lines) >= max_lines:
                    break
        if buf and len(lines) < max_lines:
            lines.append(buf)
    if len(lines) > max_lines:
        lines = lines[:max_lines]
    if lines and len(lines) == max_lines and len(text) > sum(len(x) for x in lines):
        lines[-1] = lines[-1].rstrip(" .，。") + "…"
    return lines


def _text_block(x: int, y: int, width_chars: int, lines: Iterable[str], size: int, color: str, weight: int = 400, line_gap: int = 1, max_lines: int = 6) -> str:
    output: List[str] = []
    yy = y
    for idx, line in enumerate(list(lines)[:max_lines]):
        escaped = _safe_text(line)
        output.append(
            f'<text x="{x}" y="{yy}" font-family="Arial, Microsoft YaHei, sans-serif" '
            f'font-size="{size}" font-weight="{weight}" fill="{color}">{escaped}</text>'
        )
        yy += int(size * 1.35) + line_gap
    return "\n".join(output)


def _title_block(x: int, y: int, text: str, size: int, color: str, max_chars: int = 26, max_lines: int = 2) -> str:
    lines = _wrap_text(text, max_chars=max_chars, max_lines=max_lines)
    return _text_block(x, y, max_chars, lines, size=size, color=color, weight=700, max_lines=max_lines)


def _bullet_lines(items: Sequence[str], max_item_chars: int = 55, max_items: int = 4) -> List[str]:
    lines: List[str] = []
    for item in list(items)[:max_items]:
        wrapped = _wrap_text(str(item), max_chars=max_item_chars, max_lines=2)
        if not wrapped:
            continue
        lines.append("• " + wrapped[0])
        lines.extend("  " + part for part in wrapped[1:])
    return lines


def _rounded_rect(x: int, y: int, w: int, h: int, fill: str, stroke: str = "none", sw: int = 1, rx: int = 20, opacity: Optional[float] = None) -> str:
    opacity_attr = f' fill-opacity="{opacity}"' if opacity is not None else ""
    return f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"{opacity_attr}/>'


def _section_card(x: int, y: int, w: int, h: int, title: str, items: Sequence[str], palette: Dict[str, str], accent: str) -> str:
    parts = [
        _rounded_rect(x, y, w, h, palette["panel"], palette["line"], 1, 18),
        f'<rect x="{x}" y="{y}" width="8" height="{h}" rx="4" fill="{accent}"/>',
        _text_block(x + 24, y + 32, 30, _wrap_text(title, 28, 1), 19, palette["text"], 700, max_lines=1),
        _text_block(x + 24, y + 68, 58, _bullet_lines(items, max_item_chars=48, max_items=4), 16, palette["muted"], 400, max_lines=7),
    ]
    return "\n".join(parts)


def _image_element(asset_rel: str, x: int, y: int, w: int, h: int, palette: Dict[str, str]) -> str:
    if not asset_rel:
        return "\n".join([
            _rounded_rect(x, y, w, h, "#EEF2FF", palette["line"], 1, 24),
            f'<circle cx="{x + int(w * 0.28)}" cy="{y + int(h * 0.40)}" r="70" fill="{palette["accent2"]}" fill-opacity="0.18"/>',
            f'<circle cx="{x + int(w * 0.64)}" cy="{y + int(h * 0.55)}" r="96" fill="{palette["accent"]}" fill-opacity="0.14"/>',
        ])
    safe_href = _safe_text(asset_rel)
    return "\n".join([
        _rounded_rect(x, y, w, h, palette["panel"], palette["line"], 1, 24),
        f'<image href="{safe_href}" xlink:href="{safe_href}" x="{x + 8}" y="{y + 8}" width="{w - 16}" height="{h - 16}" preserveAspectRatio="xMidYMid slice"/>',
    ])


def _decorations(style_family: str, palette: Dict[str, str]) -> str:
    if style_family == "minimal":
        return f'<line x1="76" y1="650" x2="1204" y2="650" stroke="{palette["line"]}" stroke-width="2"/>'
    if style_family == "luxury":
        return "\n".join([
            f'<circle cx="1125" cy="96" r="78" fill="{palette["accent3"]}" fill-opacity="0.10"/>',
            f'<line x1="82" y1="105" x2="214" y2="105" stroke="{palette["accent"]}" stroke-width="5"/>',
        ])
    if style_family == "playful":
        return "\n".join([
            f'<circle cx="1100" cy="96" r="38" fill="{palette["accent2"]}" fill-opacity="0.22"/>',
            f'<circle cx="1168" cy="142" r="24" fill="{palette["accent3"]}" fill-opacity="0.24"/>',
            f'<circle cx="100" cy="610" r="28" fill="{palette["accent"]}" fill-opacity="0.18"/>',
        ])
    if style_family == "festive":
        return "\n".join([
            f'<circle cx="1140" cy="90" r="72" fill="{palette["accent2"]}" fill-opacity="0.16"/>',
            f'<circle cx="1205" cy="170" r="35" fill="{palette["accent"]}" fill-opacity="0.16"/>',
            f'<path d="M76 90 C210 126, 330 54, 460 92" fill="none" stroke="{palette["accent"]}" stroke-width="4" stroke-opacity="0.35"/>',
        ])
    if style_family == "nature":
        return "\n".join([
            f'<ellipse cx="1110" cy="104" rx="118" ry="58" fill="{palette["accent3"]}" fill-opacity="0.13"/>',
            f'<ellipse cx="102" cy="615" rx="72" ry="38" fill="{palette["accent"]}" fill-opacity="0.10"/>',
        ])
    if style_family == "tech":
        return "\n".join([
            f'<rect x="1015" y="58" width="145" height="145" rx="28" fill="{palette["accent2"]}" fill-opacity="0.12"/>',
            f'<line x1="960" y1="640" x2="1198" y2="640" stroke="{palette["accent"]}" stroke-width="3" stroke-opacity="0.25"/>',
        ])
    return f'<circle cx="1130" cy="102" r="88" fill="{palette["accent2"]}" fill-opacity="0.12"/>'


def _footer(plan: Any, page_no: int, palette: Dict[str, str]) -> str:
    brand = _get(plan, "brand_or_company", "")
    theme = _get(plan, "theme", "Product Catalog")
    page_count = _get(plan, "page_count", page_no)
    text = f"{brand} · {theme} · Page {page_no}/{page_count}" if brand else f"{theme} · Page {page_no}/{page_count}"
    return f'<text x="76" y="684" font-family="Arial, Microsoft YaHei, sans-serif" font-size="14" fill="{palette["muted"]}">{_safe_text(text)}</text>'


def _slide_shell(content: str, style_family: str, palette: Dict[str, str]) -> str:
    return f'''<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="{PAGE_W}" height="{PAGE_H}" viewBox="0 0 {PAGE_W} {PAGE_H}">
  <rect width="{PAGE_W}" height="{PAGE_H}" fill="{palette['bg']}"/>
  {_decorations(style_family, palette)}
  {content}
</svg>
'''


def _asset_rel_for_slide(slide: Any, asset_by_key: Dict[str, str]) -> str:
    key = _get(slide, "image_key", "")
    return asset_by_key.get(key, "")


def _render_cover(plan: Any, slide: Any, page_no: int, asset_rel: str, palette: Dict[str, str], style_family: str) -> str:
    title = _get(slide, "title", _get(plan, "theme", "Product Catalog"))
    subtitle = _get(slide, "subtitle", "Product Catalog")
    brand = _get(plan, "brand_or_company", "")
    sections = list(_get(slide, "sections", []) or [])
    chips: List[str] = []
    for section in sections[:2]:
        chips.extend(str(x) for x in section.get("items", [])[:2])
    content = [
        _image_element(asset_rel, 720, 116, 470, 390, palette),
        f'<text x="78" y="112" font-family="Arial, Microsoft YaHei, sans-serif" font-size="18" font-weight="700" fill="{palette["accent"]}">{_safe_text(brand or "CATALOG")}</text>',
        _title_block(76, 210, title, 54, palette["text"], 20, 3),
        _text_block(80, 390, 38, _wrap_text(subtitle, 36, 2), 25, palette["muted"], 500, max_lines=2),
    ]
    y = 488
    for idx, chip in enumerate(chips[:3]):
        content.append(_rounded_rect(80, y + idx * 48, 470, 34, palette["panel"], palette["line"], 1, 17))
        content.append(f'<text x="104" y="{y + 23 + idx * 48}" font-family="Arial, Microsoft YaHei, sans-serif" font-size="16" fill="{palette["muted"]}">{_safe_text(chip)}</text>')
    content.append(_footer(plan, page_no, palette))
    return _slide_shell("\n".join(content), style_family, palette)


def _render_company(plan: Any, slide: Any, page_no: int, asset_rel: str, palette: Dict[str, str], style_family: str) -> str:
    title = _get(slide, "title", _get(plan, "brand_or_company", "Company"))
    bullets = list(_get(slide, "bullets", []) or [])
    sections = list(_get(slide, "sections", []) or [])
    content = [
        f'<text x="76" y="96" font-family="Arial, Microsoft YaHei, sans-serif" font-size="18" font-weight="700" fill="{palette["accent"]}">COMPANY INTRODUCTION</text>',
        _title_block(74, 160, title, 44, palette["text"], 24, 2),
        _text_block(78, 278, 52, _wrap_text(" ".join(bullets), 58, 4), 20, palette["muted"], 400, max_lines=4),
        _image_element(asset_rel, 720, 106, 470, 340, palette),
    ]
    y = 480
    for idx, section in enumerate(sections[:2]):
        content.append(_section_card(720 if idx == 1 else 76, y, 470 if idx == 1 else 570, 120, section.get("title", ""), section.get("items", []), palette, [palette["accent"], palette["accent2"]][idx]))
    content.append(_footer(plan, page_no, palette))
    return _slide_shell("\n".join(content), style_family, palette)


def _render_category(plan: Any, slide: Any, page_no: int, asset_rel: str, palette: Dict[str, str], style_family: str) -> str:
    title = _get(slide, "title", "Category")
    subtitle = _get(slide, "subtitle", "")
    variant = _get(slide, "layout_variant", "image_left") or "image_left"
    sections = list(_get(slide, "sections", []) or [])[:4]
    accents = [palette["accent"], palette["accent2"], palette["accent3"], palette["accent"]]
    content = [
        f'<text x="76" y="75" font-family="Arial, Microsoft YaHei, sans-serif" font-size="15" font-weight="700" fill="{palette["accent"]}">CATEGORY SHOWCASE</text>',
        _title_block(74, 132, title, 42, palette["text"], 24, 2),
        _text_block(78, 218, 52, _wrap_text(subtitle, 48, 1), 17, palette["muted"], 500, max_lines=1),
    ]
    if variant == "image_right":
        content.append(_image_element(asset_rel, 762, 112, 430, 408, palette))
        card_positions = [(76, 282, 310, 132), (410, 282, 310, 132), (76, 438, 310, 132), (410, 438, 310, 132)]
    elif variant == "image_top":
        content.append(_image_element(asset_rel, 682, 80, 510, 242, palette))
        card_positions = [(76, 360, 265, 145), (360, 360, 265, 145), (644, 360, 265, 145), (928, 360, 265, 145)]
    else:
        content.append(_image_element(asset_rel, 76, 282, 470, 292, palette))
        card_positions = [(590, 120, 290, 132), (902, 120, 290, 132), (590, 282, 290, 132), (902, 282, 290, 132)]
    if not sections:
        sections = [{"title": "Product Direction", "items": [title]}]
    for idx, section in enumerate(sections):
        x, y, w, h = card_positions[min(idx, len(card_positions) - 1)]
        content.append(_section_card(x, y, w, h, section.get("title", ""), section.get("items", []), palette, accents[idx % len(accents)]))
    content.append(_footer(plan, page_no, palette))
    return _slide_shell("\n".join(content), style_family, palette)


def render_slide_svg(plan: Any, slide: Any, page_no: int, asset_rel: str) -> str:
    style_family = str(_get(plan, "style_family", "commercial") or "commercial")
    palette = STYLE_HEX.get(style_family, STYLE_HEX["commercial"])
    layout = _get(slide, "layout", "category_showcase")
    if layout == "cover_catalog":
        return _render_cover(plan, slide, page_no, asset_rel, palette, style_family)
    if layout == "company_intro":
        return _render_company(plan, slide, page_no, asset_rel, palette, style_family)
    return _render_category(plan, slide, page_no, asset_rel, palette, style_family)


def _candidate_script_paths(root: Path, name: str) -> List[Path]:
    """兼容 PPT-master 不同安装形态。

    当前 Hugo He 的仓库把脚本放在：
        <repo>/skills/ppt-master/scripts/<name>
    早期/裁剪包或部分本地适配包可能放在：
        <repo>/scripts/<name>
    OpenClaw 适配层两种都接受，避免把完整仓库误判为缺失。
    """
    return [
        root / "skills" / "ppt-master" / "scripts" / name,
        root / "scripts" / name,
    ]


def _find_script(root: Path, name: str) -> Optional[Path]:
    for path in _candidate_script_paths(root, name):
        if path.exists():
            return path.resolve()
    return None


def discover_pptmaster_root(workspace_root: Path) -> Path:
    candidates = []
    env_root = os.getenv("PPT_MASTER_ROOT", "").strip()
    if env_root:
        candidates.append(Path(env_root).expanduser())
    candidates.extend([
        workspace_root / "vendor" / "ppt-master",
        workspace_root / "PPT-master",
        Path.home() / ".openclaw" / "workspace-PPT-Generation" / "vendor" / "ppt-master",
    ])
    checked: List[str] = []
    for root in candidates:
        root = root.expanduser()
        checked.extend(str(p) for p in _candidate_script_paths(root, "svg_to_pptx.py"))
        if _find_script(root, "svg_to_pptx.py") is not None:
            return root.resolve()
    raise PPTMasterUnavailable(
        "未找到完整 PPT-master：需要存在 svg_to_pptx.py。"
        "当前适配器会检查 skills/ppt-master/scripts/ 和 scripts/ 两种位置。"
        f"已检查：{', '.join(checked)}。"
        "请确认 PPT_MASTER_ROOT 指向 /home/cy/.openclaw/workspace-PPT-Generation/vendor/ppt-master。"
    )


def _copy_assets(assets: Sequence[Any], project_dir: Path) -> Dict[str, str]:
    asset_dir = project_dir / "images"
    asset_dir.mkdir(parents=True, exist_ok=True)
    mapping: Dict[str, str] = {}
    for asset in assets:
        key = str(_get(asset, "key", "") or "")
        path = Path(str(_get(asset, "path", "") or ""))
        if not key or not path.exists():
            continue
        suffix = path.suffix.lower() or ".png"
        target = asset_dir / f"{key}{suffix}"
        shutil.copy2(path, target)
        mapping[key] = f"../images/{target.name}"
    return mapping


def write_pptmaster_project(plan: Any, assets: Sequence[Any], run_dir: Path) -> Path:
    project_dir = run_dir / "pptmaster_project"
    if project_dir.exists():
        shutil.rmtree(project_dir)
    for sub in ["svg_output", "svg_final", "images", "assets", "exports", "backup", "notes", "sources"]:
        (project_dir / sub).mkdir(parents=True, exist_ok=True)

    asset_by_key = _copy_assets(assets, project_dir)
    plan_dict = _obj_to_dict(plan)
    slides = list(_get(plan, "slides", []) or [])

    (project_dir / "deck_plan.json").write_text(json.dumps(plan_dict, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    write_design_files(plan, project_dir)

    total_parts: List[str] = []
    manifest_slides: List[Dict[str, Any]] = []
    for idx, slide in enumerate(slides, start=1):
        raw_title = str(_get(slide, "title", f"slide {idx}") or f"slide {idx}")
        slug = re.sub(r"[^A-Za-z0-9_\-]+", "_", raw_title).strip("_").lower()[:36] or f"slide_{idx:02d}"
        filename = f"{idx:02d}_{slug}.svg"
        svg = render_slide_svg(plan, slide, idx, _asset_rel_for_slide(slide, asset_by_key))
        (project_dir / "svg_output" / filename).write_text(svg, encoding="utf-8")
        note_title = Path(filename).stem
        total_parts.append(f"# {note_title}\n\n{raw_title}\n\n---\n")
        manifest_slides.append({
            "page": idx,
            "filename": filename,
            "title": raw_title,
            "layout": _get(slide, "layout", ""),
            "layout_variant": _get(slide, "layout_variant", ""),
            "image_key": _get(slide, "image_key", ""),
        })
    (project_dir / "total.md").write_text("\n".join(total_parts), encoding="utf-8")
    (project_dir / "README.md").write_text(
        "# OpenClaw PPT-master Project\n\n"
        "该目录由 OpenClaw 商品目录册 Agent 自动生成。\n\n"
        "- `deck_plan.json`：业务规划唯一事实源。\n"
        "- `design_spec.md`：设计说明。\n"
        "- `spec_lock.md`：逐页执行硬约束。\n"
        "- `svg_output/`：进入 PPT-master 后处理和导出的 SVG 页面。\n"
        "- `images/`：SVG 生成阶段引用的图片资源，路径形态为 `../images/xxx.png`。\n"
        "- `exports/`：PPT-master 导出结果。\n",
        encoding="utf-8",
    )
    (project_dir / "openclaw_manifest.json").write_text(json.dumps({"slides": manifest_slides}, ensure_ascii=False, indent=2), encoding="utf-8")
    return project_dir


def write_design_files(plan: Any, project_dir: Path) -> None:
    style_family = str(_get(plan, "style_family", "commercial") or "commercial")
    palette = STYLE_HEX.get(style_family, STYLE_HEX["commercial"])
    requested_sections = list(_get(plan, "requested_sections", []) or [])
    slides = list(_get(plan, "slides", []) or [])
    design_spec = f"""# Design Spec

## Business Context
- Type: product catalog / 商品目录册
- Theme: {_get(plan, 'theme', '')}
- Brand: {_get(plan, 'brand_or_company', '')}
- Language: {_get(plan, 'language', '')}
- Pages: {_get(plan, 'page_count', len(slides))}

## Visual Direction
- Style family: {style_family}
- Background: {palette['bg']}
- Panel: {palette['panel']}
- Accent: {palette['accent']}, {palette['accent2']}, {palette['accent3']}
- Typography: editable sans-serif text, buyer-facing concise copy.

## Content Discipline
- 用户没有要求的采购字段不主动出现。
- 价格、MOQ、箱规、认证、交期、包装、SKU 只有用户明确要求才出现。
- 图片默认无文字、无水印、无商标、无品牌包装、无 IP 角色。
"""
    spec_lock = {
        "format": "ppt169",
        "canvas": {"width": PAGE_W, "height": PAGE_H, "viewBox": f"0 0 {PAGE_W} {PAGE_H}"},
        "style_family": style_family,
        "palette": palette,
        "language": _get(plan, "language", "en"),
        "page_count": _get(plan, "page_count", len(slides)),
        "business_rules": {
            "deck_plan_is_single_source_of_truth": True,
            "do_not_add_unrequested_procurement_fields": True,
            "requested_sections": requested_sections,
            "image_text_policy": "no text inside generated product images",
        },
        "page_layouts": [
            {
                "page": idx,
                "title": _get(slide, "title", ""),
                "layout": _get(slide, "layout", ""),
                "layout_variant": _get(slide, "layout_variant", ""),
                "rhythm": "anchor" if idx == 1 else ("dense" if len(_get(slide, "sections", []) or []) >= 3 else "breathing"),
            }
            for idx, slide in enumerate(slides, start=1)
        ],
    }
    (project_dir / "design_spec.md").write_text(design_spec, encoding="utf-8")
    (project_dir / "spec_lock.md").write_text(json.dumps(spec_lock, ensure_ascii=False, indent=2), encoding="utf-8")


def _script(root: Path, name: str) -> Path:
    path = _find_script(root, name)
    if path is None:
        checked = ", ".join(str(p) for p in _candidate_script_paths(root, name))
        raise PPTMasterUnavailable(f"PPT-master 缺少脚本：{name}。已检查：{checked}")
    return path


def _run(cmd: Sequence[str], cwd: Path, timeout: int) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    py_path = str(cwd)
    if "PYTHONPATH" in env and env["PYTHONPATH"]:
        env["PYTHONPATH"] = py_path + os.pathsep + env["PYTHONPATH"]
    else:
        env["PYTHONPATH"] = py_path
    return subprocess.run(
        list(cmd),
        cwd=str(cwd),
        env=env,
        text=True,
        capture_output=True,
        timeout=timeout,
    )


def _env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None or value == "":
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def _run_first_success(candidates: Sequence[Tuple[List[str], Path]], timeout: int, step_name: str, log_path: Path) -> subprocess.CompletedProcess[str]:
    errors: List[str] = []
    for cmd, cwd in candidates:
        proc = _run(cmd, cwd=cwd, timeout=timeout)
        with log_path.open("a", encoding="utf-8") as f:
            f.write(f"\n\n## {step_name}\n")
            f.write("$ " + " ".join(cmd) + f"  # cwd={cwd}\n")
            f.write("--- stdout ---\n" + (proc.stdout or "") + "\n")
            f.write("--- stderr ---\n" + (proc.stderr or "") + "\n")
            f.write(f"--- returncode: {proc.returncode} ---\n")
        if proc.returncode == 0:
            return proc
        tail_stdout = (proc.stdout or '')[-1200:]
        tail_stderr = (proc.stderr or '')[-1200:]
        errors.append(
            f"cmd={' '.join(cmd)} cwd={cwd} rc={proc.returncode} "
            f"stdout={tail_stdout} stderr={tail_stderr}"
        )
    raise PPTMasterPipelineError(f"PPT-master 步骤失败：{step_name}；log={log_path}；" + "；".join(errors))


def _copy_or_find_export(project_dir: Path, expected_out: Path) -> None:
    if expected_out.exists() and expected_out.stat().st_size > 0:
        return
    candidates = sorted((project_dir / "exports").glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        candidates = sorted(project_dir.glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        raise PPTMasterPipelineError("PPT-master 执行完成但未找到导出的 pptx。")
    expected_out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(candidates[0], expected_out)


def run_pptmaster_pipeline(project_dir: Path, pptx_path: Path, pptmaster_root: Path) -> Dict[str, Any]:
    timeout = int(os.getenv("PPT_MASTER_TIMEOUT", "240"))
    python_bin = os.getenv("PPT_MASTER_PYTHON", sys.executable or "python3")
    log_path = project_dir / "pptmaster_pipeline.log"
    log_path.write_text("# PPT-master pipeline log\n", encoding="utf-8")

    quality_script = _find_script(pptmaster_root, "svg_quality_checker.py")
    quality_hard_gate = _env_bool("PPT_MASTER_QUALITY_HARD_GATE", False)
    if quality_script is not None and not _env_bool("PPT_MASTER_SKIP_QUALITY", False):
        try:
            _run_first_success(
                [
                    ([python_bin, str(quality_script), str(project_dir), "--format", "ppt169"], pptmaster_root),
                    ([python_bin, str(quality_script), str(project_dir / "svg_output"), "--format", "ppt169"], pptmaster_root),
                ],
                timeout=timeout,
                step_name="svg_quality_checker",
                log_path=log_path,
            )
        except PPTMasterPipelineError as exc:
            if quality_hard_gate:
                raise
            with log_path.open("a", encoding="utf-8") as f:
                f.write("\n[WARN] svg_quality_checker failed but PPT_MASTER_QUALITY_HARD_GATE is disabled; continuing to finalize/export.\n")
                f.write(str(exc) + "\n")

    total_md_split = _find_script(pptmaster_root, "total_md_split.py")
    if total_md_split is not None and (project_dir / "total.md").exists():
        try:
            _run_first_success(
                [
                    ([python_bin, str(total_md_split), str(project_dir)], pptmaster_root),
                    ([python_bin, str(total_md_split)], project_dir),
                ],
                timeout=timeout,
                step_name="total_md_split",
                log_path=log_path,
            )
        except PPTMasterPipelineError:
            # notes 不是商品目录册第一版的关键路径；保留日志但不阻塞导出。
            with log_path.open("a", encoding="utf-8") as f:
                f.write("\n[WARN] total_md_split failed, ignored for catalog export.\n")

    finalize_script = _script(pptmaster_root, "finalize_svg.py")
    _run_first_success(
        [
            ([python_bin, str(finalize_script), str(project_dir)], pptmaster_root),
            ([python_bin, str(finalize_script)], project_dir),
        ],
        timeout=timeout,
        step_name="finalize_svg",
        log_path=log_path,
    )

    export_script = _script(pptmaster_root, "svg_to_pptx.py")
    pptx_path.parent.mkdir(parents=True, exist_ok=True)
    _run_first_success(
        [
            # 官方默认流程：导出到 <project>/exports/，再由适配器复制到 OpenClaw 约定路径。
            ([python_bin, str(export_script), str(project_dir)], pptmaster_root),
            ([python_bin, str(export_script)], project_dir),
            # 兼容部分旧版/裁剪版脚本支持 -o 的情况。
            ([python_bin, str(export_script), str(project_dir), "-o", str(pptx_path), "--only", "native", "--no-notes"], pptmaster_root),
            ([python_bin, str(export_script), "-o", str(pptx_path), "--only", "native", "--no-notes"], project_dir),
            ([python_bin, str(export_script), str(project_dir), "-o", str(pptx_path)], pptmaster_root),
            ([python_bin, str(export_script), "-o", str(pptx_path)], project_dir),
        ],
        timeout=timeout,
        step_name="svg_to_pptx",
        log_path=log_path,
    )
    _copy_or_find_export(project_dir, pptx_path)
    return {
        "backend": "pptmaster",
        "pptmaster_root": str(pptmaster_root),
        "project_dir": str(project_dir),
        "log_path": str(log_path),
        "pptx_path": str(pptx_path),
    }


def build_with_pptmaster(plan: Any, pptx_path: Path, assets: Sequence[Any], run_dir: Path, workspace_root: Path) -> Dict[str, Any]:
    pptmaster_root = discover_pptmaster_root(workspace_root)
    project_dir = write_pptmaster_project(plan, assets, run_dir)
    return run_pptmaster_pipeline(project_dir, pptx_path, pptmaster_root)
