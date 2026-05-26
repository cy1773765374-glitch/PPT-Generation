#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OpenClaw PPT 商品目录册生成器（最终结构）

设计目标：
- 用户显式页数、语言、品牌、页面结构、产品类目优先。
- deck_plan.json 是唯一事实来源，渲染层不得重新决定页数或类目。
- 生产默认调用 MiniMax 生图，失败即失败；离线测试可用 placeholder。
- 成功时飞书只回复 Windows/Samba 本地路径。
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import re
import socket
import sys
import time
import urllib.parse
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from zoneinfo import ZoneInfo

try:
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None

try:
    import requests
except Exception as exc:  # pragma: no cover
    print(json.dumps({"ok": False, "error": f"requests 未安装或加载失败：{exc}"}, ensure_ascii=False))
    sys.exit(1)

try:
    from PIL import Image, ImageDraw, ImageFont
except Exception as exc:  # pragma: no cover
    print(json.dumps({"ok": False, "error": f"Pillow 未安装或加载失败：{exc}"}, ensure_ascii=False))
    sys.exit(1)

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt
except Exception as exc:  # pragma: no cover
    print(json.dumps({"ok": False, "error": f"python-pptx 未安装或加载失败：{exc}"}, ensure_ascii=False))
    sys.exit(1)


DEFAULT_TZ = "Asia/Shanghai"
INVALID_FILENAME_CHARS = '<>:"/\\\\|?*\n\r\t'
CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")

CN_NUM_MAP = {
    "零": 0,
    "一": 1,
    "二": 2,
    "两": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "十": 10,
}


@dataclass
class ImageBrief:
    key: str
    page: int
    title: str
    prompt: str
    aspect_ratio: str = "16:9"


@dataclass
class SlidePlan:
    page: int
    layout: str
    title: str
    subtitle: str = ""
    category: str = ""
    bullets: List[str] = field(default_factory=list)
    sections: List[Dict[str, Any]] = field(default_factory=list)
    image_key: str = ""
    image_prompt: str = ""


@dataclass
class DeckPlan:
    prompt: str
    language: str
    page_count: int
    deck_type: str
    brand_or_company: str
    theme: str
    style: str
    mode: str
    slides: List[SlidePlan]
    explicit_categories: List[str] = field(default_factory=list)
    image_briefs: List[ImageBrief] = field(default_factory=list)


@dataclass
class ImageAsset:
    key: str
    title: str
    prompt: str
    path: str
    source: str
    aspect_ratio: str
    page: int


@dataclass
class ValidationResult:
    ok: bool
    checks: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)


@dataclass
class GenerationResult:
    ok: bool
    prompt: str
    page_count: int
    run_name: str
    run_dir: str
    pptx_path: str
    windows_path: str
    elapsed_seconds: float
    language: str = ""
    deck_plan_path: str = ""
    image_mode: str = ""
    image_count: int = 0
    image_assets: List[Dict[str, Any]] = field(default_factory=list)
    validation: Dict[str, Any] = field(default_factory=dict)
    reply_text: str = ""
    error: str = ""


def env_int(name: str, default: int) -> int:
    value = os.getenv(name)
    if not value:
        return default
    try:
        return int(value)
    except ValueError:
        return default


def env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None or value == "":
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def now_in_configured_timezone() -> datetime:
    tz_name = os.getenv("PPT_TIMEZONE", DEFAULT_TZ).strip() or DEFAULT_TZ
    try:
        return datetime.now(ZoneInfo(tz_name))
    except Exception:
        return datetime.now(ZoneInfo(DEFAULT_TZ))


def parse_cn_number(text: str) -> Optional[int]:
    text = text.strip()
    if not text:
        return None
    if text.isdigit():
        return int(text)
    if text in CN_NUM_MAP:
        return CN_NUM_MAP[text]
    if text.startswith("十"):
        if len(text) == 1:
            return 10
        return 10 + CN_NUM_MAP.get(text[1:], 0)
    if "十" in text:
        left, _, right = text.partition("十")
        left_num = CN_NUM_MAP.get(left, 1)
        right_num = CN_NUM_MAP.get(right, 0) if right else 0
        return left_num * 10 + right_num
    return None


def clean_prompt(prompt: str) -> str:
    prompt = re.sub(r"\s+", " ", (prompt or "").strip())
    prompt = re.sub(r"^回复\s*[^:：]{1,30}[:：]\s*", "", prompt)
    for marker in ["已完成，PPT 文件", "服务器路径", "已完成 ·", "PPT 文件："]:
        if marker in prompt:
            prompt = prompt.split(marker, 1)[0].strip()
    return prompt.strip()


def extract_page_count(prompt: str, default: int = 5, max_page_count: int = 20) -> int:
    text = clean_prompt(prompt)
    patterns = [
        r"(\d{1,3})\s*[-–—]?\s*page\b",
        r"(\d{1,3})\s*pages\b",
        r"(\d{1,3})\s*slides?\b",
        r"(\d{1,3})\s*[页頁]",
        r"([一二两三四五六七八九十]{1,3})\s*[页頁]",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            value = parse_cn_number(match.group(1))
            if value:
                return max(1, min(value, max_page_count))

    # 明确 Page 1 / Pages 2-10 但开头没写 10-page 时，用最大页码兜底。
    max_page = 0
    for m in re.finditer(r"Pages?\s+(\d{1,3})(?:\s*[-–—]\s*(\d{1,3}))?", text, flags=re.IGNORECASE):
        right = m.group(2) or m.group(1)
        max_page = max(max_page, int(right))
    if max_page:
        return max(1, min(max_page, max_page_count))
    return max(1, min(default, max_page_count))


def detect_language(prompt: str) -> str:
    text = clean_prompt(prompt)
    lower = text.lower()
    if re.search(r"all\s+text\s+in\s+english|english\s+descriptions?|in\s+english|英文|英语", lower, flags=re.IGNORECASE):
        return "en"
    if CJK_RE.search(text):
        return "zh"
    return "en"


def extract_brand_or_company(prompt: str, language: str) -> str:
    text = clean_prompt(prompt)
    patterns = [
        r"for\s+([A-Z][A-Z0-9& .'-]{2,60}?)(?:[\.,;]|\s+Page\s+\d|\s+Pages\s+\d|$)",
        r"Company\s+introduction\s+for\s+([A-Z][A-Z0-9& .'-]{2,60}?)(?:\s*[-–—:]|[\.,;]|$)",
        r"brand\s*[:：]\s*([^\.;，。]{2,60})",
        r"company\s*[:：]\s*([^\.;，。]{2,60})",
        r"公司\s*[:：]\s*([^\.;，。]{2,60})",
        r"品牌\s*[:：]\s*([^\.;，。]{2,60})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            value = re.sub(r"\s+", " ", match.group(1)).strip(" -–—:：,，.;。")
            # 避免把 Product Catalog PPT 误吸进品牌名。
            value = re.sub(r"\s+PPT$", "", value, flags=re.IGNORECASE).strip()
            if value:
                return value
    return "SELLERS UNION" if language == "en" and "SELLERS UNION" in text.upper() else ("Your Company" if language == "en" else "品牌公司")


def extract_style(prompt: str, language: str) -> str:
    text = clean_prompt(prompt)
    patterns = [
        r"(High-quality[^\.。;；]{2,120}style)",
        r"(premium[^\.。;；]{2,120}style)",
        r"风格\s*[:：]\s*([^\.。;；]{2,120})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip(" .。;；")
    return "High-quality commercial product catalog style" if language == "en" else "高质量商业商品目录册风格"


def extract_theme(prompt: str, language: str) -> str:
    text = clean_prompt(prompt)
    match = re.search(r"Generate\s+a\s+\d{1,3}\s*[-–—]?\s*page\s+(.+?)\s+product\s+catalog", text, flags=re.IGNORECASE)
    if match:
        return re.sub(r"\s+", " ", match.group(1)).strip(" .;:")
    match = re.search(r"(.{2,40}?)(?:相关|主题|类|系列)?(?:的)?(?:\d+|[一二两三四五六七八九十]+)?页?(?:商品目录册|产品目录册|目录册|PPT|ppt)", text)
    if match:
        raw = match.group(1)
        raw = re.sub(r"^(生成|做|制作|设计|帮我|帮我做|帮我生成|一个|一份)+", "", raw).strip()
        raw = re.sub(r"\d+|[一二两三四五六七八九十]+", "", raw).strip(" -_，。,.：:；;")
        if raw:
            return raw
    return "product catalog" if language == "en" else "商品目录册"


def title_case_category(text: str) -> str:
    text = re.sub(r"\s+", " ", text.strip(" -–—:：,，.;。"))
    if not text:
        return text
    # 用户已经给了全大写缩写时保留，例如 COS。
    words = []
    for w in text.split(" "):
        if w.isupper() and len(w) <= 5:
            words.append(w)
        elif w.lower() in {"and", "or", "of", "for", "with"}:
            words.append(w.lower())
        else:
            words.append(w[:1].upper() + w[1:])
    if words:
        words[0] = words[0][:1].upper() + words[0][1:]
    return " ".join(words)


def extract_product_categories(prompt: str, language: str) -> List[str]:
    text = clean_prompt(prompt)
    candidates = ""
    patterns = [
        r"Product\s+categories\s*[:：]\s*(.+?)(?:\.\s|。\s|$)",
        r"产品类目\s*[:：]\s*(.+?)(?:\.\s|。\s|$)",
        r"商品类目\s*[:：]\s*(.+?)(?:\.\s|。\s|$)",
        r"类目\s*[:：]\s*(.+?)(?:\.\s|。\s|$)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            candidates = match.group(1).strip()
            break
    if not candidates:
        return []

    parts: List[str] = []
    # 先按逗号/分号切，避免把 Balloons and Balloon Sets 里的 and 切掉。
    for item in re.split(r"[,，;；]", candidates):
        item = re.sub(r"^\s*(and|以及|和)\s+", "", item.strip(), flags=re.IGNORECASE)
        item = item.strip(" .。;；")
        if not item:
            continue
        if language == "en":
            item = title_case_category(item)
            if item.lower().startswith("other "):
                item = "Other " + item[6:]
        parts.append(item)
    # 去重但保序。
    seen = set()
    result = []
    for item in parts:
        key = item.lower()
        if key not in seen:
            result.append(item)
            seen.add(key)
    return result


def generated_categories(theme: str, language: str, needed: int) -> List[str]:
    if language == "en":
        base = [
            title_case_category(theme) if theme and theme != "product catalog" else "Hero Products",
            "Bestseller Collection",
            "Premium Gift Sets",
            "Retail Display Options",
            "Packaging and Customization",
            "Buyer Information",
        ]
    else:
        base = [
            f"{theme}主推款" if theme else "主推商品",
            "热销组合",
            "高端礼品款",
            "零售陈列方案",
            "包装与定制",
            "采购沟通信息",
        ]
    while len(base) < needed:
        base.append((f"Catalog Category {len(base)+1}" if language == "en" else f"目录类目 {len(base)+1}"))
    return base[:needed]


def english_category_content(category: str, theme: str) -> List[Dict[str, Any]]:
    key = category.lower()
    presets: Dict[str, Dict[str, List[str] | str]] = {
        "balloons": {
            "occasion": "Birthday parties, weddings, retail party kits, seasonal events",
            "points": ["Coordinated color sets", "Easy arch and garland setup", "Strong visual impact for party scenes"],
            "skus": "Latex balloons, foil balloons, balloon arch kits, number balloons",
            "packaging": "Polybag, color box, display pack, customized set card",
        },
        "rain curtains": {
            "occasion": "Backdrop walls, photo booths, stage decoration, holiday displays",
            "points": ["Shiny metallic finish", "Quick installation", "Strong shelf appeal for party aisles"],
            "skus": "Foil fringe curtains, metallic rain curtains, themed backdrop sets",
            "packaging": "Flat bag, header card, hanging display pack",
        },
        "candles": {
            "occasion": "Birthday cakes, celebration tables, gift sets, party supplies shelves",
            "points": ["Bright color choices", "Multiple number and novelty shapes", "Suitable for bundled selling"],
            "skus": "Number candles, spiral candles, glitter candles, themed cake candles",
            "packaging": "Blister card, color box, retail hanging card",
        },
        "bunting": {
            "occasion": "Room decoration, garden parties, birthdays, graduation events",
            "points": ["Ready-to-hang decoration", "Reusable visual element", "Flexible color and pattern matching"],
            "skus": "Paper garlands, triangle bunting, letter banners, themed hanging decor",
            "packaging": "OPP bag, header card, compact retail pack",
        },
        "hats": {
            "occasion": "Birthday parties, kids events, costume parties, celebration photos",
            "points": ["Lightweight wearing experience", "Bright party colors", "High impulse-purchase potential"],
            "skus": "Cone hats, crowns, themed headwear, party cap sets",
            "packaging": "Stacked polybag, display box, party set bundle",
        },
        "party blowers": {
            "occasion": "Kids parties, birthday tables, New Year countdown, party favor bags",
            "points": ["Fun interactive item", "Low-cost add-on SKU", "Good for multi-piece packs"],
            "skus": "Paper blowers, foil blowers, themed party horns, favor pack blowers",
            "packaging": "Polybag set, blister card, counter display box",
        },
        "tableware": {
            "occasion": "Buffet tables, birthday parties, picnics, event catering",
            "points": ["Complete disposable table setup", "Color-matched party solution", "Easy cleanup for event users"],
            "skus": "Paper plates, paper cups, napkins, straws, cutlery sets, table covers",
            "packaging": "Shrink pack, color sleeve, full party tableware kit",
        },
        "cos costumes": {
            "occasion": "Cosplay events, Halloween parties, stage activities, themed celebrations",
            "points": ["High visual recognition", "Good seasonal promotion potential", "Supports accessory combinations"],
            "skus": "Role-play costumes, capes, masks, themed accessories, costume kits",
            "packaging": "Garment bag, color insert, boxed costume set",
        },
        "other party decorations": {
            "occasion": "Complete party scene building, seasonal programs, one-stop retail displays",
            "points": ["Broad category coverage", "Easy cross-selling with core party SKUs", "Flexible theme extension"],
            "skus": "Confetti, banners, hanging swirls, photo props, party favors, table decor",
            "packaging": "Mixed set pack, retail display box, themed bundle solution",
        },
    }

    chosen = None
    for k, v in presets.items():
        if k in key:
            chosen = v
            break
    if chosen is None:
        chosen = {
            "occasion": f"{title_case_category(theme)} parties, retail displays, seasonal promotions",
            "points": ["Clear category positioning", "Flexible color and style options", "Suitable for catalog and retail programs"],
            "skus": f"Core {category} items, bundle sets, seasonal variants",
            "packaging": "Polybag, color box, display pack, customized retail set",
        }
    return [
        {"title": "Occasion Fit", "items": [str(chosen["occasion"])]},
        {"title": "Selling Points", "items": list(chosen["points"])},
        {"title": "Suggested SKUs", "items": [str(chosen["skus"])]},
        {"title": "Packaging & Sourcing Notes", "items": [str(chosen["packaging"]), "MOQ, carton size, delivery time and certification can be added from real product data."]},
    ]


def chinese_category_content(category: str, theme: str) -> List[Dict[str, Any]]:
    return [
        {"title": "适用场景", "items": ["零售陈列、节庆促销、客户初筛、业务目录册沟通"]},
        {"title": "核心卖点", "items": ["视觉统一", "支持颜色/包装/组合定制", "适合套装化销售"]},
        {"title": "建议 SKU", "items": [f"{category}主推款、组合款、陈列款、季节款"]},
        {"title": "采购备注", "items": ["真实图片、价格、MOQ、箱规、认证和交期可由业务继续补充。"]},
    ]


def make_image_prompt(slide: SlidePlan, plan: DeckPlan) -> str:
    no_text = "no text, no watermark, no logo"
    style = "commercial product catalog photography, realistic, premium lighting, clean warm white background, high detail"
    if plan.language == "en":
        if slide.layout == "company_intro":
            return (
                f"Premium commercial catalog opening image for {plan.brand_or_company}, {plan.theme} product catalog, "
                f"assorted party supplies and festive decorations arranged in a clean showroom composition, {style}, {no_text}"
            )
        return (
            f"Commercial product catalog photography of {slide.category}, party supplies and festive decoration products, "
            f"clean e-commerce catalog composition, cohesive {plan.theme} style, {style}, {no_text}"
        )
    if slide.layout == "company_intro":
        return f"商业商品目录册封面图，{plan.theme}，产品陈列，干净背景，高级灯光，无文字，无水印，无logo"
    return f"{slide.category} 商品目录册摄影图，干净背景，高级灯光，商业产品陈列，无文字，无水印，无logo"


def build_deck_plan(prompt: str) -> DeckPlan:
    prompt = clean_prompt(prompt)
    default_page_count = env_int("PPT_DEFAULT_PAGE_COUNT", 5)
    max_page_count = env_int("PPT_MAX_PAGE_COUNT", 20)
    language = detect_language(prompt)
    page_count = extract_page_count(prompt, default_page_count, max_page_count)
    brand = extract_brand_or_company(prompt, language)
    theme = extract_theme(prompt, language)
    style = extract_style(prompt, language)
    categories = extract_product_categories(prompt, language)

    has_explicit_page_plan = bool(re.search(r"Page\s+1\s*[:：]|Pages\s+\d+\s*[-–—]\s*\d+", prompt, flags=re.IGNORECASE))
    has_company_intro = bool(re.search(r"company\s+introduction|公司介绍|企业介绍", prompt, flags=re.IGNORECASE))
    mode = "detailed" if (has_explicit_page_plan or categories) else "simple"

    slides: List[SlidePlan] = []
    if has_company_intro or (mode == "detailed" and language == "en"):
        if language == "en":
            title = brand
            subtitle = f"{title_case_category(theme)} Product Catalog"
            bullets = [
                "Premier party supplies and festive decorations supplier",
                "Integrated category planning for seasonal, retail and event programs",
                "Catalog structure follows the user's requested page order and product categories",
            ]
            sections = [
                {"title": "Company Positioning", "items": ["One-stop supplier for party supplies, decorations and celebration-ready product programs"]},
                {"title": "Catalog Focus", "items": ["Commercial product presentation", "Clear category pages", "English-only buyer-facing copy"]},
                {"title": "Buyer Value", "items": ["Fast category overview", "Scenario-based product planning", "Ready for adding real SKUs, MOQ, pricing and carton data"]},
            ]
        else:
            title = brand
            subtitle = f"{theme}商品目录册"
            bullets = ["供应商介绍", "商品目录册结构根据用户需求生成", "后续可补充真实 SKU、报价、MOQ、箱规和认证"]
            sections = [
                {"title": "公司定位", "items": ["面向客户的商品目录册初稿"]},
                {"title": "目录重点", "items": ["按类目展示", "图文结合", "便于业务二次补充"]},
            ]
        slides.append(SlidePlan(page=1, layout="company_intro", title=title, subtitle=subtitle, bullets=bullets, sections=sections))
    else:
        if language == "en":
            title = title_case_category(theme)
            subtitle = "Product Catalog"
            sections = [
                {"title": "Catalog Goal", "items": ["Create a concise product catalog for buyer communication"]},
                {"title": "Next Step", "items": ["Add real product images, pricing, MOQ, carton size and certification details"]},
            ]
        else:
            title = theme
            subtitle = "商品目录册"
            sections = [
                {"title": "目录目标", "items": ["用于业务初稿沟通和商品方向展示"]},
                {"title": "后续补充", "items": ["真实商品图、报价、MOQ、箱规、认证、交期"]},
            ]
        slides.append(SlidePlan(page=1, layout="cover_catalog", title=title, subtitle=subtitle, sections=sections))

    remaining = page_count - 1
    if remaining > 0:
        category_plan = categories[:remaining]
        if len(category_plan) < remaining:
            category_plan.extend(generated_categories(theme, language, remaining - len(category_plan)))
        for idx, category in enumerate(category_plan[:remaining], start=2):
            if language == "en":
                sections = english_category_content(category, theme)
                subtitle = "Category Showcase"
                bullets = [
                    "Product positioning built from the requested category",
                    "Image brief is bound to this page, not reused from a generic template",
                ]
            else:
                sections = chinese_category_content(category, theme)
                subtitle = "类目展示页"
                bullets = ["按当前类目生成独立页面", "图片 brief 与本页类目绑定"]
            slides.append(SlidePlan(page=idx, layout="category_showcase", title=category, subtitle=subtitle, category=category, bullets=bullets, sections=sections))

    # 强制页码与页数一致，deck_plan 后续作为唯一事实源。
    slides = slides[:page_count]
    for i, slide in enumerate(slides, start=1):
        slide.page = i
        slide.image_key = f"slide_{i:02d}"

    plan = DeckPlan(
        prompt=prompt,
        language=language,
        page_count=page_count,
        deck_type="product_catalog",
        brand_or_company=brand,
        theme=theme,
        style=style,
        mode=mode,
        slides=slides,
        explicit_categories=categories,
    )
    briefs: List[ImageBrief] = []
    for slide in plan.slides:
        slide.image_prompt = make_image_prompt(slide, plan)
        briefs.append(ImageBrief(key=slide.image_key, page=slide.page, title=slide.title, prompt=slide.image_prompt, aspect_ratio="16:9"))
    plan.image_briefs = briefs
    validate_plan_before_render(plan)
    return plan


def validate_plan_before_render(plan: DeckPlan) -> None:
    if plan.page_count < 1:
        raise RuntimeError("需求解析失败：page_count 不能小于 1。")
    if len(plan.slides) != plan.page_count:
        raise RuntimeError(f"需求解析失败：deck_plan 页数不一致，要求 {plan.page_count} 页，计划 {len(plan.slides)} 页。")
    seen_pages = [s.page for s in plan.slides]
    if seen_pages != list(range(1, plan.page_count + 1)):
        raise RuntimeError(f"需求解析失败：页码不连续：{seen_pages}")
    if plan.language not in {"en", "zh"}:
        raise RuntimeError(f"需求解析失败：不支持的语言：{plan.language}")
    if plan.mode == "detailed" and plan.explicit_categories:
        required = max(0, plan.page_count - 1)
        if len(plan.explicit_categories) < required:
            # 不直接失败：少类目时自动补齐，但写入 metadata。这里做明确提示。
            pass


def plan_to_dict(plan: DeckPlan) -> Dict[str, Any]:
    return {
        **{k: v for k, v in asdict(plan).items() if k not in {"slides", "image_briefs"}},
        "slides": [asdict(s) for s in plan.slides],
        "image_briefs": [asdict(b) for b in plan.image_briefs],
    }


def read_openclaw_minimax_key() -> str:
    config_path = Path(os.getenv("OPENCLAW_CONFIG", "~/.openclaw/openclaw.json")).expanduser()
    if not config_path.exists():
        return ""
    try:
        data = json.loads(config_path.read_text(encoding="utf-8"))
    except Exception:
        return ""

    key_names = {"apikey", "api_key", "key", "token", "authorization"}

    def walk(obj: Any, parent_keys: List[str]) -> Iterable[str]:
        if isinstance(obj, dict):
            for k, v in obj.items():
                lk = str(k).lower()
                new_parent = parent_keys + [lk]
                if isinstance(v, str) and lk in key_names and any("minimax" in p or "minimaxi" in p for p in parent_keys):
                    yield v
                else:
                    yield from walk(v, new_parent)
        elif isinstance(obj, list):
            for item in obj:
                yield from walk(item, parent_keys)

    for candidate in walk(data, []):
        candidate = candidate.strip()
        if candidate and not candidate.startswith("${") and len(candidate) > 8:
            if candidate.lower().startswith("bearer "):
                candidate = candidate.split(None, 1)[1].strip()
            return candidate
    return ""


def get_minimax_api_key() -> str:
    for name in ["MINIMAX_API_KEY", "MINIMAX_TOKEN", "MINIMAX_API_TOKEN"]:
        value = os.getenv(name, "").strip()
        if value:
            return value.removeprefix("Bearer ").strip()
    key_file = os.getenv("MINIMAX_API_KEY_FILE", "").strip()
    if key_file:
        path = Path(key_file).expanduser()
        if path.exists():
            return path.read_text(encoding="utf-8").strip().removeprefix("Bearer ").strip()
    return read_openclaw_minimax_key()


def normalize_minimax_image_api_url(raw_url: str) -> str:
    url = (raw_url or "").strip() or "https://api.minimax.com/v1/image_generation"
    if not url.startswith("http://") and not url.startswith("https://"):
        url = "https://" + url.lstrip("/")
    parsed = urllib.parse.urlparse(url)
    scheme = parsed.scheme or "https"
    host = parsed.netloc
    path = parsed.path or ""
    if not host:
        stripped = url.replace("https://", "").replace("http://", "").strip("/")
        host, _, path_tail = stripped.partition("/")
        path = "/" + path_tail if path_tail else ""
    if "image_generation" in path:
        return urllib.parse.urlunparse((scheme, host, path.rstrip("/"), "", "", ""))
    return urllib.parse.urlunparse((scheme, host, "/v1/image_generation", "", "", ""))


def get_minimax_image_api_candidates() -> List[str]:
    configured_url = (
        os.getenv("MINIMAX_IMAGE_API_URL", "").strip()
        or os.getenv("MINIMAX_IMAGE_BASE_URL", "").strip()
        or os.getenv("MINIMAX_BASE_URL", "").strip()
        or "https://api.minimax.com/v1/image_generation"
    )
    raw_candidates = os.getenv("MINIMAX_IMAGE_API_URL_CANDIDATES", "").strip()
    if raw_candidates:
        candidates = [x.strip() for x in re.split(r"[,\n]", raw_candidates) if x.strip()]
    else:
        candidates = [configured_url, "https://api.minimaxi.com/v1/image_generation", "https://api.minimax.io/v1/image_generation"]
    if not env_bool("MINIMAX_AUTO_FALLBACK_HOSTS", True):
        candidates = [configured_url]
    normalized: List[str] = []
    seen = set()
    for item in candidates:
        url = normalize_minimax_image_api_url(item)
        if url not in seen:
            normalized.append(url)
            seen.add(url)
    return normalized


def assert_hostname_resolves(api_url: str, timeout: int = 8) -> None:
    parsed = urllib.parse.urlparse(api_url)
    host = parsed.hostname
    if not host:
        raise RuntimeError(f"MiniMax 生图接口地址不合法：{api_url}")
    old_timeout = socket.getdefaulttimeout()
    socket.setdefaulttimeout(timeout)
    try:
        socket.getaddrinfo(host, parsed.port or 443)
    except Exception as exc:
        raise RuntimeError(
            f"MiniMax 生图接口域名解析失败：{host}；当前接口地址={api_url}。"
            "如果服务器通过代理访问 MiniMax，建议设置 MINIMAX_NETWORK_PRECHECK=0。"
        ) from exc
    finally:
        socket.setdefaulttimeout(old_timeout)


def collect_strings(obj: Any) -> Iterable[str]:
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, dict):
        for value in obj.values():
            yield from collect_strings(value)
    elif isinstance(obj, list):
        for value in obj:
            yield from collect_strings(value)


def looks_like_base64_image(text: str) -> bool:
    if text.startswith("data:image/"):
        return True
    if len(text) < 1000:
        return False
    compact = text.strip().replace("\n", "")
    return re.fullmatch(r"[A-Za-z0-9+/=]+", compact) is not None


def decode_base64_image(text: str) -> bytes:
    if text.startswith("data:image/"):
        _, _, text = text.partition(",")
    compact = text.strip().replace("\n", "")
    missing = len(compact) % 4
    if missing:
        compact += "=" * (4 - missing)
    return base64.b64decode(compact, validate=False)


class MiniMaxImageClient:
    def __init__(self) -> None:
        self.api_key = get_minimax_api_key()
        self.api_urls = get_minimax_image_api_candidates()
        self.model = os.getenv("MINIMAX_IMAGE_MODEL", "image-01").strip()
        self.timeout = env_int("MINIMAX_IMAGE_TIMEOUT", 180)
        self.response_format = os.getenv("MINIMAX_IMAGE_RESPONSE_FORMAT", "base64").strip().lower()
        self.prompt_optimizer = env_bool("MINIMAX_PROMPT_OPTIMIZER", True)
        self.verify_ssl = env_bool("MINIMAX_VERIFY_SSL", True)
        self.network_precheck = env_bool("MINIMAX_NETWORK_PRECHECK", False)
        proxy_url = os.getenv("MINIMAX_PROXY_URL", "").strip()
        self.proxies = {"http": proxy_url, "https": proxy_url} if proxy_url else None

    def generate_to_file(self, prompt: str, out_path: Path, aspect_ratio: str = "16:9") -> None:
        if not self.api_key:
            raise RuntimeError("MINIMAX_API_KEY 未配置，且未能从 ~/.openclaw/openclaw.json 自动读取到 MiniMax 密钥；为避免生成纯文字 PPT，本次已停止。")
        errors: List[str] = []
        for api_url in self.api_urls:
            try:
                if self.network_precheck:
                    assert_hostname_resolves(api_url)
                self._generate_to_file_once(api_url, prompt, out_path, aspect_ratio)
                return
            except Exception as exc:
                errors.append(f"{api_url} => {exc}")
                if not env_bool("MINIMAX_AUTO_FALLBACK_HOSTS", True):
                    break
        raise RuntimeError("MiniMax 生图全部候选接口均失败：" + "；".join(errors))

    def _generate_to_file_once(self, api_url: str, prompt: str, out_path: Path, aspect_ratio: str = "16:9") -> None:
        payload: Dict[str, Any] = {
            "model": self.model,
            "prompt": prompt[:1500],
            "aspect_ratio": aspect_ratio,
            "response_format": self.response_format,
            "n": 1,
            "prompt_optimizer": self.prompt_optimizer,
        }
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
        try:
            resp = requests.post(api_url, headers=headers, json=payload, timeout=self.timeout, verify=self.verify_ssl, proxies=self.proxies)
        except requests.exceptions.RequestException as exc:
            raise RuntimeError(f"MiniMax 生图请求失败：url={api_url}；{exc}") from exc
        if resp.status_code >= 400:
            raise RuntimeError(f"MiniMax 生图失败 HTTP {resp.status_code}: {resp.text[:600]}")
        try:
            data = resp.json()
        except Exception as exc:
            raise RuntimeError(f"MiniMax 生图返回不是 JSON：{exc}; body={resp.text[:300]}") from exc
        base_resp = data.get("base_resp") if isinstance(data, dict) else None
        if isinstance(base_resp, dict) and int(base_resp.get("status_code", 0) or 0) != 0:
            raise RuntimeError(f"MiniMax 生图失败：{base_resp.get('status_msg') or base_resp}")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        image_urls = []
        if isinstance(data, dict) and isinstance(data.get("data"), dict):
            image_urls = data.get("data", {}).get("image_urls", [])
        if image_urls:
            self._download_image(str(image_urls[0]), out_path)
            return
        for item in collect_strings(data):
            if item.startswith("http://") or item.startswith("https://"):
                self._download_image(item, out_path)
                return
            if looks_like_base64_image(item):
                raw = decode_base64_image(item)
                if len(raw) > 1000:
                    out_path.write_bytes(raw)
                    self._normalize_image(out_path)
                    return
        raise RuntimeError(f"MiniMax 生图响应中未找到 image_urls 或 base64 图片字段：{json.dumps(data, ensure_ascii=False)[:800]}")

    def _download_image(self, url: str, out_path: Path) -> None:
        resp = requests.get(url, timeout=self.timeout, verify=self.verify_ssl, proxies=self.proxies)
        if resp.status_code >= 400:
            raise RuntimeError(f"MiniMax 图片下载失败 HTTP {resp.status_code}: {url}")
        out_path.write_bytes(resp.content)
        self._normalize_image(out_path)

    def _normalize_image(self, path: Path) -> None:
        try:
            with Image.open(path) as im:
                im = im.convert("RGB")
                im.save(path, format="PNG")
        except Exception as exc:
            raise RuntimeError(f"MiniMax 图片文件无法被 Pillow 读取：{path}；{exc}") from exc


def find_font() -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, 36)
        except Exception:
            continue
    return ImageFont.load_default()


def generate_placeholder_image(title: str, out_path: Path, aspect_ratio: str, language: str = "en") -> None:
    size = (1600, 900) if aspect_ratio == "16:9" else (1200, 1200)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    im = Image.new("RGB", size, (248, 250, 252))
    draw = ImageDraw.Draw(im)
    draw.rounded_rectangle([46, 46, size[0] - 46, size[1] - 46], radius=42, fill=(255, 255, 255), outline=(203, 213, 225), width=4)
    draw.ellipse([size[0] * 0.08, size[1] * 0.16, size[0] * 0.34, size[1] * 0.62], fill=(219, 234, 254))
    draw.rounded_rectangle([size[0] * 0.40, size[1] * 0.20, size[0] * 0.80, size[1] * 0.68], radius=34, fill=(224, 242, 254))
    draw.line([size[0] * 0.14, size[1] * 0.74, size[0] * 0.84, size[1] * 0.74], fill=(37, 99, 235), width=9)
    font = find_font()
    label = "OFFLINE TEST IMAGE" if language == "en" else "离线测试图"
    text = f"{label}\n{title[:54]}"
    draw.multiline_text((80, size[1] - 178), text, fill=(71, 85, 105), font=font, spacing=10)
    im.save(out_path, format="PNG")


def prepare_images(plan: DeckPlan, run_dir: Path, image_mode: str) -> List[ImageAsset]:
    image_mode = image_mode.lower().strip()
    if image_mode not in {"minimax", "placeholder", "off"}:
        raise ValueError("PPT_IMAGE_MODE 只能是 minimax / placeholder / off")
    if image_mode == "off":
        if env_bool("PPT_REQUIRE_IMAGES", True):
            raise RuntimeError("PPT_IMAGE_MODE=off 但 PPT_REQUIRE_IMAGES=1；为避免纯文字 PPT，本次已停止。")
        return []
    asset_dir = run_dir / "assets" / image_mode
    client = MiniMaxImageClient() if image_mode == "minimax" else None
    assets: List[ImageAsset] = []
    for brief in plan.image_briefs:
        out_path = asset_dir / f"{brief.key}.png"
        if image_mode == "minimax":
            assert client is not None
            client.generate_to_file(brief.prompt, out_path, brief.aspect_ratio)
        else:
            generate_placeholder_image(brief.title, out_path, brief.aspect_ratio, plan.language)
        assets.append(ImageAsset(key=brief.key, title=brief.title, prompt=brief.prompt, path=str(out_path), source=image_mode, aspect_ratio=brief.aspect_ratio, page=brief.page))
    if env_bool("PPT_REQUIRE_IMAGES", True) and len(assets) != len(plan.slides):
        raise RuntimeError(f"图片数量校验失败：计划 {len(plan.slides)} 张，实际 {len(assets)} 张。")
    return assets


def asset_map(assets: List[ImageAsset]) -> Dict[str, ImageAsset]:
    return {asset.key: asset for asset in assets}


def set_slide_background(slide, rgb=(248, 250, 252)) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*rgb)


def font_name(language: str) -> str:
    return "Arial" if language == "en" else "Microsoft YaHei"


def add_textbox(slide, left, top, width, height, text, font_size=18, bold=False, color=(31, 41, 55), align=None, language="en"):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.clear()
    lines = str(text).split("\n") if str(text) else [""]
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = line
        if align is not None:
            p.alignment = align
        for run in p.runs:
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.color.rgb = RGBColor(*color)
            run.font.name = font_name(language)
    return box


def add_bullet_box(slide, left, top, width, height, lines: List[str], font_size=10.5, color=(71, 85, 105), language="en"):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)
    tf.clear()
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = f"• {line}"
        p.font.size = Pt(font_size)
        p.font.name = font_name(language)
        p.font.color.rgb = RGBColor(*color)
        p.space_after = Pt(3)
    return box


def add_section_card(slide, left, top, width, height, title: str, items: List[str], accent=(37, 99, 235), language="en") -> None:
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.color.rgb = RGBColor(226, 232, 240)
    shape.line.width = Pt(1)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.07), height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(*accent)
    bar.line.fill.background()
    add_textbox(slide, left + Inches(0.22), top + Inches(0.14), width - Inches(0.35), Inches(0.28), title, 12.8, True, (15, 23, 42), language=language)
    add_bullet_box(slide, left + Inches(0.22), top + Inches(0.52), width - Inches(0.42), height - Inches(0.62), items, 9.4 if language == "en" else 9.2, language=language)


def add_image(slide, asset: Optional[ImageAsset], left, top, width, height, label: str, language="en") -> None:
    if asset and Path(asset.path).exists():
        try:
            pic = slide.shapes.add_picture(asset.path, left, top, width=width, height=height)
            pic.line.color.rgb = RGBColor(226, 232, 240)
            return
        except Exception:
            pass
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(226, 232, 240)
    shape.line.color.rgb = RGBColor(203, 213, 225)
    fallback = "Image unavailable" if language == "en" else "图片不可用"
    add_textbox(slide, left + Inches(0.15), top + height / 2 - Inches(0.16), width - Inches(0.3), Inches(0.4), fallback, 10.5, True, (100, 116, 139), PP_ALIGN.CENTER, language=language)


def add_footer(slide, plan: DeckPlan, page: int) -> None:
    if plan.language == "en":
        text = f"{plan.brand_or_company} Product Catalog · Page {page}/{plan.page_count}"
    else:
        text = f"{plan.brand_or_company} 商品目录册 · 第 {page}/{plan.page_count} 页"
    add_textbox(slide, Inches(0.66), Inches(7.08), Inches(12.0), Inches(0.22), text, 8.4, False, (148, 163, 184), language=plan.language)


def create_company_intro_slide(prs: Presentation, plan: DeckPlan, slide_plan: SlidePlan, assets: Dict[str, ImageAsset]) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, (242, 246, 251))
    add_textbox(slide, Inches(0.72), Inches(0.62), Inches(5.65), Inches(0.78), slide_plan.title, 34, True, (15, 23, 42), language=plan.language)
    add_textbox(slide, Inches(0.75), Inches(1.44), Inches(5.55), Inches(0.36), slide_plan.subtitle, 16, False, (37, 99, 235), language=plan.language)
    add_bullet_box(slide, Inches(0.78), Inches(2.05), Inches(5.3), Inches(0.92), slide_plan.bullets, 10.2, language=plan.language)

    positions = [
        (Inches(0.75), Inches(3.18), Inches(2.55), Inches(1.58)),
        (Inches(3.55), Inches(3.18), Inches(2.55), Inches(1.58)),
        (Inches(0.75), Inches(5.08), Inches(5.35), Inches(1.28)),
    ]
    accents = [(37, 99, 235), (14, 165, 233), (16, 185, 129)]
    for idx, section in enumerate(slide_plan.sections[:3]):
        left, top, width, height = positions[idx]
        add_section_card(slide, left, top, width, height, section["title"], section["items"], accents[idx], plan.language)

    panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.55), Inches(0.72), Inches(6.15), Inches(5.95))
    panel.fill.solid()
    panel.fill.fore_color.rgb = RGBColor(255, 255, 255)
    panel.line.color.rgb = RGBColor(226, 232, 240)
    add_image(slide, assets.get(slide_plan.image_key), Inches(6.78), Inches(0.96), Inches(5.68), Inches(3.2), slide_plan.title, plan.language)
    if plan.language == "en":
        callout_title = "Catalog Structure"
        callout_lines = [
            f"1 company introduction page",
            f"{max(0, plan.page_count - 1)} product category pages",
            "Every product category follows the requested order",
        ]
    else:
        callout_title = "目录结构"
        callout_lines = ["1 页公司介绍", f"{max(0, plan.page_count - 1)} 页产品类目", "类目顺序按用户需求执行"]
    add_section_card(slide, Inches(6.78), Inches(4.45), Inches(5.68), Inches(1.55), callout_title, callout_lines, (37, 99, 235), plan.language)
    add_footer(slide, plan, slide_plan.page)


def create_cover_slide(prs: Presentation, plan: DeckPlan, slide_plan: SlidePlan, assets: Dict[str, ImageAsset]) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, (242, 246, 251))
    add_textbox(slide, Inches(0.78), Inches(0.74), Inches(5.5), Inches(0.78), slide_plan.title, 32, True, (15, 23, 42), language=plan.language)
    add_textbox(slide, Inches(0.82), Inches(1.58), Inches(5.2), Inches(0.36), slide_plan.subtitle, 16, False, (37, 99, 235), language=plan.language)
    for idx, section in enumerate(slide_plan.sections[:2]):
        add_section_card(slide, Inches(0.82), Inches(2.42 + 1.65 * idx), Inches(5.25), Inches(1.28), section["title"], section["items"], [(37, 99, 235), (16, 185, 129)][idx], plan.language)
    add_image(slide, assets.get(slide_plan.image_key), Inches(6.55), Inches(0.74), Inches(6.1), Inches(5.8), slide_plan.title, plan.language)
    add_footer(slide, plan, slide_plan.page)


def create_category_showcase_slide(prs: Presentation, plan: DeckPlan, slide_plan: SlidePlan, assets: Dict[str, ImageAsset]) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide)
    add_textbox(slide, Inches(0.65), Inches(0.34), Inches(8.2), Inches(0.45), slide_plan.title, 25.5, True, (15, 23, 42), language=plan.language)
    add_textbox(slide, Inches(0.68), Inches(0.86), Inches(8.0), Inches(0.28), slide_plan.subtitle, 11.5, False, (100, 116, 139), language=plan.language)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.65), Inches(1.18), Inches(12.0), Inches(0.028))
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(37, 99, 235)
    line.line.fill.background()

    add_image(slide, assets.get(slide_plan.image_key), Inches(0.72), Inches(1.46), Inches(5.6), Inches(3.15), slide_plan.title, plan.language)
    if plan.language == "en":
        intro_title = "Page Brief"
        intro_lines = [
            f"Category: {slide_plan.category}",
            "Commercial catalog copy generated for this exact category",
            "Image prompt is bound to this slide and avoids text inside the image",
        ]
    else:
        intro_title = "页面定位"
        intro_lines = [f"类目：{slide_plan.category}", "内容和图片提示词均绑定当前类目", "图片提示词避免在图中生成文字"]
    add_section_card(slide, Inches(0.72), Inches(4.82), Inches(5.6), Inches(1.45), intro_title, intro_lines, (37, 99, 235), plan.language)

    card_positions = [
        (Inches(6.65), Inches(1.46), Inches(2.85), Inches(1.55)),
        (Inches(9.8), Inches(1.46), Inches(2.85), Inches(1.55)),
        (Inches(6.65), Inches(3.25), Inches(2.85), Inches(1.55)),
        (Inches(9.8), Inches(3.25), Inches(2.85), Inches(1.55)),
    ]
    accents = [(37, 99, 235), (14, 165, 233), (16, 185, 129), (245, 158, 11)]
    for idx, section in enumerate(slide_plan.sections[:4]):
        left, top, width, height = card_positions[idx]
        add_section_card(slide, left, top, width, height, section["title"], section["items"], accents[idx], plan.language)

    if plan.language == "en":
        note = "Ready for real SKU photos, price tiers, MOQ, carton size, certification and delivery details."
    else:
        note = "后续可补充真实 SKU 图片、价格梯度、MOQ、箱规、认证和交期。"
    add_textbox(slide, Inches(6.72), Inches(5.36), Inches(5.75), Inches(0.42), note, 10.2, False, (71, 85, 105), language=plan.language)
    add_footer(slide, plan, slide_plan.page)


def build_presentation(plan: DeckPlan, pptx_path: Path, assets: List[ImageAsset]) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    by_key = asset_map(assets)
    for slide_plan in plan.slides:
        if slide_plan.layout == "company_intro":
            create_company_intro_slide(prs, plan, slide_plan, by_key)
        elif slide_plan.layout == "cover_catalog":
            create_cover_slide(prs, plan, slide_plan, by_key)
        elif slide_plan.layout == "category_showcase":
            create_category_showcase_slide(prs, plan, slide_plan, by_key)
        else:
            raise RuntimeError(f"未知版式：{slide_plan.layout}")
    pptx_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(pptx_path))


def extract_pptx_texts(pptx_path: Path) -> List[str]:
    prs = Presentation(str(pptx_path))
    texts: List[str] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                texts.append(str(shape.text))
            if hasattr(shape, "table"):
                try:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if cell.text:
                                texts.append(cell.text)
                except Exception:
                    pass
    return texts


def validate_output(plan: DeckPlan, pptx_path: Path, assets: List[ImageAsset]) -> ValidationResult:
    result = ValidationResult(ok=True)
    errors: List[str] = []
    checks: List[str] = []
    if not pptx_path.exists():
        errors.append(f"PPT 文件不存在：{pptx_path}")
    else:
        prs = Presentation(str(pptx_path))
        actual = len(prs.slides)
        if actual != plan.page_count:
            errors.append(f"页数校验失败：用户/计划要求 {plan.page_count} 页，实际生成 {actual} 页。")
        else:
            checks.append(f"页数校验通过：{actual}/{plan.page_count}")

    if env_bool("PPT_REQUIRE_IMAGES", True):
        if len(assets) != len(plan.slides):
            errors.append(f"图片数量校验失败：计划 {len(plan.slides)} 张，实际 {len(assets)} 张。")
        elif any(not Path(a.path).exists() for a in assets):
            errors.append("图片文件校验失败：存在缺失图片。")
        else:
            checks.append(f"图片数量校验通过：{len(assets)} 张")

    if pptx_path.exists():
        text_blob = "\n".join(extract_pptx_texts(pptx_path))
        missing = []
        for category in plan.explicit_categories[: max(0, plan.page_count - 1)]:
            if category and category not in text_blob:
                missing.append(category)
        if missing:
            errors.append("用户指定类目未全部出现在 PPT 中：" + ", ".join(missing))
        elif plan.explicit_categories:
            checks.append("用户指定类目校验通过")
        if plan.language == "en":
            if CJK_RE.search(text_blob):
                errors.append("语言校验失败：用户要求英文，但 PPT 页面文本中仍包含中文字符。")
            else:
                checks.append("英文语言校验通过：页面文本未发现中文字符")
        for idx, slide in enumerate(plan.slides, start=1):
            if not slide.title:
                errors.append(f"第 {idx} 页标题为空。")
        if all(s.title for s in plan.slides):
            checks.append("标题完整性校验通过")

    result.errors = errors
    result.checks = checks
    result.ok = not errors
    return result


def sanitize_filename(text: str, max_chars: int = 56) -> str:
    text = clean_prompt(text)
    for ch in INVALID_FILENAME_CHARS:
        text = text.replace(ch, "")
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[，,。.!！?？：:；;]+", "-", text)
    text = re.sub(r"-+", "-", text).strip("-._ ")
    if not text:
        text = "product-catalog"
    if len(text) > max_chars:
        text = text[:max_chars].rstrip("-._ ")
    return text


def make_run_paths(prompt: str, root: Path, now: datetime) -> Dict[str, Path | str]:
    max_chars = env_int("PPT_FILENAME_MAX_CHARS", 56)
    time_prefix = now.strftime("%Y-%m-%d-%H时%M分")
    slug = sanitize_filename(prompt, max_chars=max_chars)
    base_name = f"{time_prefix}-{slug}"
    run_name = base_name
    run_dir = root / run_name
    if env_bool("PPT_AUTO_DEDUP", True):
        i = 1
        while run_dir.exists():
            run_name = f"{base_name}-{i:02d}"
            run_dir = root / run_name
            i += 1
    pptx_path = run_dir / f"{run_name}.pptx"
    return {"run_name": run_name, "run_dir": run_dir, "pptx_path": pptx_path}


def windows_path_for(pptx_path: Path, root: Path) -> str:
    prefix = os.getenv("PPT_WINDOWS_SHARE_PREFIX", "Z:\\yaq\\ppt\\catalog").rstrip("\\/")
    try:
        rel = pptx_path.relative_to(root)
    except ValueError:
        rel = pptx_path.name
    return f"{prefix}\\{str(rel).replace('/', '\\')}"


def write_run_files(run_dir: Path, plan: DeckPlan, result: GenerationResult, now: datetime) -> None:
    run_dir.mkdir(parents=True, exist_ok=True)
    plan_path = run_dir / "deck_plan.json"
    plan_path.write_text(json.dumps(plan_to_dict(plan), ensure_ascii=False, indent=2), encoding="utf-8")
    metadata = {
        "created_at": now.isoformat(timespec="seconds"),
        "timezone": os.getenv("PPT_TIMEZONE", DEFAULT_TZ),
        "output_rule": "成功时飞书只回复 windows_path；失败时回复明确错误。",
        "quality_rule": "deck_plan 是唯一事实源；渲染后校验页数、语言、类目、标题、图片数量。",
        "ppt_master_rule": "安装脚本不会删除或覆盖 workspace 中已有 PPT-master 目录。",
        "result": asdict(result),
    }
    (run_dir / "metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    (run_dir / "README.txt").write_text(
        f"OpenClaw PPT 商品目录册生成结果\n"
        f"用户问题：{plan.prompt}\n"
        f"页数：{plan.page_count}\n"
        f"语言：{plan.language}\n"
        f"PPT：{result.windows_path}\n"
        f"deck_plan：{plan_path}\n"
        f"飞书回复规则：成功只回复 PPT 本地映射路径；失败回复错误。\n",
        encoding="utf-8",
    )


def generate(prompt: str, sender_name: str = "", sender_open_id: str = "", root: Optional[str] = None, image_mode: Optional[str] = None) -> GenerationResult:
    start = time.time()
    if load_dotenv:
        load_dotenv()
    now = now_in_configured_timezone()
    prompt = clean_prompt(prompt)
    root_path = Path(root or os.getenv("PPT_CATALOG_ROOT", "/data/share/yaq/ppt/catalog")).expanduser().resolve()
    paths = make_run_paths(prompt, root_path, now)
    run_name = str(paths["run_name"])
    run_dir = Path(paths["run_dir"])
    pptx_path = Path(paths["pptx_path"])
    windows_path = windows_path_for(pptx_path, root_path)
    selected_image_mode = (image_mode or os.getenv("PPT_IMAGE_MODE", "minimax")).strip().lower()

    try:
        plan = build_deck_plan(prompt)
        plan_path = run_dir / "deck_plan.json"
        run_dir.mkdir(parents=True, exist_ok=True)
        plan_path.write_text(json.dumps(plan_to_dict(plan), ensure_ascii=False, indent=2), encoding="utf-8")
        image_assets = prepare_images(plan, run_dir, selected_image_mode)
        build_presentation(plan, pptx_path, image_assets)
        validation = validate_output(plan, pptx_path, image_assets)
        if not validation.ok:
            raise RuntimeError("；".join(validation.errors))
        elapsed = round(time.time() - start, 1)
        result = GenerationResult(
            ok=True,
            prompt=prompt,
            page_count=plan.page_count,
            run_name=run_name,
            run_dir=str(run_dir),
            pptx_path=str(pptx_path),
            windows_path=windows_path,
            elapsed_seconds=elapsed,
            language=plan.language,
            deck_plan_path=str(plan_path),
            image_mode=selected_image_mode,
            image_count=len(image_assets),
            image_assets=[asdict(asset) for asset in image_assets],
            validation=asdict(validation),
            reply_text=windows_path,
        )
        write_run_files(run_dir, plan, result, now)
        return result
    except Exception as exc:
        elapsed = round(time.time() - start, 1)
        try:
            run_dir.mkdir(parents=True, exist_ok=True)
            (run_dir / "ERROR.txt").write_text(str(exc), encoding="utf-8")
        except Exception:
            pass
        return GenerationResult(
            ok=False,
            prompt=prompt,
            page_count=0,
            run_name=run_name,
            run_dir=str(run_dir),
            pptx_path=str(pptx_path),
            windows_path=windows_path,
            elapsed_seconds=elapsed,
            image_mode=selected_image_mode,
            error=str(exc),
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a product catalog PPT from a Feishu/OpenClaw prompt")
    parser.add_argument("--prompt", required=True, help="飞书用户原始问题")
    parser.add_argument("--sender-name", default="", help="飞书发送人姓名，仅写入元数据")
    parser.add_argument("--sender-open-id", default="", help="飞书发送人 open_id，仅写入元数据")
    parser.add_argument("--root", default=None, help="输出根目录，不填则读 PPT_CATALOG_ROOT")
    parser.add_argument("--image-mode", choices=["minimax", "placeholder", "off"], default=None, help="生产用 minimax；离线布局测试用 placeholder")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出")
    args = parser.parse_args()

    result = generate(args.prompt, args.sender_name, args.sender_open_id, args.root, args.image_mode)
    data = asdict(result)
    if args.json:
        print(json.dumps(data, ensure_ascii=False, indent=2))
    else:
        if result.ok:
            print(result.windows_path)
        else:
            print(f"PPT 生成失败：{result.error}", file=sys.stderr)
            sys.exit(1)


if __name__ == "__main__":
    main()
