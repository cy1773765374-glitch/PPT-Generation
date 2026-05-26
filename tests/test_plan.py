#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import os
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from generate_catalog_ppt import (  # noqa: E402
    build_deck_plan,
    detect_language,
    extract_page_count,
    extract_product_categories,
    generate,
    sanitize_filename,
)
from pptx import Presentation  # noqa: E402


SELLERS_UNION_PROMPT = (
    "Generate a 10-page festive party theme product catalog PPT for SELLERS UNION. "
    "Page 1: Company introduction for SELLERS UNION - a premier party supplies and festive decorations supplier. "
    "Pages 2-10: Each page features one product category with party-related images and English descriptions. "
    "Product categories: Balloons and Balloon Sets, Rain Curtains, Candles, Bunting Garlands, Hats, Party Blowers, "
    "Disposable Party Tableware, COS Costumes, and other party decorations. All text in English. "
    "High-quality commercial product catalog style."
)


def test_detailed_english_plan():
    plan = build_deck_plan(SELLERS_UNION_PROMPT)
    assert plan.language == "en"
    assert plan.page_count == 10
    assert len(plan.slides) == 10
    assert plan.slides[0].layout == "company_intro"
    assert plan.slides[0].title == "SELLERS UNION"
    assert plan.slides[1].category == "Balloons and Balloon Sets"
    assert plan.slides[-1].category == "Other Party Decorations"
    assert len(plan.image_briefs) == 10


def test_extract_categories_does_not_split_balloon_and_sets():
    cats = extract_product_categories(SELLERS_UNION_PROMPT, "en")
    assert cats[0] == "Balloons and Balloon Sets"
    assert len(cats) == 9


def test_page_count_patterns():
    assert extract_page_count("Generate a 10-page catalog") == 10
    assert extract_page_count("Generate 12 pages catalog") == 12
    assert extract_page_count("生成一个厨房餐具相关的3页的商品目录册PPT") == 3
    assert extract_page_count("做一份厨房餐具三页PPT") == 3


def test_language_detection():
    assert detect_language(SELLERS_UNION_PROMPT) == "en"
    assert detect_language("生成一个厨房餐具相关的3页的商品目录册PPT") == "zh"


def test_filename_keeps_normal_english_letters():
    name = sanitize_filename("Generate a 10-page festive party theme product catalog PPT for SELLERS UNION")
    assert "Generate" in name
    assert "party" in name


def test_placeholder_generation_exact_10_pages_and_english_text():
    with tempfile.TemporaryDirectory() as td:
        os.environ["PPT_REQUIRE_IMAGES"] = "1"
        result = generate(SELLERS_UNION_PROMPT, root=td, image_mode="placeholder")
        assert result.ok, result.error
        prs = Presentation(result.pptx_path)
        assert len(prs.slides) == 10
        text = "\n".join(shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text"))
        assert "商品" not in text
        assert "Balloons and Balloon Sets" in text
        assert "Other Party Decorations" in text
        assert result.validation["ok"] is True


if __name__ == "__main__":
    test_detailed_english_plan()
    test_extract_categories_does_not_split_balloon_and_sets()
    test_page_count_patterns()
    test_language_detection()
    test_filename_keeps_normal_english_letters()
    test_placeholder_generation_exact_10_pages_and_english_text()
    print("ok")
