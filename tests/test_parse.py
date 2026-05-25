#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from generate_catalog_ppt_v11 import extract_page_count, extract_category, sanitize_filename


def test_page_count():
    assert extract_page_count("生成一个厨房餐具相关的3页的商品目录册PPT", 5, 20) == 3
    assert extract_page_count("做一份厨房餐具三页PPT", 5, 20) == 3
    assert extract_page_count("做厨房餐具目录册", 5, 20) == 5
    assert extract_page_count("做厨房餐具100页PPT", 5, 20) == 20


def test_category():
    assert extract_category("生成一个厨房餐具相关的3页的商品目录册PPT") == "厨房餐具"


def test_filename():
    name = sanitize_filename("生成一个厨房餐具相关的3页的商品目录册PPT")
    assert "/" not in name
    assert "\\" not in name


if __name__ == "__main__":
    test_page_count()
    test_category()
    test_filename()
    print("ok")
