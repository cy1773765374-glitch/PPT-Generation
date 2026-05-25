#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from generate_catalog_ppt_v14 import extract_page_count, extract_category, sanitize_filename, make_run_paths, normalize_minimax_image_api_url


def test_page_count():
    assert extract_page_count("生成一个厨房餐具相关的3页的商品目录册PPT", 5, 20) == 3
    assert extract_page_count("做一份厨房餐具三页PPT", 5, 20) == 3
    assert extract_page_count("做厨房餐具目录册", 5, 20) == 5
    assert extract_page_count("做厨房餐具100页PPT", 5, 20) == 20


def test_category():
    assert extract_category("生成一个厨房餐具相关的3页的商品目录册PPT") == "厨房餐具"


def test_filename():
    name = sanitize_filename("回复 陈玉: 生成一个厨房餐具相关的3页的商品目录册PPT")
    assert "/" not in name
    assert "\\" not in name
    assert "回复" not in name


def test_minimax_url_normalize():
    assert normalize_minimax_image_api_url("https://api.minimax.com/v1/image_generation") == "https://api.minimax.com/v1/image_generation"
    assert normalize_minimax_image_api_url("https://api.minimax.com") == "https://api.minimax.com/v1/image_generation"
    assert normalize_minimax_image_api_url("api.minimax.com") == "https://api.minimax.com/v1/image_generation"


def test_run_name_time_format():
    now = datetime(2026, 5, 25, 16, 35, tzinfo=ZoneInfo("Asia/Shanghai"))
    paths = make_run_paths("生成一个厨房餐具相关的3页的商品目录册PPT", Path("/tmp/ppt"), now)
    assert str(paths["run_name"]).startswith("2026-05-25-16时35分-")
    assert "-1635-" not in str(paths["run_name"])


if __name__ == "__main__":
    test_page_count()
    test_category()
    test_filename()
    test_minimax_url_normalize()
    test_run_name_time_format()
    print("ok")
