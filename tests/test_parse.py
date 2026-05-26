#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from test_plan import *  # noqa

if __name__ == "__main__":
    test_detailed_english_plan()
    test_extract_categories_does_not_split_balloon_and_sets()
    test_page_count_patterns()
    test_language_detection()
    test_filename_keeps_normal_english_letters()
    test_placeholder_generation_exact_10_pages_and_english_text()
    print("ok")
