#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""兼容入口：旧集成如果仍调用 generate_catalog_ppt_v12 文件，会自动转到 v1.5 实现。"""
from generate_catalog_ppt_v15 import main

if __name__ == "__main__":
    main()
