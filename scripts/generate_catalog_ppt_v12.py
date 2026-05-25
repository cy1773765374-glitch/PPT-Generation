#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""兼容入口：旧集成如果仍调用 v12 文件，会自动转到 v1.3 实现。"""
from generate_catalog_ppt_v13 import main

if __name__ == "__main__":
    main()
