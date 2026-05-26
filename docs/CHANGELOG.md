# CHANGELOG

## 最终版

- 主入口统一为 `scripts/generate_catalog_ppt.py`，删除 v12/v13/v14/v15 旧入口。
- 新增 `deck_plan.json` 唯一事实源。
- 支持英文详细需求：`10-page`、`Page 1`、`Pages 2-10`、`Product categories`、`All text in English`。
- 英文需求下页面文本不再含中文硬编码字段。
- 用户指定类目逐页展开，不再使用“核心款 A / 升级款 B / 组合款 C / 展示款 D”兜底模板。
- 每页生成独立 image brief，MiniMax prompt 与当前页类目绑定。
- 新增硬校验：页数、语言、类目、标题、图片数量。
- 安装脚本保留 `.env`、`.venv`、`PPT-master/`。
