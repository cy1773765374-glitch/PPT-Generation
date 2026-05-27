# CHANGELOG

## PPT-master OpenClaw 适配版

- 新增 `scripts/openclaw_pptmaster_adapter.py`，把 OpenClaw 的 `DeckPlan` 转为 PPT-master 项目目录。
- 新增 `vendor/ppt-master/` 作为默认 PPT-master 完整仓库位置，适配用户服务器路径 `/home/cy/.openclaw/workspace-PPT-Generation/vendor/ppt-master/`。
- 主入口新增 `PPT_RENDER_BACKEND=auto|pptmaster|python_pptx`。
- 默认按 `auto` 优先调用 PPT-master；生产可设置 `PPT_RENDER_BACKEND=pptmaster` 和 `PPT_MASTER_STRICT=1` 强制走 PPT-master。
- 每次生成保留 `pptmaster_project/design_spec.md`、`spec_lock.md`、`svg_output/`、`assets/`、`pptmaster_pipeline.log`，方便排障。
- 调用 PPT-master 的 `svg_quality_checker.py`、`finalize_svg.py`、`svg_to_pptx.py` 导出 native editable PPTX。
- 安装脚本明确保留 `vendor/ppt-master/` 和 `PPT-master/`，不覆盖用户本地完整 PPT-master 仓库。

## 内容自适应优化版

- 新增内容字段白名单：价格、MOQ、箱规、认证、包装、交期、SKU 等只有用户明确要求时才展示。
- 新增否定约束识别：支持“不要价格 MOQ 箱规”“只要产品方向”“without price/MOQ”等。
- 新增自适应风格引擎：根据 minimal/luxury/playful/festive/nature/tech 等风格词切换配色与版式。
- 类目数量覆盖页数时不再强制生成封面；公司介绍页只在用户明确要求时生成。
- 类目页不再固定生成 Page Brief、Suggested SKUs、Packaging & Sourcing Notes 等模板化内容。
- 新增覆盖上述规则的单元测试。

## 最终版

- 主入口统一为 `scripts/generate_catalog_ppt.py`，删除 v12/v13/v14/v15 旧入口。
- 新增 `deck_plan.json` 唯一事实源。
- 支持英文详细需求：`10-page`、`Page 1`、`Pages 2-10`、`Product categories`、`All text in English`。
- 英文需求下页面文本不再含中文硬编码字段。
- 用户指定类目逐页展开，不再使用“核心款 A / 升级款 B / 组合款 C / 展示款 D”兜底模板。
- 每页生成独立 image brief，MiniMax prompt 与当前页类目绑定。
- 新增硬校验：页数、语言、类目、标题、图片数量。
- 安装脚本保留 `.env`、`.venv`、`PPT-master/`。
## fix3 - 类目完整性校验修正

- 修复 PPT-master 导出后长类目被换行、拆分 text run 或 `&/and` 表达差异导致的误判失败。
- 类目硬校验改为以 deck_plan 中的 category page 计划为准，PPTX 文本抽取仅作为辅助检查。
- 覆盖 SELLERS UNION 10 页 party catalog 示例中的 `Balloons and Balloon Sets`、`Disposable Party Tableware` 长类目。
- 测试数：26 passed。
