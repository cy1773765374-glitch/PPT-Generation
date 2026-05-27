# workspace-PPT-Generation · PPT-master OpenClaw 适配版

本包是在“内容自适应优化版”基础上继续改造的 PPT-master 接入版本。主入口仍然是 OpenClaw/飞书可调用的 `scripts/generate_catalog_ppt.py`，但渲染链路默认变为：

```text
用户描述 / 飞书消息
→ CatalogIntent / DeckPlan / SlidePlan
→ MiniMax 或 placeholder 图片资产
→ pptmaster_project/design_spec.md + spec_lock.md + svg_output/*.svg
→ vendor/ppt-master/scripts/svg_quality_checker.py
→ vendor/ppt-master/scripts/finalize_svg.py
→ vendor/ppt-master/scripts/svg_to_pptx.py
→ native editable PPTX
```

如果服务器上没有完整 PPT-master，本地测试会按 `auto` 策略降级到旧的 `python-pptx` 渲染器；生产环境建议设置 `PPT_RENDER_BACKEND=pptmaster`，让 PPT-master 缺失或失败时直接报错，避免误以为已经走了 PPT-master。

## 这版解决的问题

1. 不再只是保留 `PPT-master/` 空目录，而是新增 `scripts/openclaw_pptmaster_adapter.py`，真正生成 PPT-master 项目并调用 vendor 脚本。
2. 默认查找你当前服务器路径：`/home/cy/.openclaw/workspace-PPT-Generation/vendor/ppt-master/`。
3. `generate_catalog_ppt.py` 新增 `PPT_RENDER_BACKEND=auto|pptmaster|python_pptx`。
4. 每次生成都会在输出目录下留下 `pptmaster_project/`，包括 `design_spec.md`、`spec_lock.md`、`svg_output/`、`assets/`、`pptmaster_pipeline.log`。
5. 内容规划仍然保持之前的规则：用户没要求的字段不强塞，类目和页数按用户描述走，风格根据用户描述变化。
6. 图片仍然默认走 MiniMax；离线测试可用 `PPT_IMAGE_MODE=placeholder`。
7. PPT-master 不负责理解飞书业务；OpenClaw 层负责商品目录册规划，PPT-master 只负责 SVG 后处理与 PPTX 导出。

## 目录结构

```text
workspace-PPT-Generation/
├── AGENTS.md
├── TOOLS.md
├── README.md
├── .env.example
├── .gitignore
├── install_update.sh
├── vendor/
│   └── ppt-master/
│       └── README.md                # 占位说明；实际完整仓库保留在服务器本地
├── templates/
│   ├── catalog_styles/
│   └── catalog_page_types/
├── scripts/
│   ├── generate_catalog_ppt.py       # OpenClaw / 飞书主入口
│   ├── openclaw_pptmaster_adapter.py # PPT-master 适配层
│   ├── check_minimax_image_api.py
│   ├── feishu_reply_formatter.py
│   └── run_local_test.sh
├── docs/
│   ├── FEISHU_INTEGRATION.md
│   └── CHANGELOG.md
├── examples/
│   └── sample_request.json
└── tests/
    ├── test_plan.py
    └── test_parse.py
```

## 覆盖部署

```bash
mkdir -p ~/.openclaw/workspace-PPT-Generation

tar -xzvf workspace-PPT-Generation-pptmaster-openclaw.tar.gz \
  -C ~/.openclaw/workspace-PPT-Generation \
  --strip-components=1

cd ~/.openclaw/workspace-PPT-Generation
bash install_update.sh ~/.openclaw/workspace-PPT-Generation
```

`install_update.sh` 不会删除或覆盖：

```text
.env
.venv/
vendor/ppt-master/
PPT-master/
```

所以你现在的完整 PPT-master 目录：

```text
/home/cy/.openclaw/workspace-PPT-Generation/vendor/ppt-master/
```

会被保留。

## 推荐 .env 配置

```bash
PPT_CATALOG_ROOT=/data/share/yaq/ppt/catalog
PPT_WINDOWS_SHARE_PREFIX=Z:\\yaq\\ppt\\catalog
PPT_TIMEZONE=Asia/Shanghai

# 生产建议强制 pptmaster；本地调试可用 auto。
PPT_RENDER_BACKEND=pptmaster
PPT_MASTER_ROOT=/home/cy/.openclaw/workspace-PPT-Generation/vendor/ppt-master
PPT_MASTER_STRICT=1
PPT_MASTER_TIMEOUT=240
PPT_MASTER_SKIP_QUALITY=0

# 生图
PPT_IMAGE_MODE=minimax
PPT_REQUIRE_IMAGES=1
MINIMAX_API_KEY=你的 MiniMax API Key
MINIMAX_IMAGE_API_URL=https://api.minimax.com/v1/image_generation
MINIMAX_NETWORK_PRECHECK=0
```

如果 `~/.openclaw/openclaw.json` 已配置 MiniMax key，也可以不写 `MINIMAX_API_KEY`，脚本会尝试自动读取。

## 主入口

```bash
python scripts/generate_catalog_ppt.py \
  --prompt "Generate a 10-page festive party theme product catalog PPT for SELLERS UNION. Page 1: Company introduction for SELLERS UNION - a premier party supplies and festive decorations supplier. Pages 2-10: Each page features one product category with party-related images and English descriptions. Product categories: Balloons and Balloon Sets, Rain Curtains, Candles, Bunting Garlands, Hats, Party Blowers, Disposable Party Tableware, COS Costumes, and other party decorations. All text in English. High-quality commercial product catalog style." \
  --json
```

成功 JSON 中重点看：

```json
{
  "ok": true,
  "page_count": 10,
  "language": "en",
  "image_count": 10,
  "render_backend": "pptmaster",
  "pptmaster_project_dir": "/data/share/yaq/ppt/catalog/.../pptmaster_project",
  "pptmaster_log_path": "/data/share/yaq/ppt/catalog/.../pptmaster_project/pptmaster_pipeline.log",
  "validation": {"ok": true},
  "reply_text": "Z:\\yaq\\ppt\\catalog\\...\\xxx.pptx"
}
```

飞书最终只发送 `reply_text`。

## 渲染后端说明

### 1. `PPT_RENDER_BACKEND=pptmaster`

生产推荐。必须找到完整 PPT-master，并且 `svg_quality_checker.py / finalize_svg.py / svg_to_pptx.py` 成功，否则直接失败。

### 2. `PPT_RENDER_BACKEND=auto`

默认值。优先尝试 PPT-master；如果 vendor 缺失或执行失败，则生成 `PPT_MASTER_FALLBACK.txt` 并降级到 `python-pptx`。适合本地开发和测试。

### 3. `PPT_RENDER_BACKEND=python_pptx`

完全不走 PPT-master，只保留旧渲染器。用于排障对比。

## PPT-master 项目产物

每次运行会生成：

```text
输出目录/
├── deck_plan.json
├── metadata.json
├── README.txt
├── xxx.pptx
└── pptmaster_project/
    ├── design_spec.md
    ├── spec_lock.md
    ├── deck_plan.json
    ├── openclaw_manifest.json
    ├── total.md
    ├── assets/
    ├── svg_output/
    ├── svg_final/
    ├── exports/
    └── pptmaster_pipeline.log
```

`spec_lock.md` 会固化这些规则：

```text
- deck_plan 是唯一事实源
- 页数、类目、语言不能由渲染层重算
- 用户没要求的价格、MOQ、箱规、认证、交期、包装、SKU 不主动出现
- 图片不出现文字、水印、商标、品牌包装、IP 角色
- 每页 layout_variant 不同，避免所有页面长得一样
```

## 内容自适应规则

- 用户明确要求公司介绍时，才生成公司介绍页。
- 类目数量已经覆盖页数时，不强制增加封面。
- 用户没有要求的采购字段不主动出现，例如价格、MOQ、箱规、认证、交期、包装。
- 用户明确要求的字段优先展示；字段较多时，不用“产品定位”挤掉用户点名字段。
- 支持否定约束：`不要价格MOQ箱规`、`只要产品方向`、`without price/MOQ` 等。
- 支持风格词识别：`极简/留白/minimal`、`高端/轻奢/premium/luxury`、`活泼/儿童/playful`、`节庆/派对/festive`、`自然/环保/nature`、`科技/智能/tech`。

## 离线测试

离线只验证解析、排版和校验，不调用 MiniMax。没有完整 vendor 时会降级到 `python-pptx`：

```bash
PPT_RENDER_BACKEND=auto PPT_IMAGE_MODE=placeholder PPT_REQUIRE_IMAGES=1 \
python scripts/generate_catalog_ppt.py \
  --prompt "Generate a 3-page playful stationery product catalog PPT. Product categories: Notebooks, Pens, Erasers. All text in English." \
  --root /tmp/ppt_catalog_test \
  --json
```

强制验证 PPT-master：

```bash
PPT_RENDER_BACKEND=pptmaster PPT_MASTER_STRICT=1 PPT_IMAGE_MODE=placeholder PPT_REQUIRE_IMAGES=1 \
python scripts/generate_catalog_ppt.py \
  --prompt "Generate a 3-page playful stationery product catalog PPT. Product categories: Notebooks, Pens, Erasers. All text in English." \
  --root /tmp/ppt_catalog_test \
  --json
```

单元测试：

```bash
pytest -q
```

当前包内离线测试结果：`20 passed`。

## MiniMax 预检

```bash
python scripts/check_minimax_image_api.py
python scripts/check_minimax_image_api.py --live --out /tmp/check_minimax.png
```

如果这里失败，优先排查服务器到 MiniMax 图像接口的 DNS、代理、网络或 API Key。

## 飞书回复规则

成功时只回复一行：

```text
Z:\yaq\ppt\catalog\...\xxx.pptx
```

失败时回复：

```text
PPT 生成失败：具体错误
```

不要回复服务器 Linux 路径、耗时、页数或“已完成”。
