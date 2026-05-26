# workspace-PPT-Generation 最终版

本包是在 1.5 压缩包基础上直接重构后的最终交付包，不再沿用 v1.6、v1.7 这种中间版本命名。

## 解决的问题

1. 详细英文需求不再被压成固定 5 页模板。
2. `10-page`、`Pages 2-10`、`Product categories:` 会进入 `deck_plan.json`，后续渲染只能按该计划执行。
3. 用户要求 `All text in English` 时，PPT 页面文本禁止出现中文硬编码字段。
4. 用户显式给出的 9 个产品类目会逐页展开，不再生成“核心款 A / 升级款 B / 组合款 C / 展示款 D”。
5. 每页都有独立 image brief，MiniMax 生图 prompt 与当前页面类目绑定。
6. 输出前做硬校验：页数、语言、类目、标题、图片数量，不通过则失败，不返回“已完成”。
7. 保留 `PPT-master/`：安装脚本不会删除或覆盖已有 `PPT-master` 目录。

## 目录结构

```text
workspace-PPT-Generation/
├── AGENTS.md
├── TOOLS.md
├── README.md
├── .env.example
├── .gitignore
├── install_update.sh
├── PPT-master/
│   └── README.md
├── scripts/
│   ├── generate_catalog_ppt.py
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

tar -xzvf workspace-PPT-Generation-final.tar.gz \
  -C ~/.openclaw/workspace-PPT-Generation \
  --strip-components=1

cd ~/.openclaw/workspace-PPT-Generation
bash install_update.sh ~/.openclaw/workspace-PPT-Generation
```

`install_update.sh` 会清理旧的 `generate_catalog_ppt_v12.py / v13.py / v14.py / v15.py` 入口，避免飞书继续调用旧逻辑；同时不会删除 `.env`、`.venv` 和 `PPT-master/`。

## 安装依赖

如果不使用安装脚本，也可以手动执行：

```bash
cd ~/.openclaw/workspace-PPT-Generation
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

编辑 `.env`，重点确认：

```bash
PPT_CATALOG_ROOT=/data/share/yaq/ppt/catalog
PPT_WINDOWS_SHARE_PREFIX=Z:\\yaq\\ppt\\catalog
PPT_TIMEZONE=Asia/Shanghai
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

成功 JSON 中应包含：

```json
{
  "ok": true,
  "page_count": 10,
  "language": "en",
  "image_count": 10,
  "validation": {"ok": true},
  "reply_text": "Z:\\yaq\\ppt\\catalog\\...\\xxx.pptx"
}
```

飞书最终只发送 `reply_text`。

## 离线测试

离线只验证解析、排版和校验，不调用 MiniMax：

```bash
PPT_IMAGE_MODE=placeholder PPT_REQUIRE_IMAGES=1 python scripts/generate_catalog_ppt.py \
  --prompt "Generate a 10-page festive party theme product catalog PPT for SELLERS UNION. Page 1: Company introduction for SELLERS UNION - a premier party supplies and festive decorations supplier. Pages 2-10: Each page features one product category with party-related images and English descriptions. Product categories: Balloons and Balloon Sets, Rain Curtains, Candles, Bunting Garlands, Hats, Party Blowers, Disposable Party Tableware, COS Costumes, and other party decorations. All text in English. High-quality commercial product catalog style." \
  --root /tmp/ppt_catalog_test \
  --json
```

或者直接：

```bash
bash scripts/run_local_test.sh
python tests/test_plan.py
```

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
