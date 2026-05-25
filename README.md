# workspace-PPT-Generation v1.2

## 版本定位

v1.2 是基于 v1.1 的修正版，重点解决：

1. 飞书回复太啰嗦：现在只回复 `Z:\...pptx` 本地映射盘路径。
2. 文件名时间错误/不清楚：现在固定 `Asia/Shanghai`，格式为 `YYYY-MM-DD-HH时MM分-用户询问的问题`。
3. PPT 只有文字：现在生产默认调用 MiniMax `image-01` 生图，并把封面图、商品图插入 PPT。
4. 避免纯文字退化：默认 `PPT_REQUIRE_IMAGES=1`，MiniMax 不可用就失败，不继续生成干巴巴的文字版。
5. 输出结构保持：PPT 文件直接放在任务目录一级，不再进入 `project/exports/`。

---

## 推荐解压方式

```bash
mkdir -p ~/.openclaw/workspace-PPT-Generation

tar -xzvf workspace-PPT-Generation-v1.2.tar.gz \
  -C ~/.openclaw/workspace-PPT-Generation \
  --strip-components=1
```

---

## 安装依赖

```bash
cd ~/.openclaw/workspace-PPT-Generation
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

编辑 `.env`：

```bash
PPT_CATALOG_ROOT=/data/share/yaq/ppt/catalog
PPT_WINDOWS_SHARE_PREFIX=Z:\\yaq\\ppt\\catalog
PPT_TIMEZONE=Asia/Shanghai
PPT_IMAGE_MODE=minimax
PPT_REQUIRE_IMAGES=1
MINIMAX_API_KEY=你的 MiniMax API Key
```

如果你已经在 `~/.openclaw/openclaw.json` 配了 MiniMax key，也可以不写 `MINIMAX_API_KEY`，脚本会尝试自动读取。

---

## 生产测试

```bash
cd ~/.openclaw/workspace-PPT-Generation
source .venv/bin/activate

python scripts/generate_catalog_ppt_v12.py \
  --prompt "生成一个厨房餐具相关的3页的商品目录册PPT" \
  --sender-name "陈玉" \
  --sender-open-id "ou_6e056040c28331827575c0061644569c" \
  --json
```

成功 JSON 中应包含：

```json
{
  "ok": true,
  "page_count": 3,
  "image_mode": "minimax",
  "image_count": 5,
  "reply_text": "Z:\\yaq\\ppt\\catalog\\2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT\\2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT.pptx"
}
```

飞书最终只发送 `reply_text`。

---

## 离线连通性测试

无 MiniMax key 时可以只验证 PPT 布局：

```bash
PPT_IMAGE_MODE=placeholder python scripts/generate_catalog_ppt_v12.py \
  --prompt "生成一个厨房餐具相关的3页的商品目录册PPT" \
  --root /tmp/ppt_catalog_test \
  --json
```

注意：`placeholder` 只用于本地测试，不用于飞书生产。

---

## 飞书回复格式

v1.2 成功时只回复：

```text
Z:\yaq\ppt\catalog\2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT\2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT.pptx
```

不要再回复服务器路径、页数、耗时。
