# workspace-PPT-Generation v1.4

## 版本定位

v1.4 是针对 v1.2 线上测试失败的修正版，重点解决：

1. **MiniMax 生图域名策略**：默认使用 `https://api.minimax.com/v1/image_generation`，不再自动改写域名。
2. **失败时找不到原因**：生图失败时不会生成伪 PPT，但会在任务目录下写入 `ERROR.txt`，方便排查。
3. **飞书成功回复仍保持一行路径**：成功后仍只回复 `Z:\...pptx`，不回复服务器路径、页数、耗时。
4. **PPT 必须有图**：生产默认 `PPT_IMAGE_MODE=minimax`、`PPT_REQUIRE_IMAGES=1`，MiniMax 不可用就失败，不退化成纯文字 PPT。
5. **兼容旧入口**：如果 Feishu/OpenClaw 旧配置仍调用 `scripts/generate_catalog_ppt_v12.py`，会自动转到 v1.4 实现。

---

## 推荐解压方式

```bash
mkdir -p ~/.openclaw/workspace-PPT-Generation

tar -xzvf workspace-PPT-Generation-v1.4.tar.gz \
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

如果你已经在 `~/.openclaw/openclaw.json` 配了 MiniMax key，也可以不写 `MINIMAX_API_KEY`，脚本会尝试自动读取。

---

## 先做 MiniMax 接口预检

只检查 key + 域名解析：

```bash
python scripts/check_minimax_image_api.py
```

执行一次真实生图请求：

```bash
python scripts/check_minimax_image_api.py --live --out /tmp/check_minimax.png
```

如果这里失败，说明不是 PPT 代码问题，而是服务器到 MiniMax 图像接口的 DNS、代理、网络或 API Key 问题。

---

## 生产测试

```bash
cd ~/.openclaw/workspace-PPT-Generation
source .venv/bin/activate

python scripts/generate_catalog_ppt_v13.py \
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
  "reply_text": "Z:\\yaq\\ppt\\catalog\\2026-05-25-18时32分-生成一个厨房餐具相关的3页的商品目录册PPT\\2026-05-25-18时32分-生成一个厨房餐具相关的3页的商品目录册PPT.pptx"
}
```

飞书最终只发送 `reply_text`。

---

## 飞书回复格式

成功时只回复一行：

```text
Z:\yaq\ppt\catalog\2026-05-25-18时32分-生成一个厨房餐具相关的3页的商品目录册PPT\2026-05-25-18时32分-生成一个厨房餐具相关的3页的商品目录册PPT.pptx
```

不要再回复服务器路径、页数、耗时。

---

## 离线布局测试

仅用于确认 PPT 布局，不用于飞书生产：

```bash
PPT_IMAGE_MODE=placeholder python scripts/generate_catalog_ppt_v13.py \
  --prompt "生成一个厨房餐具相关的3页的商品目录册PPT" \
  --root /tmp/ppt_catalog_test \
  --json
```

生产环境不要把 `PPT_IMAGE_MODE` 改成 `placeholder`，否则就不是 MiniMax 生图版本。
