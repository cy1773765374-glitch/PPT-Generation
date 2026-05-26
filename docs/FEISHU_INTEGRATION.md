# 飞书接入说明

## 目标

飞书用户输入详细商品目录册需求时，Agent 生成符合页数、语言、类目顺序和图片要求的 PPT，并只回复 Windows/Samba 本地路径。

## OpenClaw 调用方式

推荐主入口：

```bash
cd ~/.openclaw/workspace-PPT-Generation
source .venv/bin/activate

python scripts/generate_catalog_ppt.py \
  --prompt "$USER_TEXT" \
  --sender-name "$SENDER_NAME" \
  --sender-open-id "$OPEN_ID" \
  --json
```

再交给 formatter：

```bash
python scripts/generate_catalog_ppt.py --prompt "$USER_TEXT" --json \
  | python scripts/feishu_reply_formatter.py
```

## 成功回复

只回复：

```text
Z:\yaq\ppt\catalog\...\xxx.pptx
```

不要附加服务器路径、页数、耗时或“已完成”。

## 失败回复

```text
PPT 生成失败：具体错误
```

如果页数、语言、类目、标题或图片数量校验失败，也必须走失败回复。

## 生产环境变量

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

如果服务器通过代理访问 MiniMax，可设置：

```bash
MINIMAX_PROXY_URL=http://127.0.0.1:17897
```

## 详细需求示例

```text
Generate a 10-page festive party theme product catalog PPT for SELLERS UNION. Page 1: Company introduction for SELLERS UNION - a premier party supplies and festive decorations supplier. Pages 2-10: Each page features one product category with party-related images and English descriptions. Product categories: Balloons and Balloon Sets, Rain Curtains, Candles, Bunting Garlands, Hats, Party Blowers, Disposable Party Tableware, COS Costumes, and other party decorations. All text in English. High-quality commercial product catalog style.
```

应生成 10 页：

```text
1. SELLERS UNION company introduction
2. Balloons and Balloon Sets
3. Rain Curtains
4. Candles
5. Bunting Garlands
6. Hats
7. Party Blowers
8. Disposable Party Tableware
9. COS Costumes
10. Other Party Decorations
```

## 输出文件

每次任务目录下至少包含：

```text
xxx.pptx
assets/
deck_plan.json
metadata.json
README.txt
```

其中 `deck_plan.json` 是排查页数、语言、类目、图片 prompt 的关键文件。

## PPT-master 保留

部署脚本不会删除或覆盖：

```text
~/.openclaw/workspace-PPT-Generation/PPT-master/
```

不要使用会删除该目录的同步命令。
