# 飞书接入说明 v1.3

## 预检 MiniMax

部署后先执行：

```bash
python scripts/check_minimax_image_api.py
```

如果要真实验证生图：

```bash
python scripts/check_minimax_image_api.py --live --out /tmp/check_minimax.png
```

只有这里通过后，飞书触发才会真正生成带图 PPT。

## 目标

飞书用户输入：

```text
生成一个厨房餐具相关的3页的商品目录册PPT
```

Agent 调用：

```bash
python scripts/generate_catalog_ppt_v13.py \
  --prompt "生成一个厨房餐具相关的3页的商品目录册PPT" \
  --sender-name "陈玉" \
  --sender-open-id "ou_xxx" \
  --json
```

然后把 JSON 交给 formatter：

```bash
python scripts/generate_catalog_ppt_v13.py --prompt "$USER_TEXT" --json \
  | python scripts/feishu_reply_formatter.py
```

## v1.3 飞书回复规则

成功时只回复一行本地路径：

```text
Z:\yaq\ppt\catalog\2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT\2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT.pptx
```

不要回复：

```text
已完成，PPT 文件：
服务器路径：
已完成 · 3页 · 耗时 12.6s
```

原因：OpenClaw/飞书外层可能已经有耗时提示，脚本再回复会导致内容重复、显得很乱。

## MiniMax 生图规则

生产环境 `.env` 使用：

```bash
PPT_IMAGE_MODE=minimax
PPT_REQUIRE_IMAGES=1
MINIMAX_API_KEY=你的 MiniMax API Key
MINIMAX_IMAGE_API_URL=https://api.minimax.io/v1/image_generation
MINIMAX_NETWORK_PRECHECK=1
MINIMAX_IMAGE_MODEL=image-01
```

如果 `.env` 没写 `MINIMAX_API_KEY`，脚本会尝试读取：

```text
~/.openclaw/openclaw.json
```

中的 MiniMax provider key。

如果 MiniMax 密钥缺失或接口失败，脚本会返回失败，不会继续生成纯文字 PPT。

## 输出目录规则

PPT 文件直接位于任务目录下一级：

```text
/data/share/yaq/ppt/catalog/2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT/2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT.pptx
```

禁止回退到：

```text
project/exports/catalog_lite_v1.pptx
```

## 时间规则

使用：

```bash
PPT_TIMEZONE=Asia/Shanghai
```

文件名时间格式：

```text
YYYY-MM-DD-HH时MM分-用户询问的问题
```

例如：

```text
2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT
```
