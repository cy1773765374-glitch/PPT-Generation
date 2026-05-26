# 工具说明：PPT 商品目录册生成

## 主脚本

```text
scripts/generate_catalog_ppt.py
```

作用：

1. 解析飞书用户输入。
2. 识别页数、语言、品牌、主题、产品类目、显式页面结构。
3. 生成 `deck_plan.json` 作为唯一事实来源。
4. 为每页生成独立图片 brief。
5. 调用 MiniMax `image-01` 生成图片。
6. 使用 python-pptx 渲染 PPTX。
7. 做页数、语言、类目、标题、图片数量校验。
8. 输出 JSON，供 OpenClaw/飞书回复。

## 常用调用

```bash
python scripts/generate_catalog_ppt.py \
  --prompt "$USER_TEXT" \
  --sender-name "$SENDER_NAME" \
  --sender-open-id "$OPEN_ID" \
  --json
```

## 参数

| 参数 | 必填 | 说明 |
|---|---:|---|
| `--prompt` | 是 | 飞书用户原始问题 |
| `--sender-name` | 否 | 发送人姓名，仅写入元数据 |
| `--sender-open-id` | 否 | 发送人 open_id，仅写入元数据 |
| `--root` | 否 | 输出根目录，不填则读 `PPT_CATALOG_ROOT` |
| `--image-mode` | 否 | `minimax` / `placeholder` / `off`，生产使用 `minimax` |
| `--json` | 否 | 使用 JSON 输出，推荐开启 |

## 输出 JSON 字段

```json
{
  "ok": true,
  "prompt": "...",
  "page_count": 10,
  "language": "en",
  "run_dir": "/data/share/yaq/ppt/catalog/...",
  "pptx_path": "/data/share/yaq/ppt/catalog/.../...pptx",
  "windows_path": "Z:\\yaq\\ppt\\catalog\\...\\...pptx",
  "deck_plan_path": "/data/share/yaq/ppt/catalog/.../deck_plan.json",
  "image_mode": "minimax",
  "image_count": 10,
  "validation": {"ok": true, "checks": [], "errors": []},
  "reply_text": "Z:\\yaq\\ppt\\catalog\\...\\...pptx"
}
```

## 飞书 formatter

```bash
python scripts/generate_catalog_ppt.py --prompt "$USER_TEXT" --json \
  | python scripts/feishu_reply_formatter.py
```

成功时只输出：

```text
Z:\yaq\ppt\catalog\...\xxx.pptx
```

失败时输出：

```text
PPT 生成失败：具体错误
```

## MiniMax 预检

```bash
python scripts/check_minimax_image_api.py
python scripts/check_minimax_image_api.py --live --out /tmp/check_minimax.png
```

## 离线测试

```bash
bash scripts/run_local_test.sh
python tests/test_plan.py
```

离线测试使用 `placeholder` 图片，只验证需求解析、PPT 页数、英文文本、类目和质量校验。
