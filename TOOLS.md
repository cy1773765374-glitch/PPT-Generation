# 工具说明：PPT 商品目录册 v1.2

## 主脚本

```text
scripts/generate_catalog_ppt_v12.py
```

作用：

1. 解析飞书用户输入。
2. 识别目录册主题、页数、商品类目。
3. 使用 `Asia/Shanghai` 生成任务目录名。
4. 调用 MiniMax `image-01` 生成封面图和商品图。
5. 生成带图文排版的 PPTX。
6. 输出 JSON 结果，供 OpenClaw Agent 回复飞书。

## 常用参数

```bash
python scripts/generate_catalog_ppt_v12.py \
  --prompt "生成一个厨房餐具相关的3页的商品目录册PPT" \
  --sender-name "陈玉" \
  --sender-open-id "ou_xxx" \
  --json
```

### 参数说明

| 参数 | 必填 | 说明 |
|---|---:|---|
| `--prompt` | 是 | 飞书用户原始问题 |
| `--sender-name` | 否 | 发送人姓名，用于元数据 |
| `--sender-open-id` | 否 | 发送人 open_id，用于元数据 |
| `--root` | 否 | 输出根目录，不填则读 `PPT_CATALOG_ROOT` |
| `--image-mode` | 否 | `minimax` / `placeholder` / `off`；生产使用 `minimax` |
| `--json` | 否 | 使用 JSON 格式输出，推荐开启 |

## 输出 JSON 字段

```json
{
  "ok": true,
  "prompt": "生成一个厨房餐具相关的3页的商品目录册PPT",
  "category": "厨房餐具",
  "page_count": 3,
  "run_name": "2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT",
  "run_dir": "/data/share/yaq/ppt/catalog/2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT",
  "pptx_path": "/data/share/yaq/ppt/catalog/.../...pptx",
  "windows_path": "Z:\\yaq\\ppt\\catalog\\...\\...pptx",
  "reply_text": "Z:\\yaq\\ppt\\catalog\\...\\...pptx",
  "image_mode": "minimax",
  "image_count": 5
}
```

## 飞书 formatter

```bash
python scripts/generate_catalog_ppt_v12.py --prompt "$USER_TEXT" --json \
  | python scripts/feishu_reply_formatter.py
```

formatter 成功时只输出：

```text
Z:\yaq\ppt\catalog\...\xxx.pptx
```

## 失败时输出

```json
{
  "ok": false,
  "error": "MINIMAX_API_KEY 未配置..."
}
```

Agent 收到失败结果时，不要回复“已完成”，应回复：

```text
PPT 生成失败：错误信息
```
