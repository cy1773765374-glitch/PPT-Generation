# 工具说明：PPT 商品目录册 v1.1

## 主脚本

```text
scripts/generate_catalog_ppt_v11.py
```

作用：

1. 解析飞书用户输入。
2. 识别目录册主题、页数、商品类目。
3. 生成任务目录。
4. 生成 PPTX。
5. 输出 JSON 结果，供 OpenClaw Agent 回复飞书。

## 常用参数

```bash
python scripts/generate_catalog_ppt_v11.py \
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
| `--json` | 否 | 使用 JSON 格式输出，推荐开启 |

## 输出 JSON 字段

```json
{
  "ok": true,
  "prompt": "生成一个厨房餐具相关的3页的商品目录册PPT",
  "category": "厨房餐具",
  "page_count": 3,
  "run_name": "2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT",
  "run_dir": "/data/share/yaq/ppt/catalog/2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT",
  "pptx_path": "/data/share/yaq/ppt/catalog/.../...pptx",
  "windows_path": "Z:\\yaq\\ppt\\catalog\\...\\...pptx",
  "elapsed_seconds": 13.1
}
```

## 失败时输出

```json
{
  "ok": false,
  "error": "错误信息"
}
```

Agent 收到失败结果时，不要回复“已完成”，应回复：

```text
PPT 生成失败：错误信息
```
