# workspace-PPT-Generation v1.1

## 版本定位

v1.1 是基于 v1.0 的“飞书触发生成商品目录册 PPT”增强版，重点解决：

1. 飞书中用户说“3页”，脚本必须真正生成 3 页，而不是固定 fallback 页数。
2. 输出 PPT 不再放到 `project/exports/` 二级目录，而是直接放到任务目录一级。
3. 任务目录和 PPT 文件名改为：`YYYY-MM-DD-HHMM-用户询问的问题`。
4. 飞书回复内容返回最终 PPT 文件路径，而不是项目中间目录。
5. 生成内容从“待确认占位模板”升级为“按类目动态生成商品目录册初稿”。

---

## 推荐目录

解压到 OpenClaw Agent 工作区，例如：

```bash
mkdir -p ~/.openclaw/workspace-PPT-Generation

tar -xzvf workspace-PPT-Generation-v1.1.tar.gz \
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

根据实际路径修改 `.env`：

```bash
PPT_CATALOG_ROOT=/data/share/yaq/ppt/catalog
PPT_WINDOWS_SHARE_PREFIX=Z:\\yaq\\ppt\\catalog
PPT_DEFAULT_PAGE_COUNT=5
PPT_MAX_PAGE_COUNT=20
```

---

## 本地测试

```bash
cd ~/.openclaw/workspace-PPT-Generation
source .venv/bin/activate

python scripts/generate_catalog_ppt_v11.py \
  --prompt "生成一个厨房餐具相关的3页的商品目录册PPT" \
  --sender-name "陈玉" \
  --sender-open-id "ou_6e056040c28331827575c0061644569c" \
  --json
```

预期输出类似：

```json
{
  "ok": true,
  "run_dir": "/data/share/yaq/ppt/catalog/2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT",
  "pptx_path": "/data/share/yaq/ppt/catalog/2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT/2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT.pptx",
  "windows_path": "Z:\\yaq\\ppt\\catalog\\2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT\\2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT.pptx",
  "page_count": 3
}
```

---

## 飞书回复建议格式

Agent 在飞书里回复时建议只回业务人员能理解的内容：

```text
已完成，PPT 文件：
Z:\yaq\ppt\catalog\2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT\2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT.pptx

服务器路径：
/data/share/yaq/ppt/catalog/2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT/2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT.pptx

已完成 · 3页 · 耗时 13.1s
```

---

## v1.1 与 v1.0 的关键差异

| 项目 | v1.0 | v1.1 |
|---|---|---|
| 页数 | 固定 fallback，多数情况不解析用户页数 | 解析 `3页/三页/5页` 等表达 |
| 输出目录 | `任务目录/project/exports/catalog_lite_v1.pptx` | `任务目录/任务名.pptx` |
| 任务命名 | 时间 + open_id + 商品目录册 | 日期-时分-用户询问的问题 |
| 内容 | 待命名、待识别、待确认 | 按类目生成商品系列、卖点、规格 |
| 飞书返回 | 服务器路径 | Windows 共享路径 + 服务器路径 |
| 业务可用性 | 链路验证 | 第一版可交付目录册草稿 |

