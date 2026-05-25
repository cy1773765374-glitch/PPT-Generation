# Tools for workspace-PPT-Generation

## 核心命令

```bash
python scripts/run_ppt_generation.py --user cy --text "根据描述生成商品目录册PPT：..."
```

可带多张图片：

```bash
python scripts/run_ppt_generation.py   --user cy   --text "根据图片做商品目录册"   --image /path/to/1.jpg   --image /path/to/2.png
```

## 输出

脚本标准输出中包含：

```text
RUN_DIR=...
PPTX_PATH=...
REPLY_TEXT=...
```

飞书 Agent 只需要把 `REPLY_TEXT` 发回用户。

## 环境变量

| 变量 | 默认值 | 说明 |
|---|---|---|
| PPT_OUTPUT_ROOT | /data/share/yaq/ppt | 最终输出根目录 |
| PPT_GENERATION_MODE | auto | auto/catalog/general |
| PPT_MASTER_DIR | vendor/ppt-master | 原版 PPT-master 目录 |
| PPT_ALLOW_FALLBACK | true | 原版 PPT-master 不存在时是否启用 lite fallback |
| TZ | Asia/Shanghai | 运行目录时间命名 |

## 可选安装上游 PPT-master

```bash
bash scripts/install_upstream_ppt_master.sh
```
