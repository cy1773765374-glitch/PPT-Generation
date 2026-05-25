# workspace-PPT-Generation v1.0

这是一个给 OpenClaw 使用的 PPT 生成工作区。

## 定位

- 保留 PPT-master 原始通用能力，不把它改成只能做商品目录册。
- 在 OpenClaw 外层新增任务识别与 Profile 机制。
- 当用户在飞书里提出“商品目录册 / 产品册 / 商品册 / 产品目录 / 根据实物或图片做 PPT”等需求时，加载 `profiles/catalog/` 下的目录册专用提示。
- 商品目录册 PPT 的内部内容不写死，由 MiniMax 根据用户输入动态规划。
- 第一版只处理飞书文字和图片输入。
- 第一版不主动把 PPTX 上传回飞书，只回复服务器保存路径。

## 输出目录

固定输出根目录：

```bash
/data/share/yaq/ppt
```

每次任务会创建一个独立运行目录：

```bash
/data/share/yaq/ppt/catalog/2026-05-25-153012-feishu-user-商品目录册/
/data/share/yaq/ppt/general/2026-05-25-153012-feishu-user-通用PPT/
```

目录中会保存：

```text
input/message.txt              # 原始飞书文本
input/images/                  # 原始图片
product_input.json             # 结构化输入
catalog_outline.md             # 页面规划与内容大纲
prompt_bundle.md               # 给模型/PPT-master 的提示集合
project/                       # PPT-master 或 fallback 项目目录
project/exports/*.pptx         # 最终 PPTX
README.md                      # 本次任务说明
```

## 推荐安装

```bash
cd /home/cy/.openclaw
tar -xzvf workspace-PPT-Generation-v1.0.tar.gz
cd workspace-PPT-Generation

bash install.sh
```

安装脚本会创建 `/data/share/yaq/ppt`，并建立 Python 虚拟环境。

## 可选：接入原版 PPT-master

本压缩包默认不内置完整上游 PPT-master 仓库，避免包体过大。工作区保留 `vendor/ppt-master/` 位置和 adapter。

如果要接入上游 PPT-master：

```bash
cd /home/cy/.openclaw/workspace-PPT-Generation
bash scripts/install_upstream_ppt_master.sh
```

接入后，Agent 可以继续读取并遵守：

```text
vendor/ppt-master/skills/ppt-master/SKILL.md
```

## 本地测试

```bash
cd /home/cy/.openclaw/workspace-PPT-Generation
source .venv/bin/activate
python scripts/run_ppt_generation.py   --user cy   --text "根据这个产品描述生成一个商品目录册PPT：冰蓝水花包装盒，适合夏季饮品包装，清爽、年轻、可定制印刷。"
```

成功后会输出：

```text
PPTX_PATH=/data/share/yaq/ppt/.../project/exports/xxx.pptx
REPLY_TEXT=已完成，保存路径：...
```

## 第一版边界

第一版不做：

- 飞书文件上传回传。
- 飞书卡片交互。
- 多轮确认。
- Excel、PDF、Word、网页输入。
- 自动联网调研。
- 视频、旁白、动画。

这些能力后续可以作为 profile 或 mode 扩展。
