# SKILL Extension: OpenClaw PPT Generation

本扩展不替代 PPT-master 原始 SKILL。

当运行在 OpenClaw + 飞书场景时，额外遵守：

1. 用户输入来自飞书。
2. 结果保存到 `/data/share/yaq/ppt`。
3. 飞书只回复保存路径。
4. 商品目录册任务加载 `profiles/catalog`。
5. 通用 PPT 任务加载 `profiles/general`，同时遵守 PPT-master 原始工作流。
6. 商品目录册内部内容动态生成，不固定页数和章节。
