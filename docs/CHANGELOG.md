# CHANGELOG

## v1.2

- 飞书回复改为只返回 Windows/Samba 本地路径，例如：
  `Z:\yaq\ppt\catalog\2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT\2026-05-25-16时35分-生成一个厨房餐具相关的3页的商品目录册PPT.pptx`
- 不再回复服务器路径、页数、耗时，也不主动添加“已完成”。
- 任务目录与 PPT 文件名时间改为 `YYYY-MM-DD-HH时MM分-用户询问的问题`。
- 固定使用 `PPT_TIMEZONE=Asia/Shanghai`，避免服务器 UTC 时间导致文件名显示为 `09:35`。
- 生产默认调用 MiniMax `image-01` 生图，并将封面图、商品图插入 PPT。
- 默认 `PPT_REQUIRE_IMAGES=1`，如果 MiniMax 密钥缺失或生图失败，脚本直接失败，避免继续生成纯文字 PPT。
- 新增 `PPT_IMAGE_MODE=placeholder` 离线测试模式，仅用于本地连通性测试，不用于飞书生产。
- 继续保持 PPT 文件直接放在任务目录一级，不使用 `project/exports/`。

## v1.1

- 新增用户页数解析：支持 `3页`、`三页`、`五页` 等表达。
- 修改输出目录：PPT 文件直接放在任务目录一级。
- 修改任务命名：`YYYY-MM-DD-HHMM-用户询问的问题`。
- PPT 文件名与任务目录名保持一致。
- 新增 Windows 共享路径回显：例如 `Z:\yaq\ppt\catalog\...`。
- 移除 v1.0 中 `project/exports/catalog_lite_v1.pptx` 的输出结构。
- fallback 内容升级：按商品类目生成商品系列、卖点、规格。
