# vendor/ppt-master

这里是用户服务器上的 PPT-master 完整仓库保留位置。

目标路径：

```text
/home/cy/.openclaw/workspace-PPT-Generation/vendor/ppt-master/
```

本压缩包不会覆盖该目录，`install_update.sh` 也显式排除了 `vendor/ppt-master/`。

生产环境必须确保该目录中至少存在：

```text
scripts/svg_quality_checker.py
scripts/finalize_svg.py
scripts/svg_to_pptx.py
```

主程序默认通过以下环境变量指向这里：

```bash
PPT_MASTER_ROOT=/home/cy/.openclaw/workspace-PPT-Generation/vendor/ppt-master
PPT_RENDER_BACKEND=pptmaster
PPT_MASTER_STRICT=1
```
