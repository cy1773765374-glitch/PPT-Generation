# 飞书接入说明 v1.1

## 目标

飞书用户输入：

```text
生成一个厨房餐具相关的3页的商品目录册PPT
```

Agent 应调用：

```bash
python scripts/generate_catalog_ppt_v11.py \
  --prompt "生成一个厨房餐具相关的3页的商品目录册PPT" \
  --sender-name "陈玉" \
  --sender-open-id "ou_6e056040c28331827575c0061644569c" \
  --json
```

然后从 JSON 中读取结果，并回复飞书。

---

## Agent 回复逻辑

### 成功

```text
已完成，PPT 文件：
Z:\yaq\ppt\catalog\2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT\2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT.pptx

服务器路径：
/data/share/yaq/ppt/catalog/2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT/2026-05-25-1632-生成一个厨房餐具相关的3页的商品目录册PPT.pptx

已完成 · 3页 · 耗时 13.1s
```

### 失败

```text
PPT 生成失败：具体错误信息
```

---

## 关键注意事项

1. 不要再回复 `project/exports/catalog_lite_v1.pptx`。
2. 不要把用户原始问题只塞进“商品描述”。
3. 用户写了几页，最终 PPT 就必须生成几页。
4. Windows 路径给业务人员看，Linux 路径用于服务器排查。
5. 如果后续接入飞书云盘上传，则在此基础上增加 upload_file 步骤，不影响本地输出结构。
