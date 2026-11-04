---
title: "文字定位与受限替换"
description: "在原字体几何可证明的范围内定位和替换 OFD 正文，使用可信字体身份并验证新文件。"
sidebar:
  order: 120
  label: "文字定位与替换"
---

文字定位依据原嵌入字体、实际字形和基线证据，面向正文字面匹配。替换只适用于受支持的 Regular 原字体、简单 Unicode scalar、水平正向和仅平移的文字区域。

## 推荐流程

1. 读取文档并查询字面匹配及 `CanReplace` 标志。
2. 检查候选区域、原字体身份和替换文字是否适合原空间。
3. 生成一个**新文件**，不要直接覆盖唯一输入副本。
4. 重新读取输出，核对原文、替换结果和实际页面。

没有匹配、文本相同产生的 Noop 与真正写入是不同结果。模板、批注、复杂 shaping、超出原区域或无法证明原字体几何的替换会拒绝，不会用近似字体或缩放文字冒充成功。

## CLI 替换

`replace-text` 要求至少一个显式审核过的原嵌入字体身份，格式为 `SHA256:BYTECOUNT`：

```sh
ofdrw replace-text input.ofd output.ofd \
  --find "旧文本" --replace "新文本" \
  --trusted-font SHA256:BYTECOUNT --pages 1,3-5
```

- SHA-256 为 64 位十六进制，BYTECOUNT 为原字体字节数；多个身份可重复传 `--trusted-font`。
- `--replace ""` 表示删除匹配文字。
- `--pages` 是一基页号或范围，页面位置不能重复。
- 输入和输出必须是不同路径；不支持的几何或原区域空间不足会拒绝整个批次。

CLI 不会根据本机同名字体猜测信任身份。调用方应先从来源包和业务字体清单确认实际原字体，再提供哈希与字节数。

### CLI 预览（main，尚未发布）

替换前可用只读的 `locate` 查看会匹配哪些位置，参数与上面的 `--find`、`--trusted-font`、`--pages` 含义相同，且从不写文件：

```sh
ofdrw locate input.ofd --find "旧文本" --trusted-font SHA256:BYTECOUNT --pages 1,3-5 --json
```

每个命中给出一基页号、匹配文本、毫米字形边界和 `canReplace`；SDK 无法度量的来源（例如未信任或未嵌入字体）直接报错，而不是返回不完整结果。退出码见 [CLI 参考](https://whynpc9.github.io/ofdrw.net/reference/cli/)。

## 验证与限制

替换后分别检查：

- 新文件能否读取，替换范围和页码是否正确。
- 原始未替换文字是否保持不变。
- 文字是否裁切、重叠、错位或产生重复绘制。
- 目标阅读器的显示、Copy/Find 行为；这些 UI 行为需要独立验证。

完整回归、原子文件提交和失败保护的范围见[定位替换验收](https://github.com/whynpc9/ofdrw.net/blob/main/docs/evidence/issue06-owned-save/README.md)。任意复杂文字重排、自动换字体和跨区域重排不在该入口的承诺内。
