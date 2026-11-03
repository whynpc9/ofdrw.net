---
title: "文字定位与受限替换"
---

现有文字定位依据原嵌入字体、实际字形和基线证据，面向正文字面匹配。替换适用于受支持的 Regular 原字体、简单 Unicode scalar、水平正向及仅平移的文字区域。

建议流程：读取文档 → 查询匹配及可替换标志 → 检查候选区域 → 生成新文件 → 重新读取并检查实际页面。不要直接覆盖唯一输入副本。

模板、批注、复杂 shaping、超出原区域或无法证明原字体几何的替换会拒绝。没有匹配或文本相同的 Noop 与真正写入是不同结果。

完整回归、文件原子提交与失败保护的范围见[定位替换验收](https://github.com/whynpc9/ofdrw.net/blob/main/docs/evidence/issue06-owned-save/README.md)。阅读器 Copy/Find 行为仍须在实际 UI 中独立验证。
