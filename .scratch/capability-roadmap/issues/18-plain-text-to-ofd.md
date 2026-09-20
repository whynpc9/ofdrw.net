# 纯文本 → OFD

**What to build:** 把纯文本文件经公开流式布局生成 OFD，可设字号与页尺寸。

**Blocked by:** 01 公开流式布局：Paragraph / Span、折行、分页

**Status:** ready-for-agent

- [ ] 文本文件可生成多页 OFD；字号与页尺寸可配
- [ ] API 与 CLI 可用
- [ ] 中英样例可抽取原文

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P1-08

## What to build

给「只有一份 .txt」的调用方一条最短生成路径：读文本、用已公开的段落引擎分页、写出 OFD。不要再手写坐标，也不要绕进 DOCX。

## Acceptance criteria

- [ ] API 与 CLI 将 UTF-8 文本转为 OFD；可设字号、页宽高、边距
- [ ] 长文本自动分页；抽出文本与源（除约定的换行规范化）一致
- [ ] 中英样例 Preview 抽查不乱码、不裁半行
- [ ] 更新功能对照与 CLI 帮助

## Blocked by

- 01 公开流式布局：Paragraph / Span、折行、分页
