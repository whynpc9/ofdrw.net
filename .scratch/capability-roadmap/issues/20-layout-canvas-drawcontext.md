# 布局层 Canvas DrawContext

**What to build:** 在公开流式布局同一页上，用画布画页眉线、票面框等，并与段落/表格混排。

**Blocked by:** 01 公开流式布局：Paragraph / Span、折行、分页；04 类 Graphics2D 绘图 API（issue #3）

**Status:** ready-for-agent

- [ ] 布局层画布与流式块可同页混用
- [ ] 页眉/票面框线样例可 Preview 验收
- [ ] 教程有混用对照

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P1-03

## What to build

只排段落不够画票据。调用方需要在同一页既流式放字，又用 04 的绘图 API 画固定框线。`DrawContext` 是布局层的画布，不是第二套底层对象模型。

## Acceptance criteria

- [ ] 同一页可同时包含流式段落（或表格，若 16 已完成则可用）与画布绘制的线框
- [ ] 样例含页眉线和票面框；Preview 中框线与文字对齐，无重叠错位
- [ ] 画布绘制结果仍是 OFD 路径/文字，可被现有 SVG/PDF 导出看见
- [ ] 更新教程实现对照

## Blocked by

- 01 公开流式布局：Paragraph / Span、折行、分页
- 04 类 Graphics2D 绘图 API（issue #3）
