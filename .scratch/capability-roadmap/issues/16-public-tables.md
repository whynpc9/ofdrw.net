# 公开表格与单元格

**What to build:** 在流式布局上公开表格/单元格，生成带边框、底色、对齐的表。跨页策略有文档（允许有限：例如整行放到下一页）。

**Blocked by:** 01 公开流式布局：Paragraph / Span、折行、分页

**Priority:** P0

**Status:** blocked

- [ ] 可生成带边框/底色/对齐的表
- [ ] 跨页策略有文档，并有对应样例
- [ ] Preview 视觉验收通过（Native OFD → PDF → Preview）

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P0-02

## What to build

段落能流式分页之后，调用方还要能排表，而不是手画路径当格子。交付公开的表/单元格元素：边框、底色、水平垂直对齐，并写入 OFD 路径与文字。跨页不必一次做完所有 Word 拆行策略，但必须写明实际策略（例如行不可拆时整行下页），并用样例证明没有重叠或丢格。

## Acceptance criteria

- [ ] 公开 API 能生成至少含合并单元格或多种对齐之一的表，带边框与底色
- [ ] 跨页样例的行为与文档一致；无重叠、裁切半格、重影
- [ ] Preview 验收走本次生成的 Native OFD；局部单元格样式不扩散到整表
- [ ] DOCX Native 表格路径仍通过既有测试（若已改用同一引擎）
- [ ] 更新功能对照与教程

## Blocked by

- 01 公开流式布局：Paragraph / Span、折行、分页
