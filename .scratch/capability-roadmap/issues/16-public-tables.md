# 公开表格与单元格

**What to build:** 在流式布局上公开表格/单元格，生成带边框、底色、对齐的表。跨页策略有文档（允许有限：例如整行放到下一页）。

**Blocked by:** 01 公开流式布局：Paragraph / Span、折行、分页

**Priority:** P0

**Status:** implemented (local acceptance passed; stacked PR #10 remains open and unmerged)

- [x] 可生成带边框/底色/对齐的表
- [x] 跨页策略有文档，并有对应样例
- [x] Preview 视觉验收通过（Native OFD → PDF → Preview）

## Execution

- 起点：Issue01 PR #9 `5d71dc9e4fb3ccfa4dd028386436e58cd261bc39`；按用户授权以未合并依赖建立 stacked PR，base=`codex/public-flow-layout`，分支 `codex/public-tables`。
- 交付：[PR #10](https://github.com/whynpc9/ofdrw.net/pull/10)，base=`codex/public-flow-layout`，源码复验基线 `26f33e91d716c863bed09e5feb3c95a0a423b00f`。最新 head 的 CI 与复审状态以 PR 读回为准；首轮三个意见及复审非正跨度定位均已修复，继续最新head复审。
- 原始可下载产物：[PDF/PNG、OFD档案与完整性校验](../../../docs/validation/public-tables-evidence/README.md)，干净checkout可复核。
- 公开契约与限制：[public-tables.md](../../../docs/public-tables.md)。
- 验收和 review 状态：[2026-10-02 记录](../../../docs/validation/public-tables-2026-10-02.md)。300/300独立回归、11包隔离消费与31f0b75本次生成11页Preview复验通过（后续仅无页面影响的失败诊断变更）；最终review闭合证据保留在PR。

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P0-02

## What to build

段落能流式分页之后，调用方还要能排表，而不是手画路径当格子。交付公开的表/单元格元素：边框、底色、水平垂直对齐，并写入 OFD 路径与文字。跨页不必一次做完所有 Word 拆行策略，但必须写明实际策略（例如行不可拆时整行下页），并用样例证明没有重叠或丢格。

## Acceptance criteria

- [x] 公开 API 能生成至少含合并单元格或多种对齐之一的表，带边框与底色
- [x] 跨页样例的行为与文档一致；无重叠、裁切半格、重影
- [x] Preview 验收走本次生成的 Native OFD；局部单元格样式不扩散到整表
- [x] DOCX Native 表格路径仍通过既有测试（若已改用同一引擎）
- [x] 更新功能对照与教程

## Blocked by

- 01 公开流式布局：Paragraph / Span、折行、分页
