# 强类型对象补齐：大纲、注释、动作、裁剪、组合对象

**What to build:** 按使用频率把大纲/书签、注释、动作、裁剪、组合对象做成可读写的强类型，其余继续走 `Raw` / `Preserved*`。往返不丢未知节点。

**Blocked by:** None (can start immediately)

**Priority:** P1

**Status:** implemented (bounded functional/package/Preview acceptance passed; no merge or release)

- [x] 大纲/书签、注释、动作、裁剪、组合对象有读写往返测试
- [x] 未知节点仍不丢
- [x] 更新功能对照与教程（大纲/注释相关课）

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P1-10

## What to build

常用读写闭环已经覆盖文字、路径、图片等，但大纲、注释、动作、裁剪和组合对象仍偏原始 XML。这张票按使用频率补强类型，让调用方能建目录树、读写注释、设动作、表达裁剪和组合，而不必拼接 XML。没建模的节点继续保留，不能为了「类型漂亮」丢掉扩展。

## Acceptance criteria

- [x] 大纲/书签可写可读，往返后层次与目标页还在
- [x] 注释、动作、裁剪、组合对象有最小可构建样本，读写往返通过
- [x] 含未知子节点的包再保存后，这些节点仍在
- [x] 更新功能对照；教程里对大纲/注释的实现对照不再写「只能 Raw」

源码候选、公开样例、功能结果与独立视觉验收状态见 [验收记录](../../../docs/evidence/issue09-typed-objects/README.md)。最小范围为 GoTo、局部 appearance、每 Clip 一个 filled Path Area 与共享无环组合；多 Area 可保留但不导出，复杂 DrawParam、group opacity、替代图形及未建模 XML 仍保留并拒绝不完整处理。本次 5 份 PDF 共 8 页已在 Preview 逐页验收；功能、包消费及 a1cebe3 的双 bot/CI 已闭合。后续文档提交的当前 head 状态见 PR #20；未合并或发布。

## Blocked by

- None (can start immediately)
