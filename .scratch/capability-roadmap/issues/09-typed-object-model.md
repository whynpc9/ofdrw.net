# 强类型对象补齐：大纲、注释、动作、裁剪、组合对象

**What to build:** 按使用频率把大纲/书签、注释、动作、裁剪、组合对象做成可读写的强类型，其余继续走 `Raw` / `Preserved*`。往返不丢未知节点。

**Blocked by:** None (can start immediately)

**Status:** ready-for-agent

- [ ] 大纲/书签、注释、动作、裁剪、组合对象有读写往返测试
- [ ] 未知节点仍不丢
- [ ] 更新功能对照与教程（大纲/注释相关课）

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P1-10

## What to build

常用读写闭环已经覆盖文字、路径、图片等，但大纲、注释、动作、裁剪和组合对象仍偏原始 XML。这张票按使用频率补强类型，让调用方能建目录树、读写注释、设动作、表达裁剪和组合，而不必拼接 XML。没建模的节点继续保留，不能为了「类型漂亮」丢掉扩展。

## Acceptance criteria

- [ ] 大纲/书签可写可读，往返后层次与目标页还在
- [ ] 注释、动作、裁剪、组合对象有最小可构建样本，读写往返通过
- [ ] 含未知子节点的包再保存后，这些节点仍在
- [ ] 更新功能对照；教程里对大纲/注释的实现对照不再写「只能 Raw」

## Blocked by

- None (can start immediately)
