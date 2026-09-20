# 附件增删

**What to build:** 对已经存在的 OFD 增加或删除附件（`Attachments.xml` 及文件载荷），行为与合并选项 `IncludeAttachments` 一致。

**Blocked by:** None (can start immediately)

**Status:** ready-for-agent

- [ ] 可向已有包添加附件并读回
- [ ] 可删除指定附件及无引用载荷
- [ ] 与合并是否带附件的约定一致

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P1-07

## What to build

生成时已经能 `AddAttachment`，缺的是打开一份现成包再增删附件。调用方应能挂上合同附件或拿掉过期文件，保存后阅读器看得到或看不到对应附件，且不破坏页面内容。合并时 `IncludeAttachments` 的取舍要和这套增删语义一致，避免「合并丢掉、编辑又加回来」各说各话。

## Acceptance criteria

- [ ] API 对已有包添加附件（名称、文件名、载荷）；保存后再读能列出并取出字节
- [ ] API 删除指定附件后，清单与载荷都去掉；仍被其他结构引用的共享对象不误删
- [ ] 合并 `IncludeAttachments` 为 true/false 时，与编辑后的附件集合行为在文档中对齐
- [ ] 有往返测试；更新功能对照

## Blocked by

- None (can start immediately)
