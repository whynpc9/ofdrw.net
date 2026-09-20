# OFD-A 只读检查器

**What to build:** 对标档案规则引擎（GB/T 42133）做只读检查和违规报告。不宣称档案合规认证；转换管道另议，不在这张票。

**Blocked by:** None (can start immediately)

**Status:** ready-for-agent

- [ ] 对样例包输出结构化违规/通过项
- [ ] 文档写明不是档案合规认证
- [ ] 不改转换默认行为

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P2-05

## What to build

给归档前检查一条机器可读报告：字体嵌入、加密、签名范围之类能静态看到的规则先做只读扫描。目标是「这份包违反了哪些可检查项」，不是发证。不必在这张票里改 DOCX/PDF 转换去「自动变合规」。

## Acceptance criteria

- [ ] API（可选 CLI）对故意违规的样本和较干净的样本分别给出报告，条目可定位到规则/路径
- [ ] 报告只读，不改包
- [ ] 文档与功能对照写明：检查器不是 GB/T 42133 认证，也不是长期保存承诺
- [ ] 默认转换管道行为不变

## Blocked by

- None (can start immediately)
