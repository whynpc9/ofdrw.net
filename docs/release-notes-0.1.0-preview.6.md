# 0.1.0-preview.6

本次预览版包含自 0.1.0-preview.5 以来的文档转换、资源隔离、包消费验证及教程更新。

## 主要变化

- 改进 Native DOCX → OFD 的结构化排版、表格、字体、局部样式与页眉页脚；默认输出保留 OpenXML 原始文本。
- 增加转换结果诊断和源页映射，改进 DualLayer 原文定位及统一页选择行为。
- 修复页面删除、合并资源映射、字体隔离及 SVG 几何/资源处理。
- 加强输入与渲染预算、外部进程取消/超时处理及 CLI 原子输出。
- 发布前从隔离缓存消费同一批 11 个 NuGet 包，校验包依赖、版本和 SHA256 清单；稳定跨平台字体与 CI。
- 增加 OFD 格式渐进教程及 OFD-H 征求意见稿差异说明。

## 验证与能力边界

发布准备阶段 Release 回归 84/84、包清单测试 5/5、11 个候选包消费 E2E 通过。确定性 generated-layout.docx 的显式 Native 和默认模式各两页，经 OFD → PDF → macOS Preview 检查通过。

SDK 目标为 netstandard2.0/netstandard2.1，CLI 需要 .NET 10。复杂 Word 文档仍需业务样例验证；附属文字不保证复现 Word 页底脚注布局，无法可靠定位的原文不提供虚假的准确页级映射。详细行为见 [转换契约](conversion-contracts.md)。
