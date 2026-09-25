# 0.1.0-preview.8

本次预览版包含 preview.7 之后的 OFD-H 阅读器兼容性和 DOCX DualLayer 分页映射改进。

## 主要变化

- DualLayer DOCX → OFD 使用与 Native 相同的 OFD 命名空间和文档元数据；资源清单带 `ofd` 前缀，改善旧 OFD-H 阅读器中的页面显示。
- PDF 栅格页图铺白底并以 RGB PNG 写入，避免透明页图在旧阅读器中被合成为黑底。
- Native 文本在 `TextObject` 上写出字重和斜体属性；名称字体增加可见的模拟粗体和斜体。OFD → PDF 与 OFD → SVG 同步识别这些属性。
- DualLayer 改进跨页表格单元格的原文定位；正文整体没有渲染锚点时拒绝转换，个别短段落缺少锚点时保留原文、标记页码不可靠并拒绝选页转换。

## 验证范围

本地发布准备验证：127/127 Release 回归、5/5 包清单测试、11 个候选包隔离缓存消费和 6 页 macOS Preview 检查均通过；详细范围与产物见 `artifacts/release-ready-preview8/report.md`。真实 OFD-H 阅读器的兼容性仍需在目标阅读器上单独验收；macOS Preview 检查路径为 DOCX → OFD → PDF → Preview。

SDK 目标为 netstandard2.0/netstandard2.1，CLI 需要 .NET 10。复杂 Word 文档的布局和字体保真边界见 [转换契约](conversion-contracts.md)。
