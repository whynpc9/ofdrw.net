# 0.1.0-preview.9

发布版本：`0.1.0-preview.9`。功能源码基线：`e9b237bc85e4dee06c1f690bea48553a1f5a22e2`。

## 主要变化

- 增加流式排版、公共表格、图片导入导出，以及文档合并、拆分、附件与签名清理等工具。
- 增加纯文本 OFD、字体子集、类型化对象模型、严格文字定位及受保护的文字替换与保存；保留未知 XML 和附件描述信息，无法证明安全的修改会拒绝执行。
- 改进 DOCX native/default 转换的跨平台字体身份、英文比例度量及混合粗体表现，补充 Windows、Linux 和 macOS 的包消费及视觉验证路径。
- 加固资源预算、路径安全、转换子进程所有权及退出清理。
- 仓库另有 PDF 矢量转换、SkiaSharp 绘制适配、口令密码与 SES 互操作扩展；这些扩展有独立包消费门禁，不属于默认 11 包发布集合。

## 验证与边界

本地 Release 回归 3075/3075、真实字体专项 68/68、Python 检查 34/34 及默认 11 包隔离消费通过。确定性 DOCX 的 native/default 各两页经 OFD → PDF → PNG 和 macOS Preview 检查；本次用户明确授权以 Codex 的实际检查替代人工签收，检查人如实记录为 Codex。仅对这些样例页面作出视觉结论。依赖还原仍报告 ImageSharp 2.1.13 的 NU1902/NU1903 警告，本轮未升级该依赖。详细产物与记录在 `artifacts/release-ready-preview9/`，发布 CI 留存同批包和验证页。

默认发布工作流仅发布 `scripts/verify-package-artifacts.py` 定义的 11 个产品。SDK 目标为 netstandard2.0/netstandard2.1，CLI 需要 .NET 10。复杂 Word 文档和目标 OFD 阅读器的保真边界见 [转换契约](conversion-contracts.md)。
