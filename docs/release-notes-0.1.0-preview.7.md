# 0.1.0-preview.7

本次预览版包含自 preview.6 以来的 CJK 样式修复与字体加载改进。

## 主要变化

- 修复 OFD → PDF 模拟粗体的错位重影：补偿 SixLabors.Fonts 1.0.1 的 em 框居中偏移，并采用 PDFsharp 的顶部对齐基线；补绘仍只写路径，不重复语义文本。
- 对未嵌入字体的 PDF 导出后备路径，也检查实际字体文件的字重/斜体；修复宿主 resolver 返回常规替代字体但不提供模拟标志时的样式丢失。
- Native DOCX → OFD 保留 SimSun/宋体的字体名称和 Bold 标志，不再用无 Bold 的 SimHei 替换。PDF 导出按实际字体文件决定是否模拟粗体/斜体。
- 保留前序合入的 CJK 样式资源区分、按内容暂存配置字体、viewer-local CJK 的 PDF 替代字体加载，以及 SixLabors.Fonts 1.0.1 接口兼容修复。
- 名称字体的宿主探测仅用于请求粗体/斜体的面，并保护 TTC、无效字节、宿主异常及可选注册预算不足的回退；内嵌字体仍严格验证。
- 仓库打包默认版本、安装示例和包消费脚本统一为 preview.7。

## 验证与能力边界

发布前验证通过：113/113 Release 回归、5/5 包清单测试、11 包隔离缓存消费以及 6 页 macOS Preview 检查。验证范围包括 Release 全量回归、像素级粗体对齐及字体替代测试、包清单测试、11 包隔离缓存消费，以及 Native/default/显式宋体的 DOCX → OFD → PDF → macOS Preview 逐页检查。验证日志、页面图和范围限制记录在 `artifacts/release-ready-preview7/report.md`；这些本机产物不随 NuGet 包发布。

SDK 目标为 netstandard2.0/netstandard2.1，CLI 需要 .NET 10。viewer-local CJK 仍依赖目标机器可用字体；PDF 自动加载器目前使用固定目录与 TTF 文件名，不能保证任意 OTF/TTC 或 DOCX FontDirectories 的替代字体可用。字体替代保留样式语义，但不保证字体造型与原 Word 相同。直接打开 Native OFD 的宋体加粗表现尚未在真实 OFD 阅读器验收；忽略 Bold 标志的阅读器可能仍显示常规宋体。复杂文档、图片和页眉页脚需单独业务样例验收，详见 [转换契约](conversion-contracts.md)。
