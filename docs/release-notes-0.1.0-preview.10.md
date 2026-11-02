# 0.1.0-preview.10 — candidate / 未发布

这是候选发布说明，**尚未发布，也未据此宣称全部验收通过**。候选版本 `0.1.0-preview.10`；产品代码含 PR91 合并 `5d83fb38` 的 N03/N12 修复，文档从 `d447a0d5` 继续冻结。最终源码 SHA、包清单及页面收据由冻结后的验收记录绑定。当前公开安装版本仍为 `0.1.0-preview.9`。

## 相比 preview.9

- 扩展已有 Flow/图片/Canvas、复杂表格分页、文档操作、类型化对象和有界字体出口；固定字体的测量、HarfBuzz 塑形及 run/span/显式字号/定位查询有增量实现。通用段落 bidi、自动断行/重塑形、多字体分页和 writer 全链路仍未完成，不能称完整 R22。
- 默认发布集合新增 `Ofdrw.Net.Text.HarfBuzz`，由原 11 包变为 12 包，实际清单以 `verify-package-artifacts.py:PACKAGES` 为准。PDF→OFD 对未验证 RTL 语义明确拒绝，仍可显式使用纯页面图路径；有界 R02、归档与各字体出口维持各自契约。
- **N03 正字号输出契约**：创建新文本时，如有限正字号按既有 `0.###` 序列化会变为零，输出前明确拒绝，避免无效 `Size="0"`；普通字号及坐标精度不变。typed writer 返回 `TextLayoutException/UnsupportedOutput`，primitive package writer 抛 `NotSupportedException`。外来 `SourceXml` 合法原始精度继续保留；不是全局增加精度或固定最小字号。详见 [N03](validation/cn-en-release-n03.md)。
- **N12 发布绑定**：源码 fingerprint 对合法 LFS pointer/已下载载荷统一绑定相同 oid/size，并严格拒绝错哈希/大小/非法 index；普通跟踪文件继续按实际字节绑定。保留全部发布门、测试和 tag 工作流，不以旧 preview.9 验收凭证放行本版。
- 提供可复现的中英文真实 Regular/Bold [API 字体部署教程](tutorials/cn-en-font-deployment.md)及页眉页脚/内嵌图/选页的有限验收资料。新候选只按下面的具名配置评估。

## 受限字体配置与兼容性

候选 profile **`cn-en-physical-cjk-400-700-v1`**：显式目录提供固定 `Noto Sans CJK SC` 静态 Regular 400 / Bold 700，以及 `Noto Sans` Regular；Native API 设置严格 `NativeEmbeddedBoldBehavior=Throw`、显式 fallback 和 `UseInstalledMicrosoftOfficeFonts=false`。源字体哈希、OFL 来源/许可证及两模式完整示例见[当前计划](next-steps-2026-11-02.md)与教程。源字体不随 NuGet 分发；fontTools 仅用于离线资料准备，不是应用运行时依赖。

不承诺 Noto Sans 真实 Bold、物理 Italic、任意字体面、任意系统字体、复杂 Word 版式或跨机像素一致。CLI 仍使用默认 TextStroke 策略，没有严格粗体开关；严格 profile 使用 API。OFD 原文和 PDF 重复文本检查、native/default 逐页实际 PNG/Preview 仍是候选必验项。

## 依赖、资产与运行范围

- 普通 SDK 为 `netstandard2.0/netstandard2.1`；CLI/E2E 要求 .NET 10。目标框架声明不证明所有宿主、RID 或发布方式均已运行。
- HarfBuzz 适配包固定 `HarfBuzzSharp` 及 macOS/Win32/Linux native-assets **14.2.1.301**；可选塑形调用需要相符的原生库。默认字体查询/有界 writer 能力不等于自动通用段落排版。实际加载的 native 模块、OS/RID/TFM 由候选运行收据另核对。
- 文档/PDF 链路沿用 OpenXML 3.5.1、PdfSharpCore/MigraDocCore.Rendering 1.3.67、PdfPig 0.1.15、Docnet.Core 2.6.0、SixLabors.Fonts 1.0.1 与 ImageSharp 2.1.13；完整直接/传递依赖及许可由还原后 Q01 对照 40 项受审清单核验。历史还原报告 ImageSharp NU1902/NU1903，未由本轮修复消除。
- 现有 Linux/Windows/macOS 工作流和既有 RID/原生资产门继续执行；不新增 musl、AOT、trimming、single-file 或任意架构支持承诺。Word/LibreOffice 直接 PDF/双层路径有独立外部依赖，不能冒充 native 链路。
- Skia、PDF vector、SES interop、password crypto 是独立可选扩展，不在默认 12 包发布循环。仓库 MIT、依赖与外部工具/字体许可见 [THIRD-PARTY-NOTICES](../THIRD-PARTY-NOTICES.md)。

## 预览版已知限制

- **Regular-only 合成粗体**：原未修改样例在默认 TextStroke 下的表头“量”字仍有字腔堵塞，历史视觉 FAIL 保留；真实 Bold 配置不等于该缺陷已修复。本版承诺限定到上述物理配置。
- **macOS Preview Find**：macOS 26.6.2 (25G83) 上，原生成 PDF 的完整页末日期/部分整词可能同时显示 Found on 1 page 与 Not found，无法导航/高亮；根因未定。本版不承诺 Preview 通用搜索兼容，也不把 PDFKit/AX/文本提取成功当作 UI 通过。
- **原始 Unicode Copy 未验证**：原始剪贴板字节验收未完成；粘贴到搜索框不构成复制保真证据。其他阅读器的 Copy/Find 同样未验证，不能宣称通过。
- R02 完整外来映射、R20 完整归档符合性、通用 R22、RTL/Arabic/Indic 扩展与高级字体/新部署组合保持既有有限/未完成边界。普通测试绿不解除这些限制。

用户于 2026-11-02 明确接受以上具名范围/限制用于新的预览候选；旧验收失败不改为 PASS。profile 内新的视觉/原文/功能失败与强制工程门失败仍阻断。最终候选测试、精确包、PNG/Preview 与 Copy/Find 的实际结果应读取新候选收据；[此前独立验证](validation/cn-en-release-local-validation.md)仅为旧源码/产物证据。

## 发布状态

2026-11-02 01:13:57 UTC 核对官方 NuGet 的全部 12 个默认 ID 与远端 tag：11 个既有 ID 最新 preview.9，无 preview.10；HarfBuzz ID 404，没有 preview.10 tag。404 不证明新 ID 所有权或推送权限。版本未预留，实际发布前重查。尚无本版 tag、公开包、签名或公开源安装回读结论；README 安装命令保持 preview.9。
