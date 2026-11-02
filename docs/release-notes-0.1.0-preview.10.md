# 0.1.0-preview.10 — 已发布预览版

**2026-11-02 发布状态：**`0.1.0-preview.10` 已在 [GitHub Release](https://github.com/whynpc9/ofdrw.net/releases/tag/v0.1.0-preview.10) 公开（非草稿、prerelease），12 个主 NuGet 包已公开回读。产品标签/源码为 `8f27f92b51b70a9b01082da35ff1cfe989d46123`，包含 PR91 的 N03/N12 修复；发布控制器为 `3568c00bb52408f1151095823a1337cb97cb25dc`。两者职责不同，包源码身份为前者。成功工作流、公开包/签名/独立消费及页面验收见[最终发布记录](validation/preview10-publication-2026-11-02.md)。本版通过结论只覆盖以下具名配置及实际检查页面，历史 FAIL 与未完成项保留。

## 相比 preview.9

- 扩展已有 Flow/图片/Canvas、复杂表格分页、文档操作、类型化对象和有界字体出口；固定字体的测量、HarfBuzz 塑形及 run/span/显式字号/定位查询有增量实现。通用段落 bidi、自动断行/重塑形、多字体分页和 writer 全链路仍未完成，不能称完整 R22。
- 默认发布集合新增 `Ofdrw.Net.Text.HarfBuzz`，由原 11 包变为 12 包，实际清单以 `verify-package-artifacts.py:PACKAGES` 为准。PDF→OFD 对未验证 RTL 语义明确拒绝，仍可显式使用纯页面图路径；有界 R02、归档与各字体出口维持各自契约。
- **N03 正字号输出契约**：创建新文本时，如有限正字号按既有 `0.###` 序列化会变为零，输出前明确拒绝，避免无效 `Size="0"`；普通字号及坐标精度不变。typed writer 返回 `TextLayoutException/UnsupportedOutput`，primitive package writer 抛 `NotSupportedException`。外来 `SourceXml` 合法原始精度继续保留；不是全局增加精度或固定最小字号。详见 [N03](validation/cn-en-release-n03.md)。
- **N12 发布绑定**：源码 fingerprint 对合法 LFS pointer/已下载载荷统一绑定相同 oid/size，并严格拒绝错哈希/大小/非法 index；普通跟踪文件继续按实际字节绑定。保留全部发布门、测试和 tag 工作流，不以旧 preview.9 验收凭证放行本版。
- 提供可复现的中英文真实 Regular/Bold [API 字体部署教程](tutorials/cn-en-font-deployment.md)及页眉页脚/内嵌图/选页的有限验收资料。本版 DOCX 页面通过结论限定为下面的具名配置。

## 受限字体配置与兼容性

发布 profile **`cn-en-physical-cjk-400-700-v1`**：显式目录提供固定 `Noto Sans CJK SC` 静态 Regular 400 / Bold 700，以及 `Noto Sans` Regular；Native API 设置严格 `NativeEmbeddedBoldBehavior=Throw`、显式 fallback 和 `UseInstalledMicrosoftOfficeFonts=false`（只关闭特定 Office 字体暂存，不代表禁用全部系统字体解析）。源字体哈希、OFL 来源/许可证及两模式完整示例见[受限发布配置](release-profile-cn-en.md)与教程。源字体不随 NuGet 分发；fontTools 仅用于离线资料准备，不是应用运行时依赖。

不承诺 Noto Sans 真实 Bold、物理 Italic、任意字体面、任意系统字体、复杂 Word 版式或跨机像素一致。CLI 仍使用默认 TextStroke 策略，没有严格粗体开关；严格 profile 使用 API。F01 `generated-layout.docx` 经公开包独立消费，在显式 native 与默认模式各生成两页；189 个非空白源字符保持，PDF 无重复文本，四页 PNG 和真实 macOS Preview 检查通过。Preview 检查由用户授权的 `gpt-6.1-terra` High AI reviewer 完成，不称人工签收；仅覆盖这些配置、样例及四页。

## 依赖、资产与运行范围

- 普通 SDK 为 `netstandard2.0/netstandard2.1`；CLI/E2E 要求 .NET 10。目标框架声明不证明所有宿主、RID 或发布方式均已运行。
- HarfBuzz 适配包固定 `HarfBuzzSharp` 及 macOS/Win32/Linux native-assets **14.2.1.301**；可选塑形调用需要相符的原生库。默认字体查询/有界 writer 能力不等于自动通用段落排版。公开源消费已运行 HarfBuzz 版本/塑形调用，未独立观察其原生库实际加载路径，不能由包内文件推定。
- 文档/PDF 链路沿用 OpenXML 3.5.1、PdfSharpCore/MigraDocCore.Rendering 1.3.67、PdfPig 0.1.15、Docnet.Core 2.6.0、SixLabors.Fonts 1.0.1 与 ImageSharp 2.1.13；完整直接/传递依赖及许可由还原后 Q01 对照 40 项受审清单核验。历史还原报告 ImageSharp NU1902/NU1903，未由本轮修复消除。
- 现有 Linux/Windows/macOS 工作流和既有 RID/原生资产门继续执行；不新增 musl、AOT、trimming、single-file 或任意架构支持承诺。Word/LibreOffice 直接 PDF/双层路径有独立外部依赖，不能冒充 native 链路。
- Skia、PDF vector、SES interop、password crypto 是独立可选扩展，不在默认 12 包发布循环。仓库 MIT、依赖与外部工具/字体许可见 [THIRD-PARTY-NOTICES](../THIRD-PARTY-NOTICES.md)。

## 预览版已知限制

- **Regular-only 合成粗体**：原未修改样例在默认 TextStroke 下的表头“量”字仍有字腔堵塞，历史视觉 FAIL 保留；真实 Bold 配置不等于该缺陷已修复。本版承诺限定到上述物理配置。
- **macOS Preview Find**：macOS 26.6.2 (25G83) 上，原生成 PDF 的完整页末日期/部分整词可能同时显示 Found on 1 page 与 Not found，无法导航/高亮；根因未定。本版不承诺 Preview 通用搜索兼容，也不把 PDFKit/AX/文本提取成功当作 UI 通过。
- **原始 Unicode Copy 未验证**：原始剪贴板字节验收未完成；粘贴到搜索框不构成复制保真证据。其他阅读器的 Copy/Find 同样未验证，不能宣称通过。
- R02 完整外来映射、R20 完整归档符合性、通用 R22、RTL/Arabic/Indic 扩展与高级字体/新部署组合保持既有有限/未完成边界。普通测试绿不解除这些限制。

用户于 2026-11-02 明确接受以上具名范围/限制用于本预览版；旧验收失败不改为 PASS。profile 内新的视觉/原文/功能失败与强制工程门失败仍阻断。发布验证结果及其证据身份见[最终发布记录](validation/preview10-publication-2026-11-02.md)；公开消费时未重测 Preview Find 或原始 Unicode Copy。[此前独立验证](validation/cn-en-release-local-validation.md)仅为旧源码/产物证据。

## 发布状态

GitHub Release 于 **2026-11-02 05:27:19 UTC（北京时间 13:27:19）**公开。[Trusted Publishing 工作流 41391145165](https://github.com/whynpc9/ofdrw.net/actions/runs/41391145165)成功上传 12 个主包及 11 个符号包；Converter 元包无 snupkg。全部主包版本、内部依赖、产品源码身份和 NuGet 仓库签名已验证；全新公开源缓存的 API、HarfBuzz 与 CLI 消费完成。CLI 文本提取会丢终端 LF，不宣称字节完全一致。安装命令使用 `0.1.0-preview.10`。

历史观察：2026-11-02 01:13:57 UTC 的版本占用检查当时仍为 11 个既有 ID 最新 preview.9、HarfBuzz ID 404、没有 preview.10 tag；这是发布前快照，不是当前状态。首次发布因过期 API key 在首次上传返回 403，无包发布；后续通过 PR96/PR97 在保留原标签/产品源码的条件下恢复，详见最终发布记录。
