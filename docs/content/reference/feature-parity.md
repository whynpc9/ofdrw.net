---
title: "Ofdrw.Net 与 OFDRW 功能对照"
description: "按能力域比较 Ofdrw.Net 与 Java OFDRW，标明已支持、部分支持、扩展点和验证边界。"
sidebar:
  order: 50
  label: "功能对照"
---

维护基线为[公开产品源码 `4970c1c5`](https://github.com/whynpc9/ofdrw.net/tree/v0.2.0)（`0.2.0`）。固定样例的发布验收范围见[0.2.0 发布说明](https://whynpc9.github.io/ofdrw.net/releases/020/)。Preview Find 和原始 Unicode Copy 的限制仍适用。

本表描述有限实现；逐入口/出口、输入语义与失败策略见[能力契约](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/)。功能已实现、实际视觉通过、互操作和公开发布是不同结论；不构成 GB/T 33190 或密码应用规范认证。状态中的“已支持”仅限所列 API/输入，“部分支持”保留明确拒绝和未验证范围，“扩展点”需要调用方提供实现。

## 能力矩阵

| 能力域 | OFDRW | Ofdrw.Net 当前状态 | 结论 |
| --- | --- | --- | --- |
| OFD 核心模型 | `ofdrw-core` 提供较完整标准数据结构 | 页面、图层、模板、文本游程、路径、图片、字体、附件、自定义标签、保留未知 XML/包条目 | **部分支持**；常用读写闭环已具备，完整标准对象模型尚未覆盖 |
| 容器与打包 | `ofdrw-pkg` | 标准入口/文档/页面/资源引用，附件、模板和扩展条目；ZIP 路径、数量、展开大小和压缩比限制 | **已支持** |
| 大纲、书签、动作与注释 | `ofdrw-core` | typed 层次/页目标、GoTo、局部 Text/Image/Path Appearance；未知动作/XML 保留 | **部分支持**；动作仅序列化；导入图元编辑、未证明闭包的合并/混合拒绝。见[教程](https://whynpc9.github.io/ofdrw.net/reference/core-model/) |
| 裁剪与组合对象 | `ofdrw-core` | typed Path Clip/Area、共享 CompositeGraphicUnit 与 nested DAG placement；读写、独立克隆和 PDF/SVG | **部分支持**；每 Clip 一个填充 Path Area，可包含多子路径，多个 Clip 相交；多 Area 可往返，导出拒绝一般 union；Text clip、group opacity/DrawParam 等保留并拒绝不安全导出。见[路径教程](https://whynpc9.github.io/ofdrw.net/reference/core-model/) |
| 文档读取 | `ofdrw-reader`，含反序列化及签章读取 | 多图层/模板/资源解析，未知内容无损保留，结构检查 | **部分支持**；默认首个 `DocBody`，可显式[枚举与选择](https://whynpc9.github.io/ofdrw.net/guides/multi-docbody/) |
| 文档生成 | 布局、Canvas、Graphics2D、多种块元素 | 低层 builder、公开 FlowDocument/Paragraph/Span、Table/Row/Cell/ImageBlock/CanvasBlock 与原生 OfdGraphics API | **部分支持**；中英段落/局部样式/分页、固定列宽/水平合并/整行分页、不可拆图片/Canvas 布局块及页/片段回调已实现；浮动布局与复杂 Word 版式仍未覆盖，见[Flow 契约](https://whynpc9.github.io/ofdrw.net/guides/flow-layout/) |
| 段落排版 | Java 侧布局接口 | Unicode 17 双向、断行、字素簇分析，HarfBuzz 塑形、分页和簇映射 OFD writer | **部分支持**。独立入口和已验收脚本见[段落排版](https://whynpc9.github.io/ofdrw.net/reference/text-paragraph-layout/) |
| 字体 | `ofdrw-font`，字体生成与字形处理 | 按全包用字生成静态 TrueType 子集、复合/GSUB/UVS 字形闭包、TTC 选面、内容身份载荷复用，缺字明确失败 | **部分支持**；TTF/glyf 子集已实现；CFF/变量/AAT/RTL 与保留未知 XML 的字体以全量和诊断保护，详见[字体子集契约](https://whynpc9.github.io/ofdrw.net/reference/font-subset-and-reuse/) |
| OFD → PDF | `ofdrw-converter` 导出能力 | 模板、图层、字体、文本游程、矢量路径、图片、变换、颜色/透明度和裁剪原点 | **部分支持（预览）**；CGTransform 仅接纳[原 GID 有界候选](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/r02-cgtransform-profile.md)和 0.2.0 段落 writer 产生的受限簇映射，逐出口范围见[契约](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/) |
| PDF → OFD | `ofdrw-converter` 转换能力 | 默认双层；可选 Pdf.Vector 包经 19/04 输出受限 native PathObject/TextObject，校验原文与嵌入字体字形 | **部分支持**；[矢量模式](https://whynpc9.github.io/ofdrw.net/guides/pdf-vector/) 支持明确页框/路径/TrueType profile；压缩流、clip/图像/effects/复杂字体等明确整页回退或失败；默认 CLI 仍双层，扫描件无 OCR，不承诺任意 PDF 保真/阅读顺序标记 |
| DOCX → PDF / OFD | 上游不提供 Office 文档渲染 | 默认直接将 DOCX/OpenXML 原文写为原生 OFD `TextObject`，完全跳过 PDF；可选 `DualLayer` 使用 Word/LibreOffice/BuiltIn 页面图像，并继续以 OpenXML 原文作为语义层 | **已支持（预览）**；原文可直接抽取，Native 支持常见段落、表格、内嵌图片与按页页眉页脚；附属正文带标签追加，DualLayer 按实际页定位原文，不承诺复杂 Word 版式和浮动对象保真 |
| OFD → SVG | 上游推荐/生态具备 SVG 预览方案 | 自包含单页 SVG，支持模板、路径、文本、颜色、变换和内嵌图片 | **部分支持（预览）**；CGTransform 仅接纳[原 GID 有界候选](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/r02-cgtransform-profile.md)和 0.2.0 段落 writer 产生的受限簇映射，其它语义不从 PDF/HTML 结果推定 |
| OFD ↔ PNG/JPEG | 图片导出与一图一页导入 | `Converter.Pdf` 的 `OfdToImageConverter` / `ImageToOfdConverter` 及 CLI；指定零基页/ppm导出、PNG默认/JPEG、自然或固定页等比居中导入、有界暂存 | **有限实现**；PDF/PDFium 预览保真边界，视觉门状态见 [Issue 02 验证记录](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/issue02/README.md) |
| 纯文本 → OFD | `TextConverter` | `PlainTextOfdConverter` 与 `text-to-ofd` CLI：严格 UTF-8、有限预算、Flow 分页 | **部分支持**；首版源码及测试入口已存在，复杂 shaping/字体嵌入不由此推定；见[契约](https://whynpc9.github.io/ofdrw.net/guides/plain-text-ofd/) |
| HTML 导出 | 转换模块覆盖 HTML 方向 | `OfdToHtmlConverter` 与 `ofd-to-html` CLI：有界、自包含静态 SVG 页面、严格预检 | **部分支持**；首版源码及测试入口已存在；外来复杂图元及各出口分别验收，见[教程](https://whynpc9.github.io/ofdrw.net/guides/html-preview/)与[契约](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/) |
| OFD-A 归档检查 | `ofdrw-archive` | 有限只读 archive preflight；已有有限标准 Rule/Handler，完整条款覆盖尚未完成 | **部分支持**；结构/有限预检不代表档案合规或认证，见[预检契约](https://whynpc9.github.io/ofdrw.net/guides/archival-preflight/)与[能力契约](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/) |
| 文本提取 | 上游提供文字抽取 | 页面和模板文本提取，API 与 CLI | **已支持** |
| 关键词定位与替换 | 文本位置查询及文档工具 | 正文字面匹配、原嵌入字体字形/基线/advance/字形框证据；受限替换 API 与原子文件 CLI | **部分支持**；Regular 原字体、简单 scalar、水平正向与仅平移替换；模板/批注、复杂 shaping 和超出原区域的替换明确拒绝。见[教程](https://whynpc9.github.io/ofdrw.net/guides/keyword-replacement/)和[验收证据](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/evidence/issue06-owned-save/README.md) |
| 文档工具 | `ofdrw-tool` 提供合并、裁剪、重组、混合等 | 页重排、删除、裁剪、自包含合并；指定层文字/图片水印、按列表 Split、带模板/可见注释的 Mix、全部 DocBody 清签，API 与 CLI | **部分支持**；03 实现范围见[工具教程](https://whynpc9.github.io/ofdrw.net/guides/document-tools/)，复杂未知引用明确失败、共享资源保守保留；附件已有逻辑移除与受限显式物理清理，见[附件教程](https://whynpc9.github.io/ofdrw.net/reference/core-model/)；视觉/review 状态见[03 证据](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/evidence/document-tools/README.md) |
| 数字签名 | `ofdrw-sign` | 标准签名目录/引用生成、签名值 provider、签名值 verifier | **扩展点** |
| 摘要校验 | 上游签章验证链路 | 内置 SM3、SHA-1、SHA-256，逐引用恒定时间比较，篡改检测 | **已支持** |
| SES/SM2 电子签章 | `ofdrw-sign` + `ofdrw-gm` | 可注册厂商 `IOfdSignedValueVerifier` / `IOfdSignatureProvider` | **扩展点**；核心无默认 SES/SM2 实现和印章外观验证 |
| SES dev/interop 测试容器 | 固定上游 SES V1/V4 测试容器 | 可选 `Ofdrw.Net.Signatures.SesInterop`：常见字段解析、SM2 测试自签/双层验签、完整测试证书固定 | **测试扩展**；仅显式注册，默认 verifier/CLI 不给自签 FullyValid；核心能力标志仍 false。无阅读器互认、法律效力或合格 TSA，范围见[教程](https://whynpc9.github.io/ofdrw.net/guides/ses-interop/)与[证据](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/evidence/ses-interop/README.md) |
| GM/T 0099 密码应用 / 可选加密扩展 | `ofdrw-crypto` | 核心能力标志仍 false；独立 `Crypto.Password` 保留私有 password v2，新增显式 Java 2.4.0 兼容的口令/SM2 证书同类多接收人及签名 `OFDEntries` 有限子集 | **核心未支持 / 可选有限扩展**；v2 有 scrypt/HMAC；标准子集的 Java 原生清单只证明路径，逐文件内容仅在本包的签名摘要扩展存在时可验；非通用 GM/T/厂商互认或生产 PKI，见[契约](https://whynpc9.github.io/ofdrw.net/reference/password-crypto/)及[R14 证据](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/evidence/r14-standard-encryption/README.md) |
| CLI/工具分发 | `ofdrw-tool` 及各模块工具 | 转换、SVG、PNG/JPEG 图片进出、文本提取/定位、归档预检、页面与附件管理、合并、签名验证的 .NET tool | **已支持** |
| Graphics2D 兼容层 | `ofdrw-graphics2d` | Layout 公开 OfdGraphics/OfdPen/OfdBrush/OfdFont/OfdGraphicsPath/OfdMatrix，输出原生 PathObject/TextObject/ImageObject；R09 新增图片裁切和可配置线型、显式栅格矩形渐变/图片 tile | **部分支持**；04 原语绘图已实现，无 Skia 必需依赖；R09 的源与出口边界见[设计](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/r09-graphics-design.md)，外来原生渐变/图案仍不支持；19 已实现受限合作 producer 事件适配（不透明捕捉 SKCanvas，验收见独立证据）；05 已有有限 TrueType 子集实现，其他字体格式继续遵循[契约](https://whynpc9.github.io/ofdrw.net/reference/font-subset-and-reuse/) |
| 兼容性与发布 | Maven 多模块、长期演进 | `netstandard2.0/2.1` SDK + `net10.0` CLI，NuGet 包消费 E2E | **部分支持**；仓库采用 MIT 并附第三方声明；发布候选需通过同一批 NuGet 包的消费验证 |

## 后续范围与验证

剩余问题、被动观察、取消项与未来候选统一见[现行路线图](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/roadmap.md)。已实现的命名区域、Canvas、多 DocBody、有限签名/加密/归档能力不再重复列为新任务。Graphics 原语与字体子集分别见[绘图契约](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/graphics-design-contract.md)、[字体契约](https://whynpc9.github.io/ofdrw.net/reference/font-subset-and-reuse/)及各自验收记录，不从单一接口推定整个上游 issue 已关闭。

实际结果按各 `validation/`、`evidence/` 报告中的原源码、输入和页范围解释；旧失败不重标通过。R00 PageBlock 的[原漏绘样本](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/evidence/r00-pageblock/README.md)仍保留，后续 R01 的[修复记录](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/evidence/r01-pageblock-paint/README.md)另证其有限结果。新生成产物须按[发布门禁](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/release-gates.md)与根 `AGENTS.md` 重新验收。

依赖健康（2026-10-15 原记录）：[PDF 项目](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Pdf/Ofdrw.Net.Converter.Pdf.csproj)与[DOCX 项目](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Docx/Ofdrw.Net.Converter.Docx.csproj)仍固定 `SixLabors.ImageSharp 2.1.13`；维护者 [GHSA-wmxv-xphr-5c9g](https://github.com/SixLabors/ImageSharp/security/advisories/GHSA-wmxv-xphr-5c9g)（2026-09-15 发布，Moderate，CVSS 5.3）列出的受影响区间包含该版本。图片导入在 Identify 前只准入 PNG/JPEG；DOCX、OFD 与签章位图在 Identify/解码前拒绝 BigTIFF 的两种字节序，保留经典 TIFF 等既有格式。此边界修复不代表全部 ImageSharp 风险已消除；未做全量传递依赖/deprecated 审计，版本、目标框架与许可未变。

生产签章与可选加密的边界见[签名保护范围](https://whynpc9.github.io/ofdrw.net/reference/signature-scope-v12/)、[SES 教程](https://whynpc9.github.io/ofdrw.net/guides/ses-interop/)与[可选加密](https://whynpc9.github.io/ofdrw.net/reference/password-crypto/)。核心能力标志不因可选包或测试容器而提升。
