---
title: "Ofdrw.Net 能力契约与范围索引"
description: "按入口、出口和输入语义查阅支持状态、回退、失败输出与各工作单元的能力边界。"
sidebar:
  order: 60
  label: "能力契约"
---

`0.3.1` 的骑缝章修复及验收范围见[0.3.1 发布说明](/releases/031/)。先前交付能力见[0.3.0 发布说明](/releases/030/)；[0.2.0 发布说明](/releases/020/)保留固定 DOCX 和段落样例的原版本结论。Preview Find、原始 Copy 和段落 PDF 选择几何仍有独立限制。

维护基线为 `0.3.1` 产品标签源码。本页按功能说明已交付能力的输入与出口契约，并显式标注候选和未验证范围。后续工作见[现行路线图](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/roadmap.md)。

## 设计与判定

能力的键是「入口及出口、输入语义、策略」的组合。`supported` 仅指本行列出的有限实现和已有测试入口；仍须按本行的视觉、互操作、包消费门分别验收。`unsupported` 表示源码明确拒绝或没有该入口；没有实现不能凭格式名推断会正确处理。`unverified` 表示有模型、保留或静态候选，但尚无当前样本证明导出或互操作。保留 XML、可编辑、可呈现、可跨文档迁移是四个独立结论。历史失败照原记录保留。

失败出口的共同规则：文件型 CLI 以目标旁临时文件和 `AtomicOutput` 提交，异常或取消不得把临时产物报告为成功；调用者传入的任意 `Stream` **不保证回滚**，复制或 flush 失败可能留下前缀，调用者必须丢弃输出。未来新入口应预检目标出口能力、在最终提交前完成转换；显式 fallback 必须写进结果/诊断并以整页或明确单元替换，不能把失败前的半成品混入成功结果。该规则不把现有所有流 API 宣称为原子写入。

### 已有出口与输入语义

| 入口 → 出口、输入 | 状态与实际边界 | 默认/显式 fallback；失败输出 | 源码与验证入口 |
| --- | --- | --- | --- |
| DOCX → Native OFD；常见 OpenXML 段落、表格、图片、页眉页脚与原文 | `supported` 有限布局；默认模式为 Native。复杂 Word 浮动、域、跨页表格不保证保真；不支持内容依 `UnsupportedFeatureBehavior` 诊断或失败 | 默认不会悄悄改走 DualLayer。嵌入常规面加粗默认 `TextStroke` 并报告 `DOCX_EMBEDDED_BOLD_SIMULATED`；可设 `Throw`。CLI 原子提交，流输出按上段处理 | [`DocxConversionOptions`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Docx/DocxConversionOptions.cs)、[`DocxToOfdConverter`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Docx/Converters/DocxToOfdConverter.cs)、[`转换约定`](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/) |
| DOCX → DualLayer OFD；同源 OpenXML 原文 | `supported` 可选；页面图依赖 Word/LibreOffice/BuiltIn 渲染，文字语义从 OpenXML 映射，未定位正文且选页时拒绝 | 渲染引擎按配置回退；视觉页必须栅格化，不能用 PDF 抽字冒充原文。`ConvertWithResultAsync` 报实际引擎、诊断和页映射 | 同上；[`DocxToPdfConverter`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs) |
| PDF → OFD，默认 `PdfToOfdConverter` | `supported` 页面图加受限透明文字；扫描件无 OCR，不能保证阅读顺序。语义抽取遇未验证 RTL/bidi 字符默认明确拒绝；可显式选纯页面图，或自 0.3.0 起选用受限的 `VisualLineProfile`（F06） | 尝试配置的页面栅格器；失败明确抛出，不写空白成功页。RTL 语义边界与纯图选项见[转换契约](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/#pdf-导入的文字顺序边界)。文件 CLI 原子提交 | [`PdfToOfdConverter`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Pdf/Converters/PdfToOfdConverter.cs)、[`PdfToOfdOptions`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Pdf/PdfToOfdOptions.cs) |
| PDF → 原生矢量 OFD，可选包 | `supported` 为受限的不透明路径、固定字体 Unicode 文字、路径裁切、画笔及混合页 raw RGB 图像；Form XObject、复杂字体/透明度等仍拒绝 | `UnsupportedPagePolicy=Fail` 默认；显式 `RasterizePage` 才整页图像/双层回退。预算、取消及不可表示页不转成静默回退；结果逐页标注 native/raster | [`PdfVectorToOfdOptions`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Pdf.Vector/PdfVectorToOfdOptions.cs)、[`PdfVectorToOfdConverter`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Pdf.Vector/PdfVectorToOfdConverter.cs)、[`教程`](https://whynpc9.github.io/ofdrw.net/guides/pdf-vector/) |
| OFD → PDF；Reader 投影的页面/模板、文字、路径、图片 | `supported` 有限预览；R01 的普通页和模板嵌套 `PageBlock` Path/Text/Image 已按原顺序绘制。文字 Fill/Stroke/HScale 与受限 DrawParam/ColorSpace 投影可呈现；仅 8-bit GRAY/RGB/无 ICC 的简单 CMYK 及有界调色板。未知资源、外来原生渐变/图案等拒绝；R09 的显式矩形栅格填充是普通 ImageObject，须按图片出口分别验收。R02 一对一候选接纳 BMP 独立字符、完整原 GID 映射、允许嵌入的静态 TrueType `glyf`；多 `TextCode` 全局位置已在 PR #58 的同输入包与 Preview 验证。PR #59 另接纳单个无 Delta 的 2–8 BMP 字符源片段映射一个原 GID；PR #60 接纳单字符到两个原 GID，PR #62 接纳单个完整带显式定位的 1:N/N:M 源跨度，均有独立包消费与 Preview 显示/复制记录；PR #62 的 Preview 搜索仍有歧义 | 无通用自动栅格 fallback；未解析的颜色引用明确失败，包括签章内嵌 OFD；API `Stream` 可能留下前缀，不将部分 PDF 当成功 | [`OfdToPdfConverter`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs)、[`Reader`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Reader/Readers/OfdReader.cs)、[`R02 候选`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/r02-cgtransform-profile.md) |
| OFD → 独立 SVG；同类图元 | `supported` 有限预览；上述 `PageBlock` 与文字/路径颜色样本按独立 SVG 页面检查。字体、复杂颜色与 CGTransform 保留各自有限边界 | 无通用自动 fallback；未知图元或未解析颜色拒绝；目标流失败语义同上 | [`OfdToSvgConverter`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs)、[`R01 证据`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/evidence/r01-pageblock-paint/README.md) |
| OFD → 静态 HTML；选中页的内嵌 SVG | `supported` 的有限安全预览；R01 允许同一受限 `PageBlock`/颜色/文字 Stroke/HScale 样本，与独立 SVG **分别验收**。不执行动作、脚本或外部资源；HTML 预检继续拒绝未知属性、子节点及不支持的 raw 绘图 | 生成至有界 spool 后复制；生成失败不触碰公共流，最终复制/flush 失败公共流可留前缀；CLI 原子提交 | [`OfdToHtmlConverter`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Svg/Converters/OfdToHtmlConverter.cs)、[`HtmlRenderContext`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Svg/Internal/HtmlRenderContext.cs)、[`R01 证据`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/evidence/r01-pageblock-paint/README.md) |
| OFD → PNG/JPEG；单页预览 | `supported` 的有限图像导出，复用 PDF 预览路径；继承 R01 已验证的有限 `PageBlock`/颜色 profile 与拒绝边界；不是独立证明外来 OFD 所有图元正确 | 不支持或超预算明确失败；CLI 文件提交仍按原子约定 | [`OfdToImageConverter`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Pdf/ImageIO/OfdToImageConverter.cs)、[`R01 证据`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/evidence/r01-pageblock-paint/README.md) |
| PNG/JPEG → OFD；一图一页 | `supported` 仅白名单输入；尺寸/像素预算与居中几何有约束，非通用图片格式导入 | 非白名单或超限拒绝；不默默转空白页 | [`ImageToOfdConverter`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Pdf/ImageIO/ImageToOfdConverter.cs) |
| UTF-8 纯文本 → OFD；公开 Flow | `supported` 首版严格 UTF-8、有限预算与分页；没有通用字体嵌入或复杂脚本 shaping 承诺 | 不合法输入、溢出或超限拒绝；CLI 原子提交。已有 Flow 度量不等于真实 face 测量 | [`PlainTextOfdConverter`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Layout/PlainTextOfdConverter.cs)、[`PlainTextOfdTests`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/tests/Ofdrw.Net.Layout.Tests/PlainTextOfdTests.cs) |
| 已有 OFD → 保存/编辑/合并；未知源对象 | `supported` 仅各 API 约定的对象/资源闭包；未知 XML 保留不等于可导出或可迁移 | 合并默认拒绝未能安全映射的 raw；仅显式 `SkipUnsupportedRawElements` 可跳过，必须读取结果诊断。签名改写失效另行报告 | [`转换约定`](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/)、[`功能对照`](https://whynpc9.github.io/ofdrw.net/reference/feature-parity/) |
| 外来 OFD 显式 `CGTransform` → PDF/SVG/HTML | 保存源 XML：`supported`。[R02 有界能力](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/r02-cgtransform-profile.md)覆盖独立 BMP 字符的一对一原 GID、多个全局递增 `TextCode`、单个完整无 Delta 源片段到一个原 GID；单字符到两个原 GID 的 PR #60 子集已完成本地、包消费、PDF/SVG 页面及 Preview 证据。PR #62 另验收单个完整带显式字形槽定位的 1:N/N:M 源跨度，包括 Tamil 2:3、包消费、独立读者文本、PDF/SVG 页面、Preview 显示和复制以及精确 head CI/审查；Preview 搜索仍有歧义。未经修改的外来文件与更广映射仍须分别验收 | 范围外明确拒绝；无自动文字代换/光栅 fallback。局部 SVG 外来页面核对不等于全部出口放行 | [`PDF`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs)、[`SVG`](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs) |
| SES/SM2、GM/T 0099、归档 | 摘要引用完整性与可插拔 provider 已有；SES 测试容器和 password v2 是显式可选包；通用标准加密/生产 PKI/档案认证 `unsupported` 或 `unverified` | 自签、摘要匹配、外观可见均不自动升级为生产 `FullyValid` 或认证 | [`功能对照`](https://whynpc9.github.io/ofdrw.net/reference/feature-parity/)、[`能力路线`](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/roadmap.md) |

### 有界能力索引

下表合并旧开工矩阵；只列可复用入口与仍有效的限制，不把已经交付的单元继续标为 `new`。源码、功能测试、实际页面和外部互操作分别以链接内具名记录为准，文档整理不新增验收结论。

| 原单元 | 现有范围与维护入口 | 保留边界 |
| --- | --- | --- |
| R00/R01 容器与绘图投影 | 普通页/模板嵌套 PageBlock、文字/颜色、附件动作；[修复记录](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/evidence/r01-pageblock-paint/README.md) | 未建模图元和无效颜色拒绝；未知 XML 保留不代表可呈现 |
| R02 显式 GID 映射 | [全局受限映射](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/r02-cgtransform-profile.md)与显式 [GroupedLocal](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/r02-grouped-textcode.md)，共享 PDF/SVG/HTML 投影 | 默认保持 LegacyConcatenatedGlobal；非 BMP、越段、广泛自然外来互操作未解锁，完整 R02 未验收 |
| R03 命名区域 | [持久命名、页引用及回填](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/r03-named-areas-design.md) | 重名、丢失和溢出明确失败，不自动重新分页 |
| R04/R19 Flow 与表格 | [图片/Canvas/回调/Span](https://whynpc9.github.io/ofdrw.net/guides/flow-layout/)、[有限复杂表格分页](https://whynpc9.github.io/ofdrw.net/guides/public-tables/) | 固定块/合并行组及拆行/重复表头遵守各自规则；不等于 DOCX 任意表格保真 |
| R05 Merge/Mix | [文档工具](https://whynpc9.github.io/ofdrw.net/guides/document-tools/)的共享/嵌套资源、裁剪/组合与批注闭包 | 未闭合引用默认拒绝，跳过必须显式且报告 |
| R06 PDF 矢量 | [受限路径/字体/画笔/裁切/raw RGB](https://whynpc9.github.io/ofdrw.net/guides/pdf-vector/) | 默认 Fail；显式整页回退不保留半成品矢量，不承诺任意 PDF |
| R07 导航与附件导入 | [PDF 书签、附件迁移](https://whynpc9.github.io/ofdrw.net/guides/pdf-import-navigation-attachments/) | 默认关闭；页映射、冲突与失败诊断独立判定 |
| R08 字体后端 | [四出口能力表](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/r08-four-outlet-ledger.md)与各固定 profile | TrueType、CFF、symbol、非 BMP 各出口分别验收；变量/彩色/任意字体不推定支持 |
| R09 绘图上下文 | [图片裁切、画笔和显式栅格矩形填充](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/r09-graphics-design.md) | 外来原生渐变/图案和任意路径渐变仍拒绝，栅格填充不是矢量渐变 |
| R10 图像与区域 | [JBIG2/ImageMask 与受限区域提取](https://whynpc9.github.io/ofdrw.net/reference/r10-restricted-image-region-profile/) | 超预算、未知解码/几何拒绝或 Unknown，不作任意区域定位保证 |
| R11 签章外观 | [普通章/骑缝外观](https://whynpc9.github.io/ofdrw.net/reference/core-model/) | 可见外观不证明密码学有效、信任或 SVG 签章投影支持 |
| R12 SES | [V5 与固定 Java 双向互操作子集](https://whynpc9.github.io/ofdrw.net/guides/ses-interop/) | 默认 verifier 不注册测试容器；真实设备、PKI/TSA/法律效力未验证 |
| R13 签名范围 | [References、续签与可变目录](https://whynpc9.github.io/ofdrw.net/reference/signature-scope-v12/) | 无签名值 provider 不得假锁定；摘要匹配不等于生产信任 |
| R14 加密 | [私有 v2 与显式标准交换子集](https://whynpc9.github.io/ofdrw.net/reference/password-crypto/) | 核心标志不变；混合接收人/多层及完整 GM/T/厂商互认未支持 |
| R15 多 DocBody | [枚举、选择、导出与隔离编辑](https://whynpc9.github.io/ofdrw.net/guides/multi-docbody/) | 默认首文档；签名、共享资源和跨文档引用的编辑受限 |
| R16 标准对象 | [DocInfo/权限/版本/动作编辑](https://whynpc9.github.io/ofdrw.net/guides/standard-object-editing/) | 未知节点保留；声明权限不等于执行策略 |
| R17/R20 归档 | [有限 Rule/Handler](https://whynpc9.github.io/ofdrw.net/guides/archival-preflight/)、[条款覆盖账本](https://whynpc9.github.io/ofdrw.net/reference/archival-standard-coverage/) | Unknown 不升格 Pass；完整 R20 未验收，不宣称认证 |
| R18 附件清理 | [逻辑移除与显式物理清理](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/r18-attachment-cleanup-design.md) | 只删除来源/独占载荷与引用闭包可证明的对象，不能证明时拒绝 |
| 文字查询与输出 | [固定字体测量](https://whynpc9.github.io/ofdrw.net/reference/text-prelayout/)、[类型化输出](https://whynpc9.github.io/ofdrw.net/reference/text-shaped-output/)、[混排查询](https://whynpc9.github.io/ofdrw.net/reference/text-mixed-paragraph/) | 0.2.0 新增[独立段落排版与 OFD writer](https://whynpc9.github.io/ofdrw.net/reference/text-paragraph-layout/)，Flow/纯文本/DOCX 自 0.3.0 起可显式启用最终文字排版引擎，默认仍使用原入口。[阅读器验证记录](https://github.com/whynpc9/ofdrw.net/blob/main/docs/validation/text-reader-method-validation.md)保留有限配置的反例与边界 |

## 证据与门禁

[原 R00 PageBlock 漏绘](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/evidence/r00-pageblock/README.md)是历史负面证据；[R01 修复记录](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/evidence/r01-pageblock-paint/README.md)不回写它的原结论。`validation/` 与 `evidence/` 中各次测试、包消费、PNG/Preview 和 Copy/Find 仅绑定原源码及产物，不能移作新候选凭证；[文档索引](https://whynpc9.github.io/ofdrw.net/)提供旧开工/调度记录的固定提交入口。

每个影响页面的变更仍按根 `AGENTS.md` 和[发布门禁](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/release-gates.md)保存源码 SHA、输入/输出哈希、真实模式、功能/包消费及全部受影响页的 PNG/Preview 结果。未检查的互操作标为 unverified；重新生成后重新打开产物；视觉缺陷未解决不得标整体通过。发布记录和 Q05 JSON 中的身份保持原意，不自动绑定到后续 main 文档树。
