# Ofdrw.Net 与 OFDRW 功能对照

更新时间：2026-10-21（R00 契约核对）

本表沿用[历史开发计划](capability-roadmap-development-plan.md)固定的 OFDRW 比较快照；原状态核对固定在 `d0f88c7321a57443a4cb0ec92631664f66ed8cc0`。R00 在 `6aa8183ea31e193f500fb9d9717677b15e86c5d3` 重核源码入口与仓库文件，逐出口、输入和 fallback 边界见 [v1.2 能力契约](capability-contract-v12.md)。历史 ticket 状态和验收记录见 [ticket 索引](../.scratch/capability-roadmap/issues/README.md)。R00 除定向功能测试外，重生成并用 PNG/Preview 复查了一个 PageBlock 缺陷样本；该样本视觉未通过；源码、测试入口、合并、视觉和发布状态不能互相替代。本表不代表 GB/T 33190 或密码应用规范的合规认证。

状态说明：

- **已支持**：存在公开 API/CLI，且有自动化测试或真实样本验证。
- **部分支持**：主流程可用，但数据模型、保真度或算法覆盖仍有限。
- **扩展点**：协议编排已实现，关键算法需要调用方或厂商实现。
- **未支持**：当前没有可用实现。

## 能力矩阵

| 能力域 | OFDRW | Ofdrw.Net 当前状态 | 结论 |
| --- | --- | --- | --- |
| OFD 核心模型 | `ofdrw-core` 提供较完整标准数据结构 | 页面、图层、模板、文本游程、路径、图片、字体、附件、自定义标签、保留未知 XML/包条目 | **部分支持**；常用读写闭环已具备，完整标准对象模型尚未覆盖 |
| 容器与打包 | `ofdrw-pkg` | 标准入口/文档/页面/资源引用，附件、模板和扩展条目；ZIP 路径、数量、展开大小和压缩比限制 | **已支持** |
| 大纲、书签、动作与注释 | `ofdrw-core` | typed 层次/页目标、GoTo、局部 Text/Image/Path Appearance；未知动作/XML 保留 | **部分支持**；动作仅序列化；导入图元编辑、未证明闭包的合并/混合拒绝。见[教程](tutorials/09-templates-attachments-tags.md) |
| 裁剪与组合对象 | `ofdrw-core` | typed Path Clip/Area、共享 CompositeGraphicUnit 与 nested DAG placement；读写、独立克隆和 PDF/SVG | **部分支持**；每 Clip 一个填充 Path Area，可包含多子路径，多个 Clip 相交；多 Area 可往返，导出拒绝一般 union；Text clip、group opacity/DrawParam 等保留并拒绝不安全导出。见[路径教程](tutorials/07-paths.md) |
| 文档读取 | `ofdrw-reader`，含反序列化及签章读取 | 多图层/模板/资源解析，未知内容无损保留，结构检查 | **部分支持**；默认首个 `DocBody`，可显式[枚举与选择](multi-docbody.md) |
| 文档生成 | 布局、Canvas、Graphics2D、多种块元素 | 低层 builder、公开 FlowDocument/Paragraph/Span、Table/Row/Cell/ImageBlock/CanvasBlock 与原生 OfdGraphics API | **部分支持**；中英段落/局部样式/分页、固定列宽/水平合并/整行分页、不可拆图片/Canvas 布局块及页/片段回调已实现；浮动布局与复杂 Word 版式仍未覆盖，见[Flow 契约](flow-layout.md) |
| 字体 | `ofdrw-font`，字体生成与字形处理 | 按全包用字生成静态 TrueType 子集、复合/GSUB/UVS 字形闭包、TTC 选面、内容身份载荷复用，缺字明确失败 | **部分支持**；TTF/glyf 子集已实现；CFF/变量/AAT/RTL 与保留未知 XML 的字体以全量和诊断保护，详见[字体子集契约](font-subset-and-reuse.md) |
| OFD → PDF | `ofdrw-converter` 导出能力 | 模板、图层、字体、文本游程、矢量路径、图片、变换、颜色/透明度和裁剪原点 | **部分支持（预览）**；CGTransform 仅接纳[原 GID 有界候选](validation/r02-cgtransform-profile.md)，其它明确拒绝；逐出口范围见[契约](capability-contract-v12.md) |
| PDF → OFD | `ofdrw-converter` 转换能力 | 默认双层；可选 Pdf.Vector 包经 19/04 输出受限 native PathObject/TextObject，校验原文与嵌入字体字形 | **部分支持**；[矢量模式](tutorials/18-pdf-vector-mode.md) 支持明确页框/路径/TrueType profile；压缩流、clip/图像/effects/复杂字体等明确整页回退或失败；默认 CLI 仍双层，扫描件无 OCR，不承诺任意 PDF 保真/阅读顺序标记 |
| DOCX → PDF / OFD | 上游不提供 Office 文档渲染 | 默认直接将 DOCX/OpenXML 原文写为原生 OFD `TextObject`，完全跳过 PDF；可选 `DualLayer` 使用 Word/LibreOffice/BuiltIn 页面图像，并继续以 OpenXML 原文作为语义层 | **已支持（预览）**；原文可直接抽取，Native 支持常见段落、表格、内嵌图片与按页页眉页脚；附属正文带标签追加，DualLayer 按实际页定位原文，不承诺复杂 Word 版式和浮动对象保真 |
| OFD → SVG | 上游推荐/生态具备 SVG 预览方案 | 自包含单页 SVG，支持模板、路径、文本、颜色、变换和内嵌图片 | **部分支持（预览）**；CGTransform 仅接纳[原 GID 有界候选](validation/r02-cgtransform-profile.md)，其它明确拒绝；其它语义不从 PDF/HTML 结果推定 |
| OFD ↔ PNG/JPEG | 图片导出与一图一页导入 | `Converter.Pdf` 的 `OfdToImageConverter` / `ImageToOfdConverter` 及 CLI；指定零基页/ppm导出、PNG默认/JPEG、自然或固定页等比居中导入、有界暂存 | **已实现（本票源码）**；PDF/PDFium 预览保真边界，视觉门状态见 [Issue 02 验证记录](validation/issue02/README.md) |
| 纯文本 → OFD | `TextConverter` | `PlainTextOfdConverter` 与 `text-to-ofd` CLI：严格 UTF-8、有限预算、Flow 分页 | **部分支持**；首版源码及测试入口已存在，复杂 shaping/字体嵌入不由此推定；见[契约](plain-text-ofd.md) |
| HTML 导出 | 转换模块覆盖 HTML 方向 | `OfdToHtmlConverter` 与 `ofd-to-html` CLI：有界、自包含静态 SVG 页面、严格预检 | **部分支持**；首版源码及测试入口已存在；外来复杂图元及各出口分别验收，见[教程](html-preview.md)与[契约](capability-contract-v12.md) |
| OFD-A 归档检查 | `ofdrw-archive` | 有限只读 archive preflight；完整标准 Rule/Handler 与条款覆盖尚无可确认实现 | **部分支持**；结构/有限预检不代表档案合规或认证，见[预检契约](archival-preflight.md)与[本轮契约](capability-contract-v12.md) |
| 文本提取 | 上游提供文字抽取 | 页面和模板文本提取，API 与 CLI | **已支持** |
| 关键词定位与替换 | 文本位置查询及文档工具 | 正文字面匹配、原嵌入字体字形/基线/advance/字形框证据；受限替换 API 与原子文件 CLI | **部分支持**；Regular 原字体、简单 scalar、水平正向与仅平移替换；模板/批注、复杂 shaping 和超出原区域的替换明确拒绝。见[教程](tutorials/16-keyword-replacement.md)和[验收证据](evidence/issue06-owned-save/README.md) |
| 文档工具 | `ofdrw-tool` 提供合并、裁剪、重组、混合等 | 页重排、删除、裁剪、自包含合并；指定层文字/图片水印、按列表 Split、带模板/可见注释的 Mix、全部 DocBody 清签，API 与 CLI | **部分支持**；03 实现范围见[工具教程](tutorials/15-document-tools.md)，复杂未知引用明确失败、共享资源保守保留；附件已有逻辑移除与受限显式物理清理，见[附件教程](tutorials/09-templates-attachments-tags.md)；视觉/review 状态见[03 证据](evidence/document-tools/README.md) |
| 数字签名 | `ofdrw-sign` | 标准签名目录/引用生成、签名值 provider、签名值 verifier | **扩展点** |
| 摘要校验 | 上游签章验证链路 | 内置 SM3、SHA-1、SHA-256，逐引用恒定时间比较，篡改检测 | **已支持** |
| SES/SM2 电子签章 | `ofdrw-sign` + `ofdrw-gm` | 可注册厂商 `IOfdSignedValueVerifier` / `IOfdSignatureProvider` | **扩展点**；核心无默认 SES/SM2 实现和印章外观验证 |
| SES dev/interop 测试容器 | 固定上游 SES V1/V4 测试容器 | 可选 `Ofdrw.Net.Signatures.SesInterop`：常见字段解析、SM2 测试自签/双层验签、完整测试证书固定 | **测试扩展**；仅显式注册，默认 verifier/CLI 不给自签 FullyValid；核心能力标志仍 false。无阅读器互认、法律效力或合格 TSA，范围见[教程](tutorials/16-ses-dev-interop.md)与[证据](evidence/ses-interop/README.md) |
| GM/T 0099 密码应用 / 可选加密扩展 | `ofdrw-crypto` | 核心能力标志仍 false；独立 `Crypto.Password` 保留私有 password v2，新增显式 Java 2.4.0 兼容的口令/SM2 证书同类多接收人及签名 `OFDEntries` 有限子集 | **核心未支持 / 可选有限扩展**；v2 有 scrypt/HMAC；标准子集的 Java 原生清单只证明路径，逐文件内容仅在本包的签名摘要扩展存在时可验；非通用 GM/T/厂商互认或生产 PKI，见[契约](password-crypto.md)及[R14 证据](evidence/r14-standard-encryption/README.md) |
| CLI/工具分发 | `ofdrw-tool` 及各模块工具 | 转换、SVG、PNG/JPEG图片进出、文本提取、重排、合并、签名验证的 .NET tool | **已支持** |
| Graphics2D 兼容层 | `ofdrw-graphics2d` | Layout 公开 OfdGraphics/OfdPen/OfdBrush/OfdFont/OfdGraphicsPath/OfdMatrix，输出原生 PathObject/TextObject/ImageObject；R09 新增图片裁切和可配置线型、显式栅格矩形渐变/图片 tile | **部分支持**；04 原语绘图已实现，无 Skia 必需依赖；R09 的源与出口边界见[设计](r09-graphics-design.md)，外来原生渐变/图案仍不支持；19 已实现受限合作 producer 事件适配（不透明捕捉 SKCanvas，验收见独立证据）；05 已有有限 TrueType 子集实现，其他字体格式继续遵循[契约](font-subset-and-reuse.md) |
| 兼容性与发布 | Maven 多模块、长期演进 | `netstandard2.0/2.1` SDK + `net10.0` CLI，NuGet 包消费 E2E | **部分支持**；仓库采用 MIT 并附第三方声明；发布候选需通过同一批 NuGet 包的消费验证 |

## 本轮补强结果

本轮优先关闭了影响真实 OFD 使用的高杠杆缺口：

1. 从“只认识少量对象”提升为图层、模板、文字游程、路径、字体、图片、
   附件和未知扩展的损失可控读写。
2. 增加资源引用、全局对象 ID、结构检查和有界 ZIP 加载，降低错误包、
   路径穿越和解压炸弹风险。
3. 提升 OFD → PDF 保真度，并新增自包含 SVG 导出。
4. 增加文本提取、页编辑和文档合并 API/CLI。
5. 增加 DOCX → PDF/OFD 的公开 API、CLI、隔离外部进程和无 Office 的 BuiltIn 兜底，并以无隐私样例覆盖自动回退、资源限制与 NuGet 消费链路。
6. 增加签名生成编排、SM3/摘要校验和可插拔签名值验证，且明确区分
   “引用完整性有效”与“密码学签名完全有效”。
7. 将 PDF 解析依赖迁移到官方 `PdfPig` 包，移除无用包引用，并补齐
   NuGet 消费端 E2E、第三方许可清单和发布许可闸门。

## 后续工作

可执行 TODO、批次顺序和无密码厂商前提下的加密/签章边界见
[能力路线图](capability-roadmap.md)。

本对照表只记录现状。路线图已取代下面这类「愿望清单」写法，避免把生产 SES
内置或 GM/T 0099 核心支持写成默认承诺。

## 历史验证记录与本次核对

下列构建、样本和 11 包数字保留原验证快照，不能推定为 R00 的 `6aa8183` 上重新执行。各能力后续测试/页面的提交、产物和检查范围以 ticket 索引所链接的原证据为准；失败和未完成记录保持不变。

- Release 构建：解决方案全部项目成功。
- 单元/集成测试：以两页非隐私样例逐字核对 OpenXML 原文与 Native/DualLayer OFD 抽取结果；另覆盖真实 LibreOffice、BuiltIn、自动兜底和外部资源拒绝。
- NuGet：11 个本地包成功生成，并由独立消费项目在不安装 Word/LibreOffice 的环境覆盖 BuiltIn DOCX、PDF、SVG 和签名转换。
- 上游样本：`999.ofd` 的 5 页、图层、模板、路径、文字对象、注释保留、
  SVG 视觉输出和 20/20 个 SM3 保护引用通过验证。
- PDF 互操作：3 个上游 PDF 样本完成 PDF → OFD → PDF 视觉冒烟。
- 依赖健康（2026-10-15 核对）：[PDF 项目](../src/Ofdrw.Net.Converter.Pdf/Ofdrw.Net.Converter.Pdf.csproj)与[DOCX 项目](../src/Ofdrw.Net.Converter.Docx/Ofdrw.Net.Converter.Docx.csproj)仍固定 `SixLabors.ImageSharp 2.1.13`；维护者 [GHSA-wmxv-xphr-5c9g](https://github.com/SixLabors/ImageSharp/security/advisories/GHSA-wmxv-xphr-5c9g)（2026-09-15 发布，Moderate，CVSS 5.3）列出的受影响区间包含该版本。图片导入在 Identify 前只准入 PNG/JPEG；DOCX、OFD 与签章位图在 Identify/解码前拒绝 BigTIFF 的两种字节序，保留经典 TIFF 等既有格式。此边界修复不代表全部 ImageSharp 风险已消除；未做全量传递依赖/deprecated 审计，版本、目标框架与许可未变。

详细行为与兼容边界见 [转换、编辑与资源约定](conversion-contracts.md)。

issue #3 的逐项状态见 [实现对照](issue-3-implementation.md)，不因 04 单票而宣称整个 issue 完成。

## 合并、视觉和发布边界

| 状态层 | 本次可以确认的内容 | 未执行的门 |
| --- | --- | --- |
| 本地源码 | `6aa8183` 的公开实现；R00 定向运行纯文本 35/35、HTML 29/29，见[本轮契约](capability-contract-v12.md) | 未重跑全套回归，不扩大有限 profile |
| 已合并 | R00 起始时 `origin/main=6aa8183`；旧能力按已有历史记录和本轮源码判定 | R00 分支在本表编写时尚未合并；后续票不凭旧 open/unmerged 文案推断状态 |
| 实际视觉 | 历史记录只证明当时具名产物与页码；R00 [缺陷样本](evidence/r00-pageblock/README.md) 新 PDF 第 1 页已用 PNG/Preview 确认漏绘 | 此次检查是负面证据，不是其它文档的视觉通过 |
| 已发布版本 | 本地消费包只证明该次包消费；[preview.9 页面接受记录](release-preview-acceptance.json)在其固定源码指纹及具名页面上为 `accepted` | R00 未查询新的 NuGet 版本、未打包/发布；历史接受记录不延伸到新产物 |

## 11 票：独立口令扩展

`Ofdrw.Net.Crypto.Password`（net8.0）提供本包 profile v2 的随机盐/有界 scrypt、
独立 IV 的 SM4-CBC/PKCS#7 与先验 HMAC、显式部分页/条目加密及精确恢复；拒绝v1。
核心 `SupportsGmT0099EncryptionEnvelope=false` 保持不变，不进入默认 Converter。
私有 v2 仍仅自闭环；显式 R14 入口支持固定 Java 2.4.0 的口令或 SM2 证书同类多接收人、
签名清单及本包逐文件摘要扩展，仍不支持混合接收人或多层加密。无厂商互认、非商用密码认证、非长期保存格式。
契约与限制见[可选加密扩展](password-crypto.md)，
功能/页面/review 状态见[本票证据](evidence/password-crypto/README.md)。
