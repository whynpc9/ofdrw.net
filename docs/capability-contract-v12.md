# Ofdrw.Net v1.2 能力契约与路线基线（R00）

日期：2026-10-21。核对源码基线：`6aa8183ea31e193f500fb9d9717677b15e86c5d3`。研究报告 v1.2 的 .NET 基线是较早的 `e9b237bc85e4dee06c1f690bea48553a1f5a22e2`；报告中的候选缺陷和建议不视为已复现或已验收。此契约面向**当前源码**，不是已发布 NuGet 的能力声明。后续票合并时按新 head、实际出口和样本更新对应行。

## 设计与判定

能力的键是「入口及出口、输入语义、策略」的组合。`supported` 仅指本行列出的有限实现和已有测试入口；仍须按本行的视觉、互操作、包消费门分别验收。`unsupported` 表示源码明确拒绝或没有该入口；没有实现不能凭格式名推断会正确处理。`unverified` 表示有模型、保留或静态候选，但尚无当前样本证明导出或互操作。保留 XML、可编辑、可呈现、可跨文档迁移是四个独立结论。历史失败照原记录保留。

失败出口的共同规则：文件型 CLI 以目标旁临时文件和 `AtomicOutput` 提交，异常或取消不得把临时产物报告为成功；调用者传入的任意 `Stream` **不保证回滚**，复制或 flush 失败可能留下前缀，调用者必须丢弃输出。未来新入口应预检目标出口能力、在最终提交前完成转换；显式 fallback 必须写进结果/诊断并以整页或明确单元替换，不能把失败前的半成品混入成功结果。该规则不把现有所有流 API 宣称为原子写入。

### 已有出口与输入语义

| 入口 → 出口、输入 | 状态与实际边界 | 默认/显式 fallback；失败输出 | 源码与验证入口 |
| --- | --- | --- | --- |
| DOCX → Native OFD；常见 OpenXML 段落、表格、图片、页眉页脚与原文 | `supported` 有限布局；默认模式为 Native。复杂 Word 浮动、域、跨页表格不保证保真；不支持内容依 `UnsupportedFeatureBehavior` 诊断或失败 | 默认不会悄悄改走 DualLayer。嵌入常规面加粗默认 `TextStroke` 并报告 `DOCX_EMBEDDED_BOLD_SIMULATED`；可设 `Throw`。CLI 原子提交，流输出按上段处理 | [`DocxConversionOptions`](../src/Ofdrw.Net.Converter.Docx/DocxConversionOptions.cs)、[`DocxToOfdConverter`](../src/Ofdrw.Net.Converter.Docx/Converters/DocxToOfdConverter.cs)、[`转换约定`](conversion-contracts.md) |
| DOCX → DualLayer OFD；同源 OpenXML 原文 | `supported` 可选；页面图依赖 Word/LibreOffice/BuiltIn 渲染，文字语义从 OpenXML 映射，未定位正文且选页时拒绝 | 渲染引擎按配置回退；视觉页必须栅格化，不能用 PDF 抽字冒充原文。`ConvertWithResultAsync` 报实际引擎、诊断和页映射 | 同上；[`DocxToPdfConverter`](../src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs) |
| PDF → OFD，默认 `PdfToOfdConverter` | `supported` 页面图加受限透明文字；扫描件无 OCR，不能保证阅读顺序 | 尝试配置的页面栅格器；失败明确抛出，不写空白成功页。文件 CLI 原子提交 | [`PdfToOfdConverter`](../src/Ofdrw.Net.Converter.Pdf/Converters/PdfToOfdConverter.cs)、[`PdfToOfdOptions`](../src/Ofdrw.Net.Converter.Pdf/PdfToOfdOptions.cs) |
| PDF → 原生矢量 OFD，可选包 | `supported` 仅不透明路径和固定字体 Unicode 文字的有限 profile；clip、复杂 XObject/字体/透明度等逐项看拒绝路径 | `UnsupportedPagePolicy=Fail` 默认；显式 `RasterizePage` 才整页图像/双层回退。预算、取消及不可表示页不转成静默回退；结果逐页标注 native/raster | [`PdfVectorToOfdOptions`](../src/Ofdrw.Net.Converter.Pdf.Vector/PdfVectorToOfdOptions.cs)、[`PdfVectorToOfdConverter`](../src/Ofdrw.Net.Converter.Pdf.Vector/PdfVectorToOfdConverter.cs)、[`教程`](tutorials/18-pdf-vector-mode.md) |
| OFD → PDF；Reader 投影的页面/模板、文字、路径、图片 | `supported` 有限预览；普通 `PageBlock` 包含路径时已复现**成功导出但漏绘**；文字 DrawParam/颜色空间等仍待逐项最小样本复现。CGTransform 明确拒绝 | 无通用自动栅格 fallback；不支持语义应失败。API `Stream` 可能留下前缀，不将部分 PDF 当成功 | [`OfdToPdfConverter`](../src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs)、[`Reader`](../src/Ofdrw.Net.Reader/Readers/OfdReader.cs) |
| OFD → 独立 SVG；同类图元 | `supported` 有限预览；普通 `PageBlock` 路径已复现**成功导出但漏绘**；其它文字/颜色语义仍 `unverified`；CGTransform 明确拒绝 | 无通用自动 fallback；目标流失败语义同上 | [`OfdToSvgConverter`](../src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs) |
| OFD → 静态 HTML；选中页的内嵌 SVG | `supported` 的源码及测试入口，有限安全预览；与独立 SVG **分别验收**。不执行动作、脚本或外部资源；不支持的 raw/绘图语义由 HTML 预检拒绝 | 生成至有界 spool 后复制；生成失败不触碰公共流，最终复制/flush 失败公共流可留前缀；CLI 原子提交 | [`OfdToHtmlConverter`](../src/Ofdrw.Net.Converter.Svg/Converters/OfdToHtmlConverter.cs)、[`HtmlRenderContext`](../src/Ofdrw.Net.Converter.Svg/Internal/HtmlRenderContext.cs)、[`HtmlPreviewTests`](../tests/Ofdrw.Net.Converter.Pdf.Tests/HtmlPreviewTests.cs) |
| OFD → PNG/JPEG；单页预览 | `supported` 的有限图像导出，复用 PDF 预览路径；继承 PDF 出口已确认的普通 `PageBlock` 漏绘风险，不是独立证明外来 OFD 所有图元正确 | 不支持或超预算明确失败；CLI 文件提交仍按原子约定 | [`OfdToImageConverter`](../src/Ofdrw.Net.Converter.Pdf/ImageIO/OfdToImageConverter.cs)、[`历史验收`](validation/issue02/README.md) |
| PNG/JPEG → OFD；一图一页 | `supported` 仅白名单输入；尺寸/像素预算与居中几何有约束，非通用图片格式导入 | 非白名单或超限拒绝；不默默转空白页 | [`ImageToOfdConverter`](../src/Ofdrw.Net.Converter.Pdf/ImageIO/ImageToOfdConverter.cs) |
| UTF-8 纯文本 → OFD；公开 Flow | `supported` 首版严格 UTF-8、有限预算与分页；没有通用字体嵌入或复杂脚本 shaping 承诺 | 不合法输入、溢出或超限拒绝；CLI 原子提交。已有 Flow 度量不等于真实 face 测量 | [`PlainTextOfdConverter`](../src/Ofdrw.Net.Layout/PlainTextOfdConverter.cs)、[`PlainTextOfdTests`](../tests/Ofdrw.Net.Layout.Tests/PlainTextOfdTests.cs) |
| 已有 OFD → 保存/编辑/合并；未知源对象 | `supported` 仅各 API 约定的对象/资源闭包；未知 XML 保留不等于可导出或可迁移 | 合并默认拒绝未能安全映射的 raw；仅显式 `SkipUnsupportedRawElements` 可跳过，必须读取结果诊断。签名改写失效另行报告 | [`转换约定`](conversion-contracts.md)、[`功能对照`](feature-parity.md) |
| 外来 OFD 显式 `CGTransform` → PDF/SVG/HTML | 保存源 XML：`supported`；呈现导出：`unsupported`。不能用保存成功声称视觉成功 | 出口拒绝；无自动文字代换/光栅 fallback。R02 逐后端放行前保留护栏 | [`PDF`](../src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs)、[`SVG`](../src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs) |
| SES/SM2、GM/T 0099、归档 | 摘要引用完整性与可插拔 provider 已有；SES 测试容器和 password v2 是显式可选包；通用标准加密/生产 PKI/档案认证 `unsupported` 或 `unverified` | 自签、摘要匹配、外观可见均不自动升级为生产 `FullyValid` 或认证 | [`功能对照`](feature-parity.md)、[`能力路线`](capability-roadmap.md) |

### 本轮 23 项的交付及验收矩阵

下表是**开工基线**，不是进度勾选。`limited` 表示有相关前置或局部实现；`candidate` 表示报告提出的静态候选，等待最小 fixture；`new` 表示目标能力尚未验收。每票要在自己的源码 head 上记录输入哈希、输出哈希、正反样本、功能测试、PNG/Preview、包消费及外部互操作中实际适用的门。R00 的文档核对不会替其它票关闭这些门。

| 票 | 当前判定及后续验收焦点 | 放行前失败/fallback 边界 |
| --- | --- | --- |
| R00 契约 | 本页与旧文档状态对齐；核对基线 SHA、源码入口、历史视觉记录 | 记录流与 CLI 差异，不把旧验收当本次产物 |
| R01 PageBlock/文字/颜色 | 普通页嵌套 Path 漏绘已复现；模板/Text/Image 及其语义仍待样本，各出口独立视觉与文本语义 | 未建模图元/颜色须拒绝，不能静默漏绘 |
| R02 CGTransform | 源 XML 保留 `limited`，PDF/SVG 呈现 `unsupported`；GID、原文、Delta 与实际字体交叉验收 | 各后端保留拒绝直到各自通过 |
| R03 命名区域 | `new`；持久区域名、页引用，重开后文字/图片回填及区域外几何 | 重名、丢失、溢出明确失败 |
| R04 流式块/回调/Span | Flow/Graphics 前置 `limited`；图片/Canvas 测量分页、最终回调、下划线/字距/字重 | 异常不提交假成功，旧 Flow 兼容模式保留 |
| R05 复杂 Merge/Mix | 常见图元 `limited`；共享/嵌套资源、Clip/Composite/批注 ID 与 ownership | 未闭合引用默认拒绝，跳过只可显式选择并报告 |
| R06 PDF 矢量 | 有限 profile `limited`；Rotation/CropBox、clip、画笔、XObject、字体逐特性正反样本 | 默认 Fail；显式回退整页，不能保留该页半成品矢量 |
| R07 PDF 导航/附件 | 导入迁移 `new`；层级页映射、全局/页附件名与哈希 | 关闭迁移无残留；非法目标/冲突明确诊断 |
| R08 字体后端 | TTF 子集、字体身份 `limited`；按保存/SVG/PDF/轮廓四出口核查非 BMP、symbol、Type1/CFF | 后端未验前不删现有拒绝；PR #25 已在基线 |
| R09 绘图上下文 | Path/Text/状态栈 `limited`；DrawImage、源裁切、dash/cap/join、渐变逐子集 | 不支持 brush/effect 拒绝或明确回退 |
| R10 专用图像/区域提取 | JBIG2/ImageMask `new`；字符级 CTM/Delta 矩形筛选 `candidate` | 超预算/未知解码或几何语义显式失败/Unknown |
| R11 签章外观 | 外来外观显示 `limited`；新生成普通章/骑缝章各页和签后重读 | 视觉外观不等于签名值或信任结论 |
| R12 SES V5/外部签名 | provider 与 V1/V4 测试容器 `limited`；显式 V5 与 GB/T 35275 单签名者 profile 已与固定版本 Java OFDRW 双向验证，仍为 `limited` | 错证书/容器拒绝；PKI/TSA/法律效力不推定，默认验签仍 `Unsupported`；完整标准符合性和真实设备互操作未验证 |
| R13 签名范围 | 默认引用策略 `limited`；Relative/Parameters、续签与漏保护文件 | 默认保守，显式策略不得无签名值假锁定 |
| R14 标准加密 | 私有 password v2 `limited`；标准 profile、多 UserInfo、OFDEntries 与交叉样本 | v2 保留原义；错误口令/额外条目拒绝 |
| R15 多 DocBody | Reader 枚举和显式选择、PDF/SVG/HTML/图像导出；隔离无签名目录下的选中文档编辑 `limited`，详见[多文档契约](multi-docbody.md) | 默认仍首文档；未选条目与 DocBody 保留；签名、共享资源和跨文档引用的编辑拒绝，不猜映射或自动清签 |
| R16 强类型对象 | 少量 typed `limited`；DocInfo/权限/版本/动作读改写 | 未知节点保留；权限声明不等于执行策略 |
| R17 归档 Rule/Handler | 结构预检 `limited`；具名条款、Pass/Fail/Unknown、独立副本 | Unknown 不升格 Pass，修复后重检 |
| R18 附件清理 | 附件读写及部分编辑 `limited`；逻辑/物理移除、引用闭包 | 无法证明无引用则不物理删除 |
| R19 跨页复杂表格 | 整行分页 `limited`；RowSpan、拆行、重复表头、图片单元格 | 超出有限规则明确失败，逐页视觉验收 |
| R20 完整归档规则 | `new`；逐条标准版本、条款、fixture 与独立核准 | 缺规范保持 Unknown/阻断，不声称认证 |
| R21 真实字体测量 | 新查询 `limited`：固定静态 TrueType face、HarfBuzz 单 run GID/cluster/advance/ink、有限 Latin/Han 预排版及严格 ASCII OFD 出口；旧 Flow 仍用原度量 | 缺 native/字体或超出已验 profile 明确失败；混合字体、段落 bidi 和通用 GID 输出留待 R22，详见 [R21 profile](text-prelayout.md) |
| R22 HarfBuzz 闭环 | `new`；单 run → OFD GID/原文 → 段落 bidi/断行 → PDF/SVG 分别视觉与语义 | 各阶段各出口单独放行；缺平台资产/规范/独立阅读器仍未验收 |

## 证据与门禁

本次 R00 不修改产品代码；生成的页面仅用于下述缺陷复现。定向功能测试：`PlainTextOfdTests` 35/35、`HtmlPreviewTests` 29/29（.NET 10.0.401、Release、单节点、禁用 build servers，共同源码基线如页首）。HTML 首次 28/29：LFS 未物化的 `style-metrics.ttf` 被当成字体输入；仅获取该文件后复跑为 29/29。这是测试前置条件，不记录成产品缺陷已修复。NuGet restore 对现有 ImageSharp 2.1.13 输出 `NU1902/NU1903` 依赖告警，本票未升级依赖。文档链接检查和 `git diff --check` 通过。普通 `PageBlock` 最小路径样本复现 PDF/SVG 漏绘、HTML 明确拒绝；新 PDF 第 1 页的 PNG 与 macOS Preview 均确认蓝色矩形缺失，已关窗，见[复现与哈希](evidence/r00-pageblock/README.md)。这是**视觉缺陷证据**，不是视觉通过。**本次仅对上述缺陷样本完成自动渲染、PNG 与 Preview 复查；没有完成其它功能的视觉或包消费验收**。已有 [`preview.9 接受记录`](release-preview-acceptance.json) 仅证明其列出的 `generated-layout.docx`、源码指纹、两种 OFD/PDF 哈希及各 1–2 页在 2026-10-18 的实际 Preview 查看；不得泛化到所有 DOCX 或当前未生成的 OFD。[`Issue 02 历史记录`](validation/issue02/README.md) 中早期斜体失败及后续修复/复验仍按各自基线保留，不能删除失败来“改善”总状态。

每个后续视觉票按仓库 `AGENTS.md` 留存：源码 SHA、fixture 及 SHA-256、模式、OFD/PDF/PNG 哈希、功能和包消费结果、全部受影响页 PNG 检查范围，以及实际 `DOCX → native/default OFD → PDF → macOS Preview` 或其它准确链路、页码、发现与关窗记录。独立阅读器未检查时标 `unverified`。任何新产物重生成后重开 Preview；视觉缺陷未解决时不得标整体通过。R00 不改调度文件，不授权 tag、NuGet 或 Release。
