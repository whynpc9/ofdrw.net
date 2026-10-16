# 能力路线图 tickets

来源是 [能力路线图](../../../docs/capability-roadmap.md)。一票一文件；本索引是**当前状态的唯一领票入口**，文件编号不是开工顺序。核对基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15。

票文件和[历史开发计划](../../../docs/capability-roadmap-development-plan.md)内的旧 `Status`、PR open/unmerged、阻塞和验收勾选保留为当时快照；与本索引不一致时以本索引的固定源码证据复核。`implemented-source` 表示已有实现及测试入口，禁止从零重派；不表示所有原票目标、回归、视觉或发布门关闭。`ready-for-agent` 表示未实现且源码前置可用，只是可按需求领取。原票验收条件继续有效。

## 当前开工入口

1. 已实现项按原证据的限制、失败记录和未完成门开展修复/补验，不重新开发同一 API。
2. 18（纯文本）、17（占位）、20（Canvas）、22（骑缝外观）的源码前置已具备，但自身未实现；18 仍须交付 UTF-8/换行/参数/API/CLI 薄适配，不再等待从零开发 Flow。
3. 07、10、13 按调用方需求决定；08 HTML、14 OFD-A 保持未来方向，P2-06 尚未建票。15 现有工程门需要对实际候选执行，不能从检查代码存在推断发布通过。

## 索引与本地证据

链接中的测试只表明基线有相应断言，本次未运行。未实现结论由相关现有入口及 `src/`/CLI 静态检索核对，复跑记录见[本次验证](../../../artifacts/plan-026/README.md)。

| 票 | 优先级 / 路线图 | 当前源码状态与剩余范围 | 源码、测试及历史验收入口 |
| --- | --- | --- | --- |
| [01 公开流式布局](01-public-flow-layout.md) | P0 / P0-01 | implemented-source；公开段落/Span、分页已有 | [FlowDocument](../../../src/Ofdrw.Net.Layout/FlowDocument.cs)、[共享段落核心](../../../src/Ofdrw.Net.Layout/Internal/Flow/FlowParagraphLayout.cs)、[测试](../../../tests/Ofdrw.Net.Layout.Tests/FlowDocumentTests.cs)、[历史页面记录](../../../docs/validation/flow-layout-2026-09-29.md) |
| [16 公开表格](16-public-tables.md) | P0 / P0-02 | implemented-source；固定列宽、水平合并、整行分页；纵向合并拒绝 | [Table/Row/Cell](../../../src/Ofdrw.Net.Layout/Table.cs)、[测试](../../../tests/Ofdrw.Net.Layout.Tests/TableTests.cs)、[历史验收](../../../docs/validation/public-tables-2026-10-02.md) |
| [17 区域占位](17-area-holders.md) | P0 / P0-03 | ready-for-agent；未实现；01 源码前置已具备 | 现有 [FlowBlock](../../../src/Ofdrw.Net.Layout/FlowDocument.cs) 仅 Paragraph/Table，无命名占位与回填 API；原票定义验收 |
| [02 图片进出](02-ofd-image-io.md) | P0 / P0-04、P0-05 | implemented-source；保留 PDF/PDFium 预览保真边界 | [导出](../../../src/Ofdrw.Net.Converter.Pdf/ImageIO/OfdToImageConverter.cs)、[导入](../../../src/Ofdrw.Net.Converter.Pdf/ImageIO/ImageToOfdConverter.cs)、[测试](../../../tests/Ofdrw.Net.Converter.Pdf.Tests/ImageIoTests.cs)、[历史集成/页面记录](../../../docs/validation/issue02/README.md) |
| [03 文档工具](03-document-tools.md) | P0 / P0-06–P0-09 | implemented-source；水印、Split、Mix、清签已有；未知引用保守拒绝 | [水印](../../../src/Ofdrw.Net.Layout/Editing/OfdWatermark.cs)、[拆分](../../../src/Ofdrw.Net.Layout/Editing/OfdDocumentSplitter.cs)、[混合](../../../src/Ofdrw.Net.Layout/Editing/OfdDocumentMixer.cs)、[清签](../../../src/Ofdrw.Net.Layout/Editing/OfdSignatureCleaner.cs)、[测试](../../../tests/Ofdrw.Net.Packaging.Tests/DocumentToolTests.cs)、[历史证据](../../../docs/evidence/document-tools/README.md) |
| [04 OfdGraphics](04-ofd-graphics-api.md) | P1 / P1-01 | implemented-source；原生毫米图元、状态与受限裁剪 | [API](../../../src/Ofdrw.Net.Layout/Graphics/OfdGraphics.cs)、[测试](../../../tests/Ofdrw.Net.Layout.Tests/GraphicsContractTests.cs)、[历史证据](../../../docs/evidence/graphics/README.md) |
| [05 字体子集与复用](05-font-subset-and-reuse.md) | P1 / P1-02 | implemented-source；TrueType/glyf profile，非任意字体子集 | [子集实现](../../../src/Ofdrw.Net.Core/Fonts/TrueTypeSubset.cs)、[内容复用](../../../src/Ofdrw.Net.Packaging/OfdResourceCatalog.cs)、[测试](../../../tests/Ofdrw.Net.Packaging.Tests/FontSubsetTests.cs)、[契约](../../../docs/font-subset-and-reuse.md)、[历史证据](../../../docs/evidence/font-subset/README.md) |
| [20 布局 Canvas](20-layout-canvas-drawcontext.md) | P1 / P1-03 | ready-for-agent；未实现；01/04 源码前置已具备 | 现有 [FlowBlock](../../../src/Ofdrw.Net.Layout/FlowDocument.cs) 无 CanvasBlock；可复用 [Graphics](../../../src/Ofdrw.Net.Layout/Graphics/OfdGraphics.cs)，原票验收仍需完成 |
| [21 PDF 矢量](21-pdf-to-ofd-vectors.md) | P1 / P1-04 | implemented-source；受限可选 profile，非任意 PDF；04/19 已有源码 | [转换器](../../../src/Ofdrw.Net.Converter.Pdf.Vector/PdfVectorToOfdConverter.cs)、[测试](../../../tests/Ofdrw.Net.Converter.Pdf.Vector.Tests/VectorTests.cs)、[历史失败/修复/页面证据](../../../docs/evidence/pdf-vectors/README.md) |
| [06 定位并替换](06-keyword-locate-and-replace.md) | P1 / P1-05、P1-06 | implemented-source；严格几何与原字体证据，有限原位替换 | [定位](../../../src/Ofdrw.Net.Reader/Keywords/OfdKeywordLocator.cs)、[替换](../../../src/Ofdrw.Net.Layout/Editing/OfdTextReplacer.cs)、[测试](../../../tests/Ofdrw.Net.Packaging.Tests/TextReplacementOwnedSessionTests.cs)、[历史证据](../../../docs/evidence/issue06-owned-save/README.md) |
| [07 附件增删](07-attachment-add-remove.md) | P1 / P1-07 | ready-for-agent；现有包独立增删服务未实现 | 已有 [附件模型](../../../src/Ofdrw.Net.Core/Models/OfdAttachment.cs) 和 [Writer](../../../src/Ofdrw.Net.Packaging/OfdPackageWriter.cs) 附件写入，不等于原票的安全编辑 API |
| [18 纯文本 → OFD](18-plain-text-to-ofd.md) | P1 / P1-08 | ready-for-agent；薄适配器未实现，01 源码前置已具备 | 已有 [FlowDocument](../../../src/Ofdrw.Net.Layout/FlowDocument.cs) 与[测试](../../../tests/Ofdrw.Net.Layout.Tests/FlowDocumentTests.cs)；[CLI](../../../src/Ofdrw.Net.Cli/Program.cs) 无纯文本转换入口 |
| [08 OFD → HTML](08-ofd-to-html.md) | P1 / P1-09 | future / 未实现；不默认开工 | 现有 [SVG 转换器](../../../src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs) 是可复用前置，不等于 HTML API/CLI |
| [09 强类型对象](09-typed-object-model.md) | P1 / P1-10 | implemented-source（最小范围）；未知 XML 所有权与不支持闭包继续拒绝 | [注释模型](../../../src/Ofdrw.Net.Core/Models/OfdAnnotation.cs)、[裁剪](../../../src/Ofdrw.Net.Core/Models/OfdClips.cs)、[组合](../../../src/Ofdrw.Net.Core/Models/OfdComposite.cs)、[测试](../../../tests/Ofdrw.Net.Packaging.Tests/TypedClipCompositeTests.cs)、[历史证据](../../../docs/evidence/issue09-typed-objects/README.md) |
| [10 多 DocBody](10-multi-docbody.md) | P1 / P1-11 | ready-for-agent；完整多文档选择/重写未实现 | [Reader](../../../src/Ofdrw.Net.Reader/Readers/OfdReader.cs)、[Writer](../../../src/Ofdrw.Net.Packaging/OfdPackageWriter.cs) 仍以首文档路径为主，不能从未知条目保留推定多文档支持 |
| [19 SkiaSharp 适配](19-skiasharp-graphics-adapter.md) | P1 / P1-12 | implemented-source；合作 producer 事件，不拦截任意 SKCanvas | [适配器](../../../src/Ofdrw.Net.Graphics.SkiaSharp/OfdSkiaAdapter.cs)、[测试](../../../tests/Ofdrw.Net.Graphics.SkiaSharp.Tests/AdapterTests.cs)、[历史范围/页面记录](19-skiasharp-graphics-adapter.md) |
| [11 口令加解密](11-gmt0099-password-crypto.md) | P2 / P2-01 | implemented-source；可选私有 password profile v2，拒绝 v1；非通用 GM/T 互操作 | [Envelope](../../../src/Ofdrw.Net.Crypto.Password/OfdPasswordEnvelope.cs)、[v2 测试](../../../tests/Ofdrw.Net.Crypto.Password.Tests/PasswordV2SecurityTests.cs)、[历史 v2 证据](../../../docs/evidence/password-crypto/v2/README.md)；早期 v1 证据不提升为 v2 通过 |
| [12 SES 自签](12-ses-self-sign-interop.md) | P2 / P2-02 | implemented-source；仅可选 dev/interop，默认不提升 FullyValid | [显式测试验证器](../../../src/Ofdrw.Net.Signatures.SesInterop/SesTestSignedValueVerifier.cs)、[测试](../../../tests/Ofdrw.Net.Signatures.SesInterop.Tests/SesTests.cs)、[历史证据](../../../docs/evidence/ses-interop/README.md) |
| [22 骑缝外观](22-riding-seal-appearance.md) | P2 / P2-03 | ready-for-agent；未实现；06 源码前置已具备；外观不等于签名 | 现有[定位](../../../src/Ofdrw.Net.Reader/Keywords/OfdKeywordLocator.cs)与[图片水印](../../../src/Ofdrw.Net.Layout/Editing/OfdWatermark.cs)可复用，尚无骑缝切片/定位服务 |
| [13 锁定签名](13-lock-and-continue-sign.md) | P2 / P2-04 | ready-for-agent；约定保护范围/继续签模式未实现 | 现有[签名服务](../../../src/Ofdrw.Net.Signatures/Signing/OfdSignatureService.cs)和[选项](../../../src/Ofdrw.Net.Signatures/Signing/OfdSignatureOptions.cs)不等于原票保护范围；非首 DocBody 必须与 10 协调 |
| [14 OFD-A 检查器](14-ofd-a-checker.md) | P2 / P2-05 | future / 未实现；不宣称档案认证 | [现有结构检查](../../../src/Ofdrw.Net.Reader/Readers/OfdReader.cs)不是 GB/T 42133 规则引擎，原票仅定义未来范围 |
| [15 发布闸门](15-release-engineering-gates.md) | Q / Q-01–Q-05 | partial-source；检查代码/模板已有，候选验收未完成 | [API 文档检查](../../../scripts/check-public-api-docs.py)、[门禁测试](../../../scripts/tests/test_release_gates.py)、[候选记录](../../../docs/release-preview-acceptance.json)仍为 `not-reviewed`；不能列为从零开发或全绿 |

## 四种状态分别记录

| 状态层 | 本次核对结果 |
| --- | --- |
| 本地源码能力 | 上表固定基线的实现及测试入口已核对；没有执行 .NET 功能测试，已有有限范围不变 |
| 已合并 | 本次未查询实时 GitHub；既有[集成快照](../../../docs/evidence/roadmap-merge-round-2026-10-15/README.md)及各票旧 open/unmerged/merge 描述只保留原时点含义，不推断当前远端状态 |
| 实际视觉验收 | 只引用原记录中的具名样例、哈希和页码；本次未生成页面、未做自动渲染/PNG 复查或 macOS Preview，不宣称 `d0f88c7` 新视觉通过 |
| 已发布版本 | 未查询 NuGet、未打包/发布；本地包消费记录不证明公开发布；候选 `not-reviewed` 原样保留 |

生产 SES/信任、证书来源、OCSP、合格 TSA 和 P2-06 不因这些可选包实现进入默认排期。当前剩余实现与质量门应分别领取，源码前置已具备不自动解除验收条件。
