# Ofdrw.Net 能力路线图开发计划

> 计划基线：2026-09-28；评审修订：2026-09-29。本文件是设计与实施安排，不表示任何能力票已经完成。需求边界以 [22 张 issue 票](../.scratch/capability-roadmap/issues/README.md) 和 [能力路线图](capability-roadmap.md) 为准；现状以本次核对的源码为准。

## 1. 基线、目标与证据

- 本仓库基线：`ae28394`（合入能力票文档的提交）。编写本计划前，工作区已有 `README.md`、三个测试文件和 `docs/test-coverage.md` 的未提交改动；本计划新增本文，并在 `docs/capability-roadmap.md` 添加回链。这些既有改动不作为本计划的已实现能力证据。
- 上游基线：[ofdrw/ofdrw `5fe9c4276c64e40b455e6ea649b695adf8a9a734`](https://github.com/ofdrw/ofdrw/tree/5fe9c4276c64e40b455e6ea649b695adf8a9a734)。比较的是这一份 Java 源码快照，不把上游 API 名称、测试存在或 README 描述直接算作 .NET 已实现/互操作已通过。
- 目标：按依赖交付 P0 日常生成与工具，随后补 P1 绘图、字体和保真能力；P2 只交付明确隔离的测试/检查扩展；Q-01～Q-05 成为发布前的可重复闸门。
- 每票完成必须同时满足该票的验收条款、公开 API/CLI 文档、功能对照更新，以及相应的测试和实际页面检查。`Status: blocked` 只是当前依赖状态，不能按编号直接开工；实际开工前复核依赖票已合入的实现和回归证据。

### 1.1 源码核对得到的关键差异

| 主题 | 上游实现 | 当前 .NET 实现 | 对计划的影响 |
| --- | --- | --- | --- |
| 流式排版 | `ofdrw-layout` 的 `Paragraph`/`Span`、`StreamingLayoutAnalyzer`、`ParagraphRender` 将段拆分、定位和渲染分层 | `BuiltInOfdRenderer` 内有折行、行内样式、表格、分页，但类型是 DOCX 转换内部模型；公开 `OfdDocumentBuilder` 仍按页放图元 | 01 先提取与 DOCX 无关的测量/分页核心，再公开 API；Native DOCX 回接同一核心，不能只包一层 DTO |
| 表格 | 上游 `element.canvas.Cell` 是固定尺寸 Canvas 单元，不是完整的行列/跨页表格引擎 | BuiltIn DOCX 路径已有列宽、`GridSpan` 水平合并、行测量和跨页处理；未映射 `vMerge`/`RowSpan` | 16 要设计公开 `Table/Row/Cell` 与确定性跨页策略，明确纵向合并的首版行为；不能声称直接移植上游表格引擎 |
| 字体 | `ofdrw-layout/engine/ResManager` 做资源缓存/注册；本次快照未发现通用 TTF/OTF 子集化写出链路 | `OfdPackageWriter.BuildFonts` 登记字体，`OfdResourceCatalog` 按内容哈希命名载荷；仍可能嵌入整份字体 | 05 的子集化是新研发任务；先做格式/许可/字形映射可行性验证，再接写包与渲染回归 |
| 多文档 | 上游 `OFD.getDocBodies()`、`getDocBody(int)` 和 Reader 的 `numOfDoc` 路径可选择文档 | `OfdReader` 取首个 `DocRoot`；`OfdPackageWriter` 只重建一个 `DocBody` | 10 要调整数据模型和读写接口，防止保存时静默丢掉后续 DocBody |
| PDF→OFD | 上游 `PDFConverter` 通过 PDFBox `renderPageToGraphics` 把绘制事件送到 `OFDPageGraphics2D` | 当前 `PdfToOfdConverter` 默认页面 PNG + 透明文本；PdfPig 提供页数、尺寸与文本语义 | 21 需要可取得 PDF 路径/文字绘制信息的桥接验证；仅凭现有 PdfPig 使用方式不足以承诺完整矢量模式，先做样例探针 |
| 签名/保护 | 上游 `OFDSigner` 构建保护引用，`SignCleaner` 遍历所有 DocBody 清签 | .NET 有签名值提供者、引用摘要验证、重写后清理失效签名，但没有独立“清签”命令 | 03/13 共享包级引用与清理逻辑；“摘要仍匹配”与 `FullyValid` 必须分别报告 |

上游 `OFDSigner.toBeDigestFileList()` 的继续签模式只特别跳过 `Signatures.xml`，并允许调用方提供保护文件过滤器；它并未直接实现 13 票要求的“排除全部约定 `Annots`/`Signs` 载荷”。因此 13 需要自行定义目录边界和回归矩阵。上游 `layout/edit/Attachment` 主要是附件对象构造，不能把它当成完整的现有包删除实现。

源文件入口：[上游布局](https://github.com/ofdrw/ofdrw/tree/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-layout/src/main/java/org/ofdrw/layout)、[上游转换](https://github.com/ofdrw/ofdrw/tree/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-converter/src/main/java/org/ofdrw/converter)、[上游签章](https://github.com/ofdrw/ofdrw/tree/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-sign/src/main/java/org/ofdrw/sign)、[上游档案检查](https://github.com/ofdrw/ofdrw/tree/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-archive/src/main/java/org/ofdrw/archive/check)。本仓库入口：`src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs`、`src/Ofdrw.Net.Layout/Builders/OfdDocumentBuilder.cs`、`src/Ofdrw.Net.Reader/Readers/OfdReader.cs`、`src/Ofdrw.Net.Packaging/OfdPackageWriter.cs`、`src/Ofdrw.Net.Packaging/OfdPackagePruner.cs`。

## 2. 架构约束与共用基础

1. **模型边界。** `Core` 放稳定的文档/图元/资源值类型；`Layout` 负责测量、分页、绘图编排和编辑；`Packaging` 负责 OFD 路径、ID、资源引用与 ZIP 写入；`Reader` 负责读取/定位；`Converter.*` 只适配格式；`Cli` 调用公开服务，不复制格式逻辑。新可选密码/interop 包保持独立，不加入默认 `Ofdrw.Net.Converter` 元包。
2. **保留未知结构。** 现在的 `PreservedEntries`、`SourceXml`、`Preserved*` 是往返兼容基础。改写一个节点时明确其所有权及其引用闭包；未知 XML 不得因强类型化、拆分或混合而静默丢失。所有新包路径经既有内部 `OfdPackagePath` 规范化；包级编辑对输入、输出和错误采用原子写入。
3. **坐标与页面。** 公开 API 统一以毫米表达几何；CLI 页号维持一基，内部 `OfdPage.Index` 零基，所有入口集中转换。布局测量与渲染用同一字体选择和字形宽度策略，禁止先排一版、落盘另算一版。
4. **变更与签名。** 默认“重写包导致既有签名失效”的安全行为保持。03/06/07/09/10/13 对包的字节变更必须走共同的签名处理策略；只有 13 明确配置的保护范围允许追加约定的注释/签名目录，正文改变仍使引用摘要失配。
5. **容量与取消。** 继续沿用 `OfdPackageLoadOptions` 的压缩包/展开量预算。图片 ppm、字体子集、PDF 矢量、HTML 内嵌、关键词全页扫描和密码扩展再设各自像素、图元、字形、时间及输出大小限额；公开异步入口传递 `CancellationToken`。
6. **包级写入缺口。** `OfdPackageWriter.BuildEntries` 当前固定写一个 DocBody；内部 `OfdPackagePruner` 清理资源/失效签名，内部 `OfdResourceCatalog` 可按哈希复用载荷。多文档和包级编辑票在这些共用结构上定义清楚谁可重建、谁只可保留，避免各票各写一套 ZIP/XML 手术。这些内部类型不是调用方扩展点；若有票需要公开路径或资源操作，须单独审查公开 API。

### 2.1 拟定公开契约的首轮审查点

这些是实施起点，名字与签名在各票 API 设计评审后固定；调用方能观察的单位、默认值和错误语义必须先固定。内部 `OfdPackagePath`、`OfdPackagePruner`、`OfdResourceCatalog` 仅供实现复用，不计入公开契约。

| 票组 | 拟定契约 | 审查要点 |
| --- | --- | --- |
| 01/16/17/20 | `FlowDocument` + `Paragraph`/`Span`/`Table`/`AreaHolder`/`CanvasBlock`，统一 `Render` 成 `OfdDocumentPackage` | 元素是否可重复渲染，测量是否纯函数，CJK 字形按 1 em 近似的策略，分页上限及溢出诊断；16 首版若不支持 `RowSpan>1`，公开 API 明确拒绝、DOCX `vMerge` 明确诊断，不得静默拆成独立单元格；画布层序 |
| 02/08/18/21 | 各格式转换器的 `Stream`/文件入口和 Options；CLI 仅做参数映射与原子输出 | 页号一基/零基换算，默认 PNG、双层默认，输出大小和无内容行为；02 在单页 SVG→位图与 OFD→PDF→位图之间选定实现链路，由转换器统一负责 ppm/DPI 到像素的换算和预算 |
| 03/06/07/09/10 | `OfdDocumentPackage` 上的受控编辑与读取服务；编辑结果包含资源/签名处理报告 | 哪些 XML/载荷被重写、哪些保留，失效签名是否移除，未知引用导致的拒绝条件 |
| 22 | 复用 06 的定位结果，提供骑缝/对开页面外观写入入口 | 图像对象及图层顺序、相邻页裁剪几何、旋转页/不同页尺寸的失败或适配行为；不生成签名值 |
| 04/05/19 | `OfdGraphics`、字体选择/子集策略、可选 Skia 适配 | 绘图始终产原生图元，字体缺字如何回退，核心包依赖不扩张 |
| 11/12/13/14 | 独立扩展入口、显式注册/模式、结构化报告 | 不把自签和摘要匹配提升为默认 `FullyValid`，不让检查器改写原包 |

15 是发布工程闸门，按 §5 验收，不引入业务公开 API。

## 3. 依赖、批次与交付顺序

| 批次 | 票 | 开始条件与交付检查点 |
| --- | --- | --- |
| A：基础闸门 | 15（Q），与所有功能并行 | 先固化无隐私语料、坏包、许可和验收记录模板；发布候选必须五项全绿 |
| B：公开生成 | 01 → 16 | 01 的公开 API、Native DOCX 共核回归和两页视觉样例先完成，随后进入跨页表格；17/18 在 01 后可按需求领取 |
| C：日常进出/工具 | 02、03 | 02 与 03 可同批并行；03 保持一票四功能，内部按“水印→拆分→Mix→清签”提交可验证增量，不擅自改票范围 |
| D：issue #3 | 04、05；04 → 19 | 04+05 才满足 issue #3 的完成定义；19 是 21 的前置可选适配。20 在 01 和 04 均完成后即可按需求启动，不必等 05 |
| E：无厂商密码上限 | 11、12 | 与路线图建议顺序一致，排在保真/检索批次前；独立包与显式启用，不修改核心能力标志和默认验签判定 |
| F：保真与检索 | 04 + 19 → 21；06、09 | 21 待两个依赖和 PDF 绘制事件探针成立再开工；06 先定位再原位替换；09 先固定未知节点样本 |
| G：其余按需求 | 07、08、10、13、14，以及 17、18、20、22 | 不默认全部开工；17/18 依赖 01，20 依赖 01+04，22 依赖 06。10 涉及多文档保存，须先固定样本；13/14 保持可选范围 |

批次 B→F 的建议开工顺序与路线图 §8 及 ticket 索引一致；某票已有明确依赖时仍以依赖完成为准。无强制依赖的同批工作可以并行，但修改同一 `OfdPackageWriter`/`OfdReader`/`OfdPackagePruner` 的票应先约定数据结构和合并顺序。每票先提交最小 API 设计及失败行为，再做实现、样例、回归、文档和验收记录。预计投入只用于排序：小票约 2–5 开发日，中票约 5–10 日，01/05/09/10/11/12/16/21 属于需要原型和多轮验证的 10 日以上票；不把这些粗估当发布时间承诺。路线图 P2-06（证书加密等）没有 ticket，不在本计划实施范围。

### 3.1 里程碑出口

- **M1（01）：** 外部调用方仅引用公开 Layout 包即可生成两页样例；DOCX Native 接同一排版核心，既有 Native/default/DualLayer 测试与本次产物视觉记录齐备。未达到此点，16/17/18/20 不启动实现；20 还须等待 04 完成。
- **M2（02/03/16）：** P0 图片进出、四种文档工具和公开表格全部可通过 API/CLI（16 无 CLI 要求）复现；页面图、资源、签名清理和跨页表格分别有证据。03 任一子项未完成则 03 保持未完成。
- **M3（04/05）：** 不引用 Skia 的绘图样例 + 子集/复用对照均通过后，才可评估 issue #3 已完成；05 若只实现去重仍为部分完成。04 单独完成且 M1 已通过时，20 可按需求启动，无须等待 05 或整个 M3。
- **M4（21/06/09）：** 若选择推进 21，先完成 19，再以明确 PDF 样本集证明矢量路径与文字；06、09 分别完成定位替换与强类型对象验收。这是路线图第 5 步的出口，不代表全部 P1 完成。07、08、10、18、20 等票按各自依赖与调用方需求推进；未选做的票继续标记未实现。
- **M5（P2/Q）：** 可选包独立发布判断与五道 Q 闸门完成；P2 的自测通过仍不自动成为生产签章、档案认证或目标阅读器互认。

## 4. 逐票实施方案

下表每行都对应 `.scratch/capability-roadmap/issues/` 的一张票。行序沿用 ticket 索引的依赖分组；实际开工顺序看 §3。上游列给出本次实际核对的入口；落地列是拟实施设计，不表示现有行为。每票原文验收条件仍全部有效。

| 票 | 上游源码参考 | .NET 变更方案与关键验收 |
| --- | --- | --- |
| [01 公开流式布局](../.scratch/capability-roadmap/issues/01-public-flow-layout.md) | `ofdrw-layout/element/Paragraph.java`、`Span.java`、`engine/StreamingLayoutAnalyzer.java`、`engine/render/ParagraphRender.java` | 在 `Layout` 提取 DOCX BuiltIn 的样式游程、字形测量、行分割、基线及分页为内部核心；公开 `Paragraph`/`Span`、页尺寸/边距和排版入口。DOCX Native 适配成同一输入，保留现有页眉、页脚、换页语义。验证至少两页中英、局部样式不扩散、原文完整、异常空白页为零，并逐页 Preview。 |
| [16 公开表格](../.scratch/capability-roadmap/issues/16-public-tables.md) | `element/canvas/Cell.java`、`CellContentDrawer.java`；上游无可直接复用的完整跨页表格模型 | 基于 01 的块测量/分页建立 `Table/Row/Cell`，复用 BuiltIn 的列宽、水平 `GridSpan`、行高与边框写出。首版定义整行不可拆、过高行明确失败或受控拆分；纵向 `RowSpan>1` 若未实现则公开 API 明确拒绝，DOCX `vMerge` 进入共享核心时必须明确诊断或支持，不能静默变成独立单元格。带底色、对齐、水平合并和跨页样例，检查边框与文字无重影。 |
| [17 区域占位](../.scratch/capability-roadmap/issues/17-area-holders.md) | `element/AreaHolderBlock.java`、`areaholder/*`、`engine/render/AreaHolderBlockRender.java` | 01 布局时登记唯一名称、页路径、毫米框和目标 PageBlock ID；扩展数据类似上游 `AreaHolderBlocks.xml`，但先明确与未知扩展的兼容规则。回填只修改目标区域并执行溢出策略；比对回填前后页数和区域外坐标，签名变化按包级策略报告。 |
| [02 OFD 图片进出](../.scratch/capability-roadmap/issues/02-ofd-image-io.md) | `converter/export/ImageExporter.java`、`converter/ofdconverter/ImageConverter.java` | 新 API 与 `ofd-to-image`/`image-to-ofd` CLI；设计评审先在逐页自包含 SVG 栅格化与 OFD→PDF→位图之间选定导出链路，转换器统一换算 ppm/像素并检查预算，默认 PNG；导入解析 PNG/JPEG 尺寸并按页大小等比居中。校验页号、格式、像素预算和原子输出；导出图须含可辨文字/图形，导入包再读的页尺寸/图元正确。 |
| [03 文档工具](../.scratch/capability-roadmap/issues/03-document-tools.md) | `layout/edit/Watermark*.java`、`tool/merge/OFDMerger.java` 的 `addMix`、`sign/SignCleaner.java` | `Layout.Editing` 建立水印、按页拆分、Mix、显式清签四个独立 API，CLI 同名命令。水印写指定图层且导出/合并可见；拆分复用 Merger 资源映射并清未引用资源；Mix 以首源页尺寸、按输入顺序叠层并处理模板/注释；清签遍历全部声明并删除清单/值/外观的无引用载荷，验签返回“无签名”。四条均留独立样例。 |
| [04 OfdGraphics](../.scratch/capability-roadmap/issues/04-ofd-graphics-api.md) | `graphics2d/OFDPageGraphics2D.java`、`OFDGraphics2DDrawParam.java` | 在 `Layout` 公开 OFD 原语画笔、填充、路径、字体、矩阵和状态栈；先定义毫米坐标、矩阵乘法顺序、裁剪/不支持操作的失败行为。线/矩形/路径写 `PathObject`，文字写可抽取的 `TextObject`；无 Skia 依赖的票面示例往返并用 PDF/SVG/Preview 检查。 |
| [05 字体子集与复用](../.scratch/capability-roadmap/issues/05-font-subset-and-reuse.md) | `layout/engine/ResManager.java` 的资源缓存、`font/Font.java`；子集化没有现成可移植实现 | 先做 TTF/OTF/TTC、复合字形、字形 ID 与 Unicode 映射探针，选定可维护的子集库/授权；当前 Native 嵌入字节来自 `BuiltInOfdRenderer` 调用 `PdfFontRegistry.GetOriginalFont`，抽取布局时应将字体字节解析收敛为一个共享来源，再以全包实际用字集合子集化一次、按哈希复用并更新 OFD 字体引用。缺字回退或明确失败；同一样例记录全量/子集字节和视觉，验证 PDF/SVG/抽取与许可说明。 |
| [20 布局 Canvas](../.scratch/capability-roadmap/issues/20-layout-canvas-drawcontext.md) | `element/canvas/Canvas.java`、`DrawContext.java`、`engine/render/CanvasRender.java` | 在 01 的块序列中加入有边界的画布块，内部委托 04 的图元绘制；画布只占已测量区域，页眉/票面框和段落同页。若 16 完成再加表格混排样例；坐标、层序、溢出写入教程并检查 PDF/SVG/Preview。 |
| [21 PDF 矢量](../.scratch/capability-roadmap/issues/21-pdf-to-ofd-vectors.md) | `converter/ofdconverter/PDFConverter.java` 通过 PDFBox `renderPageToGraphics` 到 `OFDPageGraphics2D` | 保持现有双层默认。先证明所选 PDF 解析/渲染桥能提供路径绘制、文字矩阵、图像和裁剪事件，记录无法表达的效果；再增显式 `Vector`/混合选项，经 04/19 输出 OFD 原语。对同一 PDF 比较双层/矢量对象类型、文本、视觉和体积；扫描页保留图片或明确失败，不产空白页。 |
| [06 定位并替换](../.scratch/capability-roadmap/issues/06-keyword-locate-and-replace.md) | `reader/keyword/KeywordExtractor.java`、`layout/DocContentReplace.java` | 在 `Reader` 将 TextCode、Dx/Dy、字号、CTM 和嵌套页块重建为字符/字形毫米框，支持跨游程关键字并返回页/框/源对象定位；`Layout.Editing` 按定位结果做受控原位替换，测量新文字并定义超框失败/缩放策略。对重复关键字、多行、变换文字、无命中和已签包测试，抽取结果与视觉同时验收。 |
| [07 附件增删](../.scratch/capability-roadmap/issues/07-attachment-add-remove.md) | `layout/edit/Attachment.java`、`reader/OFDReader.java` 附件读取 | 给已有包提供添加/删除 API；使用 `OfdAttachment` 和现有 `Attachments.xml` 写入路径，按名称/身份定位、规范化载荷路径并只清无引用文件。与 `IncludeAttachments` 两种合并选项做往返矩阵，页面内容和未知扩展不变。 |
| [18 纯文本→OFD](../.scratch/capability-roadmap/issues/18-plain-text-to-ofd.md) | `converter/ofdconverter/TextConverter.java` | 01 的薄适配器：UTF-8 读取、换行规范化、字号/页尺寸/边距参数与 CLI；多页原文抽取对比，验证中文字体回退及 Preview 无裁半行。不在转换器里再建第二套分页器。 |
| [08 OFD→HTML](../.scratch/capability-roadmap/issues/08-ofd-to-html.md) | `converter/export/HTMLExporter.java`、`SVGMaker` | 复用 `OfdToSvgConverter`，按源页顺序生成自包含 SVG 并嵌入单 HTML，做 HTML/SVG 转义和输出大小预算；API 输出路径、CLI 按现有命令风格提供。浏览器打开多页样例检查文字、路径、图像与顺序；声明是预览，文字选择不保证。 |
| [09 强类型对象](../.scratch/capability-roadmap/issues/09-typed-object-model.md) | `ofdrw-core` 的 `bookmark/*`、`annotation/*`、`action/*`、`pageDescription/clips/*`、`compositeObj/*` | 以大纲/书签→注释→动作→裁剪→组合对象分阶段建模；每类先固定 XML 所有权、ID/路径和保留未知子节点策略，再加读写映射。每类有最小样例及 round trip；含扩展节点的旧包保存后比较其 XML/资源引用，不为强类型整洁而丢未知数据。 |
| [10 多 DocBody](../.scratch/capability-roadmap/issues/10-multi-docbody.md) | `core/basicStructure/ofd/OFD.java`、`reader/OFDReader.java` 的 `numOfDoc`、`pkg/container/OFDDir.java` | 新文档集合/选中索引 API，Reader 逐 DocBody 解析各自 DocRoot、资源、页面和签名；旧单文档 `ReadAsync` 语义保持首文档。Writer 明确默认单文档和显式追加策略，按文档隔离 ID/资源/签名路径；多文档样本在读取、选择、保存后不丢第二文档。签名写入/验证能否指定非首文档须与 13 单独定义，不能从上游多文档读取推断已支持多文档签章。 |
| [19 SkiaSharp 适配](../.scratch/capability-roadmap/issues/19-skiasharp-graphics-adapter.md) | `graphics2d/OFDPageGraphics2D.java` 的事件到 OFD 映射 | 独立可选包把 Skia 的受支持线、矩形、路径、文字操作投到 04；列出不支持的 shader/滤镜/混合模式并给显式失败或局部栅格回退。核心 Layout/Converter 不新增强制 Skia 依赖；与 04 同图形样例比较对象类型和外观。 |
| [11 GM/T 0099 口令](../.scratch/capability-roadmap/issues/11-gmt0099-password-crypto.md) | `crypto/OFDEncryptor.java`、`OFDDecryptor.java`、`enryptor/UserPasswordEncryptor.java` | 独立可选包实现条目筛选、文件密钥、口令包装、`Encryptions.xml` 和解密闭环；先用固定测试向量核验 SM3/KDF/SM4 参数和错误口令，后验证只加密部分页。输出原子化，清理临时明文；默认 Converter 依赖与核心 `SupportsGmT0099EncryptionEnvelope=false` 不变。文档明确非厂商互认/认证、非长期保存。 |
| [12 SES 自签互操作](../.scratch/capability-roadmap/issues/12-ses-self-sign-interop.md) | `gm/ses/v1|v4/*`、`sign/signContainer/SESV1Container.java`/`SESV4Container.java` | 独立 dev/interop 包解析常见 SES V1/V4 字段，用测试密钥产出并验证自己的 `SignedValue.dat`；只通过显式注册的 `IOfdSignedValueVerifier` 接入。测试默认 `OfdSignatureVerifier`/CLI 对该包仍非 `FullyValid`，引用完整性可单独成立；不改变 `SupportsBuiltInSesSm2Verification=false`，不声称阅读器互认/法律效力。 |
| [22 骑缝外观](../.scratch/capability-roadmap/issues/22-riding-seal-appearance.md) | `sign/stamppos/RidingStampPos.java`、`CuttingRideStampPos.java` | 复用 06 的页/毫米定位与图像裁剪，按相邻页边和对开比例生成连续的页面外观；用拼页图验证接缝、旋转页和不同页尺寸。只写普通页面图像/外观，不生成 `SignedValue.dat`，不得报告生产签章有效。 |
| [13 锁定/继续签](../.scratch/capability-roadmap/issues/13-lock-and-continue-sign.md) | `sign/OFDSigner.java` 的保护引用构建；`core/signatures/range/References.java` | 在 `OfdSignatureService` 增显式保护范围模式，构建引用时排除约定的 `Annots`/`Signs` 条目；在签名值提供者缺失时拒绝“已锁定”输出。上游 `OFDSigner` 仅针对首个文档，10 完成后本票必须按目标 DocBody 界定目录和引用范围，不能照搬上游单文档列表。测试加注释后原引用摘要匹配、改正文后失配、继续签后各签名的引用范围；验签状态仍由真实验证器决定。 |
| [14 OFD-A 检查器](../.scratch/capability-roadmap/issues/14-ofd-a-checker.md) | `archive/check/OFDArchiveChecker.java`、`ArchiveRule.java`、各 `rule/*` | 独立只读规则集，先列可静态判断的单文档、外部资源、字体嵌入、加密、附件/签名范围等规则及适用条件；每项返回规则 ID、严重度、包内路径和证据。坏包与故意违规/较干净样本出结构化报告；不修改包、不改变转换默认、不宣称 GB/T 42133 认证。 |
| [15 发布闸门](../.scratch/capability-roadmap/issues/15-release-engineering-gates.md) | 上游 `ofdrw-archive` 可作规则样本参考；本票以本仓库工程链路为准 | Q-01 许可/第三方声明清单及同批 NuGet 消费；Q-02 票据、发票、模板、异常包像素差阈值；Q-03 固定坏包/模糊输入与资源预算；Q-04 新增公开 API 的 CS1591 门禁及旧债清理顺序；Q-05 Native OFD→PDF→Preview 的记录模板。五项全做才可宣称发布闸门完成。 |

## 5. 各批次验收与发布判定

### 5.1 测试矩阵

| 层级 | 必查内容 | 证据 |
| --- | --- | --- |
| 单元/往返 | 新 API 的成功、空值/越界/超限、未知 XML 保留、ID 与资源引用、签名失效/保护范围、取消 | 对应 `tests/` 中有意义的断言；不是只断言文件存在 |
| 转换集成 | Native/default/DualLayer DOCX、PDF 双层默认、OFD→SVG/PDF、CLI 页号与退出码、包级编辑后重读 | 受影响测试项目；共享布局/字体/转换链路变更跑全套回归及本地包消费 E2E |
| 互操作 | 选取上游生成的无隐私样本与本库生成样本，双向读/导出；P2 仅测试自闭环和明确外部样本 | 记录上游提交、样本哈希、成功/失败范围；没有真实目标阅读器证据时不作互认结论 |
| 视觉 | 中英、比例字体、局部粗斜体/颜色、表格底色/边框、对齐、分页；票涉及图像/签名则加对应页面 | 每次重新生成本次 OFD，经 `OFD → PDF → macOS Preview` 逐页查看，记录页码和关键区域；PNG 辅助，不替代 Preview |
| 发布 | Q-01～Q-05、所有包一致版本/许可证、同批包消费、可追溯回滚 | CI/脚本结果、预览产物、Preview 验收记录；绿测不等于视觉或生产验收 |

建议测试落点：01/04/16/17/20/22 新建 `tests/Ofdrw.Net.Layout.Tests`；05 放字体子集、Packaging 与转换器视觉回归测试；19 放独立适配包测试；02/08/18/21 放转换器测试与 `tests/Ofdrw.Net.Cli.Tests`；03/06/07/09/10 放 Core/Packaging/Reader 相关往返测试；11/12/13 放 Signatures/扩展包测试；14 放只读检查器测试；15 由 `scripts/`、CI 和 `artifacts/` 中的留存记录共同验证。新增测试项目须进入 `Ofdrw.Net.sln` 和 CI，不能只有本地通过。

运行 .NET 构建/测试时遵守根 `AGENTS.md`：先设置可写 `DOTNET_CLI_HOME`、`DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1`、`DOTNET_CLI_TELEMETRY_OPTOUT=1`，单节点 `-m:1 /nodeReuse:false /p:UseSharedCompilation=false`；SDK 支持时再加 `--disable-build-servers`。本计划仅改文档，编写时未运行构建、转换或 Preview，故没有新的功能或视觉验收结论。

### 5.2 每票交付物与关闭条件

1. 一份简短 API/数据契约（含页号、单位、失败行为、资源限额、旧 API 兼容），实现与对应测试。
2. 无隐私可复现样例、生成命令和保留的 OFD/PDF/页面 PNG；涉及视觉的票附 Preview 记录：源码基线、样例、模式、实际查看链路、逐页检查范围、遗留缺陷和明显体积变化。
3. 更新 [功能对照](feature-parity.md)、相应 [教程](tutorials/README.md) 与 CLI 帮助；状态只写实际实现与实际验证到的范围。
4. 审核依赖边界：功能票合入不自动解除被阻塞票的质量门；发布需要 Q 五道闸门齐备。发现视觉缺陷时不宣称整体通过；Preview 不可用时记录未完成项。

发布候选应从固定源码提交构建并核对所有 NuGet 包版本和内容清单，先完成本地包消费 E2E，再执行发布工作流与公开源安装回读。若任一视觉或互操作门未过，保留上一个已验证包/标签供回滚，修复后由新提交和新版本重走构建、验收与发布；不覆盖已发布的同版本包。需要生产部署的下游由其所有者另做环境验收，本计划的库级验证不代替下游签收。

### 5.3 主要技术风险与预设处理

- **01/16 字体测量差异：** `BuiltInOfdRenderer.Advance` 通过 PdfSharpCore `XGraphics.CreateMeasureContext`/`XFont` 测量，CJK 字形按 1 em 近似；抽取时先固定可替换的测量接口和该近似的默认/配置策略，再比较改造前后的中英基准页、DOCX Native 分页与文本。不能以提取文本正确替代视觉。
- **05 子集化：** 字体格式、复合字形、许可和文字映射可能超出原估；先交探针与可回退的全量嵌入选项，子集化不通过 PDF/SVG/Preview 前不切默认。
- **03/09/10 包级改写：** `PreservedEntries` 会让未知引用难以安全裁剪；保守保留优于误删，若票要求“无残留”则通过明确引用闭包处理已知结构，对未知引用给诊断或拒绝。
- **10/13 多文档签名边界：** 上游 `OFDSigner` 类注释限定首文档。非首 DocBody 的签名值、注释和 `Signs` 路径在 .NET 侧必须单独设计、测试并与首文档隔离；不能用上游 `OFDReader` 可选择文档来推定签名也可选择。
- **21 矢量提取：** 当前 PdfPig 路径只支持文字语义层；若绘制事件桥不能稳定取得路径/文字参数，先完成探针与技术选型，不以“原 PDF 可选中文字”推定矢量写出可行。
- **11/12/13 安全表述：** 测试自签、摘要匹配、页面外观和证书/设备互认是不同层级；默认 CLI/功能对照保持区分，任何能力标志更改均单独审查。

## 6. 开工前需要锁定的少量设计选择

这些选择应在对应票的 API 设计阶段确定，不阻挡无依赖批次启动：

1. 01/16 的公开排版对象是可复用不可变描述，还是一次性可变 Builder；需给出并发/复用语义。
2. 16 的过高单行、跨页重复表头与合并跨页单元格首版如何失败或降级；先选择可验证的有限策略。
3. 05 选用的子集库及字体授权/再分发方式；未确认前保持全量嵌入作为旧行为。
4. 10 显式多文档写入的 API 是文档集合写入还是追加操作；默认仍写一个 DocBody。
5. 21 路径事件来源和字体回退模式；由样例探针决定是否可实施当前票的矢量范围。

## 7. 上游源码定位索引

下列链接固定在本计划的上游提交，便于实施时核对真实调用链；表格里的简写路径均相对于对应模块的 `src/main/java/org/ofdrw/`，转换器等包名不同处以链接为准。

| 对应票 | 固定源码入口 | 核对重点 |
| --- | --- | --- |
| 01/16/17/20 | [段落](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-layout/src/main/java/org/ofdrw/layout/element/Paragraph.java)、[流式分页](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-layout/src/main/java/org/ofdrw/layout/engine/StreamingLayoutAnalyzer.java)、[Canvas Cell](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-layout/src/main/java/org/ofdrw/layout/element/canvas/Cell.java)、[占位渲染](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-layout/src/main/java/org/ofdrw/layout/engine/render/AreaHolderBlockRender.java) | 上游分页段、固定单元、扩展区块的实际边界 |
| 02/08/18 | [图片导出](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-converter/src/main/java/org/ofdrw/converter/export/ImageExporter.java)、[图片导入](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-converter/src/main/java/org/ofdrw/converter/ofdconverter/ImageConverter.java)、[HTML 导出](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-converter/src/main/java/org/ofdrw/converter/export/HTMLExporter.java)、[文本导入](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-converter/src/main/java/org/ofdrw/converter/ofdconverter/TextConverter.java) | 格式进出及 HTML 内嵌 SVG |
| 03/07 | [水印](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-layout/src/main/java/org/ofdrw/layout/edit/Watermark.java)、[混合与按页选择](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-tool/src/main/java/org/ofdrw/tool/merge/OFDMerger.java)、[清签](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-sign/src/main/java/org/ofdrw/sign/SignCleaner.java)、[附件构造](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-layout/src/main/java/org/ofdrw/layout/edit/Attachment.java) | 文件重组、资源引用和签名处理 |
| 04/05/19/21 | [Graphics2D](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-graphics2d/src/main/java/org/ofdrw/graphics2d/OFDPageGraphics2D.java)、[资源管理](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-layout/src/main/java/org/ofdrw/layout/engine/ResManager.java)、[PDF 桥接](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-converter/src/main/java/org/ofdrw/converter/ofdconverter/PDFConverter.java) | 原生绘制事件、字体资源复用、PDF 渲染桥 |
| 06/09/10 | [关键字定位](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-reader/src/main/java/org/ofdrw/reader/keyword/KeywordExtractor.java)、[替换](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-layout/src/main/java/org/ofdrw/layout/DocContentReplace.java)、[DocBody 集合](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-core/src/main/java/org/ofdrw/core/basicStructure/ofd/OFD.java)、[Reader](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-reader/src/main/java/org/ofdrw/reader/OFDReader.java) | 字形坐标、内容替换、非首文档选择；09 的强类型目录见 `ofdrw-core` |
| 11/12/13/14/22 | [口令加密](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-crypto/src/main/java/org/ofdrw/crypto/enryptor/UserPasswordEncryptor.java)、[SES V4 容器](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-sign/src/main/java/org/ofdrw/sign/signContainer/SESV4Container.java)、[保护引用](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-sign/src/main/java/org/ofdrw/sign/OFDSigner.java)、[档案检查](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-archive/src/main/java/org/ofdrw/archive/check/OFDArchiveChecker.java)、[骑缝定位](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-sign/src/main/java/org/ofdrw/sign/stamppos/RidingStampPos.java) | 可选密码/规则引擎与页面外观；不得照搬其对外效力表述 |
