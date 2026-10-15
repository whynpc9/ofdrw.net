# Ticket 21：整页原始图像回退设计门禁

状态：设计与实包探针完成；产品修复、独立验证、Preview 验收均未完成。基线 `e7dc5651eee90e8be32f3a3da1b56c7dccdc4676`，2026-10-08，限定当前 3a07 worktree。仅修改本文件及 `artifacts/pdf-vectors/image-fallback-design/`。未打开 GUI、未修改产品、未提交或推送。

## 结论与证据

建议有条件采用：在可选 vector 转换器显式 `RasterizePage` 的回退分支内，保留严格单幅全页 RGB 图像的原始采样网格，并将源 `/Interpolate` 的有效布尔值写成图像 SourceXml 扩展属性。保持 `IsNative=false`、一个 ImageObject、零 PathObject、零 TextObject；不将扫描图像归类为 native，也不改变默认 dual converter、19 或字体契约。

必须同时解决两个实际门禁：完整 token/EOF 识别，以及 Core 精确属性白名单。仅靠 PdfPig 返回的四个 operations、或仅改 PDF/SVG exporter，均不足。

| 本次证据 | 实际结果 |
| --- | --- |
| `source-and-export-facts.json` | 原始 fallback-source.pdf SHA256 `40567d531b000b269f5c6a8e3af7d49b54cc0e6b8be66a2096d1fd0e6beae7a2`；page 3 `q 420 0 0 595 0 0 cm /Im1 Do Q`；object 8 为 DeviceRGB/8、2×2、原始 12 字节 `d2283c1e82be1e82bed2283c`，省略 Interpolate。 |
| 同一 source facts | Resources 除 `/Im1` 外有**未使用** `/Font /F1` 和 `/ExtGState /GS1`（`ca=.4`, `BM=Multiply`）。不能因这些声明存在就拒绝此页，也不应加载它们。 |
| `parser-probe.json`，真实 PdfPig 0.1.15 | `q cm Do Q 99`、`q cm Do Q /dangling` 返回四个操作，尾随操作数未进 factory；`99 q` 也被 Reflection factory 接受；未知 operator 在 factory 返回 null 时从结果消失。 |
| `nupkg-probe.json`，真实 r1 nupkg（nuspec repository commit=e7dc565） | `false/true/invalid` 属性均通过 Writer→Reader、内部 CloneElement 的 namespace remap 保留；带属性的公共 Merge、Mix、无归档基线 Split 均抛 `NotSupportedException`，不带属性时成功。归档读回后的 Split→Writer→Reader 成功保留属性（含现阶段未知的 invalid）；这不是非法值应被 exporter 接受的依据。克隆探针通过反射调用已有内部函数，仅用于设计证据。 |
| 四份 `false/true/absent/invalid.pdf`、对应 SVG | 当前 exporter 忽略所有新标记；PDF 均输出 2×2 `/Interpolate true`，SVG 均无 image-rendering。证明需要 exporter 修复，不能声称 r1 已实现提示。 |
| `mixed-interpolation.pdf` | 同一 PNG 字节分别经全新 XImage 设置 false/true/false，三页 pdfimages 为 no/yes/no，像素字节相同；本例未发生图像缓存串用，无需修改共享 cache。 |
| `orientation-source.pdf`、`orientation-source-true.pdf` 与 `mixed-interpolation.pdf` page 1/2 | 新增非对称红/绿/蓝/黄 2×2 原始 PDF，省略（有效 false）和显式 true 分别与直接 PdfSharp false/true 输出的 Poppler 144 DPI PNG（840×1190）逐像素相等，且 decoded RGB 字节一致。说明该正向全页 CTM 无需翻转原始像素行；这不是产品 importer 验收，也没有跨 renderer 比较 true。 |

协调者已有 `interpolation-diagnostic/` 和 `image-page-probe/` 证据：GS/Poppler 原始页面为硬边四块，Preview 源页面平滑；直接 XImage false 的 Poppler 与原始源页相同。此处未重新执行 GS/Preview，不替代已有记录，也不通过将 fixture 改为 true 来消除差异。

## 受限识别与回退边界

建议内部入口 `TryCreateOriginalImageFallback(options, remainingImageBytes, token, out page, out encodedBytes, out diagnostic)`；方法可放在已有 VectorPageContext 或专用内部 helper。不新增公开 option/model property。只在 native 转换因 `NotSupportedException` 被拒绝、且 policy=RasterizePage 时调用；Fail 保持原行为。重新打开严格 PdfPig 上下文后仅读取字典和原始流，不调用普通 GetPage、GetImages、font loader 或 renderer。返回 false 表示不匹配，继续现有整页 dual 回退；预算、取消、I/O 和内部错误不得降格为不匹配。

必须一次通过以下条件，才分配 PNG 并产生页：

| 层级 | 必须验证／明确拒绝 |
| --- | --- |
| 文档颜色/可选内容 | 新 profile 检查 `document.Structure.Catalog.CatalogDictionary`（PdfPig 0.1.15 公开 API）；有 OutputIntents、OCProperties、AcroForm/NeedsRendering 时保守不命中。不要只检查页级 ColorSpace 就声称 DeviceRGB 无文档级颜色/可见性影响；不调用可能解码 ICC/表单流的高层 getter。该检查只限制新 profile，不改既有 native/dual 的全局契约。 |
| 页树 | 与现有逻辑一致，按最近祖先解析继承 MediaBox/CropBox/Rotate/Resources，沿 Parent 有 token 检查及 MaxStackDepth/循环边界。UserUnit 为页自身属性；缺省 1，显式值只接受 1。不把非继承键错误继承。 |
| 页几何 | MediaBox 必须四个有限数 `[0 0 W H]`，W/H 正；CropBox 缺省或精确等于 MediaBox；Rotate 缺省或 0。不接受负向、旋转、非零原点、裁切页框；mm 换算有限、正且 Writer 可表达。不要用浮点 epsilon 将欠覆盖或越界“吸附”到全页。 |
| 页渲染状态 | `/Annots`、`/Group` 只要存在即拒绝，包括空值；额外页面绘制容器、可选内容或不能证明无影响的渲染扩展拒绝。仅查看 MediaBox 的静态页面，不承诺 TrimBox/BleedBox/ArtBox 所定义的生产裁切。 |
| 内容流 | 首版最窄可只接受一个直接或间接 StreamToken；数组、多流走旧回退。只接受内部无过滤流，拒绝 Filter/DecodeParms/F/FFilter/FDecodeParms，即使 null/空数组。检查原始字节限额再复制；Length 一致性异常不得“修复后接受”。 |
| 内容语法 | 使用受限 PDF tokenizer 对**全流**严格匹配 `q NUM NUM NUM NUM NUM NUM cm NAME Do Q EOF`，允许标准空白和 `%` 注释；operator 的操作数数量分别为 0/6/1/0。仅 NUM token 可成为矩阵，Name token 可成为资源名；正确处理名称 #xx 转义或保守拒绝。拒绝 trailing number/name/string、额外 q/Q、未知/兼容/marked-content/inline-image 操作、clip、gs、ri、文本和额外绘制。不要用 regex 提取一段成功前缀，也不要仅检查 parser 返回集合。 |
| 矩阵 | 数值精确为 `[W 0 0 H 0 0]`，所有值有限；仅一次 cm、一次 Do、一次成对 q/Q。负缩放、轴交换、平移、shear、重复图像均不命中。 |
| 资源 | 使用有效 Resources，只解析 Do 点名的 XObject；它必须解析为 Image StreamToken。不执行或解码未使用的 Font/ExtGState/其它 XObject；纯四操作语法证明无 gs/Tf/第二次 Do。保守拒绝有效 Resources 中任何 `/ColorSpace`（包括 DefaultRGB）；不要把 unused GS1 当作当前图形状态。禁止 Form、Pattern、Shading、OC 等间接绘制路径。 |
| 图像字典 | 建议白名单仅 Type（缺省或 XObject）、Subtype=Image、Width、Height、ColorSpace、BitsPerComponent、Length、Interpolate。Width/Height 必须有限正整数、可表示为 int；ColorSpace 必须直接/间接解析为精确 Name DeviceRGB，BitsPerComponent=8。其它键全部不命中，包括 Decode（即使 identity）、Mask、SMask、ImageMask、SMaskInData、Intent、Matte、Alternates、OC、OPI、Filter/DecodeParms/F/FFilter/FDecodeParms；这样无需推测透明度、颜色重解释、外部文件或滤镜语义。 |
| 原始采样 | 必须恰好 `checked((long)width * height * 3)` 字节，拒绝截断及多余字节。先核限额、再复制；逐行 top→bottom、逐像素 RGB 写无损 PNG，不缩放、不预插值、不添加透明度、不附会为嵌入 ICC 色彩。 |
| Interpolate | 缺省规范有效值 false；显式仅接受真正 BooleanToken，false→false、true→true。拒绝 Name/String/Number/null 伪布尔值。对缺省也写明确 false 提示，否则当前 XImage 默认 true 会重新改变语义。 |

页/图像未知键的拒绝属于受限 profile，不是对所有 PDF 的有效性判断。未命中仍使用原整页回退或 Fail，不静默忽略可见内容。有效 profile 没有 PDF 原文文字，因此无需透明语义文本，不从图像做 OCR。

## 限额、取消与输出

沿用 `VectorLimits.Snapshot`、输入 staging 的 MaxInputBytes、源/输出 MaxPageCount、page selection 的次序/重复语义。使用相同的 MaxContentBytesPerPage、MaxOperationsPerPage（四操作也不能绕过设为 1–3 的限制）、MaxStackDepth；一幅输出至少占一个 MaxEventsPerPage。避免先让 native parser 早退，再让新分支遗漏原有预算。

对源像素独立检查 `width*height <= Compatibility.MaxRasterizedPixelsPerPage`，且所有 stride/长度/bitmap buffer 的 checked 运算不溢出；页面按既有 Dpi clamp 72–300 计算的目标 raster 像素上限也保留，避免同样页面在新路径绕过旧 Compatibility 限制。原始像素限制与页面目标像素限制是两回事。

原始流已在 PdfPig 的输入解析对象内；检查 Data.Length 后再 ToArray 只能约束新增复制，不能声称在 PDF scanner 的首次分配前已有硬流限额。输入 cap 和不解压是这一 profile 的已有防线。建议 source RGB 复制/bitmap 分配另以剩余 MaxTotalImageBytes 作保守上限；明确它可能比最终压缩 PNG 更严格，不把 PNG 压缩比当作内存保障。

PNG 应经有界输出 Stream 编码，写入不能超过当前剩余 MaxTotalImageBytes；若使用 SKData 编码后再检查，需先证明 worst-case RGBA/编码缓冲上界，否则只是产物检查而非编码分配限额。累计预算按每个选中输出页计入 PNG 字节，保留现有 vector 重复页计费语义；与普通 raster 页共用 imageBytes，先验余额、checked 累加，再加入 package。所有预算失败抛 InvalidDataException/明确预算异常，取消抛 OperationCanceledException，不能再次 rasterize。

取消检查位置：继承/间接解析边界、读内容、token 扫描、像素每行复制、编码前后、有界流写、添加页前、最终 Writer 前。同步 encoder 内部本身不可中断时，应记录其有界延迟，不宣称即时取消。

当前公开 API **没有独立最终 OFD/ZIP 总字节预算**：MaxTotalImageBytes 是累计图片载荷限额，不能当成整个 ZIP 限额。Writer 在完整 model/entries 构造后才创建 ZipArchive（`OfdPackageWriter.cs:43–50`）；解析/预算失败应发生在目标首次写之前，但最终写入阶段取消或 I/O 故障可能留下部分输出。若目标要求最终 ZIP 的严格大小或事务写入，需单独协调 API/临时输出方案，不能在这次受限修复中暗称已经满足。

## SourceXml 与导出映射

建议内部 Core helper：`OfdImageRenderingHints.PdfInterpolateV1` 为 `XName.Get("PdfInterpolateV1", "https://ofdrw.net/image-hints")`，`ReadPdfInterpolate(OfdImageElement)` 返回 `bool?`。只读正确 OFD ImageObject 根上的该**完整 XName**，只接受规范字符串 `"false"`/`"true"`；精确标记值为 `True`、`0`、`1`、空白或其它非法串时抛 InvalidDataException（协调者明确选择 fail-fast）。其它 namespace、同名非命名空间属性、未知版本、无提示返回 null，保持原导出默认。不要从 CTM、像素尺寸、文件名或“看上去像扫描页”猜测提示。

冲突规则：同一个 expanded XName 的重复属性本身是非法 XML，维持解析失败；合法 V1 与未知 namespace/版本同名属性共存时，仅 V1 决定当前两个 exporter，未知属性沿用既有 XML 保留/拒绝行为，不得新增 namespace 通配符。V1 放在错误根或 child 不作为导出提示，也不能被白名单认可。提示只改变采样 flag，不读任何附带 CTM/Boundary/Clips/Alpha/ResourceID，不向 model 注入结构或赋予 SourceXml 绕过现有结构验证的权力；其它合法绘制属性仍走既有模型/Writer 合同。未知属性会不会被 Merge 拒绝仍按旧契约，不为“保持默认”而放宽它。

Producer 的 SourceXml 根使用本次 Compatibility.Namespace；只放 ImageObject 和该属性，其 Boundary/CTM/ID/ResourceID 留给 Writer 现有分支重建。Writer `OfdPackageWriter.cs:408–432` 保留扩展属性；Reader `OfdReader.cs:875–893` 保存完整 ImageObject XML；Cloner `OfdModelCloner.cs:50–58,87–97` 重映射 OFD 元素 namespace 时不更改外部属性。实包已验证这些行为，不需修改 Writer/Reader。

Core 还必须在 `OfdGraphicXmlContract.IsKnownAttribute`（38–49 行）只对正确 OFD ImageObject 根 + 上述 XName + 值恰为 true/false 加入精确允许项，仿照 FauxItalic 的定位但增加合法值约束；不得放开整个 image-hints namespace。否则 `OfdDocumentMerger.cs:209–212` 的公共 Merge 会拒绝生成物（实包已复现），Reader 的 annotation/template 可扁平化判断也会受同一契约影响。其它位置的同名标记仍不授予已知属性身份。Core 已将内部访问授予 Vector/Pdf/Svg，无需新增公开 API 或 friend assembly。

Mix 经 `OfdDocumentMixer.cs:95` 调用 RequireKnownAttributes=true 的 Merge；无 archive 基线的 Split 经 `OfdDocumentSplitter.cs:45–47` Clone→Merge，所以都需要上述白名单，已实包确认。带 archive 的 Split 使用 ClonePage，Writer 重建图像 XML 后 `OfdPackagePruner` 按资源引用裁剪；本次单页实测 roundtrip 成功。Pruner `AddReferences`（454–457 行）收集属性值，canonical true/false 不是 Writer 生成的数字 resource ID，不应改为读取任意扩展结构。修复验收仍须加入多页 Split 删除页、保留/共享图片资源与提示的断言，单页探针不能证明删除资源全部正确。

| helper 结果 | PDF | SVG |
| --- | --- | --- |
| false | 在 `OfdToPdfConverter.DrawImage` 的 XImage 创建后、DrawImage 前设置 `Interpolate=false`。输出可省略 false 或显式 false，验收比较有效值。 | 建议仅该 image 设置 `image-rendering="crisp-edges"`；这是无平滑意图提示，仍需实际消费者验证。 |
| true | 同一位置显式 `Interpolate=true`。 | 建议仅该 image 设置 `image-rendering="smooth"`；消费者不支持时会退回其默认。 |
| null | 不设置，保留当前 XImage 默认。 | 不增属性或 style，保持当前输出。 |
| 精确标记值非法 | 明确失败，不降为 null 或 false。 | 同一 helper 明确失败。 |

SVG 选择来自 [W3C CSS Images 3 §5.2](https://www.w3.org/TR/css-images-3/#the-image-rendering)。`pixelated` 在非整数放大时允许额外平滑，不能将它描述为严格 nearest-neighbor；`crisp-edges` 也不规定唯一算法。若目标 SVG 1.1 renderer 只支持 optimizeSpeed/optimizeQuality，应先固定并实测消费者再决定兼容映射；[SVG 1.1 §11.7.5](https://www.w3.org/TR/SVG11/painting.html#ImageRenderingProperty) 的 optimizeSpeed 允许比 nearest-neighbor 更高质量的算法，同样不能保证源 false 像素一致。这里仅推荐现代映射，未完成 SVG renderer 兼容门禁。

PDF 改动只作用于带标记的 OfdImageElement 正文/已走同一 DrawPage 的内容；不碰无标记的印章图像路径、字体 resolver/subset/cache、19 的 producer。未知版本/namespace 的提示不参与解释，精确 V1 非法值失败，不为别的图片全局关闭插值。PNG 及其它经 SVG 的下游需单独验证，不能从 SVG 字符串推断显示结果。

## 修复后的最小验收矩阵

1. 原始 SHA fixture 保持不变；原省略、显式 false、显式 true 均保留 2×2 原始 RGB 字节和原布尔语义，整页 image、native=false、无任何文本/path；对照默认 dual 产物不变。为颜色排列另加非对称至少 2×2（最好 3×2）样例，防止翻转/stride 误判。
2. 严格负例逐项：非全页/负向/旋转/crop/origin/UserUnit、Annots/Group、DefaultRGB、Form、第二个 Do、gs、clip、额外/尾随 token、操作数个数错误、数组内容、过滤/外部流、Decode/Mask/SMask/Intent/OC、非 RGB/非 8bit、错误宽高/长度/布尔；检查不会误命中新路径。未使用 Font/GS1 的当前源应命中且不加载它们。
3. source pixels、目标页面像素、content bytes、operation count、remaining image budget、重复选择与 mixed native/raster/原图页累计限制；取消前/行复制中/写出时。故障不得变成旧 raster 成功，预写失败目标为空；不把最终 I/O 部分写归为模型事务成功。
4. 精确 namespace 标记 roundtrip/rewrite/clone namespace remap/公共 Merge/Mix/Split/多页 pruning；错误 namespace、wrong root/child、未知版本/缺省不改变采样默认，精确 V1 大小写/空白/数值/非法值失败，重复 XML 属性失败。false/true 同字节相邻页与重排/重复页不串用（本次直接 PdfSharp 小探针通过，但产品接入仍需测试）。
5. actual 修复 nupkg 独立干净缓存消费，PDF image dictionary/decoded RGB、SVG 属性及 embedded PNG 分别核对；共享变更所需全套回归、默认 11 包及可选包消费，由 Main/Low 按既有门禁完成。
6. Main 基于新 OFD→PDF→Preview 逐页查看，并记录与原 PDF 在相同 viewer 下对照结果；GS/Poppler PNG 是辅助，true 在不同 renderer 下不得要求相互逐像素相等。补做指定 DOCX Native/default 共享链回归与 Preview。任何 renderer/Preview 可见缺陷仍是待处理项。

## 当前风险及剩余工作

- 阻断实现偷懒方案：四 operations 不是完整语法证明；已有实测尾随 operand 漏检。必须新增严格 token 消费证据。
- 阻断只改两个 exporter 的方案：Core 白名单不更新会使公共 Merge 拒绝此生成物。实测有提示失败、无提示通过。
- SVG 标志是消费者相关采样意图；DeviceRGB 本身也不提供跨设备色彩完全一致保证。PDF 保留原始采样与布尔提示可改善同 viewer 对照，不承诺任意 PDF 或第三方 native OFD reader 插值一致。
- 配置内存预算不等于完整进程内存/ZIP 总字节限额；原始数据、RGBA、PNG、Writer entries 在峰值时可能共存，编码实现需选择明确上界。
- 本次功能探针构建/执行成功；曾有探针 namespace 编译错误与 Python 文件名遮蔽标准库错误，已仅在允许目录内修正，原失败日志保留。NuGet 还报告现有 ImageSharp 2.1.13 audit 警告，未改依赖。
- 本次无 GUI，因此视觉验收未执行；未运行产品全套测试、未实现生产修复、未取得 Low 独立结论。只完成本设计门禁。

探针可复现入口：`artifacts/pdf-vectors/image-fallback-design/README.md`。所有新输出为设计诊断，不可替代最终修复候选及其 Preview 验收。
