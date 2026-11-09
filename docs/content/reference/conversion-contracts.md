---
title: "转换、编辑与资源约定"
description: "DOCX/PDF 原文与分页、字体、页面编辑、资源预算、取消及输出失败的精确约定。"
sidebar:
  order: 40
  label: "转换与资源约定"
---

本页描述 `0.3.0` 的转换、编辑与资源行为。固定样例的发布验收范围见[0.3.0 发布说明](/releases/030/)；0.2.0 的历史结论仍见[原发布说明](/releases/020/)。

## 原始 DOCX 文本与分页

`DocxToOfdConverter` 默认使用 Native。正文直接来自 OpenXML，生成普通 OFD 文本、路径和图片对象。`ConvertWithResultAsync` 返回模式、实际引擎（Native 为 null）、诊断、渲染总页数、输出页对应的零基源页索引，以及原文保留和页级映射状态。

```csharp
var converter = new DocxToOfdConverter(new DocxConversionOptions());
var result = await converter.ConvertWithResultAsync(docxInput, ofdOutput, new[] { 2, 0, 2 });
foreach (var diagnostic in result.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

- 页选择保留顺序和重复页，忽略混入的越界索引。未指定、空列表或全部越界时选择所有页。这与既有 PDF 转换行为一致。
- Native 的页眉页脚按页生成，支持首页、奇偶页、节继承、页码起始值及 PAGE / NUMPAGES / SECTIONPAGES 的十进制计数。分页前预留页眉页脚区域。
- BuiltIn/Native 将脚注、尾注和批注正文以带类型和编号的内容追加到正文后，返回 `DOCX_SUPPLEMENTAL_TEXT_APPENDED`；这保证附属原文可用，不等同于 Word 原位脚注排版。
- Native/BuiltIn 列表编号（自 0.2.0 起）：支持按 numId 的多级计数器（上级前进时下级重置）、`%1`–`%9` 级别模式、`start`/`startOverride`，以及 decimal、decimalZero、大小写罗马数字、大小写字母、chineseCounting/chineseCountingThousand（1–99）、ideographDigital 和项目符号。`numId=0` 关闭编号；编号定义无法解析时省略编号并返回 `DOCX_NUMBERING_UNRESOLVED`，不支持的格式退化为十进制并返回 `DOCX_NUMBERING_FORMAT_DEGRADED`。编号为纯文本加一个空格，不渲染缩进和制表符后缀。
- Native 对 `w:u` 不为 `none` 的文字运行绘制一条实线下划线路径（与 FlowDocument 几何相同，颜色跟随文字），覆盖正文、表格和页眉页脚；不区分双线、点线、波浪线等样式，也不使用 `w:u w:color`（自 0.2.0 起）。
- Native 对 `both`/`distribute` 段落中自动换行的行做两端对齐：剩余宽度平均分配到词间空格和中日韩字符之后，行首空白不拉伸，只含一个拉丁单词的行保持左对齐；末行和显式换行不拉伸，`distribute` 的末行拉伸及中文标点挤压未实现（自 0.2.0 起）。
- 已知不支持内容按 `UnsupportedFeatureBehavior` 处理；使用 Placeholder 时应检查结果的 `OriginalTextPreserved` 和 `Diagnostics`，不能只根据 Task 正常完成判断原文完整。
- DualLayer 的视觉层始终要求栅格化。即使传入的 PDF 选项允许文本后备，该 DOCX 视觉阶段也不会以 PDF 抽取文字作为替代。
- DualLayer 根据实际渲染 PDF 的字符位置定位 OpenXML 原文，再进行选页。PDF 文字只用于定位；发出的文本仍来自 OpenXML，页码字段由原字段定义计算。无法可靠定位正文时明确失败，不提交部分语义层。
- DualLayer 输出与 Native 使用同一包轮廓：命名空间取 `DocxConversionOptions.OfdNamespace`（默认 `http://www.ofdspec.org`，与 OFD-H 轮廓及只识别该 URI 的旧阅读器一致），`DocType` 为 `OFD-H`，元数据标明 DOCX 来源。视觉阶段的 PDF→OFD 默认 `/2016` 命名空间不会泄漏到结果里；经 `IPdfToOfdConverter` 暂存再读回的包，其资源清单与原始片段也会改写到目标命名空间。独立 PDF→OFD 保持 `/2016` 默认值，可用 `PdfToOfdOptions.Namespace` 改为短 URI。资源清单 `PublicRes.xml`/`DocumentRes.xml` 与其他部件一样声明 `ofd` 前缀，而不是无前缀默认命名空间。
- 未打印的批注等优先定位到原 XML 引用的页面。完全没有页面或引用锚点的附属文本，整本输出会保留在末页并返回文档级作用域警告；此时显式选页会失败，避免把未知归属的原文错误分配给选中的页面。

Native 支持确定性常见排版，仍不承诺任意浮动对象、复杂域、复杂跨页表格或全部 Word 版式的保真。

## PDF 导入的文字顺序边界

独立 PDF→OFD 的默认透明文字层，以及栅格化失败后显式允许的文本后备，均无法从 PdfPig 默认几何排序恢复经过验证的双向逻辑原文。解析字符含固定 Unicode 17 R/AL 或相关 bidi 控制范围时，转换以 `NotSupportedException` 和 `PDF_TEXT_READING_ORDER_UNVERIFIED` 明确拒绝，不再把倒序词写成成功的语义结果。即使某个未整形输入碰巧抽出预期字串，也不据此准入。

调用方可明确选择 `TextLayerMode = PdfTextLayerMode.None` 且 `RequireRasterization = true`，仅保留页面图；这条路径不声明 RTL 原文、字形映射或阅读顺序支持。现有 LTR 抽取也不保证任意多栏/复杂 PDF 的逻辑顺序。

### 可选的受限 RTL 语义导入（F06，自 0.3.0 起）

默认值不变。设置 `RightToLeftTextImport = PdfRightToLeftTextImport.VisualLineProfile` 后，含 RTL 文字的页面改走受限的“可见行”配置，仅用于页面图上方的透明文字层：

- **几何**：只采用直接绘制的字形（文字渲染模式为填充/描边，未被裁切，非透明，不在可选内容中），按同一水平基线的字形原点排出视觉顺序，与内容流顺序无关。不可见文字层（`Tr 3`、空裁切、`ca 0` 等）只是生成方的声明，不作为证据；包括本库 OFD→PDF 导出的透明语义层在内，一律拒绝。
- **字符映射**：每个字形须有解码后的 Unicode 值，且仅含 BMP 中 bidi 类别为 L、R、AL 的字母；一个字形可映射多个同向字符（如阿拉伯文 lam-alef 连字）。
- **逻辑顺序**：由视觉顺序推出候选逻辑文本，只有仓库内经一致性测试的 Unicode 17 UAX #9 正向重排**恰好**还原所观察到的字形排列时才接受。不对 `GetWords()` 的结果重新排序。
- `RightToLeftBaseDirection` 默认 `Auto`：每行独立按 P2/P3 判定，两种基方向都成立（如以拉丁文开头、希伯来文结尾的换行段落）或都不成立时拒绝；已知段落方向时可显式设为 `LeftToRight` 或 `RightToLeft`。
- 每个逻辑字符写入 OFD 时保留其来源字形的原点：`TextCode` 为逻辑顺序，RTL 字符使用负 `DeltaX`。

以下情形仍以 `PDF_TEXT_READING_ORDER_UNVERIFIED` 拒绝且不写输出：数字、标点及其他中性/弱类别字符，组合符号（如希伯来文元音点），表现形式（U+FB1D–FDFF、U+FE70–FEFF），非 BMP 字符，`/ActualText`、`/ReversedChars` 或可选内容，旋转/竖排文字或页面旋转，重叠字形，同一行带内的不同基线（上标等），超过 3 em 的行内间隙（无法区分分栏），表单 XObject 中的文字，以及无法解析的透明度。未栅格化的文本后备页仍拒绝 RTL。该配置只证明所接受的行在 UAX #9 意义上自洽且与可见字形一致；未拼成正确外观的输入按其**实际显示**的顺序导入，不还原生成方的原始字符串。真实阅读器中的 Copy/Find 尚未验证。设计与证据见 [F06 设计记录](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/f06-pdf-rtl-import.md)。

## 字体与宿主程序

可直接操作的固定 Regular/Bold 准备和 Native/default 示例见[中英文 DOCX 字体部署教程](https://whynpc9.github.io/ofdrw.net/guides/fonts/)。

BuiltIn 按 `FontFallbackFamilies` 的配置顺序选择 `FontDirectories` 中的真实 TTF/OTF 字体，也支持从 TTC 中抽取独立面。未配置目录时扫描平台字体目录。Native 排版和嵌入使用同一物理字体面：配置了真实 Noto CJK 字体时会嵌入，拉丁字符也使用该字体的度量，不再以宿主西文字体度量替代。等价的 Noto 家族名称可以匹配真实字体，但不会把其他中文字体伪装成 Noto。 未从目录加载的 `Noto Sans CJK*` 家族（含 JP/TC/KR 等变体）保留名称字体，不嵌入宿主解析器的替代字节；这一路径的拉丁字符仍沿用 Arial 兼容度量。

宋体/黑体等系统中文族保留名称字体兼容路径：OFD 只声明 `SimSun` / `SimHei`，不嵌入替代字体，由阅读器解析本机字体；这一路径的中日韩全宽字符仍使用 1em，拉丁字符保留原兼容度量。名称字体不能保证不同机器显示一致。LibreOffice Portable 默认关闭独立 UserInstallation；仍会将配置的 CJK 字体复制到便携版用户字体目录。

OFD → PDF 根据实际选中字形文件的样式判断是否模拟粗体/斜体；对名称字体仅在请求这些样式时额外探测宿主后备面。该探测只注册独立 TTF/OTF，TTC、无效字节、读取/解析异常或可选注册预算不足会跳过注册，保留逐文本字体回退；OFD 内嵌字体失败仍严格报错。成功探测的字体会计入进程级注册预算，常规名称字体不因该探测被额外复制注册。

Native OFD 中的加粗/斜体同时写在两处：字体资源 `ofd:Font` 上的 `Bold`/`Italic` 只描述字体面；文本对象 `ofd:TextObject` 上的 CT_Text `Weight`（加粗为 `700`，常规 `400` 时省略）与 `Italic="true"` 才是规范要求的样式。仅声明 `SimSun`/`SimHei` 名称字体时，不少 OFD-H 阅读器会忽略这两类标志，因此对无 `FontFile` 的名称字体，写出器还会描边字形（`Stroke="true"` `Fill="true"` `LineWidth`≈0.045em）做假粗体，并用 CTM 错切做假斜体。透明 DualLayer 文字不描边。Native DOCX 请求粗体、但选中的嵌入字体只有常规面时，默认 `NativeEmbeddedBoldBehavior = DocxEmbeddedBoldBehavior.TextStroke`：保留真实字体字节和原始文字，以 0.025em 的文字描边补充强调，并返回 `DOCX_EMBEDDED_BOLD_SIMULATED` 诊断。提供真实粗体面时不额外描边。若要求真实粗体，设置 `DocxEmbeddedBoldBehavior.Throw`；缺少实际粗体面时转换在写出前失败。此选项仅适用于 Native 嵌入文字，不是全局缺字体策略。CLI 使用默认策略。导入的嵌入文字保留原有绘制属性，不自动获得这项补绘。`OfdReader` 读回 `Weight`/`Italic`，OFD → PDF/SVG 仍以 `Weight >= 600` 或 `Italic` 作为模拟粗体/斜体的依据。

CI 使用 `scripts/install-ci-fonts.py` 下载固定版本且校验 SHA-256 的 Noto Sans CJK SC，生成 Regular 静态 TrueType 字体，以避免操作系统镜像的字体差异。脚本依赖 `fonttools==4.59.2`；本地可用 `--directory /path/to/fonts` 生成隔离目录，再通过 `FontDirectories` 或 CLI `--font-directory` 指定。字体及 OFL 许可证仅写入验证环境，不进入 NuGet 包。

PDFsharp 的字体缓存为进程级，首次使用字体后不能替换其全局解析器。应在应用启动时、其他 PDFsharp 字体操作之前初始化：

```csharp
PdfFontRegistry.EnsureInstalled();
```

已有自定义解析器的宿主可在同样的启动阶段组合：

```csharp
GlobalFontSettings.FontResolver = PdfFontRegistry.CreateResolver(myFontResolver);
```

组合后的解析器将 SDK 注册的字体交给内容身份解析，将其他字体请求交回宿主。SDK 不会重置宿主已使用的缓存；初始化太晚时会明确报错，避免悄悄用错嵌入字体。

`RegisterFontFace` 返回与字体内容和所需样式关联的内部族名，适合文档内使用。同名字体可以有不同内容。旧 `RegisterFont(name, bytes)` 保留命名注册方式，但不允许将同一名称/样式重新绑定到不同字节。

OFD 保留原始字体字节。供 PDFsharp 使用的副本具有内容唯一的内部名称，避免依赖库的名称缓存冲突；字形和度量表保持不变。不同文档共享相同字节时只在注册表中保留一份原始载荷。

注册表默认上限为 256 MiB 字体载荷和 4096 个族/样式别名，可在启动时通过 `MaximumRegisteredFontBytes`、`MaximumRegisteredFaces` 调整。达到限额会明确失败，不驱逐仍可能被 PDFsharp 使用的字体。此限额约束 SDK 注册表，不能作为整个进程工作集的保证。

长期运行的转换服务：实测每个文档嵌入一个不同子集字体时，单进程约在第 261 个 CJK 文档（每个子集约 1 MB）或第 1,603 个拉丁文档（约 167 KB）达到字节上限，此时托管内存已达 1–2 GiB，因为 PDFsharp 自身也缓存字体且无法清除。建议在 `PdfFontRegistry.RegisteredFontBytes` 超过上限一半或处理数百个含不同嵌入字体的文档后回收工作进程，并按每进程数 GiB 规划内存。测量方法与选项评估见[PDF 字体注册表生命周期](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/pdf-font-registry-lifetime.md)。

## 页面编辑与合并

读取后保存会保留未知扩展及资源。删除页面后保存会清理被删除页面的原始条目、页批注以及可确认只被其使用的字体/图片载荷。其他页面、模板、附件或扩展引用的共享内容会保留；无法检查的扩展 XML 会触发资源保留诊断。

```csharp
OfdDocumentEditor.RemovePages(package, new[] { 1 });
var saved = await new OfdPackageWriter().WriteWithResultAsync(package, output);
```

`RemovedEntries` 与 `Diagnostics` 说明实际清理范围。删除页面不是对整个任意扩展包进行内容擦除：如果同一内容仍作为模板或其他活跃内容使用，它应继续存在。

重写已签名的包且字节发生改变时，写包器移除失效的签名声明，清理已知无引用的签名载荷，并返回 `SignaturesInvalidated`。需要签名时，应对最终输出重新签署；摘要完整性与完整密码验证仍由签名模块分别报告。

合并按资源身份重映射字体，保留常见图片 CTM、透明度、路径裁剪和原始样式 XML。不能安全重映射的资源引用或 Raw 对象默认拒绝。显式启用 `SkipUnsupportedRawElements` 时，使用 `MergeWithResult` 获取被跳过对象的诊断。

## 资源预算与失败行为

- OFD 默认压缩输入上限 512 MiB、单条目展开上限 128 MiB、总展开上限 512 MiB、条目/读取页数上限各 10000。`OfdReader.ReadAsync` 可接收 `OfdPackageLoadOptions`。
- DOCX 默认压缩输入上限 64 MiB、展开上限 256 MiB，并限制 XML 元素、图片、配置字体和页数。输入限制在复制过程中执行。
- PDF→OFD 默认输入上限 128 MiB、源/输出页数上限 10000、单页解码像素上限 4000 万、累计页图像字节上限 512 MiB。DPI 统一限制在 72–300，外部渲染也使用所选 DPI。
- OFD 读取使用内部的整文档 XML 节点/引用工作计数（默认 10,000,000 / 20,000,000），内存仍由源 XML 与物化字节预算约束。preview.10 默认 100,000 / 200,000，约 1,000 个以上文本对象的文档会被拒绝；调用方暂不能配置这两个计数。写入端另有按最坏情况估算的输出预算（`MaxMaterializedBytes` 512 MiB），约 2 万段以上的单个 FlowDocument 仍会在写入时被拒绝。
- `OfdToPdfOptions` 控制 OFD 加载和图片解码预算。渲染失败不会以丢失正文的空白后备页冒充成功。
- 默认 OFD→PDF 仍容忍无法解析或绘制的签章外观，但会通过 `ConvertWithResultAsync().Diagnostics` 报告：`PDF_SIGNATURE_METADATA_UNREADABLE`（签名元数据不可读）、`PDF_SIGNATURE_APPEARANCE_UNREADABLE`（外观载荷或嵌套 OFD 不可读）和 `PDF_SIGNATURE_APPEARANCE_SKIPPED`（外观未绘制，含页索引和签名 ID）；指定页选择时只报告所选页面上的签章；CLI `ofd-to-pdf` 将其输出到 stderr，退出码仍为 0。诊断消息不回显文档原始内容（自 0.2.0 起）。
- 嵌套签章 OFD 按调用方 `OfdToPdfOptions.PackageLoadOptions` 的限制读取，同一嵌套签章载荷在一次转换中只解析一次（自 0.2.0 起）。
- main（尚未发布）在 OFD→PDF 和 OFD→图片的 PDF 中间产物中，对带 `Clip` 的位图印章关闭图像插值，避免骑缝章裁切边缘混入相邻片区的颜色。
- `OfdPackageWriter` 与 `OfdPackageSignatureCleaner` 向调用方流写入时如中途失败或被取消，已写出的字节刻意不构成可读 ZIP（不写中央目录），避免下游把缺页的部分包当成 OFD；调用方仍应丢弃部分输出，写文件时采用先写临时文件再重命名的方式（自 0.2.0 起）。
- 外部进程并发消费有界诊断输出，支持超时、取消与进程树清理。原生库内的同步渲染在调用前后检查取消，不能把外部进程超时等同于强制终止任意原生调用。
- CLI 先验证参数，再写目标目录下的临时文件；成功关闭后才替换目标。失败或取消保留原输出。Ctrl-C 返回 130。格式转换输入输出需使用不同路径；页重排支持原位更新。

## 验证与发布

`scripts/run-converter-package-e2e.sh <version>` 构建完整包集，在新消费目录和新缓存中测试聚合包、签名包和 CLI。包源映射禁止从其他源获取 `Ofdrw.Net.*`。`--consume-only --packages-dir <dir>` 用于验证已经打好的同一批包，测试前后核对 SHA256 清单。

发布工作流先消费待发布包，再验证清单并推送；CI 保存测试结果和渲染页面。必需视觉工具失败时测试失败。自动检查覆盖非空页、确定性输出比较以及具体回归像素，但不能替代项目要求的 macOS Preview 逐页验收。
