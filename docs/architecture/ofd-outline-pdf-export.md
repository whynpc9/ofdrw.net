# OFD 大纲导出到 PDF 的设计提案

## 1. OFD 模型

`OfdDocumentPackage.Outlines` 提供有序树。每个条目包含标题、可空的初始展开状态 `Expanded`、可选叶子数提示 `Count`、有序动作及有序子条目。命名书签由同一文档的 `Bookmarks` 持有，名称区分大小写。见 `src/Ofdrw.Net.Core/Models/OfdDocumentPackage.cs:84`、`src/Ofdrw.Net.Core/Models/OfdNavigation.cs:7`、`src/Ofdrw.Net.Core/Models/OfdNavigation.cs:39`。

`OfdGoToAction` 引用一个直接目标或一个命名书签。目标绑定页对象，包含 `XYZ`、`Fit`、`FitH`、`FitV`、`FitR` 类型及可空的 `Left`、`Top`、`Right`、`Bottom` 和 `Zoom`。坐标单位为毫米。缩放允许省略、用零保留当前值，或取 0.1 到 64。其他动作包括 URI、打开附件、音频、视频，以及保留原 Action XML 的 `OfdRawAction`。事件可以是 `DO`、`PO`、`CLICK` 或省略。见 `src/Ofdrw.Net.Core/Models/OfdNavigation.cs:63`、`src/Ofdrw.Net.Core/Models/OfdNavigation.cs:101`、`src/Ofdrw.Net.Core/Models/OfdNavigation.cs:132`、`src/Ofdrw.Net.Core/Models/OfdNavigationXml.cs:437`。

## 2. PdfSharpCore 1.3.67 的写入能力

当前导出器依次创建 PDF 页、绘制内容并保存。它没有把 OFD 大纲交给 PDF 写入器。见 `src/Ofdrw.Net.Converter.Pdf/Ofdrw.Net.Converter.Pdf.csproj:21`、`src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs:133`、`src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs:176`。

对 NuGet 1.3.67 的 net8.0 DLL 做反射，并核对包元数据指向的源码。公开入口是 `PdfDocument.Outlines` 和 `PdfOutline.Outlines`，通过 `PdfOutlineCollection.Add` 添加根条目或子条目。条目可设置标题、`Opened`、`DestinationPage`、`PageDestinationType` 和五个数值属性。写入器支持五种匹配的显式目标，以 `PdfPage` 的间接引用开头。仅分组条目的 `DestinationPage` 可为空。公开构造入口提供本地目标和分组。见 [PdfDocument.cs:664](https://github.com/ststeiger/PdfSharpCore/blob/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Pdf/PdfDocument.cs#L664)、[PdfOutline.cs:170](https://github.com/ststeiger/PdfSharpCore/blob/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Pdf/PdfOutline.cs#L170-L299)、[PdfOutlineCollection.cs:79](https://github.com/ststeiger/PdfSharpCore/blob/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Pdf/PdfOutlineCollection.cs#L79-L104)、[PdfOutline.cs:488](https://github.com/ststeiger/PdfSharpCore/blob/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Pdf/PdfOutline.cs#L488-L589)。

`XYZ` 可写显式缩放。相关属性用 `double.NaN` 表达 PDF `null`，可保留阅读器已有的左、上、缩放值，以及 `FitH` 的上坐标、`FitV` 的左坐标。1.3.67 的格式串 `#.##` 有两个限制。它最多保留两位小数，并会把零及舍入为零的值格式化为空字符串。实际 DLL 的格式化探针返回 `NaN → "null"`、`0 → ""`、`0.001 → ""`、`1.2345 → "1.23"`。有效的零坐标必须写成数值零。实现前应验证低层 `PdfArray` 的序列化或采用修正的后端，确保 `/Dest` 在保存后保持完整。见 [PdfOutline.cs:556](https://github.com/ststeiger/PdfSharpCore/blob/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Pdf/PdfOutline.cs#L556-L621)、[PDF 1.7 §8.2.1 表 8.2](https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/pdfreference1.7old.pdf)。

## 3. 目标映射

以实际导出的页对象为 `p`，以 OFD 页框原点为 `(X0, Y0)`，以 OFD 页高毫米值为 `H`。建议沿用正文相对页框的约定。将绝对 OFD 坐标转换为 `x(v) = (v - X0) × 72/25.4`、`y(v) = (H - (v - Y0)) × 72/25.4`。这是新 PDF 页的左下原点点坐标。导入生成的 OFD 页起点为零。见 `src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs:133`、`src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs:489`、`src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs:676`、`src/Ofdrw.Net.Converter.Pdf/Converters/PdfToOfdConverter.cs:124`。

| OFD 类型 | PDF 目标 | 约束 |
| --- | --- | --- |
| `XYZ` | `[p /XYZ x(Left) y(Top) zoom]` | 左、上可同时为 `null`。单边坐标在 PDF 中有效，现有导入会拒绝，建议首版也拒绝。缩放省略或为零时写 `null`。 |
| `Fit` | `[p /Fit]` | 按整页适配。 |
| `FitH` | `[p /FitH y(Top)]` | 缺少 `Top` 时写 `null`。 |
| `FitV` | `[p /FitV x(Left)]` | 缺少 `Left` 时写 `null`。 |
| `FitR` | `[p /FitR x(Left) y(Bottom) x(Right) y(Top)]` | 所有坐标必须有限，且满足 `Left < Right`、`Top < Bottom`。缺边或退化矩形应报错。 |

对所有类型，建议拒绝该类型未定义的已填字段（如 `Fit` 上的显式缩放）、非有限的显式数值、单位转换后溢出、超出 OFD 范围的缩放，以及序列化精度导致的矩形退化。用 `NaN` 表示 PDF 空操作数时，只从类型模型的空值生成。PDF 操作数顺序见 [PDF 1.7 §8.2.1 表 8.2](https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/pdfreference1.7old.pdf)。坐标和缩放的 OFD 入模约束见 `src/Ofdrw.Net.Core/Models/OfdNavigationXml.cs:153`。

导入已把显示用 CropBox 和旋转折入 OFD 坐标，并在 90°/270° 时交换 `FitH` 与 `FitV`。回导以新的页框和页面方向为准，比较显示位置时允许数值误差。原 PDF 的 CropBox 偏移、旋转、空操作数的拼写和展开状态无法从导入结果还原。见 `src/Ofdrw.Net.Converter.Pdf/Internal/PdfImportMetadataMigrator.cs:183`、`src/Ofdrw.Net.Converter.Pdf/Internal/PdfImportMetadataMigrator.cs:290`、`src/Ofdrw.Net.Converter.Pdf/Internal/PdfImportMetadataMigrator.cs:351`、`src/Ofdrw.Net.Converter.Pdf/Internal/PdfImportMetadataMigrator.cs:396`。

## 4. 页面选择与多文档

导出器通过 `OfdToPdfOptions.PackageLoadOptions.DocumentIndex` 读取单个 DocBody，按页索引排序后规范化页列表。列表允许重排和重复，缺省或没有有效索引时选所有页。建议创建 PDF 页时按 OFD 页对象建立映射，重复页取输出中的第一份。命名目标解析为当前文档的显式目标；跨文档、悬空或选页排除的目标按下节报错。当前导入已采用相同的首份和排除策略。见 `src/Ofdrw.Net.Converter.Pdf/OfdToPdfOptions.cs:12`、`src/Ofdrw.Net.Reader/Readers/OfdReader.cs:162`、`src/Ofdrw.Net.Reader/Readers/OfdReader.cs:537`、`src/Ofdrw.Net.Core/Models/OfdPageSelection.cs:12`、`src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs:90`、`src/Ofdrw.Net.Converter.Pdf/Internal/PdfImportMetadataMigrator.cs:188`。

## 5. 建议的 SDK 接口与失败策略

建议新增 `OfdToPdfOptions.CopyBookmarks`，缺省为 `false`。启用后导出标题、同级顺序、层次以及一个本地 GoTo（事件为空或 `CLICK`）。无动作（包括空动作集合）且有子条目的节点作为分组保留。用 `Expanded` 初始化打开状态，缺省折叠。OFD 的 `Count` 是叶子数提示，PDF 的 `/Count` 统计条目展开时可见的后代，条目折叠时以负数保存。后端只针对 `Opened` 的新增节点给祖先递增计数，而有子项的条目仅在 `OpenCount > 0` 时写计数。例如父条目及其唯一叶子均折叠时，父条目会缺失应写的 `/Count -1`。需以实际 PDF 树计算计数。见 `src/Ofdrw.Net.Core/Models/OfdNavigation.cs:18`、[PdfOutlineCollection.cs:96](https://github.com/ststeiger/PdfSharpCore/blob/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Pdf/PdfOutlineCollection.cs#L96-L104)、[PdfOutline.cs:488](https://github.com/ststeiger/PdfSharpCore/blob/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Pdf/PdfOutline.cs#L488-L552)、[PDF 1.7 §8.2.2 表 8.4](https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/pdfreference1.7old.pdf)。

按条目验证。遇到外部动作、多动作、不支持的事件（含 `DO` 和 `PO`）、无动作的叶子、无法解读的导航、排除页或无效目标时，报告具体条目的路径与原因，并在 PDF 写入前终止整次转换。读取器会保留无法进入类型树的源条目。应同时核查源 XML 的未建模子项与未建模动作，避免遗漏。用户显式启用复制时可以据此修复输入，避免生成跳错页或缺项的大纲。失败策略与现有导入一致。建议提供 `MaxBookmarkCount = 10_000`、`MaxBookmarkDepth = 64`，根深度为 1，用有界遍历检测循环、重复节点及取消。见 `src/Ofdrw.Net.Core/Models/OfdNavigationXml.cs:50`、`src/Ofdrw.Net.Core/Models/OfdNavigationXml.cs:69`、`src/Ofdrw.Net.Core/Models/OfdNavigationXml.cs:96`、`src/Ofdrw.Net.Converter.Pdf/Internal/PdfImportMetadataMigrator.cs:169`、`src/Ofdrw.Net.Converter.Pdf/Internal/PdfImportMetadataMigrator.cs:267`、`src/Ofdrw.Net.Converter.Pdf/PdfToOfdOptions.cs:89`。

## 6. 拟议测试与验收

复用 `tests/Ofdrw.Net.Converter.Pdf.Vector.Tests/PdfMetadataMigrationTests.cs:177` 的嵌套大纲、`tests/Ofdrw.Net.Converter.Pdf.Vector.Tests/PdfMetadataMigrationTests.cs:337` 的 CropBox/旋转样例、`tests/Ofdrw.Net.Packaging.Tests/StandardObjectEditingTests.cs:108` 的多 DocBody 构造方式。用确定性的 OFD 样例覆盖五种目标、零及近零坐标、非零页框原点、命名目标、分组和层次。验证默认省略大纲，选页重排或重复后目标准确，选中其他 DocBody 时同名目标互不影响。外部动作、无动作叶子、非支持事件、未建模的兄弟或子条目、排除页、非有限值、坐标溢出、舍入退化和预算超限应失败且输出流未写入。

使用 PDF 测试工程已引用的 PdfPig，通过 `TryGetBookmarks` 回读标题、顺序、层次、输出页号和目标类型，再检查原始 `/Dest` 与 `/Count`，覆盖缩放、空操作数及展开状态。现有导入同样需要组合解析器结果与原始操作数。视觉验收打开本次 OFD 导出的 PDF，使用 macOS Preview 逐页对照正文并点击各类型书签核对落页和显示区域。保存 PNG 及实际 Preview 记录。见 `tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj:11`、`src/Ofdrw.Net.Converter.Pdf/Internal/PdfImportMetadataMigrator.cs:169`、`src/Ofdrw.Net.Converter.Pdf/Internal/PdfImportMetadataMigrator.cs:290`。

## 7. 待维护者决定

1. 首版采用哪条目标数组写入路径？建议先验证保留数值零与精度的低层数组，也验证字体修复输出分支中的间接页引用。两条保存路径见 `src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs:176`。
2. 在哪一层修正 `/Count`？建议由支持正确计数的后端序列化，并用全折叠、全展开和三层混合树断言根及每个分组的计数值与符号。见 [PdfOutline.cs:488](https://github.com/ststeiger/PdfSharpCore/blob/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Pdf/PdfOutline.cs#L488-L552)。
3. 如何确认外部 OFD 的非零页框坐标？建议用已知阅读器能正确跳转的样本，验证上述页框原点扣除规则后再承诺该输入。阅读器目前将声明的目标数值直接存入模型。见 `src/Ofdrw.Net.Core/Models/OfdNavigationXml.cs:153`。
