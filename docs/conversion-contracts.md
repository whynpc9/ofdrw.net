# 转换、编辑与资源约定

本页描述当前源码的行为。新增结果 API 随后续包版本发布；本地源码验证不代表已发布到 NuGet。

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
- 已知不支持内容按 `UnsupportedFeatureBehavior` 处理；使用 Placeholder 时应检查结果的 `OriginalTextPreserved` 和 `Diagnostics`，不能只根据 Task 正常完成判断原文完整。
- DualLayer 的视觉层始终要求栅格化。即使传入的 PDF 选项允许文本后备，该 DOCX 视觉阶段也不会以 PDF 抽取文字作为替代。
- DualLayer 根据实际渲染 PDF 的字符位置定位 OpenXML 原文，再进行选页。PDF 文字只用于定位；发出的文本仍来自 OpenXML，页码字段由原字段定义计算。无法可靠定位正文时明确失败，不提交部分语义层。
- 未打印的批注等优先定位到原 XML 引用的页面。完全没有页面或引用锚点的附属文本，整本输出会保留在末页并返回文档级作用域警告；此时显式选页会失败，避免把未知归属的原文错误分配给选中的页面。

Native 支持确定性常见排版，仍不承诺任意浮动对象、复杂域、复杂跨页表格或全部 Word 版式的保真。

## 字体与宿主程序

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
- `OfdToPdfOptions` 控制 OFD 加载和图片解码预算。渲染失败不会以丢失正文的空白后备页冒充成功。
- 外部进程并发消费有界诊断输出，支持超时、取消与进程树清理。原生库内的同步渲染在调用前后检查取消，不能把外部进程超时等同于强制终止任意原生调用。
- CLI 先验证参数，再写目标目录下的临时文件；成功关闭后才替换目标。失败或取消保留原输出。Ctrl-C 返回 130。格式转换输入输出需使用不同路径；页重排支持原位更新。

## 验证与发布

`scripts/run-converter-package-e2e.sh <version>` 构建完整包集，在新消费目录和新缓存中测试聚合包、签名包和 CLI。包源映射禁止从其他源获取 `Ofdrw.Net.*`。`--consume-only --packages-dir <dir>` 用于验证已经打好的同一批包，测试前后核对 SHA256 清单。

发布工作流先消费待发布包，再验证清单并推送；CI 保存测试结果和渲染页面。必需视觉工具失败时测试失败。自动检查覆盖非空页、确定性输出比较以及具体回归像素，但不能替代项目要求的 macOS Preview 逐页验收。
