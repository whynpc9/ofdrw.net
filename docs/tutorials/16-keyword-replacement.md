# 16. 关键词定位与受限文字替换

本功能处理所选文档的正文文字。API 页位置零基，CLI 页号一基；坐标单位为毫米，相对于页面 PhysicalBox 左上角。定位结果包含原字体的字形 ID、基线、advance 和字形框四角，字形框不是实际绘制像素的轮廓。

## 定位 API

引用 `Ofdrw.Net.Reader` 和 `Ofdrw.Net.Converter.Pdf`，由应用事先审核原嵌入字体的许可、真实风格和精确内容身份。输入文档自己给出的摘要不能成为信任依据。

```csharp
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Converter.Pdf.Geometry;

var source = await new OfdReader().ReadAsync(inputPath);
// vettedSha256 / vettedByteCount 来自应用审核过的原始字体文件。
var provider = new OfdTrustedEmbeddedFontGeometryProvider(new[] {
    new OfdTrustedEmbeddedFontIdentity(vettedSha256, vettedByteCount)
});
var query = new OfdKeywordLocator(provider).Find(source, new[] { 0 }, "AAA");
foreach (var hit in query.Hits)
    foreach (var fragment in hit.Fragments)
        Console.WriteLine($"page={hit.PageIndex}, glyph={fragment.GlyphId}, " +
            $"x={fragment.BaselineMillimeters.X}, y={fragment.BaselineMillimeters.Y}");
```

匹配为 ordinal 字面匹配，返回重叠命中，不是正则表达式。不会推断阅读顺序或空白，也不会跨页、图层或连续基线边界匹配。模板和批注外观不属于正文查询范围。默认遇到不支持的正文立即失败；定位可显式选择 `ReportIncomplete`，此时必须检查 `IsComplete` 和 `Diagnostics`，不能把部分结果当成全文结果。结构路径和 occurrence 身份只在本次源版本内有效，不是 XPath、ZIP 字节偏移或永久 ID。

## 替换 API

另引用 `Ofdrw.Net.Layout` 和 `Ofdrw.Net.Packaging`：

```csharp
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;

using var prepared = new OfdTextReplacer(provider).Replace(source, "AAA", "---");
var edited = prepared.CreatePackageCopy();
await new OfdPackageWriter().WriteAsync(edited, outputPath);
```

`Replace` 保持输入模型不变；结果只能消费一次，未消费结果应 Dispose。也可将真实 `Find` 命中构造成 `OfdTextReplacement` 列表后替换。保持源模型和字体字节稳定直到消费结束；重叠请求、过期或伪造命中、缺字、不支持的几何以及容纳不下的替换均拒绝整个批次。空字符串表示删除。

已验证的替换 profile 是原嵌入 Regular 字体、简单 Unicode scalar、水平正向推进和仅平移几何。保留未替换文字的原锚点；替换文字从原起点按原字体度量排布，不承诺与每个旧字形有相同位置。不会通过回退字体、缩小字号或缩放来强行放入原区域。旋转、复杂 shaping、合成粗斜体及共享模板/批注编辑不属于此 profile；定位本身的几何支持范围可更宽。

受影响页面的邻接图元也必须满足可证明的文字编辑条件；混合路径、图片、复合或不支持的正文可能使整个替换失败，即使关键词本身可以定位。

上述 API 的准备和随后独立 Writer 保存是两个操作，Writer 有自己的预算；任意 Stream 输出由调用方负责暂存和发布。需要已验证的文件保存链路时使用下述 CLI。

## CLI

```bash
ofdrw replace-text input.ofd output.ofd --find AAA --replace '---' \
  --trusted-font 3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa:21659544 \
  --pages 1
```

这里的身份只适用于验收样例中完整原始 Noto CJK 字体，不是任意中文字体白名单。`--trusted-font SHA256:BYTECOUNT` 可重复指定；CLI 使用 Regular 身份。不传 `--pages` 时选全部页，页号必须有效且不重复，支持 `1,3-5`。`--find` 不得为空；`--replace ''` 删除匹配文字。

输入和输出必须是不同文件，不能用符号链接绕过路径限制。失败返回非零，并保留旧输出；成功完成验证、暂存复制和原子发布后返回 0。无命中或同文本替换仍完成分析，并将原 ZIP 字节原样复制到输出。四项累计账本（storage、copy、validation work、output upper bound）分别以 512 MiB 为默认上限；它们不是进程 RSS 或 CPU 耗时保证。

## 签名重写与验证范围

改变正文后原签名不能继续证明新内容。Changed 保存按 Writer 的签名清理约定去除失效声明；签名载荷是否删除取决于引用闭包，不能据此宣称所有旧签章文件已消失。CLI 的字节精确 Noop 保留整个原 ZIP，包括原有签名；独立 API 后再次调用 Writer 不等同于这条 Noop 链路。需要签章时，对最终产物重新签署并验证，见[签名教程](10-signatures.md)。

验收包括完整原始 Noto Latin/CJK 字体的真实 CLI 保存、公开 Reader/Locator 的字形与几何复核、1245 项测试（含 65 项真实字体用例）、双 TFM 和 11 个本地包消费。视觉验收通过实际 OFD→PDF→macOS Preview 检查 Latin/CJK 前后页，并检查确定性 DOCX 的 native/default OFD 导出的两页各自结果。结论只覆盖记录的样例和页面，不承诺任意复杂 OFD/Word 文档保真。功能、视觉、原始失败和许可排除分别见[持久证据](../evidence/issue06-owned-save/README.md)；Issue09 既有范围外 multi-M 问题未在本票修复。
