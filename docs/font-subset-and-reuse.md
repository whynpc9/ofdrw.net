# 字体子集与复用（05 / issue #3）

新建包默认 `SubsetWhenSafe`：`OfdFontResource.Data` 是字节来源，`Id` / `OfdTextElement.FontResourceId` 绑定字体，`Bold` / `Italic` 保留样式。05 不增加另一套公开字体描述，04 的 `OfdFont` / `OfdGraphics` 可以继续绑定这些 Core 资源。不同名称或样式但相同字节的资源保留各自 ID，共用全包用字集合和一份 `FontFile`。同名不同内容始终分开。

```csharp
package.Fonts.Add(new OfdFontResource {
    Id = "10", FontName = "My licensed font", Data = File.ReadAllBytes("font.ttf")
});
page.Elements.Add(new OfdTextElement {
    FontResourceId = "10", FontName = "My licensed font", Text = "中文 Alpha"
});
var result = await new OfdPackageWriter().WriteWithResultAsync(package, output);
foreach (var face in result.FontEmbedding)
    Console.WriteLine($"{face.SourceBytes} -> {face.PayloadBytes}, subset={face.IsSubset}");
```

## 生产后端与支持边界

SDK 内部使用纯 .NET managed sfnt/glyf 后端，无 Python/fonttools、HarfBuzz CLI、本地动态库或额外生产包依赖。Core 的 `OpenTypeFace` 是字体表读取/重建的共同来源，DOCX 集合选面、PDF 名称隔离、SVG、coverage 与子集生成共用。Native 的测量和完整嵌入源仍来自同一个 PDF 注册源；写出时才生成独立子集，不改变调用方字节、字体资源或已计算坐标。

| 范围 | 行为 |
| --- | --- |
| 静态 TTF / glyf OpenType | cmap 4/12、非 BMP、组合字形递归闭包、GSUB 1–8 保守不动点闭包；保持 GID、hmtx/vmtx、hinting、GSUB/GPOS/GDEF 原字节 |
| cmap 14 / UVS | 保留完整 14 子表及所有非默认变体 GID；输入变体序列必须存在；默认映射仍依赖普通 cmap |
| 规范化 | 加入 NFC/NFD 码点闭包，避免浏览器规范化时丢失字形 |
| TTC | `CollectionFaceIndex` 零基显式选面，默认 0；按实际 face 内容复用；写出独立 sfnt，PDF/SVG 使用相同选面 |
| CFF/CFF2 OTF、变量字体、AAT/颜色/未知依赖表、JSTF | 全量保留，返回 `FONT_FULL_PRESERVED` 与原因；不宣称这些格式已子集化 |
| RTL/bidi 字符 | 目前缺少通用 Unicode mirror closure，保留全量并诊断；不宣称复杂文字子集支持完整 |
| 大 BMP 字符集 | 压缩格式 4 连续 delta 段并保留格式 12；格式 4 无法容纳时全量保留，避免现有 PDFsharp 拒绝 format12-only 字体 |
| SourceXml/CGTransform/Raw、模板/注释、保留包条目 | 整包保守保留全量字体，避免裁掉未知 GID/扩展引用；已有子集保持原字节；诊断会说明 coverage 未验证的未解析字体 |

保持 GID 的代价是 `maxp` 数量和 `loca`/宽度表仍覆盖原编号；被移除的 glyph 是空 loca 区间，真实轮廓减少。字体名称改成内容身份名称，版权、许可等 name 记录保留，以遵守 RFN 字体改名条件。OFD 仍保存 Unicode `TextCode`；不会以图片或轮廓替换文本实现减小。

已有 CFF PDF 路径的 Poppler 探针报告字体类型不匹配，05 未修复该旧边界，不能把可保存/可抽取视为 CFF PDF 规范兼容通过。支持表只描述实际后端能力，不承诺任意字体/Word 文档保真。

## 缺字、预算及旧行为

已绑定的嵌入字体逐码点验证，缺字在写出 ZIP 前抛出 `InvalidDataException`，包含 `U+...`。PDF/SVG 对嵌入字体也用同一 coverage 验证。调用方应将文字绑定到包含字形的配置回退资源；库不会把缺字静默绘成方框。新增未知 XML 只阻止裁剪，不能绕过已建模文字的缺字检查。没有嵌入 Data 的 name-only 字体仍由阅读器/宿主解析，无法在写包时验证系统字体。

`OfdDocumentOptions.FontEmbedding` 和 `DocxConversionOptions.FontEmbedding` 控制字节、Unicode 用字和 GSUB/复合闭包操作预算；异步写包传递取消。重叠表在复制前拒绝。`OfdPackageWriteResult.FontEmbedding` 逐内容身份记录原/新体积、别名数、保留/原 glyph 数及原因；`Diagnostics` 可观察安全保留策略。

`Mode = OfdFontEmbeddingMode.Full` 是调用方显式选择的旧全量行为，用于对照或需要完整字形编辑的文档；不验证 coverage。对子集包新增文字，应重新绑定原字体 Data，并清除不再需要的旧 Runs/SourceXml，或生成新包；已有子集不能恢复已删除字形。读写旧未知资源不主动裁剪/重编号，也不删除可能被扩展引用的旧载荷。

## 许可与固定样例

子集化不授予字体嵌入权，不改变原字体许可证。自动子集模式尊重 `OS/2.fsType`：restricted/bitmap-only 禁止新嵌入，no-subsetting 保留全量；fsType 为 0 也不能代替调用方取得授权。显式 Full 以及旧包保留行为由调用方确保有权分发原载荷。

样例使用固定、可再分发的 LXGW WenKai（OFL 1.1，RFN 子集内部改名）和 Noto Sans（OFL 1.1）。来源、版本、SHA-256 和原许可见 `e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts`，不包含专有系统字体。后端未引入第三方 subset 代码；HarfBuzz/fonttools 仅用于独立开发探针，不是 SDK 运行时。

执行 `e2e/Ofdrw.Net.FontSubset.E2E` 可生成同文档的 full/subset OFD、各页 SVG、经本次 OFD 导出的 PDF、固定字体的 generated-layout native/default DOCX，以及字节/身份对照 JSON。11 包实际消费脚本也运行此样例。功能、自动渲染、PNG 和 macOS Preview 的证据分列在[验收目录](evidence/font-subset/README.md)。

补充平面（非 BMP）字形已经验证 OFD cmap/GID/Unicode 及 SVG；现有 PDFsharp 仍按 UTF-16 code unit 处理文字，无法保留这些字符。PDF 导出现在明确抛出 `NotSupportedException`，避免成功产物中的 `��`；不宣称非 BMP PDF 已支持。含 CGTransform 的旧包保持原字体和 XML，未建模的显式 GID 不用普通 Unicode coverage 错误拒绝；PDF/SVG 导出也明确拒绝未建模的 CGTransform，避免忽略其真实字形。

05 是 issue #3 的字体部分，04 的绘图 API 仍是另一半；本票不关闭或修改 issue #3 正文。与未合并 01/16 只核对现有资源绑定，不修改其他分支；集成时必须继续遵守同源字节及坐标不重算的契约。
