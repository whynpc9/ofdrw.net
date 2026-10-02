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
| TTC | `CollectionFaceIndex` 零基显式选面，默认 0；`MaximumCollectionBytes` 限制容器，`MaximumFontBytes` 限制所选 face，提取前先检查投影输出尺寸；按实际 face 内容复用；写出独立 sfnt，PDF/SVG 使用相同选面 |
| CFF/CFF2 OTF、变量字体、AAT/颜色/未知依赖表、JSTF | 全量保留，返回 `FONT_FULL_PRESERVED` 与原因；不宣称这些格式已子集化 |
| RTL/bidi 字符 | 按 Unicode 17.0.0 Bidi_Class（含规范默认值）识别 RTL/控制字符；BOM/ZWNBSP 和 LTR 脚本不误判。缺少通用 Unicode mirror closure 的 RTL 保留全量并诊断；不宣称复杂文字子集支持完整 |
| 大 BMP 字符集 | 压缩格式 4 连续 delta 段并保留格式 12；格式 4 无法容纳时全量保留，避免现有 PDFsharp 拒绝 format12-only 字体 |
| SourceXml/CGTransform/Raw、模板/注释、保留包条目 | 整包保守保留全量字体，避免裁掉未知 GID/扩展引用；已有子集保持原字节；诊断会说明 coverage 未验证的未解析字体 |

保持 GID 的代价是 `maxp` 数量和 `loca`/宽度表仍覆盖原编号；被移除的 glyph 是空 loca 区间，真实轮廓减少。共享样式解析按 `OS/2.fsSelection` 优先、缺省时 `head.macStyle` 回退，与 DOCX catalog 一致，PDF/SVG 不重复合成已存在的粗斜体。字体名称改成内容身份名称，版权、许可等 name 记录保留，以遵守 RFN 字体改名条件。OFD 仍保存 Unicode `TextCode`；不会以图片或轮廓替换文本实现减小。

已有 CFF PDF 路径的 Poppler 探针报告字体类型不匹配，05 未修复该旧边界，不能把可保存/可抽取视为 CFF PDF 规范兼容通过。支持表只描述实际后端能力，不承诺任意字体/Word 文档保真。

## 缺字、预算及旧行为

已绑定的嵌入字体逐码点验证。Unicode 17 Default_Ignorable_Code_Point 格式控制符不要求 cmap 轮廓（UVS 序列仍验证）；RTL 控制仍触发 full 保护。普通缺字在写出 ZIP 前抛出 `InvalidDataException`，包含 `U+...`。PDF/SVG 对嵌入字体也用同一 coverage 验证。调用方应将文字绑定到包含字形的配置回退资源；库不会把缺字静默绘成方框。新增未知 XML 只阻止裁剪，不能绕过已建模文字的缺字检查。没有嵌入 Data 的 name-only 字体仍由阅读器/宿主解析，无法在写包时验证系统字体。

`OfdDocumentOptions.FontEmbedding` 和 `DocxConversionOptions.FontEmbedding` 控制字节、Unicode 用字和 GSUB/复合闭包操作预算；异步写包传递取消。重叠表在复制前拒绝。`OfdPackageWriteResult.FontEmbedding` 逐内容身份记录原/新体积、别名数、保留/原 glyph 数及原因；`Diagnostics` 可观察安全保留策略。

`Mode = OfdFontEmbeddingMode.Full` 是调用方显式选择的旧全量行为，用于对照或需要完整字形编辑的文档；不验证 coverage。对子集包新增文字，应重新绑定原字体 Data，并清除不再需要的旧 Runs/SourceXml，或生成新包；已有子集不能恢复已删除字形。读写旧未知资源不主动裁剪/重编号，也不删除可能被扩展引用的旧载荷。

## 许可与固定样例

子集化不授予字体嵌入权，不改变原字体许可证。自动子集模式尊重 `OS/2.fsType`：restricted/bitmap-only 禁止新嵌入，no-subsetting 保留全量；fsType 为 0 也不能代替调用方取得授权。Full 和未知内容保留也检查可解析字体的这些标记；无法解析的旧不透明载荷仅可保留并诊断。调用方仍须确保有权分发原载荷。

样例使用固定、可再分发的 LXGW WenKai（OFL 1.1，RFN 子集内部改名）和 Noto Sans（OFL 1.1）。来源、版本、SHA-256 和原许可见 `e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts`，不包含专有系统字体。后端未引入第三方 subset 代码；HarfBuzz/fonttools 仅用于独立开发探针，不是 SDK 运行时。

执行 `e2e/Ofdrw.Net.FontSubset.E2E` 可生成同文档的 full/subset OFD、各页 SVG、经本次 OFD 导出的 PDF、固定字体的 generated-layout native/default DOCX，以及字节/身份对照 JSON。11 包实际消费脚本也运行此样例。功能、自动渲染、PNG 和 macOS Preview 的证据分列在[验收目录](evidence/font-subset/README.md)。

补充平面（非 BMP）字形已经验证 OFD cmap/GID/Unicode 及 SVG；现有 PDFsharp 仍按 UTF-16 code unit 处理文字，无法保留这些字符。PDF 导出现在明确抛出 `NotSupportedException`，避免成功产物中的 `��`；不宣称非 BMP PDF 已支持。含 CGTransform 的旧包保持原字体和 XML，未建模的显式 GID 不用普通 Unicode coverage 错误拒绝；PDF/SVG 导出也明确拒绝未建模的 CGTransform，避免忽略其真实字形。

05 是 issue #3 的字体部分，04 的绘图 API 仍是另一半；本票不关闭或修改 issue #3 正文。与未合并 01/16 只核对现有资源绑定，不修改其他分支；集成时必须继续遵守同源字节及坐标不重算的契约。

`SourceXml` 必须是一个合法 XML 元素。无法解析的 XML 在 Full/subset 两种模式都于 ZIP 写出前明确拒绝，保留调用方原字节；“未知内容保护”不承诺把非法 XML 写成合法 OFD。

PDF 对 ZWJ/ZWNJ 的连接/连字语义明确拒绝；可安全省略的零宽格式控制才跳过绘制，OFD 的 Unicode 原文不变，显式 Runs/Delta 坐标槽位不重排。PDF 的可见文本抽取忽略可省略的格式控制符；方向格式控制及 UVS 语义由当前 PDFsharp 无法保真，因此明确拒绝这些 PDF 导出，OFD/SVG 保留原文和语义。不能把 coverage 豁免或文件生成成功视为控制字符的视觉通过。

绘制过滤只省略零宽格式控制符或已知 cmap 无字形的 default-ignorable；已映射的 Hangul filler 保留真实字宽，包括无 Delta 的字符串和每个 gap 明确 Delta 的游程。Unicode方向格式控制全部（含 U+202C）以及蒙古文 free variation selectors 的 PDF 语义明确拒绝，避免删除控制符后改变排列或变体。

Name-only 资源若通过本地发现或 host style probe 取得真实字节，注册与 coverage 一起绑定并缓存；无字形 filler 不被误当作 mapped。LRM/已弃用零宽控制按可省略策略绘制，明确拒绝范围仅含需要当前引擎未实现排列语义的九个方向格式控制，避免不必要拒绝。

发现的 name-only face 缺普通字符时，宿主默认回退获得机会；回退仍必须覆盖实际文字，无法覆盖则明确失败，绝不返回已知缺字的 face 画方框。Windows symbol cmap 3/0 格式 4 保持全量并诊断，读取支持直接字符与 F000 重映射；SVG 未使用的其它合法未建模 cmap 不阻断 CSS 资源输出，真正被选择且无法验证的编码则明确拒绝。RLM/ALM 的 PDF 方向语义仍拒绝，LRM/弃用零宽按省略策略处理。

名称字体的候选回退按配置 default、Arial 顺序探测真实字节；空宿主结果、读取/格式/cmap 或注册预算失败可尝试下一候选，TTC 使用 face 0。只有已验证覆盖全部当前文本且注册同一字节的候选才返回成功；全部失败明确拒绝，不用未验证字体绘制缺字框。未使用的合法但未建模 cmap 不阻断 PDF，实际选中后明确拒绝。Symbol 写包保全量且诊断 coverage 未验证，普通 Unicode 字体缺字检查仍严格。

Windows symbol-only cmap 的字符语义未建模：写包/读取编辑保全量并标识 coverage 未验证；PDF/SVG 在实际选中时统一 NotSupportedException，包含偶然映射成功的 A，避免半验证或缺字回退建议。未使用的 symbol 资源不阻断 PDF/SVG。字体用字绑定在子集阶段为每个文本只解析一次，并按实际 face 内容组汇总；不匹配的元素也先检查取消。

同一 symbol 边界也适用于已实际探测的名称字体：记录 unsupported 标记须位于可选探测 catch 之外，选用后拒绝，而不能退回未验证宿主绘制。默认/Arial 候选若解析为 symbol，直接跳过并尝试后续 Unicode 候选；全部不覆盖仍明确失败。后续 R14 将同一惰性验证扩展至所有实际选中的普通名称字体，包括无 FontResource 的文本；未使用资源和实际空文本不触发该验证。

所有实际非空 PDF 文本必须返回经过当前文本覆盖验证的内容字体身份。普通宿主探测失败可尝试已验证默认字体/Arial，所有候选失败明确报错；删除未经验证的绘制 fallback。物理 face 快照按内容复用，增加粗体/斜体时登记同一字节的有效样式 alias，不落回宿主；覆盖缓存只缓存 cmap 解析，逐文本验证不能省略。

宿主返回 TTC 时，主选字体需按返回 face 名与 name IDs4/6 或唯一 family IDs1/16 匹配，多面集合不凭 index0 或请求粗斜体猜选；不明/歧义按探测失败处理。配置默认的单面集合可采用唯一面。名称目录及解码工作有界，只有选定后才展开完整 face；参见 [OpenType name 规范](https://learn.microsoft.com/en-us/typography/opentype/spec/name)。
