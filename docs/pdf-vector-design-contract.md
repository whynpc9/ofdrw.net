# 21 PDF 矢量转换契约

起点：19 分支 `d878aeda57c1da79917e11d0a35bcc028bc4c1fd`。先行实证见 [Astra 探针](pdf-vector-probe-design.md)。本文件描述本票实现范围，验收结果另行记录。

## 公开范围

增加独立可选包 `Ofdrw.Net.Converter.Pdf.Vector`，公开 `PdfVectorToOfdConverter`（实现既有 `IPdfToOfdConverter`）和 `PdfVectorToOfdOptions`。显式构造该转换器即选择 vector 模式；默认 `PdfToOfdConverter`、Converter 元包、Layout 和默认 CLI 均不新增 Skia 依赖，仍按原双层契约运行。首版为 API 接入，默认 CLI 不提供 vector 开关；没有把未加载的可选实现宣传成 CLI 能力。矢量包依赖既有 PDF 转换器及 19 可选适配包，不另造字体服务。

`UnsupportedPagePolicy` 为 `Fail`（默认）或显式 `RasterizePage`。每页报告源页号、实际 native/raster disposition、路径/文字/图像数量和诊断；零基页选择继续采用既有顺序、重复和无有效选择时全页的契约。转换首先构造完整包，再写输出；输入、页数、解析流、事件、路径命令、文字、字体和累计图片均有限额。取消与预算超限是失败，不能静默通过栅格回退绕过限额。

## PDF 事件来源

通过 PdfPig 0.1.15 的公开自定义 page factory 获得页字典、资源、scanner 和 parser 上下文；在受限流/字体预检后使用公开 `BaseStreamProcessor` 的绘制回调。操作符在创建阶段经过明确白名单，不由已丢失未知操作的 Page.Paths/Page.Letters 拼接顺序。路径和文字回调按实际绘制顺序生产不可变 `SkiaDrawEvent`，交给 `OfdSkiaAdapter.Append`，最终通过 04 写原生 PathObject/TextObject。19 的入口仍是合作 producer，不捕捉任意 SKCanvas/PDF renderer。

首版只接纳零原点、未旋转、CropBox=MediaBox、UserUnit=1 的页框；内容 affine CTM 和文字矩阵受支持。接纳 move/line/cubic/close/rectangle，nonzero/evenodd 固色填充以及符合 04/19 的 butt/miter=10 描边。不同填充/描边颜色按 PDF 的 fill-then-stroke 顺序生产两个事件。路径构造途中变化 CTM、clip、dash、hairline、其它 cap/join、ExtGState、image/Form XObject、shading/pattern、标记内容、注释、透明组、device-color-space 覆盖和未知操作均明确拒绝。首版只接纳无滤镜内容/字体/ToUnicode 流；压缩和其它流过滤器采用整页回退或失败，不把压缩长度误当解压内存限额。

## 原文与字体

只接纳 Type0 Identity-H / CIDFontType2 / identity CIDToGID、嵌入固定 TrueType 且具有可用 Unicode cmap 的实际载荷。每个字符要求 PdfPig TryGetUnicode 成功、单个可表达 Unicode 字符、native SKTypeface 从同一载荷选出的 glyph ID 与源 CID 相同。缺字、字形零、连字、多字符映射、组合 shaping、竖排和不兼容子集均回退/失败。字体不按 family name 猜测，不从轮廓或 glyph ID 反推原文。

同一 Tj/TJ 的兼容字符合并为原文游程，以实际基线位置计算显式字距，连续空格写在同一 TextCode。纯空白操作和不兼容基线明确回退，避免既有 Reader 丢失单独空白对象。字体按实际 FontFile2 字节注册唯一目标资源 ID；face flags 来自实际 OS/2/head 位，与 19 一致。Text 请求保持 400/false，不将 Skia SemiBold/Oblique 分类变成附加强调；不引入 resolver、共享 cache、字体 normalization/subset 或 05 替代实现。

## 回退与证据

支持页仅含 native 路径/可见原文文字。任何范围外操作使整页暂存事件和字体丢弃，再调用现有双层转换器：仅一幅整页图与透明文字，不能同时保留可见矢量文字制造重影。扫描、真正空白或无可表达 native 内容页同样有明确 disposition；有效的纯文本页可以原生保留，不以无路径等同扫描。

验收使用同一许可明确 PDF 的 dual/vector 对照，核对原生对象、写出/读回原文、页尺寸、映射、字体字节和产物大小；实际检查 PDF/SVG/PNG，并以本次 OFD → PDF → macOS Preview 逐页验收。另运行生成样例 DOCX 显式 Native/default → OFD → PDF、全套回归、默认 11 包和可选真实 nupkg 干净缓存消费。证据必须区分功能、PNG 辅助、Preview 与 latest-head review/CI；本票不合并、不发布，不宣称任意 PDF 保真或 05 已完成。
