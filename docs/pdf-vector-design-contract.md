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

实际 Preview 发现低分辨率原图先按 DPI 放大再输出会改变同一 viewer 的插值外观。本票因此对严格纯单幅整页 raw DeviceRGB/8 或 literal Indexed/DeviceRGB/8 图像另设回退：完整 token/EOF 必须为 `q W 0 0 H 0 0 cm /name Do Q`，正向 CTM 精确覆盖页框，无 clip/gs/第二次绘制；无过滤/Decode/mask/alpha/color override/未知图片键，保留原始 RGB 采样，或把原始索引逐像素按原调色板展开为RGB8，保留分辨率为一个 ImageObject，`IsNative=false`，没有文字。保留源 Interpolate 的有效布尔值（省略=false）于精确 `{https://ofdrw.net/image-hints}PdfInterpolateV1` SourceXml 属性。PDF/SVG 只读取根 ImageObject 的合法 true/false；非法精确值失败，未知 namespace/版本没有新语义。Core 只放行该 exact XName、正确 owner 与规范值；既有未知节点不放宽，Writer/Reader/Merge/Mix/Split/pruning 的保留须实测。Indexed仅接受整数hival0..255、长度恰为3*(hival+1)的hex/string literal lookup；省略Decode时8-bit样本按[0,255]解码并夹取到[0,hival]，合法高值不得误判为损坏。palette最多768字节，拷贝前校验长度。错误hival/palette长度为输入失败；stream lookup、其它base、Decode/ICC/DefaultRGB/过滤/mask/Intent等不进入此路径。范围外仍按既有整页双层回退或Fail。严格Fail不接纳图像，输出目标保持不变。

新路径同时保留源像素与页框按 DPI 计算的像素预算；原RGB及Indexed展开RGB分配不超过剩余图片预算，PNG 经有界输出流编码并按实际载荷累计。最终 ZIP 总字节/硬进程内存或时间限额不由这些配置承诺。完整原文仅对 native 文本页成立，扫描回退没有 OCR；第三方 OFD reader 可忽略私有插值 hint，不承诺各 renderer 使用相同插值算法。细节与真实失败/实包设计证据见 [图像回退修复设计](pdf-image-fallback-repair-design.md)。

验收使用同一许可明确 PDF 的 dual/vector 对照，核对原生对象、写出/读回原文、页尺寸、映射、字体字节和产物大小；实际检查 PDF/SVG/PNG，并以本次 OFD → PDF → macOS Preview 逐页验收。另运行生成样例 DOCX 显式 Native/default → OFD → PDF、全套回归、默认 11 包和可选真实 nupkg 干净缓存消费。证据必须区分功能、PNG 辅助、Preview 与 latest-head review/CI；本票不合并、不发布，不宣称任意 PDF 保真或 05 已完成。

原生文本也在创建19事件前限制新增producer精度误差，超范围为 `TEXT_FLOAT_PRECISION` 整页policy：回调text/ctm的double叶值精确二进制组合，逐glyph基线及规范化GlyphBounds加unit-em-square四角的绝对误差≤0.0001 PDF pt；相邻非零基线位移和matrix/字号组合线性映射扭曲≤0.00001，真实零间距保持零。比较实际发送的float矩阵/字号/advances，维护一个exact prefix及既有consumer顺序double prefix，均逐基线核对；每glyph常数工作，游程整体线性，保留取消/字数/事件/内容/字体预算。原字距缓冲改存同一已检查的float，不能检查一组却发送另一组。

该文本界限从PdfPig回调已解析的double和glyph metrics开始，不能恢复上游十进制解析、SDK文本状态/CTM累计误差，不能证明不诚实字体metric以外的ink/hinting或其它renderer。既有OFD字号decimal格式及PDF导出舍入另有误差，不将producer界限写成最终OFD/PDF像素保真保证；未修改04/19字体链、font resolver/cache/subset。Astra真实旧r4包在原九页PDF上复现1pt/4pt/.25pt及累积字距/字号损失，修复必须在同SHA原文档及正常英文/CJK/shear/spacing对照实际验证。

Indexed夹取依据ISO32000-1的8.6.6.3与8.9.5，参见[PDF Association Common Objects](https://pdfa.org/download-area/cheat-sheets/CommonObjects.pdf)。未显式声明Decode的8bit Indexed样本超过hival仍是合法输入，读取最后palette项；显式Decode、其它base和stream lookup仍按原支持范围执行明确fallback/Fail。原子Fail、展开/encoded/pixel/累计图片预算和取消不放宽。
