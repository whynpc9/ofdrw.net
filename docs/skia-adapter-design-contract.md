# 19 Skia 合作式绘制事件适配契约

起始基线：`ab87bca`（`codex/ofd-graphics`）。19 是独立可选包，不替代 [04 原语契约](graphics-design-contract.md)。04+05 仍是 issue #3 完成定义；本票不实现 05 的字体服务。

## 事件获取结论

2026-10-08 的 SkiaSharp 3.119.1 实际探针确认：`SKCanvas` 的 `Draw*` 不是托管虚方法，不能靠继承截获既有调用。`DrawText(string, ...)` 在托管层先创建 `SKTextBlob`，原生重放时没有可靠的原始 Unicode 逆映射。`SKDrawable` 是 producer 自己的绘制回调，不是观察任意 canvas 的事件钩子。

真正的 `SKSvgCanvas` 能输出线、矩形、文字，但实测原文 `Skia  native text` 的两个空格变成一个；color-matrix filter 在 bitmap 实际得到红像素 `#ffff0000`，SVG 却输出默认黑矩形且没有滤镜声明。解析已丢信息的 SVG 无法恢复原文，也不能在被忽略的效果上明确失败。

原生 C++ 自定义 SkCanvas 可以收到 SKPicture 重放事件，但需要与 SkiaSharp 所用 Skia revision/编译 ABI 完全匹配的独立 bridge、平台二进制及维护门；它仍不能从 glyph blob 恢复任意原文。因此不引入自建 Skia fork，不将 native bridge 作为本票交付路线。

本票选择**调用方显式提供的合作式事件**：已有 producer 在绘制边界产生 `SkiaDrawEvent`，再调用 `OfdSkiaAdapter.Append`。事件工厂接收 Skia 类型并快照它们；它们不执行或拦截 `SKCanvas.Draw*`。迁移样例展示 producer 显式接入事件 sink，同一 producer 的 Skia sink 仍执行真实 `SKCanvas` 绘制。任意既有 `SKCanvas`、`SKPicture` 或任意 PDF 渲染器不能直接作为输入。该范围由协调方明确接受；不是透明兼容承诺。

## 输入与映射

- 公开包 `Ofdrw.Net.Graphics.SkiaSharp` 仅依赖 Layout 和固定 SkiaSharp；Layout/Converter 不新增 Skia 引用。每个事件携带自己的完整 affine `SKMatrix`；坐标经显式 `millimetersPerUnit` 缩放（默认 1，必须正数）。原点左上、Y 向下，字体 em/笔画同样缩放，矩阵次序遵循 04。
- Line/Rectangle/Path 映射到 04 `OfdGraphicsPath` + 固色 `OfdPen`/`OfdBrush`，最终是 `PathObject`。支持 move/line/quadratic/cubic/close 与 winding/even-odd。拒绝 conic/inverse fill。所有可变路径/画笔在事件创建时复制为托管值；事件不持有 native 生命周期。
- Text 必须由 producer 提供原始 Unicode、包内唯一 `fontResourceId` 和字距（对应 04 grapheme-count-minus-one）。不从 glyph IDs、SVG、font family name 反推原文或资源。文字经 04 写 `TextObject`，空格原样保留。仅单基线文本，不执行 shaping/换行。
- `SKFont` 必须有明确 Typeface、normal width、ScaleX=1、SkewX=0、非 Embolden；真正 face 的 Bold/Italic 仅用于验证目标资源声明一致。调用方负责使用与该资源相同的实际 face（样例从相同许可载荷创建）。字体通过 `SKTypeface.OpenStream` 的 SHA256 与目标资源实际 Data 校验，且检查 face Bold/Italic 与资源元数据；同名不同载荷也拒绝。流读取有 64MB 默认上限，TTC 与含 fvar 的 variable face 明确拒绝。这是逐批载荷验证，未引入字体解析/测量/resolver/cache/subset 服务；不复制 05 服务。
- 每事件完全指定状态，没有另一套 Save/Restore API。首版不接 canvas 现有裁剪状态；producer 必须将涉及 clip/layer/image 等范围外操作作为 Unsupported 事件显式拒绝，不能漏报它们。

## 失败、容量与原子性

8-bit sRGB 固色 SrcOver only；shader/color-filter/image-filter/mask-filter/path-effect、其它 blend、hairline、非 butt cap/miter join、任意 path 的非 10 miter、text stroke、透视/奇异矩阵以及范围外操作明确失败。Skia 默认 miter=4，任意 path stroke producer 必须显式设置 10。直线没有 join，任意正 miter 无差异；矩形只在 miter≥sqrt(2) 时放行（包含默认 4），保证直角不会 bevel。没有扩大 04 API。既不静默忽略，也不隐式整页栅格化。

事件入口在复制路径/文字前检查限额、finite 和不支持项；Append 受事件、原生图元、路径、文字、几何限额及取消约束。先在临时 page 通过现有 OfdGraphics 全量预演；全部成功后按序追加。在 GetEnumerator 与每次 MoveNext 前检查取消；枚举器 Dispose 也在提交前结束。晚期坏资源/超限/取消/producer cleanup 异常均不得更改目标页、字体或既有对象。不拥有 package/page；单线程，调用方不得并发修改资源。

## 验收范围

独立包消费必须是本次真实 nupkg 的 PackageReference，默认 11 产品验证仍精确 11 个；新可选包不进入默认发布 feed。共享两个 pack 文件仅从已审查 PR15 `f7cbbcf991f33b6bc9c90cc118c636dbcca78404` 移植默认 PACKAGES 隔离逻辑，不合并功能分支。

本次样例：合作 producer → Skia 事件 → native OFD → PDF/PNG → macOS Preview；另生成实际 SKCanvas PNG 与直接 04 相同图元控制。检查原文、原生对象类型、基线、线宽、填色、曲线、矩阵、分页和局部样式；复用 licensed generated-layout DOCX Native/default 回归页。仅对实际检查页作结论；不对任意 Skia/Word/PDF 或生产发布作推断。

源码依据：[SkiaSharp 3.119.1 SKCanvas](https://github.com/mono/SkiaSharp/blob/v3.119.1/binding/SkiaSharp/SKCanvas.cs)、[Skia SVG device](https://github.com/google/skia/blob/main/src/svg/SkSVGDevice.cpp)。原始实测证据、环境/哈希及独立验收将在 [19 证据](evidence/skia/README.md) 留存。

Blender：默认 null 或 SDK 的 canonical `SKBlender.CreateBlendMode(SrcOver)` 可接受；runtime-effect/arithmetic 等自定义 Blender 即使 BlendMode getter 返回 SrcOver 也明确拒绝。枚举值不是自定义混合器语义的证明。

R5 face风格：通过现有SKTypeface native table API，仅以两次单byte读取OS/2 offset62与head offset44，前置table长度检查；Resource.Bold/Italic须匹配OS/2 bit5/0，且head bit0/1须一致。缺表/短表/矛盾或读取失败明确拒绝，不修改目标资源。Skia weight>=600/Oblique分类不用于强制字体强调；embedded Text请求始终400,false，载荷本身保留SemiBold/Oblique/真BoldItalic字形，既有导出器依据真实resourceflag选择原face而不faux。此为有界Interop元数据验证，不引入05 parser/resolver/cache/subset服务；不支持字体规范化/fallback。

既有Writer会把真实resource Bold/Italic位提升为有效CT_Text强调，Reader读取写出的有效字段；R5的400,false指输入模型未提出附加强调。写出真Bold/Italic的700/true沿用04既有fileflag继承，PDF resolver此时请求与实际位相同，MustSimulate两项false；SemiBold/Oblique不被Skia分类误提升。
