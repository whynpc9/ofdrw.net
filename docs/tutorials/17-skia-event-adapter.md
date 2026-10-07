# 可选 Skia 合作式事件适配

当应用已有使用 Skia 类型的绘图 producer，并愿意在 producer 的输出边界显式记录事件时，可引用 `Ofdrw.Net.Graphics.SkiaSharp`。普通 OFD 生成仍使用 [OfdGraphics](16-native-graphics.md)。这个包不继承、替代或拦截 `SKCanvas`，也不能直接捕捉任意 `SKPicture` 或 PDF 渲染器。

```csharp
using Ofdrw.Net.Graphics.SkiaSharp;
using SkiaSharp;

// package/page 已由调用方创建。Skia 单位由调用方定义。
using var pen = new SKPaint {
    Color = SKColors.Blue, Style = SKPaintStyle.Stroke,
    StrokeWidth = 0.5f, StrokeMiter = 10
};
var events = new[] {
    SkiaDrawEvent.Line(new(12, 40), new(136, 40), pen),
    SkiaDrawEvent.Rectangle(new(12, 58, 136, 128), pen)
};
OfdSkiaAdapter.Append(package, page, events,
    new OfdSkiaAdapterOptions { MillimetersPerUnit = 1 });
```

`Line`/`Rectangle`/`Path`/`Text` 是不可变事件工厂，接收 Skia 类型并立即验证/快照。原始 native path/paint/font 随后可修改或 Dispose。事件本身不会绘制到 Skia。调用方应在 producer 边界把同一绘制参数交给 Skia sink 或事件 sink；完整可运行迁移例子见 [E2E producer](../../e2e/Ofdrw.Net.SkiaSharp.E2E/Program.cs)。原来的 `Render(SKCanvas)` 需要显式改造，不能零修改兼容。

文字事件需要原始 Unicode、明确的目标 `fontResourceId` 与 producer 提供的字距，字距数量是 graphemeCount−1、单位是 producer 用户空间。简单非 shaping 的样例在 producer 里从同一 `SKFont` 测量各文字元素；适配器不测量、不推断原文。调用方先向 `package.Fonts` 注册许可明确的完整字体字节，并从这些相同字节创建 `SKTypeface`。Text 工厂读取有界原始 face 流计算 SHA256；Append 核对唯一 ID、目标 Data 的实际 SHA256 和 Bold/Italic 声明。同名字体的其它载荷绑定会失败；不支持 name-only、TTC、variable face、多 face 或无法取得原始流的 typeface。

每个事件有可选完整 `SKMatrix`（默认 identity）；`U * M` 把用户空间变换到毫米，包含线宽/字号。六值对应 `[ScaleX, SkewY, SkewX, ScaleY, TransX, TransY]`。文字是单基线 `TextObject`，路径是 `PathObject`。批次先通过 04 在临时页生成，全部成功才追加目标；晚期资源错误、超限、枚举/Dispose 错误或取消追加零个元素；已取消时不会启动 GetEnumerator/MoveNext。

首版只支持 8-bit sRGB 纯色 SrcOver、正线宽、butt/miter，路径 M/L/Q/cubic/close 与 winding/even-odd。任意路径要求 miter=10；直线无 join，矩形 miter≥sqrt(2)，允许常见默认 4。shader/filter/pathEffect、其它 blend、hairline、conic/inverse fill、透视/奇异矩阵、synthetic font scale/skew/embolden 都明确失败。clip/layer/image/shaped blob/text-path 不在本 profile；合作 producer 遇到这些操作必须调用 `SkiaDrawEvent.Unsupported("ClipPath")` 等明确失败，不能省略状态后只提交可见图元。

[设计与真实事件入口探针](../skia-adapter-design-contract.md) 解释为何没有通过 SVG 反解析或原生 ABI 桥透明捕获。该受限入口没有解除 21 任意 PDF 绘制事件来源的前置门；issue #3 仍要求 04+05。
