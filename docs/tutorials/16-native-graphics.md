# 16 原生绘图：毫米、基线与状态

只引用 `Ofdrw.Net.Layout` 就能运行 [完整中英两页样例](../../e2e/Ofdrw.Net.Graphics.Sample/Program.cs)，不必引用 SkiaSharp、System.Drawing 或转换器。Layout 的传递依赖为 Core/Packaging。验收工具另外引用 Reader/PDF/SVG/DOCX。

```csharp
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;

var package = new OfdDocumentPackage();
var page = new OfdPage { WidthMillimeters = 148, HeightMillimeters = 210 };
package.Pages.Add(page);
var graphics = new OfdGraphics(package, page);
var blue = new OfdColor(30, 93, 166);
graphics.DrawRectangle(new OfdPen(blue, 0.5), 10, 10, 128, 30);
graphics.DrawString("票面 Ticket", new OfdFont("Noto Sans CJK SC", 5),
    new OfdBrush(blue), 15, 25); // y=25 是基线，字体大小是毫米 em
graphics.Save();
graphics.Translate(70, 65);
graphics.Rotate(-15);
graphics.FillPath(new OfdBrush(blue), new OfdGraphicsPath()
    .MoveTo(0, 0).BezierTo(5, -10, 25, 10, 30, 0)
    .LineTo(30, 15).LineTo(0, 15).Close());
graphics.Restore();
await using var output = File.Create("ticket.ofd");
await new OfdPackageWriter().WriteAsync(package, output);
```

name-only 字体依赖阅读器环境。跨机器样例由调用方显式向 `package.Fonts` 加入许可明确的字体载荷，然后给 `OfdFont` 传入该资源 ID。ID 为目标包内身份，不能从其他包原样引用；显式未知或重复 ID 立即失败。04 不新增字体子集、fallback、测量或系统字体服务，05 继续复用现有资源/风格身份。

```csharp
package.Fonts.Add(new OfdFontResource {
    Id = "regular", FontName = "Noto Sans CJK SC",
    FileName = "Noto-Regular.ttf", Data = File.ReadAllBytes(licensedFontPath)
});
var font = new OfdFont("Noto Sans CJK SC", 4, weight: 700,
    italic: true, resourceId: "regular");
```

坐标从物理页框左上角计算，正 Y 向下。六值矩阵是 `a,b,c,d,e,f`，`x'=a*x+c*y+e`。`Multiply` 返回 `left*right`；追加变换为 `current=current*operation`。先 Translate(10,0) 再 Scale(2,2)，点(1,0)落在(12,0)。正旋转角度顺时针。线宽在用户空间，非均匀缩放及 shear 同时变换笔画和几何；默认 butt cap/miter join，miter limit 10。

路径支持直线、quadratic、cubic 与闭合，OFD 的 cubic 命令是 B，close 是 C。填充默认 NonZero，`new OfdGraphicsPath(OfdFillRule.EvenOdd)` 可画镂空。纯色含 alpha；复杂混合与设备 hairline 未建模。

R09 的图片入口接受现有静态预览 profile 可验证的 PNG 或灰度/RGB baseline JPEG；它会拒绝索引/交错 PNG、未知附加块、带 EXIF/ICC/Adobe 元数据的 JPEG 或无界载荷。目标矩形单位为毫米，可选源矩形单位为原始像素；嵌入 DPI 不改坐标。图片字节在调用时快照，同一内容多次放置由 Writer 复用媒体资源。每次图片有自身的目标裁剪，并与当前裁剪栈相交；源裁切使用非插值采样，避免邻近像素混入边缘。`alpha` 为 0–255。

```csharp
var pixels = File.ReadAllBytes("quadrants.png");
graphics.DrawImage(pixels, 10, 20, 40, 30);                 // 全图
graphics.DrawImage(pixels, 60, 20, 40, 30, 8, 0, 8, 8, 180); // 源像素裁切
var dotted = new OfdPen(blue, 1, OfdLineCap.Round, OfdLineJoin.Bevel,
    dashPattern: new[] { 4d, 2d }, dashOffset: 1);
graphics.DrawLine(dotted, 10, 70, 100, 70);
```

R09 的渐变和图片图案只支持**显式栅格化的矩形填充**，输出普通 ImageObject，并非 OFD 原生 shading/pattern。`FillLinearGradientRectangle` 将两端 RGBA 颜色按 `samples`（2–4096）采样到 PNG；`FillImagePatternRectangle` 按给定毫米 tile 尺寸重复 PNG/baseline JPEG，最多 64 个放置。它们受图片大小/像素、页面图元和几何预算约束。外来 OFD 的原生渐变、图案及任意路径 brush 仍按出口契约拒绝，不会静默取单一近似颜色。详见 [设计与证据边界](../r09-graphics-design.md)。

`DrawString` 是单基线 Unicode 游程，无折行或测量。可选 `advances` 是相邻 grapheme 原点的毫米距离，严格为 graphemeCount-1 个有限值；省略时沿用实际字体字距。文本 Boundary 使用物理整页视口，避免估计字宽造成页内裁切；文字可提取，强调补绘不复制文本。Graphics 内部把游程基线归一到 (0,size)，其中 size 采用实际写出精度，通过保真矩阵恢复调用方基线；name-only 斜体的新文本由 Writer 组合明确的生成因子 F 并记录 03 提示，PDF/SVG 只应用一次字体强调，用户矩阵不被误剥离。复杂 shaping 与任意 Word 完整保真不在本票范围。

`IntersectClip(path)` 快照当时路径及变换，裁剪固定在页面坐标；后续 Translate 不移动它。多个 Clip 取交集。`Save`/`Restore` 严格按栈恢复变换与裁剪，空栈 Restore 失败；已输出对象不变。`ResetClip` 只清理当前状态。空路径裁剪明确拒绝。每次绘制/裁剪后修改传入 path 不影响已生成对象。

`OfdGraphicsOptions` 限制页面图元、单路径命令、累计文字/几何、保存状态深度、clip 数量、单张图片编码字节及像素。有限数值、派生溢出、三位小数写出后仍为正的尺寸及保真写出的精确非奇异矩阵、字体绑定和预算均在追加前检查。绘制支持 CancellationToken；失败或取消不会追加半个对象。上下文非线程安全。

运行公开样例与完整验收链路：

```sh
python3 scripts/install-ci-fonts.py --directory artifacts/graphics-fonts
./scripts/run-graphics-e2e.sh artifacts/graphics/current
```

路径/字距/object 与 clip 矩阵输出统一保真普通十进制，无科学计数法；极小非零 advances 保留。新 name-only 样例的原生包不嵌入字体，验收 PDF 明确注册同一 OFL 字体，SVG PNG 使用隔离的许可字体目录，避免意外使用专有系统字体。

生成链路为 `公开 API → 原生 OFD → PDF/SVG`，并验证本次 DOCX 显式 Native/default → OFD → PDF。脚本输出逐页 PNG、完整原文/原生对象往返检查和哈希清单。macOS Preview 逐页验收及最新 PR 评审状态见 [证据记录](../evidence/graphics/README.md)。这是一套 OFD 原语 API；流式布局、表格、布局 Canvas 和可选 Skia 适配按各自票据交付。
