# 07 路径 PathObject

上一课：[文字](06-text.md) · 下一课：[图像与资源](08-images-and-resources.md)

## 这一课结束时

能读 `AbbreviatedData` 画出矩形和直线，并区分描边与填充。这是表格线、下划线、色块、简单图标的来源。

## 规范位置

第 9 章图形（图形对象、填充规则、路径命令）。紧缩写法对应页面描述里的 AbbreviatedData。

## 最小例子

一条从对象左上到右下的线：

```xml
<ofd:PathObject ID="8"
                Boundary="20 40 80 0.5"
                LineWidth="0.2"
                Stroke="true"
                Fill="false">
  <ofd:AbbreviatedData>M 0 0 L 80 0</ofd:AbbreviatedData>
</ofd:PathObject>
```

一个填充矩形（表格单元格底色常用）：

```xml
<ofd:PathObject ID="9"
                Boundary="20 50 170 12"
                Stroke="false"
                Fill="true">
  <ofd:FillColor Value="240 240 240"/>
  <ofd:AbbreviatedData>M 0 0 L 170 0 L 170 12 L 0 12 C</ofd:AbbreviatedData>
</ofd:PathObject>
```

路径坐标在 **对象空间**（相对 Boundary 左上角）。本仓库 `BuiltInOfdRenderer` 画单元格就是这种 `M … L … L … L … C`。

## 常用命令

紧缩路径是命令字母 + 数字，空格分隔。大小写：大写为绝对坐标，小写为相对上一终点。本仓库 `OfdPathRenderer` 识别的命令包括：

| 命令 | 作用 | 参数 |
| --- | --- | --- |
| `M` | 移动到，开始子路径 | x y |
| `L` | 直线到 | x y |
| `H` / `V` | 水平 / 垂直线 | x 或 y |
| `B` | 三次贝塞尔（OFD 用 `B`，不是 SVG 的 `C`） | 两个控制点 + 终点 |
| `Q` | 二次贝塞尔 | 控制点 + 终点 |
| `A` | 椭圆弧 | rx ry φ large sweep x y |
| `C` 或 `Z` | 闭合当前子路径 | 无 |

注意：在 OFD 紧缩数据里 **`C` 表示 Close**，不是 SVG 的 cubic Bézier。三次曲线用 `B`。本仓库 SVG 导出时会把这套命令规范化成 SVG `d`。

`S` 出现在部分生成器里表示平滑曲线；需要完整命令表时以标准第 9 章为准。遇到本仓库渲染器 `default` 分支无法识别的字母，路径会画不出来，但 XML 仍应保留。

## 描边与填充

| 属性 | 含义 |
| --- | --- |
| `Stroke` | 是否描边 |
| `Fill` | 是否填充 |
| `LineWidth` | 线宽，毫米 |
| `StrokeColor` / `FillColor` | `Value="R G B"`，可选 `Alpha` |

线宽也在对象空间，再经 CTM 缩放。CTM 带放大时，视觉线宽会变粗；本仓库 PDF 导出按矩阵平均缩放补偿线宽。

填充规则（非零环绕 / 奇偶）在复杂自交路径才重要。简单矩形用默认即可。

## 和文字、图像一起用时的顺序

表格通常是：

1. `PathObject` 填充单元格背景
2. `PathObject` 画边框
3. `TextObject` 画单元格文字

图层内顺序就是绘制顺序。背景色 path 必须出现在文字之前。

## 动手

在一份转换出来的 OFD 里搜索 `PathObject`：

1. 数一数有多少用于下划线和矩形
2. 把一段 `AbbreviatedData` 按命令拆开，在 Boundary 坐标系手绘
3. 改 `Fill="true"` 后用阅读器看是否变成色块（在副本上改）

本仓库模型：

```csharp
new OfdPathElement
{
    XMillimeters = 20,
    YMillimeters = 50,
    WidthMillimeters = 170,
    HeightMillimeters = 12,
    Stroke = false,
    Fill = true,
    FillColor = new OfdColor(240, 240, 240),
    AbbreviatedData = "M 0 0 L 170 0 L 170 12 L 0 12 C"
};
```

源项目可用 `ofdrw-layout` 的 Div 边框、或 `ofdrw-graphics2d` 的 `Graphics2D` 画路径，最终仍落成 `PathObject`。调试格式时看 XML，不要只看 Java API。

## 常见坑

- **用 SVG 的 `C` 写三次贝塞尔**。OFD 阅读器会当成闭合，路径直接封口。
- **点写在页面空间**。数据必须相对 Boundary；已经把页坐标写进 AbbreviatedData 的话，图形会跑出框或缩在一角。
- **LineWidth 用像素**。`0.353` mm 约等于 1 pt，是常见默认（本仓库路径默认线宽就是 `0.353`）。
- **只 Stroke 不 Fill 的闭合框漏了最后一条边**。闭合应用 `C`/`Z`，或显式 L 回起点。
