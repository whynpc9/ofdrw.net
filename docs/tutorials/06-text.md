# 06 文字 TextObject

上一课：[坐标系与图层](05-coordinates-and-layers.md) · 下一课：[路径](07-paths.md)

## 这一课结束时

能读懂一段 `TextObject`，解释 `Font`、`Size`、`TextCode` 和 `DeltaX`，并判断这份 OFD 里的字能否被提取而不是只有一张图。

## 规范位置

第 11 章：字型、文字对象、文字定位。

这是版式文件里 **最常用** 的图元。电子发票、公文、报告的可检索正文都靠它。

## 最小例子

```xml
<ofd:TextObject ID="4"
                Boundary="20 20 80 10"
                Font="3"
                Size="4">
  <ofd:TextCode X="0" Y="4">Hello OFD</ofd:TextCode>
</ofd:TextObject>
```

对应关系：

| XML | 含义 |
| --- | --- |
| `Boundary` | 这段文字在页面上的框，毫米 |
| `Font` | `ST_RefID`，指向 `PublicRes.xml` 里的字体声明 |
| `Size` | 字号，毫米（不是磅） |
| `TextCode` | 实际 Unicode 文本 |
| `X` `Y` | 对象空间里的起点。`Y` 通常靠近字号大小，因为原点在框的左上、文字往下画 |

可选：`FillColor`、`Weight`（如 `700` 表示粗体）、`CTM`、字符间距。

本仓库写入器在没有 `Runs` 时会生成类似结构：`TextCode` 的 `X="0"`，`Y` 取 `max(字号, 1)`。颜色默认黑则省略 `FillColor`。

## TextCode 与 DeltaX

一个 `TextObject` 可以有多个 `TextCode`（多行或多次定位）。精细排版会给每个字形写位移：

```xml
<ofd:TextCode X="0.8" Y="4.25"
              DeltaX="4.95 4.91 4.95 4.91">您好世界</ofd:TextCode>
```

`DeltaX` 是 `ST_Array`：相邻字形原点在 X 方向的增量。规范还允许 `g n value` 形式的游程压缩（连续 n 个相同增量）。没有 `DeltaX` 时，阅读器按字体度量前进。

`DeltaY` 用于竖排或逐字垂直微调，正文横排较少见。

本仓库 `OfdTextRun` 把 `Text`、`XMillimeters`、`YMillimeters`、`DeltaX`、`DeltaY` 做成强类型。读入复杂对象时，完整 XML 还在 `SourceXml`，以免丢失尚未建模的字形变换。

## 字体不是 Font 属性里的名字

`Font="3"` 不是 “SimSun”。它指向资源：

```xml
<!-- PublicRes.xml -->
<ofd:Res xmlns:ofd="http://www.ofdspec.org/2016" BaseLoc="Res">
  <ofd:Fonts>
    <ofd:Font ID="3" FontName="SimSun" FamilyName="宋体">
      <ofd:FontFile>Font_3.ttf</ofd:FontFile>
    </ofd:Font>
  </ofd:Fonts>
</ofd:Res>
```

`BaseLoc="Res"` 表示字体文件相对资源 XML 所在目录再进入 `Res/`。嵌入字体才能在没有系统宋体的机器上显示中文。本仓库 `OfdFontResource.Data` 保存原始字节；当前 **不做子集化**（见功能对照）。源项目 `ofdrw-font` 负责更完整的字体处理。

只写 `FontName` 不嵌文件时，阅读器会回退到本机字体，换一台电脑就可能缺字。

## 可提取文字 vs 整页图片

判断一份 OFD 是不是“真文字”：

1. `Content.xml` 里有 `TextObject` / `TextCode`，且字符是可读 Unicode
2. 不是只有一个铺满页面的 `ImageObject`

转换器常见两种策略：

- **原生文字**：DOCX 直接写成 `TextObject`（本仓库 `DocxToOfdMode.Native` 默认如此）
- **双层**：背景层整页栅格图 + 前景透明 `TextObject` 方便搜索（PDF→OFD 默认接近这种）

抽取 API 读的是 `TextCode`，不是 OCR。本仓库 `OfdTextExtractor`；源项目 `ContentExtractor`。

```csharp
var text = new OfdTextExtractor().Extract(package, includeTemplates: true);
```

## 动手

1. 在 `Content.xml` 搜索 `TextCode`，复制一段中文确认不是乱码实体
2. 拿它的 `Font` ID 去 `PublicRes.xml` 对上 `Font` 声明
3. 若 `FontFile` 存在，确认 ZIP 里真有该文件

用 Builder 写一行字（见 [README 示例](../../README.md) 的 `OfdTextElement`）。源项目更常见的是 `new Paragraph("你好")` 交给布局引擎，由 Render 生成 `TextObject`；调试格式时仍应打开生成后的 `Content.xml`。

## 常见坑

- **字号当磅写**。五号汉字约 10.5 磅 ≈ 3.7 mm。写成 `Size="10.5"` 会大出将近三倍。
- **Y=0 导致字形顶到 Boundary 外或被裁掉**。对象空间 Y 向下，基线需要一个接近 `Size` 的正值。
- **只改了 Text 属性却留着 SourceXml**。本仓库写包时若 `SourceXml` 非空，以源 XML 为准。要让强类型字段生效需先清空 `SourceXml`。
- **DeltaX 个数和字符数不匹配**。部分阅读器会错位或丢字。
- **以为抽取失败就是文件没字**。双层文件里文字可能 `Alpha="0"` 或白色，视觉上看不见但 XML 里有。
