# 08 图像与资源

上一课：[路径](07-paths.md) · 下一课：[模板、附件与标引](09-templates-attachments-tags.md)

## 这一课结束时

能把页面上的 `ImageObject` 追到 `DocumentRes.xml` 再追到 ZIP 里的图片文件，并分清字体资源和图像资源各写在哪。

## 规范位置

7.9 资源，第 10 章图像，11.1 字型（字体作为资源的部分）。

## 资源是“索引 XML + 载荷文件”

页面图元不内嵌二进制。它们只引用 ID；ID 在资源文件里映射到包内文件。

典型分工：

```text
PublicRes.xml      字体、颜色空间、绘制参数
DocumentRes.xml    图片、音频/视频等多媒体
Res/               实际的 .ttf / .png / .jpg
```

两个索引 XML 的根都是 `Res`，常用 `BaseLoc` 指出载荷目录：

```xml
<ofd:Res xmlns:ofd="http://www.ofdspec.org/2016" BaseLoc="Res">
  <ofd:MultiMedias>
    <ofd:MultiMedia ID="10" Type="Image" Format="PNG">
      <ofd:MediaFile>Image_10.png</ofd:MediaFile>
    </ofd:MultiMedia>
  </ofd:MultiMedias>
</ofd:Res>
```

`MediaFile` 相对 `Res` 元素所在目录 + `BaseLoc`。若 `DocumentRes.xml` 在 `Doc_0/` 且 `BaseLoc="Res"`，则文件为 `Doc_0/Res/Image_10.png`。

源项目 `ofdrw-core` 资源类型在 `pageDescription` / 资源相关包；打包由 `ofdrw-pkg` 的文档容器放置文件。本仓库 `OfdResourceCatalog` 在写包时登记字体与图像。

## ImageObject

```xml
<ofd:ImageObject ID="11"
                 Boundary="20 80 40 30"
                 CTM="40 0 0 30 0 0"
                 ResourceID="10"/>
```

| 属性 | 含义 |
| --- | --- |
| `ResourceID` | 指向 `MultiMedia` 的 ID |
| `Boundary` | 显示区域 |
| `CTM` | 通常把图像单位正方形映射到 Boundary 尺寸 |
| `Alpha` | 可选透明度 |

本仓库默认 CTM 为 `width 0 0 height 0 0`。旋转、裁剪通过 `Transform` 和 `ClipsXml` 表达。

整页扫描件、PDF 栅格化页，都是一个铺满 `PhysicalBox` 的 `ImageObject`。双层文件会在它上面再叠透明文字（第 5、6 课）。

## 字体资源回顾

字体在 `PublicRes` 的 `Fonts/Font` 下，用 `FontFile` 指向文件，页面 `TextObject/@Font` 引用该 ID。本仓库 `OfdDocumentPackage.Fonts` 列表在写包时输出到 `PublicRes.xml`。

没有嵌入文件时，阅读器只能按 `FontName` 找本机字体。跨环境分发应嵌入，并注意字体许可。本仓库保留完整字体字节，不做子集化，所以中文字体文件会明显增大包体积。

## 颜色空间

简单 RGB 颜色常直接写在图元的 `FillColor/@Value` 上，不单独建资源。复杂文档会在 `PublicRes` 声明 `ColorSpace` 再引用。本仓库颜色模型目前是 RGB + Alpha（`OfdColor`）。渐变、Pattern 属于后置内容，读到时走 `SourceXml` / Raw。

## 动手

从任意含图的 OFD：

1. 在 `Content.xml` 找到 `ImageObject` 的 `ResourceID`
2. 打开 `Document.xml` 里的 `DocumentRes` 路径
3. 在资源 XML 中定位相同 ID 的 `MultiMedia`
4. 拼出 `MediaFile` 的包内路径，用 `unzip -l` 确认

本仓库写入图像：

```csharp
page.Elements.Add(new OfdImageElement
{
    XMillimeters = 20,
    YMillimeters = 80,
    WidthMillimeters = 40,
    HeightMillimeters = 30,
    FileName = "photo.png",
    MediaType = "image/png",
    Data = File.ReadAllBytes("photo.png")
});
```

源项目布局侧使用 `Img` 元素，由 Render 注册多媒体资源并生成 `ImageObject`。

## 常见坑

- **ResourceID 与 Font ID 混用**。它们共享文档 ID 空间，但必须指向正确的资源类型。
- **改了图片文件名没改 MediaFile**。索引和载荷脱节后显示空白。
- **CTM 单位正方形假设失败**。有的生成器 CTM 已含平移；再叠加 Boundary 原点时不要重复平移。
- **把 JPEG 标成 PNG**。`Format` 与真实编码不一致时部分阅读器解码失败。
- **删除页面后留下孤儿图片**。本仓库写包时会修剪仅被删页引用的图像；手工改 ZIP 容易漏。
