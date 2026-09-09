# 05 坐标系、图层与图元

上一课：[Document.xml](04-document-and-pages.md) · 下一课：[文字](06-text.md)

## 这一课结束时

能指出页面原点、图元 `Boundary` 的含义、图层绘制顺序，以及 CTM 作用在哪一层空间。

## 规范位置

8.1 坐标系统，8.5 图元对象，7.7 页对象中的 `Content` / `Layer`。

源项目布局说明也强调同一套坐标：[ofdrw-layout 坐标与图元](https://github.com/ofdrw/ofdrw/blob/master/ofdrw-layout/doc/layout/README.md)。

## 坐标

OFD 页面空间：

- 原点在页面 **左上角**
- X 向右为正，Y **向下** 为正
- 单位毫米

这与 PDF 默认用户空间（原点在左下、Y 向上）相反。从 PDF 迁坐标时必须翻转 Y。本仓库 PDF→OFD 路径按页高度做了这层转换；手写 XML 时不要把 PDF 点坐标直接贴进 `Boundary`。

规范区分设备空间、页面空间、对象空间。阅读页面 XML 时先记住两层：

1. **页面空间**：`PhysicalBox` 给出纸张。图元 `Boundary` 的 `x y` 是页面空间中的位置。
2. **对象空间**：`Boundary` 内部从 `(0, 0)` 起算。`TextCode` 的 `X/Y`、路径 `AbbreviatedData` 里的点，都在对象空间。

## Boundary 与 CTM

每个可见对象都是图元（Graphic Unit）。源项目把它们归到 `CT_GraphicUnit`。共同属性里最常用的是：

| 属性 | 作用 |
| --- | --- |
| `ID` | 对象标识 |
| `Boundary` | 对象在页面上的外接矩形（页面空间，毫米） |
| `CTM` | 可选，把对象空间映射到 Boundary 内的 6 值矩阵 `a b c d e f` |
| `Alpha` | 可选，0 透明～255 不透明 |

单位矩阵 `1 0 0 1 0 0` 表示不变换。图像几乎总会带 CTM：把图片的单位正方形缩放到 `Boundary` 的宽高，例如宽 40 mm、高 30 mm 时常写成：

```text
CTM="40 0 0 30 0 0"
```

本仓库 `OfdImageElement.Transform` 为空时，写包器就生成这种按宽高缩放的 CTM。路径坐标默认就在对象空间，再经 CTM 进页面；`OfdPathElement` 的注释写明了这一点。

图元还可以带裁剪区 `Clips`。本仓库图像模型用 `ClipsXml` 原样保留。

## 图层

页内容层次是：

```text
Page
  Content
    Layer  (可多个，按 XML 顺序绘制)
      TextObject / PathObject / ImageObject / …
```

图层 `Type` 常见取值：

| Type | 用途 |
| --- | --- |
| `Body` | 正文（默认） |
| `Background` | 背景，先画 |
| `Foreground` | 前景，后画 |

阅读器按图层顺序再按图层内图元顺序绘制，后画的盖住先画的。双层转换（页面位图 + 透明文字）就是：背景层放整页图，前景或正文层放可检索的 `TextObject`。

本仓库用 `OfdElement.LayerId` + `LayerType` 分组写 `Layer`。未指定时 `LayerType` 为 `Body`。源项目布局引擎则由 Render 把 Div 输出到对应图层。

规范还允许 `PageBlock` 嵌套，用于组合一组图元。本仓库尚未做成强类型，读到时会进 `OfdRawElement`。

## 三种常用图元

后面三课分别展开。这里只建立索引：

```text
文字  → TextObject   （第 6 课）
矢量  → PathObject   （第 7 课）
位图  → ImageObject  （第 8 课）
```

复合对象、视频等后置。遇到本仓库不认识的标签，会作为 `OfdRawElement` 整段 XML 保存，避免 round-trip 丢失。

## 动手

在 `Content.xml` 里找任意一个带 `Boundary` 的对象，在纸上画出：

1. 页面矩形（PhysicalBox）
2. 该对象的 Boundary 矩形
3. 若有 CTM，标出它缩放/平移的是对象空间而不是页面原点

本仓库对应字段：

```csharp
element.XMillimeters;      // Boundary 的 x
element.YMillimeters;      // Boundary 的 y
element.WidthMillimeters;
element.HeightMillimeters;
text.Transform;            // 可选 CTM，6 个数
```

源项目在 `CT_GraphicUnit` 上读写 `Boundary` 与 `CTM`；`ofdrw-layout` 的 Div.X/Y 指的是盒模型左上角，最终仍会变成图元 Boundary。

## 常见坑

- **Y 轴方向**。把“基线在下”的字体习惯直接套到页面 Y，文字会贴到框顶或沉底。第 6 课专门讲 `TextCode` 的 Y。
- **Boundary 为 0 宽高**。`ST_Box` 要求宽高大于 0。本仓库对未设宽度的文本会按字号估算一个框。
- **CTM 与 Boundary 重复缩放**。图像已经用 CTM 把单位方形拉到宽高，路径数据就不要再乘一遍同样的宽高。
- **图层顺序反了**。水印若放在 Body 之前会被正文盖住；放 Foreground 才会压在字上。
