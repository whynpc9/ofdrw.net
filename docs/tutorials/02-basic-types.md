# 02 命名空间与基础类型

上一课：[打开 OFD](01-open-the-package.md) · 下一课：[主入口 OFD.xml](03-ofd-xml.md)

## 这一课结束时

能解释为什么元素带 `ofd:` 前缀、属性不带前缀，并能手读 `Boundary`、`ID`、`BaseLoc` 这类属性。

## 规范位置

7.1 命名空间，7.2 字符编码，7.3 基础数据类型。

## 命名空间

标准命名空间 URI 是：

```text
http://www.ofdspec.org/2016
```

约定：

- 包内 XML **根元素**声明默认或 `ofd` 前缀指向该 URI。
- **元素**使用命名空间（写成 `ofd:TextObject` 或默认命名空间下的 `TextObject`）。
- **属性不使用命名空间**（`Boundary="…"`，不是 `ofd:Boundary`）。

源项目把它写死在 `org.ofdrw.core.Const`：

```java
public static final String OFD_NAMESPACE_URI = "http://www.ofdspec.org/2016";
public static final String OFD_VALUE = "ofd";
```

本仓库同时认识两种 URI：

| 常量 | 值 | 用途 |
| --- | --- | --- |
| `OfdConstants.StandardNamespace` | `http://www.ofdspec.org/2016` | 与标准和 ofdrw 一致 |
| `OfdConstants.Namespace` | `http://www.ofdspec.org` | `OfdDocumentOptions` 的默认值 |

现场文件两种都常见。阅读时以 **该文件根节点声明的 URI** 为准，不要假设所有文件都是 `/2016`。本仓库 `OfdReader` 按实际根节点读取；若要写出与 ofdrw 互认的包，生成时把 `Options.Namespace` 设成 `StandardNamespace`。

短 URI `http://www.ofdspec.org` 并不是笔误：电子病历轮廓征求意见稿把包内命名空间写成这一值，这也是本仓库默认值的来源。通用 33190 文件与 ofdrw 核心常量仍是 `/2016`。病历交付与通用互认的取舍见 [第 11–12 课](11-ofd-h-overview.md)。

字符编码为 UTF-8。`xs:date` / `xs:dateTime` 在源项目 `Const` 里分别格式化为 `yyyy-MM-dd` 与 `yyyy-MM-dd'T'HH:mm:ss`。

## 六个天天见到的类型

这些类型出现在几乎所有属性里。值都是 **文本**，没有独立的二进制编码。

### ST_ID / ST_RefID

无符号整数标识。`ST_ID` 在文档内应唯一；`0` 表示无效。`ST_RefID` 指向已经存在的 `ST_ID`。

页、图层、图元、字体、图片都占用 ID。`Document.xml` 的 `MaxUnitID` 记录当前用过的最大 ID，继续往文件里加对象时必须从它之后分配。本仓库 `OfdIdAllocator` 写包时会扫描已有 ID 再递增。

### ST_Loc

包内路径。以 `/` 开头表示从包根算的绝对路径，否则相对 **当前 XML 所在目录**。`.` 与 `..` 可用，但不能跳出包根。

```text
DocRoot 值为 Doc_0/Document.xml          ← 相对 OFD.xml 所在目录（包根）
Page BaseLoc 值为 Pages/Page_0/Content.xml  ← 相对 Document.xml 所在的 Doc_0/
```

源项目类型是 `org.ofdrw.core.basicType.ST_Loc`。本仓库解析在 `OfdPackagePath.Resolve`：把 `\` 换成 `/`，处理 `.` / `..`，逃出根目录则抛错。

### ST_Pos

一个点：`x y`，单位毫米，空格分隔。

```text
0 0
20.5 15
```

### ST_Box

矩形：`x y width height`，单位毫米。前两个是左上角，后两个必须大于 0。

```text
10 10 50 20
```

这是页面 `PhysicalBox`、图元 `Boundary` 的写法。源项目 `ST_Box` 与本仓库 `BuildBox` 输出的都是这四个数字。

### ST_Array

用空格分隔的一维数组，元素不能再嵌套 `ST_Array` 或 `ST_Loc`。最常见的是 6 个数的变换矩阵 CTM：

```text
1 0 0 1 0 0
```

以及颜色 `Value="255 0 0"`、字间距 `DeltaX="4.2 4.2 4.2"`。

## 实际 XML 例子

下面三份文件来自同一份最小包，六种类型都出现了。注释只为阅读，真实 OFD 里通常没有这些说明。

包根 `OFD.xml`：元素在命名空间里，`Version` / `DocType` 是普通属性；`DocRoot` 是相对包根的 `ST_Loc`。

```xml
<?xml version="1.0" encoding="UTF-8"?>
<ofd:OFD xmlns:ofd="http://www.ofdspec.org/2016"
         Version="1.0"
         DocType="OFD">
  <ofd:DocBody>
    <ofd:DocInfo>
      <ofd:Title>基础类型示例</ofd:Title>
      <ofd:CreationDate>2026-09-09</ofd:CreationDate>
    </ofd:DocInfo>
    <!-- ST_Loc：相对 OFD.xml 所在目录（包根） -->
    <ofd:DocRoot>Doc_0/Document.xml</ofd:DocRoot>
  </ofd:DocBody>
</ofd:OFD>
```

`Doc_0/Document.xml`：`MaxUnitID` 与页 `ID` 是 `ST_ID`；`PhysicalBox` 是 `ST_Box`；`PublicRes` 和 `BaseLoc` 是相对 **本文件所在目录** `Doc_0/` 的 `ST_Loc`。

```xml
<?xml version="1.0" encoding="UTF-8"?>
<ofd:Document xmlns:ofd="http://www.ofdspec.org/2016">
  <ofd:CommonData>
    <!-- ST_ID：本文件用过的最大对象号 -->
    <ofd:MaxUnitID>6</ofd:MaxUnitID>
    <ofd:PageArea>
      <!-- ST_Box：x y width height，毫米；A4 -->
      <ofd:PhysicalBox>0 0 210 297</ofd:PhysicalBox>
    </ofd:PageArea>
    <!-- ST_Loc：相对 Doc_0/ → Doc_0/PublicRes.xml -->
    <ofd:PublicRes>PublicRes.xml</ofd:PublicRes>
  </ofd:CommonData>
  <ofd:Pages>
    <!-- ID = ST_ID；BaseLoc = ST_Loc → Doc_0/Pages/Page_0/Content.xml -->
    <ofd:Page ID="2" BaseLoc="Pages/Page_0/Content.xml"/>
  </ofd:Pages>
</ofd:Document>
```

`Doc_0/Pages/Page_0/Content.xml`：图元 `ID` 是 `ST_ID`，`Font` 是指向字体的 `ST_RefID`；`Boundary` 是 `ST_Box`；`TextCode` 的 `X`/`Y` 是对象空间里的点（与 `ST_Pos` 相同写法）；`CTM`、`DeltaX`、`Value` 都是 `ST_Array`。

```xml
<?xml version="1.0" encoding="UTF-8"?>
<ofd:Page xmlns:ofd="http://www.ofdspec.org/2016">
  <ofd:Area>
    <ofd:PhysicalBox>0 0 210 297</ofd:PhysicalBox>
  </ofd:Area>
  <ofd:Content>
    <ofd:Layer ID="3" Type="Body">
      <ofd:TextObject ID="4"
                      Boundary="20 20 48 6"
                      Font="1"
                      Size="4"
                      CTM="1 0 0 1 0 0">
        <ofd:FillColor Value="32 32 32"/>
        <!-- X Y 相当于 ST_Pos；DeltaX 是 ST_Array，四个增量对应后面四个间隙 -->
        <ofd:TextCode X="0" Y="4" DeltaX="4 4 4 4">Hello</ofd:TextCode>
      </ofd:TextObject>
      <ofd:PathObject ID="5"
                      Boundary="20 30 40 0.4"
                      LineWidth="0.2"
                      Stroke="true"
                      Fill="false">
        <ofd:StrokeColor Value="0 0 0"/>
        <ofd:AbbreviatedData>M 0 0 L 40 0</ofd:AbbreviatedData>
      </ofd:PathObject>
    </ofd:Layer>
  </ofd:Content>
</ofd:Page>
```

`Doc_0/PublicRes.xml` 里被 `Font="1"` 引用的字体（`ST_RefID` → `ST_ID`）：

```xml
<?xml version="1.0" encoding="UTF-8"?>
<ofd:Res xmlns:ofd="http://www.ofdspec.org/2016" BaseLoc="Res">
  <ofd:Fonts>
    <!-- ID=1 供上面 TextObject/@Font 引用；FontFile 相对 Doc_0/ + BaseLoc -->
    <ofd:Font ID="1" FontName="SimSun">
      <ofd:FontFile>Font_1.ttf</ofd:FontFile>
    </ofd:Font>
  </ofd:Fonts>
</ofd:Res>
```

对照表：

| 出现位置 | 值 | 类型 |
| --- | --- | --- |
| `OFD.xml` 的 `DocRoot` | `Doc_0/Document.xml` | `ST_Loc`（相对包根） |
| `Page/@BaseLoc` | `Pages/Page_0/Content.xml` | `ST_Loc`（相对 `Doc_0/`） |
| `Res/@BaseLoc` | `Res` | `ST_Loc`（相对 `PublicRes.xml` 所在目录） |
| `MaxUnitID`、`Layer/@ID`、`TextObject/@ID`、`Font/@ID` | `6`、`3`、`4`、`1` | `ST_ID` |
| `TextObject/@Font` | `1` | `ST_RefID`（必须已有对应 `ST_ID`） |
| `PhysicalBox`、`Boundary` | `0 0 210 297`、`20 20 48 6` | `ST_Box` |
| `TextCode/@X` `@Y` | `0` `4` | 与 `ST_Pos` 相同的 `x y` 写法 |
| `CTM` | `1 0 0 1 0 0` | `ST_Array`（6 项，单位矩阵） |
| `DeltaX` | `4 4 4 4` | `ST_Array`（字形间距，毫米） |
| `FillColor/@Value` | `32 32 32` | `ST_Array`（RGB） |

`FontFile` 解析为 `Doc_0/Res/Font_1.ttf`：先取 `PublicRes.xml` 所在目录 `Doc_0/`，再叠 `BaseLoc="Res"`，再加上 `Font_1.ttf`。

## 单位

未另行声明时，页面尺寸、坐标、字号、线宽都以 **毫米** 计。A4 常见为 `210 297`。源项目布局文档和本仓库 `OfdPage.WidthMillimeters` 都沿用毫米，不要把像素或磅直接写进 `Boundary`。

1 英寸 = 25.4 mm。本仓库 PDF 渲染里有 `millimeters * 72 / 25.4` 的磅换算，那是输出到 PDF 时才发生的，OFD XML 里仍然是毫米。

## 动手

先在上一节的四份 XML 上把每个属性标回 `ST_*`（用上面的对照表核对）。再打开上一课生成的真实 `OFD.xml` / `Document.xml` / `Content.xml`，看生成器有没有省略 `CTM` 或 `DeltaX`——省略时阅读器按默认单位矩阵和字体度量前进，类型本身没变。

## 常见坑

- **把 Boundary 当成像素框**。阅读器按毫米映射到屏幕；在 96 DPI 下 1 mm ≈ 3.78 px，但文件里不要写像素。
- **相对路径算错基准**。`Document.xml` 里的 `PublicRes` 相对的是 `Doc_0/`，不是包根。
- **ID 冲突**。合并两份 OFD 时必须重映射 ID 和 `MaxUnitID`。本仓库 `OfdDocumentMerger` 做字体身份重映射；源项目 `ofdrw-tool` 的合并同理。
- **命名空间只写在一部分文件**。有的生成器 `OFD.xml` 用 `/2016`，页面用无前缀默认命名空间，这合法；解析时要看 **当前文件根节点** 的声明。
