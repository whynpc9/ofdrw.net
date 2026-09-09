# 04 Document.xml、页树与资源入口

上一课：[OFD.xml](03-ofd-xml.md) · 下一课：[坐标系与图层](05-coordinates-and-layers.md)

## 这一课结束时

能从 `Document.xml` 列出每一页的路径、默认页面大小，以及公共资源/文档资源文件在哪。

## 规范位置

7.5 文档根节点，7.6 页树，7.7 页对象（本课只到页文件入口），7.9 资源。

## 文件职责

`DocRoot` 指向的 `Document.xml` 是 **这一份版式文档的目录**。页面正文不写在这里。这里放：

- 文档级公共数据（最大 ID、默认页面区域、资源索引、模板页声明）
- 页树（有序的页面 ID + 页面文件位置）
- 大纲、权限、附件列表、自定义标引、扩展等可选入口。大纲树的读法见 [第 9 课](09-templates-attachments-tags.md)。

源项目包名 `org.ofdrw.core.basicStructure.doc`。本仓库写入逻辑在 `OfdPackageWriter.BuildEntries`，未建模的子节点进 `PreservedCommonDataElements` / `PreservedDocumentElements`。

## 最小例子

```xml
<?xml version="1.0" encoding="UTF-8"?>
<ofd:Document xmlns:ofd="http://www.ofdspec.org/2016">
  <ofd:CommonData>
    <ofd:MaxUnitID>7</ofd:MaxUnitID>
    <ofd:PageArea>
      <ofd:PhysicalBox>0 0 210 297</ofd:PhysicalBox>
    </ofd:PageArea>
    <ofd:PublicRes>PublicRes.xml</ofd:PublicRes>
    <ofd:DocumentRes>DocumentRes.xml</ofd:DocumentRes>
  </ofd:CommonData>
  <ofd:Pages>
    <ofd:Page ID="2" BaseLoc="Pages/Page_0/Content.xml"/>
    <ofd:Page ID="5" BaseLoc="Pages/Page_1/Content.xml"/>
  </ofd:Pages>
</ofd:Document>
```

相对路径的基准是 `Document.xml` 所在目录，通常是 `Doc_0/`。因此：

```text
PublicRes.xml                 →  Doc_0/PublicRes.xml
Pages/Page_0/Content.xml      →  Doc_0/Pages/Page_0/Content.xml
```

## CommonData 里先认这几个

### MaxUnitID

当前文档用过的最大对象 ID。新增图层、图元、字体、图片时必须大于它，并回写。漏更新会导致部分阅读器拒绝或后续编辑冲突。

### PageArea / PhysicalBox

默认物理页面区域，`ST_Box`，单位毫米。单页仍可在自己的 `Content.xml` 里覆盖 `Area`。本仓库用第一页的尺寸写这份默认 `PhysicalBox`。

规范里同一组区域还可以有 `ApplicationBox`、`ContentBox`、`BleedBox`。日常文件经常只给 `PhysicalBox`。裁切显示时以页对象自己的区域为准。

### PublicRes 与 DocumentRes

都是 `ST_Loc`，指向资源索引 XML（本身再指向字体、图片等文件）：

| 索引 | 习惯放什么 |
| --- | --- |
| `PublicRes.xml` | 字体、颜色空间、绘制参数（多页共享） |
| `DocumentRes.xml` | 图片、多媒体（这份文档自己的资源） |

可以没有图片就省略 `DocumentRes`。本仓库仅在存在图像资源时才写 `DocumentRes.xml`。资源 **文件** 的细节在 [第 8 课](08-images-and-resources.md)。

`CommonData` 里还可以声明模板页（`TemplatePage`）。发票、公文的表头底板常用它，见 [第 9 课](09-templates-attachments-tags.md)。

## 页树

`Pages/Page` 的顺序就是阅读顺序。属性：

- `ID`：该页对象的 `ST_ID`
- `BaseLoc`：该页内容 XML 的位置

页树是权威顺序。不要用 ZIP 里 `Page_0`、`Page_1` 的目录名排序——生成器完全可以不按这个命名。本仓库 `OfdDocumentEditor.ReorderPages` 改的是内存中的 `OfdPage` 列表，写回时按新顺序输出页树。

## 页文件入口长什么样

`BaseLoc` 指向的文件根元素是 `Page`，常见最小形态：

```xml
<ofd:Page xmlns:ofd="http://www.ofdspec.org/2016">
  <ofd:Area>
    <ofd:PhysicalBox>0 0 210 297</ofd:PhysicalBox>
  </ofd:Area>
  <ofd:Content>
    <ofd:Layer ID="3" Type="Body"/>
  </ofd:Content>
</ofd:Page>
```

`Content` 里的图层和图元是下一课。有的实现把一页写成单个 `Page_0.xml` 而不是目录；以 `BaseLoc` 为准。

## 动手

1. 从 `OFD.xml` 的 `DocRoot` 打开 `Document.xml`
2. 记下 `MaxUnitID` 和页数
3. 对每一项 `BaseLoc` 用上一课的路径规则解析，确认 ZIP 中存在该条目
4. 打开 `PublicRes.xml`，看它是否声明了字体

本仓库读取：

```csharp
await using var input = File.OpenRead("hello.ofd");
var package = await new OfdReader().ReadAsync(input);
Console.WriteLine(package.Pages.Count);
Console.WriteLine(package.PublicResourceLocation);
```

源项目走 `ofdrw-reader`，先解析 `OFD`，再按 `DocRoot` 反序列化 `Document`。

## 常见坑

- **页树缺页或 BaseLoc 404**。这比 XML 写错图元更快让阅读器空白。
- **只改了 Content.xml 没改页树**。复制页面文件后必须新增 `Pages/Page` 节点并分配新 ID。
- **MaxUnitID 偏小**。手工加了一个 `TextObject ID="99"` 却把 `MaxUnitID` 留在 `7`，属于不合规文件。
- **Resource 路径写成包根绝对路径却漏了前导 `/`**。相对路径会再叠一层 `Doc_0/`。
