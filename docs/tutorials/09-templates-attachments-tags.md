# 09 模板、大纲、附件与自定义标引

上一课：[图像与资源](08-images-and-resources.md) · 下一课：[数字签名](10-signatures.md)

## 这一课结束时

能识别发票/公文里反复出现的底板页、阅读器目录树、包内附件，以及挂在文档上的自定义 XML。这些不是画第一行字所必需，但在真实业务文件里很常见。

## 规范位置

7.7 页对象中的模板引用；7.8 大纲；7.5 中的书签集；第 14 章跳转动作（大纲节点被点中时常用）；第 16 章自定义标引；第 20 章附件。

## 模板页

模板是 **可被多个页面引用的页面内容**，用来放表头、表格底纹、水印底图。内容结构与普通页相同（图层 + 图元），声明在 `Document.xml` 的 `CommonData` 里，再在每一页引用。

概念上：

```text
CommonData
  TemplatePage ID="…" BaseLoc="Tpls/Tpl_0/Content.xml" ZOrder="Background"
Page
  Template TemplateID="…" ZOrder="Background"
  Content
    Layer   ← 这一页自己的字
```

`ZOrder` 为 `Background` 时先画模板再画页面内容；`Foreground` 则盖在内容上（水印常用前景模板）。

源项目布局引擎大量使用模板减少重复图元。本仓库模型是 `OfdPage.Templates`（`OfdTemplateContent`：`TemplateId`、`ZOrder`、`BaseLocation`、以及展开后的 `Elements`）。读取时会解析模板内容以便抽取文字和导出 PDF/SVG；写回时未建模的模板声明走 `PreservedCommonDataElements` / `PreservedPageElements`。

合并文档时本仓库会 **拍平** 被引用的模板，避免跨包的 `TemplateID` 冲突（见转换约定）。

抽取文字时注意是否包含模板：页眉页码往往只在模板里。

```csharp
new OfdTextExtractor().Extract(package, includeTemplates: true);
```

## 大纲

大纲是阅读器侧栏里的 **目录树**，不是画在页面上的图元。点某一项，阅读器执行该节点上的动作（几乎总是跳到某页某位置）。公文、报告、多页病历会用它；单页发票通常没有。

它写在 `Document.xml` 里，和页树并列，不单独占一个包内文件：

```text
Document.xml
  CommonData
  Pages
  Outlines          ← 可选，树
    OutlineElem
      OutlineElem   ← 可嵌套
  Bookmarks         ← 另一回事，见下文
```

源项目类型是 `org.ofdrw.core.basicStructure.outlines.Outlines` / `CT_OutlineElem`（规范图 19）。本仓库没有强类型大纲模型；`OfdReader` 会把 `Outlines` 整段放进 `PreservedDocumentElements`，写回时原样保留。功能对照里的「书签/大纲」仍属未建模项。

### 节点长什么样

每个 `OutlineElem`：

| 属性/子节点 | 要求 | 含义 |
| --- | --- | --- |
| `Title` | 必选 | 侧栏显示的标题 |
| `Count` | 可选，默认 0 | 该节点下叶节点数目的参考值；以实际子节点为准 |
| `Expanded` | 可选，默认 true | 有子节点时，初始是否展开 |
| `Actions` | 可选 | 节点被激活时依次执行的动作 |
| `OutlineElem` | 可选 | 子节点，层层嵌套成树 |

最小可点目录：

```xml
<ofd:Outlines>
  <ofd:OutlineElem Title="入院记录" Count="0">
    <ofd:Actions>
      <ofd:Action>
        <ofd:Goto>
          <ofd:Dest Type="Fit" PageID="2"/>
        </ofd:Goto>
      </ofd:Action>
    </ofd:Actions>
  </ofd:OutlineElem>
  <ofd:OutlineElem Title="病程" Expanded="true" Count="2">
    <ofd:Actions>
      <ofd:Action>
        <ofd:Goto>
          <ofd:Dest Type="XYZ" PageID="5" Left="0" Top="20" Zoom="0"/>
        </ofd:Goto>
      </ofd:Action>
    </ofd:Actions>
    <ofd:OutlineElem Title="第一次病程">
      <ofd:Actions>
        <ofd:Action>
          <ofd:Goto>
            <ofd:Dest Type="FitH" PageID="5" Top="80"/>
          </ofd:Goto>
        </ofd:Action>
      </ofd:Actions>
    </ofd:OutlineElem>
    <ofd:OutlineElem Title="手术记录">
      <ofd:Actions>
        <ofd:Action>
          <ofd:Goto>
            <ofd:Dest Type="Fit" PageID="8"/>
          </ofd:Goto>
        </ofd:Action>
      </ofd:Actions>
    </ofd:OutlineElem>
  </ofd:OutlineElem>
</ofd:Outlines>
```

`PageID` 是页对象的 `ST_RefID`，必须等于页树里某项 `Page/@ID`（[第 4 课](04-document-and-pages.md)），不是从 0 起的页码，也不是 `Page_0` 目录名。

`Dest/@Type` 声明怎么对齐目标页，常见取值与 PDF 书签同类：

| Type | 阅读器大致行为 |
| --- | --- |
| `XYZ` | 定位到 `Left`/`Top`，`Zoom` 为 0 或不出现则保持当前缩放 |
| `Fit` | 整页适应窗口 |
| `FitH` | 适合宽度，保留 `Top` |
| `FitV` | 适合高度，保留 `Left` |
| `FitR` | 适合矩形 `Left Top Right Bottom` |

坐标仍是页面空间、毫米。动作模型完整定义在第 14 章。文档/页面上的 `Action` 用 `Event`（`DO`/`PO`/`CLICK`）表示何时触发；大纲节点则是 **被点中即激活**，不少文件省略 `Event`，只保留 `Goto`。档案应用（GB/T 42133）检查大纲动作时通常只接受 `Goto`，不要在目录树上挂打开 URI、播放视频。源项目 `ofdrw-archive` 有对应规则。

`Goto` 也可以不写 `Dest`，改为引用 `Bookmarks` 里的名称。

### 大纲不是书签

规范在 7.5 另有 `Bookmarks`：一组 **命名位置**（`Name` + `Dest`），供动作按名字跳转，本身不必做成树，也不一定显示在侧栏。

```xml
<ofd:Bookmarks>
  <ofd:Bookmark Name="入院首页">
    <ofd:Dest Type="Fit" PageID="2"/>
  </ofd:Bookmark>
</ofd:Bookmarks>
```

对照：

| | 大纲 `Outlines` | 书签 `Bookmarks` |
| --- | --- | --- |
| 作用 | 给人看的目录树 | 给动作用的命名锚点 |
| 结构 | 嵌套 `OutlineElem` | 平铺的 `Bookmark` |
| 标题 | `Title` | `Name` |
| 跳转 | 节点上的 `Actions` | 条目上的 `Dest` |

用本仓库保留已有大纲（读入后再写出）：不要清空 `PreservedDocumentElements`。新建大纲目前需要自行拼 XML 放进该列表，或等后续强类型 API。源项目可直接 `document.setOutlines(...)`。

## 附件

附件让 OFD 带上 **不是版面图元** 的文件：原始 DOCX、数据 XML、电子申请表等。规范第 20 章定义附件列表和每个附件的 `FileLoc`。

`Document.xml` 出现：

```xml
<ofd:Attachments>Attachs/Attachments.xml</ofd:Attachments>
```

列表文件再给出名称、格式、是否可见、载荷位置。本仓库写入 `Doc_0/Attachs/Attachments.xml` 和 `Attachs/Attach_N_…` 载荷；模型为 `OfdAttachment`（`Name`、`MediaType`、`Data` 或外部路径）。

源项目 README 将附件操作列为独立文档，和自定义元素、Canvas、水印等一起放在 `ofdrw-layout/doc` 下；入口见 [ofdrw README](https://github.com/ofdrw/ofdrw/blob/master/README.md)。

业务上常见的用法：把可编辑源模型塞进附件，版面仍用标准图元。没有附件的外来 OFD 只能当版式文件看，不能重排。

## 自定义标引

第 16 章允许在文档上挂 **自定义语义标签**，不进入绘制管线，供检索、归档、行业元数据使用。

`Document.xml` 指向 `CustomTags` 索引，索引再列出各标签 XML 的位置。本仓库用 `OfdDocumentPackage.CustomTags`（名称 → XML 字符串）写到 `Doc_0/Tags/`。

这不是注释（第 15 章）。注释有外观，标引可以完全不可见。

## 动手

拆开一份增值税发票或带页眉的公文 OFD（使用你自己的非隐私样本）：

1. `Document.xml` 是否有 `TemplatePage`
2. 页面 XML 是否有 `Template` 引用
3. 是否存在 `Outlines`；若有，抽一项的 `PageID` 去页树核对
4. 是否存在 `Attachments` / `CustomTags`

用本仓库添加附件：

```csharp
builder.AddAttachment(new OfdAttachment
{
    Name = "source.json",
    MediaType = "application/json",
    Data = Encoding.UTF8.GetBytes("{\"demo\":true}")
});
```

## 常见坑

- **把大纲当页面目录画在第一页**。`Outlines` 不产生任何 `TextObject`；封面目录要另外排版。
- **PageID 写成页码**。必须是页对象 ID。页树是 `ID="12" BaseLoc="Pages/Page_0/Content.xml"` 时，Dest 应写 `PageID="12"`。
- **大纲和书签混用一个树**。侧栏用 `Outlines`；`Goto` 按名字跳转才用 `Bookmarks`。
- **只改页面 Content、忘了模板里还有一份相同的表头**。显示正常，抽取或替换文字时漏网。
- **ZOrder 反了**。背景模板应在正文下面；水印用前景。
- **附件路径写成绝对操作系统路径**。`FileLoc` 是包内 `ST_Loc`（或规范允许的外部定位），不是 `C:\…`。
- **自定义标引当成长文本存储却塞进图元**。标引文件可以很大；不要把它写进每个 `TextObject` 的扩展属性。
