# 01 打开 OFD：ZIP 容器与文件组织

上一课：[目录](README.md) · 下一课：[命名空间与基础类型](02-basic-types.md)

## 这一课结束时

能把一份 `.ofd` 当 ZIP 打开，认出入口文件和文档目录，并解释“OFD 是版式文档”和“PDF 也是版式文档”差在容器层。

## 规范位置

GB/T 33190 第 5 章（概述）、第 6 章（文件结构：容器方案与文件组织）。

## 核心事实

OFD（Open Fixed-layout Document）描述的是 **已经排好版的页面**：每个对象带绝对位置，文件本身不负责“插入一行后重排”。这和 Word 的流式文档不同，和 PDF 的角色相近。

容器层选择了 OOXML/ODF 同类方案：**ZIP + 包内 XML + 嵌入资源**，而不是 PDF 那种对象编号 + xref 的二进制结构。因此：

- 把 `.ofd` 改名为 `.zip` 后可以解压阅读。
- 阅读器的第一步永远是打开 ZIP，找到固定名称的入口文件 `OFD.xml`。
- 压缩已经发生在条目上，再对整个 `.ofd` 做一遍 ZIP 收益通常很小。

源项目把这一层做成虚拟容器。`ofdrw-pkg` 的 `OFDDir` 把工作目录当成包根，打包时遍历目录写入 ZIP；入口文件名写死为 `OFD.xml`。本仓库 `OfdPackageWriter` 直接往 `ZipArchive` 写条目，同样要求根上有 `OFD.xml`。

## 最小包长什么样

一份能被阅读器打开的文件，通常接近：

```text
hello.ofd          ← ZIP
├── OFD.xml        ← 包内唯一主入口，文件名不要改
└── Doc_0/
    ├── Document.xml
    ├── PublicRes.xml
    ├── DocumentRes.xml          ← 没有图片时可能省略
    ├── Pages/
    │   └── Page_0/
    │       └── Content.xml
    └── Res/
        ├── Font_1.ttf           ← 可选
        └── Image_1.png          ← 可选
```

常见变体：

- 一个包可以有多个 `Doc_N`（多文档对象）。日常文件几乎总是 `Doc_0`。本仓库读写当前以第一个 `DocBody` 为主。
- 页文件不一定叫 `Pages/Page_0/Content.xml`。页树用 `BaseLoc` 指向实际路径，有的生成器写成 `Pages/Page_0.xml`。
- 签过名的文件会多 `Signs/` 或 `Doc_0/Signs/` 一类目录，见 [第 10 课](10-signatures.md)。

源项目 `OFDDir.newDoc()` 按 `Doc_0`、`Doc_1`… 递增。本仓库默认 `OfdConstants.DefaultDocId = "Doc_0"`，页面默认写到 `Doc_0/Pages/Page_{index}/Content.xml`。

## 动手

```bash
unzip -l hello.ofd
unzip -p hello.ofd OFD.xml | head
```

检查这三件事：

1. 根目录确实有 `OFD.xml`，而不是藏在子目录里。
2. ZIP 条目路径使用 `/`，没有 `..` 逃出包根。本仓库加载时会拒绝路径穿越、限制条目数和展开大小（见 [转换约定](../conversion-contracts.md) 的资源预算）。
3. `OFD.xml` 里的 `DocRoot` 能指到真实存在的 `Document.xml`。

用本仓库生成再拆：

```csharp
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Packaging;

var builder = new OfdDocumentBuilder();
builder.AddPage(new OfdPage
{
    Index = 0,
    WidthMillimeters = 210,
    HeightMillimeters = 297,
    Elements =
    {
        new OfdTextElement
        {
            Text = "Hello OFD",
            FontName = "SimSun",
            FontSizeMillimeters = 4,
            XMillimeters = 20,
            YMillimeters = 20
        }
    }
});

await using var output = File.Create("hello.ofd");
await new OfdPackageWriter().WriteAsync(builder.Build(), output);
```

源项目等价入口是 `ofdrw-layout` 的 `OFDDoc`：加入 `Paragraph` 后关闭文档，由布局引擎生成图元再交给 `ofdrw-pkg` 打包。高层 API 不同，落盘的仍是 ZIP + `OFD.xml`。

## 常见坑

- **当成“可编辑 Word”**。OFD 没有段落重排信息。编辑器若要重排，必须另存文档模型（本仓库可用附件，见 [第 9 课](09-templates-attachments-tags.md)）。
- **忘记压缩条目的路径分隔**。包内路径是 `ST_Loc`，统一用 `/`。Windows 本地路径的 `\` 不能写进 XML。
- **入口不在根上**。部分损坏文件把 `OFD.xml` 放进文件夹；规范要求包内有且仅有这一份主入口，且文件名固定。
- **把整个 OFD 再套一层 ZIP 发给阅读器**。阅读器打开的是内层那个带 `OFD.xml` 的包。

## 和 PDF 只记这一句

PDF 用二进制对象图；OFD 用可解压的 XML 树。版式语义（固定坐标、字体、矢量、签章）两边都有，但调试 OFD 时优先 `unzip`，不要先上十六进制编辑器。
