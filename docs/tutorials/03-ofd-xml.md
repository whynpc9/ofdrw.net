# 03 主入口 OFD.xml

上一课：[基础类型](02-basic-types.md) · 下一课：[Document.xml 与页树](04-document-and-pages.md)

## 这一课结束时

能从 `OFD.xml` 读出文档类型、元数据和真正的文档根路径，并知道签名入口挂在哪一层。

## 规范位置

7.4 主入口。

## 文件职责

包内有且仅有一份 `OFD.xml`，它不描述页面内容。它只回答：

1. 这是哪一版 OFD、哪一种子集
2. 包里有几个文档对象（`DocBody`）
3. 每个文档的元数据、根节点位置、可选的版本链和签名列表

源项目类型是 `org.ofdrw.core.basicStructure.ofd.OFD` 与 `DocBody`。本仓库由 `OfdPackageWriter` 直接写该文件，读取在 `OfdReader`。

## 最小例子

```xml
<?xml version="1.0" encoding="UTF-8"?>
<ofd:OFD xmlns:ofd="http://www.ofdspec.org/2016"
         Version="1.0"
         DocType="OFD">
  <ofd:DocBody>
    <ofd:DocInfo>
      <ofd:DocID>xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx</ofd:DocID>
      <ofd:Title>Hello OFD</ofd:Title>
      <ofd:CreationDate>2026-09-08</ofd:CreationDate>
    </ofd:DocInfo>
    <ofd:DocRoot>Doc_0/Document.xml</ofd:DocRoot>
  </ofd:DocBody>
</ofd:OFD>
```

### Version

格式版本，规范取 `1.0`。

### DocType

标识文件符合哪种子集。公开约定至少包括：

| 值 | 含义 |
| --- | --- |
| `OFD` | 符合 GB/T 33190 本标准 |
| `OFD-A` | 符合 OFD 存档相关规范（档案子集，源项目 `ofdrw-archive`） |
| `OFD-H` | 电子病历版式文档轮廓（征求意见稿要求根节点取此值） |

本仓库新建文档默认 `DocType` 为 `OFD-H`（`OfdConstants.DefaultDocType`），与 [OFD-H 补充课](11-ofd-h-overview.md) 一致，不是 GB/T 33190 正文里的通用取值。与 ofdrw 示例或只认 `OFD` 的阅读器互认时，显式设置 `OfdDocumentOptions.DocType`。写病历生效件还要满足单文档、嵌字和至少一签等约束，不能只改这一个属性。

### DocBody

一个包可以有多个 `DocBody`，从而容纳多份版式文档。每个 `DocBody` 至少包含：

- `DocInfo`：标题、作者、创建日期等元数据（本仓库映射为 `OfdMetadata`）
- `DocRoot`：指向该文档的 `Document.xml`（`ST_Loc`）

可选子节点：

- `Versions`：因注释或修订产生的版本描述
- `Signatures`：指向该文档的签名列表文件，例如 `Doc_0/Signs/Signatures.xml`

没有签章时 **不应** 出现空的 `Signatures` 节点。源项目 `OFDDir.obtainDocDefault()` 在读模式下取 **最后一个** `DocBody` 的 `DocRoot`；本仓库当前以 **首个** `DocBody` 为主。读写别人的多文档包时不要假设两边策略相同。电子病历轮廓还要求 **只用** `Doc_0`，见 [第 12 课](12-ofd-h-package-profile.md)。

## 阅读顺序

```text
打开 ZIP
  → 读 OFD.xml
    → 取 DocBody.DocRoot
      → 读 Document.xml   （下一课）
    → 若有 Signatures，记下路径，验签时再读（第 10 课）
```

不要先扫 `Pages/` 目录。页文件的权威列表在 `Document.xml` 的页树里，目录名只是常见约定。

## 动手

解压后只看 `OFD.xml`，回答：

1. `DocType` 是什么
2. 有几个 `DocBody`
3. `DocRoot` 解析后对应 ZIP 里的哪一条
4. 有没有 `Signatures`

用本仓库写元数据：

```csharp
builder.SetOptions(new OfdDocumentOptions
{
    DocType = "OFD",
    Namespace = OfdConstants.StandardNamespace,
    Metadata = new OfdMetadata
    {
        Title = "Hello OFD",
        Creator = "Ofdrw.Net tutorial"
    }
});
```

源项目在 `OFD` 对象上设置 `Version` / `DocType`，再 `setOfd` 写入虚拟容器。

## 常见坑

- **修改了 Document.xml 路径却没改 DocRoot**。入口和根节点脱节后，阅读器表现为“空文档”或直接失败。
- **多个 DocBody 只处理了第一个或最后一个**。发票类文件通常只有一个；档案或修订包可能有多个。
- **把 DocInfo.DocID 当成页对象 ID**。那是文档级标识，与图元 `ST_ID` 不是一套计数器。
- **签名节点指向了不存在的文件**。写包器若删除页面导致字节变化，本仓库会清掉失效签名声明（见转换约定）；源项目有单独的签名清理示例。
