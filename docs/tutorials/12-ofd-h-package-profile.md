# 12 OFD-H 包轮廓

上一课：[OFD-H 是什么](11-ofd-h-overview.md) · 下一课：[医疗内容与资源](13-ofd-h-medical-content.md)

## 这一课结束时

能按医疗轮廓检查一份包：是否单文档、目录是否落在约定位置、XML 命名空间是否与稿一致，以及线性化/压缩有没有踩线。

## 规范位置

征求意见稿第 6 章基础格式（技术框架、一般要求、文件组织、命名、签章结构入口）。图元语义仍以 GB/T 33190 为准。

## 四层框架（和前几课的对应）

稿把 OFD 写成「容器 + 文档」，四层与本系列对应如下：

| 层 | 稿中的职责 | 本系列 |
| --- | --- | --- |
| 虚拟存储系统 | ZIP 目录与压缩 | [第 1 课](01-open-the-package.md) |
| 文档模型 | 文档、页面、大纲、资源索引 | [第 3–4 课](03-ofd-xml.md) |
| 页面内容 | 图形、图像、文字及资源 | [第 5–8 课](05-coordinates-and-layers.md) |
| 扩展特性 | 交互、安全、可扩展 | [第 9–10 课](09-templates-attachments-tags.md) 与本补充 |

## 三条硬约束

1. **承载格式就是 OFD**，不是 PDF 外壳再贴 `.ofd` 后缀。
2. **只用单文档**：`Doc_N` 固定为 `Doc_0`。GB/T 33190 允许多个 `DocBody`；病历轮廓收掉这个自由度。本仓库读写也以首个文档为主，默认 `DocumentId = "Doc_0"`。
3. **内容组织还要满足 GB/T 33190 与 GB/T 42133**。OFD-H 是加约束，不是减规范。

文件组织要求支持 **线性化**，虚拟存储用 ZIP 实现，多文件组织按 ZIP 6.2.0，可压缩内容用 **Deflate**。本仓库 `OfdDocumentOptions.EnableDeflateCompression` 默认为 true。线性化（为流式打开优化条目顺序）当前没有单独 API，不要把“能解压”说成“已线性化”。

## 命名空间

稿要求包内 XML 使用：

```text
http://www.ofdspec.org
```

标识符宜为 `ofd`；根节点声明该默认命名空间；元素带命名空间，属性不带。这与 GB/T 33190 的 `http://www.ofdspec.org/2016` 以及 ofdrw `Const.OFD_NAMESPACE_URI` **不是同一 URI**。

本仓库默认 `OfdConstants.Namespace` 正是短 URI；`StandardNamespace` 才是 `/2016`。病历交付物按本稿应写短 URI；与 ofdrw 示例或只认 `/2016` 的工具互操作时要显式选择。阅读器仍以 **该文件根节点声明** 为准（[第 2 课](02-basic-types.md)）。

## 目录怎么认

稿用一张包内规范命名图，把可选的加密、防夹带、注释、模板和必选入口画在同一棵树上。按用途而不是按文件名背：

**包根**

- 必须有且仅有 `OFD.xml`
- 文档目录固定 `Doc_0/`
- 启用防夹带时才出现 `OFDEntries.xml`（源项目 `OFDDir` 已为该文件名留常量）
- 启用加密时才出现 `Encryptions.xml`、明密文对照与解密提示文件（GM/T 0099 路径；本仓库未实现）

**文档目录内**

- `Document.xml` 文档入口
- `PublicRes.xml`：颜色空间、字型放这里
- `DocumentRes.xml`：多页共用的图像、矢量、绘制参数、模型
- `Pages/Page_N/Content.xml` 页面描述；可选 `PageRes.xml`、`Layer_M.xml`、页级 `Res`
- `Res/` 放 `Font_*.ttf`、`Image_*.png` 等载荷
- `Signs/`：签名列表与 `Sign_N`（[第 14 课](14-ofd-h-security-and-software.md)）。稿写明 **已生效** 的电子病历版式文档至少存在一个签名目录
- `Attachs/`、`Tags/`、`Annots/`、`Temps/` 仅在使用对应能力时出现
- 元数据修订才出现 `DocInfo_N.xml`；扩展说明才出现扩展信息文件

页内分层描述时，图层可以拆到 `Layer_M.xml`。本仓库目前把图层写进同一份 `Content.xml`，读到未建模文件会进 `PreservedEntries`。

模板目录稿中示例为 `Temps/Temp_M.xml`。GB/T 33190 实现里模板路径由 `TemplatePage/@BaseLoc` 决定，不必与示例目录名逐字相同，但病历包应能在 `Document.xml` 里追到模板文件。

## 最小病历包对照

```text
record.ofd
├── OFD.xml                 DocType="OFD-H"，短命名空间
└── Doc_0/
    ├── Document.xml
    ├── PublicRes.xml
    ├── DocumentRes.xml     有图才需要
    ├── Pages/Page_0/Content.xml
    ├── Res/…
    └── Signs/              生效件至少 Sign_0/
        ├── Signatures.xml
        └── Sign_0/
            ├── Signature.xml
            └── SignedValue.dat
```

本仓库 Hello World 默认带上 `OFD-H` 和短命名空间，但 **不会** 自动签一章。那只是格式默认值，不是“已生效病历”。

## 动手

```bash
unzip -l record.ofd
unzip -p record.ofd OFD.xml
```

检查清单：

1. 没有 `Doc_1`
2. 根 `DocType` 为 `OFD-H`
3. `xmlns:ofd` 是短 URI 还是 `/2016`
4. 若声称已生效，是否存在签名目录且 `OFD.xml` 的 `DocBody` 指向它
5. ZIP 条目是否 Deflate（`unzip -v` 看 method）

## 常见坑

- **多文档合并成一个包却保留多个 DocBody**。轮廓只要 `Doc_0`；多次就诊应是同一文档里的多页或附件，而不是 `Doc_1`。
- **目录名当规范路径**。权威位置仍是 XML 里的 `ST_Loc`；图上的 `Page_N` 只是惯用名。
- **防夹带/加密文件残留在未加密包里**。没有加密描述时不应出现对照表和解密种子。
- **把本仓库默认 DocType 当成已经满足第 7 章医疗要求**。还缺嵌入字体策略、签章范围、机读标引等（下一课）。
