---
title: "GB/T 48666-2026 条款强度与可判定性矩阵"
---

# GB/T 48666-2026 条款强度与可判定性矩阵

本矩阵是 `Ofdrw.Net.Medical`（OFD-H）只读评估的条款基线，修订号 `gbt48666-2026-matrix/1`。它与 `src/Ofdrw.Net.Medical/MedicalClauseMatrix.cs` 逐行一致，`MedicalContractTests.MatrixRulesAndBaselineDocumentAgree` 会在两者不一致时失败。修改任一行时同时修改两处并提升修订号。

**标准副本：**GB/T 48666—2026《电子病历版式文档技术要求》，2026-09-28 发布，2027-04-01 实施，国家密码管理局提出并归口。审阅副本为 19 页 PDF，SHA-256 `3d2f59ea5f7d03c72d2f4257e5bf3e0f110ab87214792ea526eea31f7a979e7a`。本仓库不提交该 PDF；“摘要”列是英文转述，不是规范原文。

**列含义：**

- 强度：`Shall`（应/不应）、`Should`（宜/一般）、`May`（可）、`Capability`（应支持/具备，约束生成系统的能力，单个文件只能提供证据）、`Statement`（框架或范围说明）。`Should`、`May` 与 `Capability` 的规则失败只记为 advisory，不使条款 Fail。
- 类别：`PackageDecidable`（映射规则可从包字节判定）、`PartiallyDecidable`（规则只覆盖条款一部分）、`ProcessEvidence`（需要流程、来源或机构证据）、`ExternalEvidence`（需要法规、PKI、引用标准、视觉或行业规范证据）、`SoftwareRequirement`（约束生成或阅读软件，不约束单个文件）。
- 置信度：`Unverified` 表示原文可有多种读法或与其他条款冲突，见备注。

## 标准内部的不一致与排版问题

1. **字型资源表冲突。**6.2 表 1 写明“颜色空间（ColorSpace）、字型（Font）应定义在公共资源中”（PublicRes.xml），7.2(a) 却要求“被文档页面内容使用的字型……应在文档资源列表（DocumentRes.xml）文件中注册”。两处都是“应”。本矩阵不替标准选择读法：字型位于 PageRes 时两条都判 Fail；位于 PublicRes 时 6.2(b) Pass、7.2(a) Unknown；位于 DocumentRes 时 6.2(b) Unknown、7.2(a) Pass。原因码均为 `ResourceTableConflict`。
2. **DocType 排版字符。**PDF 把取值印作 `OFD⁃H`（U+2043 连字符项目符号），这是方正飞翔排版产物。本库只写出和接受 ASCII `OFD-H`；U+2043、U+2010–U+2015、U+2212、U+FE63、U+FF0D 的变体单独报告为 `TypesetDashVariant`（Fail）。
3. **“相关要求”未列明。**6.1.2(b) 要求符合 GB/T 33190、GB/T 42133 的“相关要求”，但 GB/T 48666 自身在 DocType、Permissions、加密、动作、音视频和附件上与 GB/T 42133 的处理相反。继承账本（下文）显式列出被取代和仅在归档阶段适用的条款，其余条款按 `Inherited` 处理并标为 `Unverified`。

## 条款矩阵

| 条款 | 强度 | 类别 | 规则 | 置信度 | 摘要 | 备注 |
| --- | --- | --- | --- | --- | --- | --- |
| `5.1` | Shall | ExternalEvidence | — | Verified | Listed medical record scenarios conform or are converted |  |
| `5.2` | Statement | ProcessEvidence | — | Verified | Reference application framework: institution file system collects, generates, converts and signs |  |
| `5.3.1(a)` | May | ProcessEvidence | — | Verified | Business-system documents may be protected by e-seal or digital signature |  |
| `5.3.1(b)` | Shall | ProcessEvidence | — | Verified | Non-electronic materials become double-layer documents within the time limit |  |
| `5.3.1(c)` | Shall | ProcessEvidence | — | Verified | External records are audited, then digitized, checked or converted |  |
| `5.3.2(a)` | Should | ProcessEvidence | — | Verified | Before exchange, combine per patient and linearize, or package as a document set |  |
| `5.3.2(b)` | May | ProcessEvidence | — | Verified | Exchange may provide variants without structured data or as double-layer documents |  |
| `5.3.2(c)` | Shall | ProcessEvidence | — | Verified | Exchanged documents use a permission control mechanism |  |
| `5.3.2(d)` | Shall | ExternalEvidence | — | Verified | Exchange of personal information follows the personal information protection law |  |
| `5.3.2(e)` | May | ProcessEvidence | — | Verified | Exchange process information may be recorded as annotations |  |
| `5.3.3` | Capability | SoftwareRequirement | — | Verified | Search and location of key content, version and annotation records |  |
| `6.1.1` | Shall | PartiallyDecidable | MH-002 | Verified | Keep the OFD container + document framework |  |
| `6.1.2(a)` | Shall | PackageDecidable | MH-002 | Verified | The carrier format is OFD with a single document |  |
| `6.1.2(b)` | Shall | PartiallyDecidable | — | Unverified | Data and organization meet the relevant GB/T 33190 and GB/T 42133 requirements | “相关要求”未列明；见下方 GB/T 42133 继承账本。 |
| `6.1.3(a)` | Capability | PartiallyDecidable | MH-005 | Verified | Documents support linearization through the ZIP container |  |
| `6.1.3(b)` | Shall | PackageDecidable | MH-004 | Unverified | ZIP 6.2.0 or later; compressible content uses Deflate | 按“未压缩存储的 XML 条目不满足 Deflate 要求”理解。 |
| `6.2(a)` | Shall | PartiallyDecidable | MH-006 | Unverified | Package entries use the canonical names and folder structure of figure 3 | 表 1 称其为“典型文件”，只强制目录结构，不强制文件名。 |
| `6.2(b)` | Shall | PackageDecidable | MH-007 | Unverified | Color spaces and fonts are defined in the public resources | 6.2 表 1（PublicRes.xml）要求字型定义在公共资源；7.2(a) 要求字型在 DocumentRes 注册。字型位于 PageRes 之外时判 Unknown，待冲突澄清。 |
| `6.2(c)` | Shall | PackageDecidable | MH-003 | Verified | XML namespace matches the GB/T 33190 edition; every root declares the ofd prefix |  |
| `6.2(d)` | Should | PackageDecidable | MH-003 | Verified | Elements use the ofd prefix; attributes are unqualified |  |
| `6.3(a)` | Shall | ExternalEvidence | — | Verified | Signatures and seals comply with national law |  |
| `6.3(b)` | Shall | PartiallyDecidable | MH-014 | Verified | Cryptographic identifiers follow GB/T 33560 |  |
| `6.3(c)` | Shall | ExternalEvidence | — | Verified | Digital signatures follow GB/T 35275 |  |
| `6.3(d)` | Shall | ExternalEvidence | — | Verified | Seals and e-seal data follow GB/T 38540 |  |
| `6.3(e)` | Shall | PartiallyDecidable | MH-013 | Verified | Signature or seal protects the page description content |  |
| `6.3(f)` | Capability | SoftwareRequirement | — | Verified | Signed files accept annotations without changing protected content |  |
| `6.3(g)` | Shall | PartiallyDecidable | MH-013 | Verified | Traceable modifications are re-signed or counter-signed |  |
| `6.3(h)` | Should | ExternalEvidence | — | Verified | Signature data includes a timestamp |  |
| `7.1(a)` | Shall | PackageDecidable | MH-001 | Verified | Root DocType is OFD-H |  |
| `7.1(b)` | Shall | PartiallyDecidable | MH-006 | Verified | Fonts, images, audio, video and models are stored inside the container |  |
| `7.1(c)` | Shall | PartiallyDecidable | MH-011 | Verified | Description metadata is stored in DocInfo or DocInfo_N.xml |  |
| `7.1(d)` | Shall | PackageDecidable | MH-006 | Verified | Custom tags are stored in the Tags folder |  |
| `7.1(e)` | Shall | PackageDecidable | MH-006 | Verified | Attachments are stored in the Attachs folder |  |
| `7.1(f)` | Shall | PackageDecidable | MH-006 | Verified | Annotations are stored in the Annots folder |  |
| `7.2(a)` | Shall | PackageDecidable | MH-007 | Unverified | Fonts, draw parameters and shared resources are registered in DocumentRes; single-page resources in PageRes | 6.2 表 1（PublicRes.xml）要求字型定义在公共资源；7.2(a) 要求字型在 DocumentRes 注册。字型位于 PageRes 之外时判 Unknown，待冲突澄清。 |
| `7.2(b)` | Shall | PartiallyDecidable | MH-008 | Verified | All fonts needed for display and print are embedded unless exempt |  |
| `7.2(c)` | Shall | PackageDecidable | MH-009 | Verified | Image resources needed for display and print are in the file |  |
| `7.2(d)` | Shall | PartiallyDecidable | MH-009 | Verified | Images are BMP, JPEG, JPEG2000, TIFF, PNG or JBIG2; converted DICOM keeps clinical identifiers |  |
| `7.2(e)` | Should | PartiallyDecidable | MH-010 | Verified | Associated objects within the institution limit (about 512 MB) are embedded, larger ones external |  |
| `7.2(f)` | Shall | PartiallyDecidable | MH-010 | Verified | Embedded audio uses MP3 or another common format |  |
| `7.2(g)` | Shall | PartiallyDecidable | MH-010 | Verified | Embedded video uses MP4 or another common format |  |
| `7.2(h)` | Shall | ExternalEvidence | — | Verified | Embedded media clarity meets medical accuracy needs |  |
| `7.3` | Shall | ExternalEvidence | — | Verified | Rendering matches or approximates the traditional carrier |  |
| `7.3(a)` | Shall | PartiallyDecidable | MH-011 | Verified | Metadata identifies patient, responsible party, creation time and keywords |  |
| `7.3(b)` | Shall | PartiallyDecidable | MH-011 | Verified | File identification gives file type, secrecy level and document identifier |  |
| `7.3(c)` | Should | ProcessEvidence | — | Verified | Structured data is stored as custom tags or attachments consistent with the pages |  |
| `7.3(d)` | Should | PartiallyDecidable | MH-013 | Verified | Document carries effectiveness information (signature or seal) |  |
| `7.3(e)` | Should | ProcessEvidence | — | Verified | Scanned images on the background layer, OCR text on the body layer |  |
| `7.3(f)` | Shall | ProcessEvidence | — | Verified | Embedded associated objects have internal jump actions; external ones have external links |  |
| `7.3(g)` | Shall | PartiallyDecidable | MH-015 | Verified | Metadata entry, annotation, combination and signing are not prohibited |  |
| `7.4(a)` | Capability | SoftwareRequirement | — | Verified | Attachments can carry non-document materials such as DICOM, audio and video |  |
| `7.4(b)` | Capability | SoftwareRequirement | — | Verified | Attachments can be internal or external |  |
| `7.4(c)` | Should | PartiallyDecidable | MH-010 | Verified | Materials over the institution limit or needing special tools are external attachments |  |
| `7.5(a)` | Shall | PartiallyDecidable | MH-012 | Verified | Machine-readable information is semi-structured text such as XML or JSON |  |
| `7.5(b)` | Shall | PartiallyDecidable | MH-012 | Verified | Machine-readable files support GB 18030 characters and use UTF-8 |  |
| `7.5(c)` | Shall | ExternalEvidence | — | Verified | Multi-part machine-readable files list header, metadata and identification information | 标准未规定 schema，字段映射需行业规范。 |
| `7.5(d)` | Shall | ExternalEvidence | — | Verified | Organization supports fast retrieval and follows health-industry standards |  |
| `7.6` | Shall | ProcessEvidence | — | Verified | Versioned records keep all historical versions or store changes as internal attachments |  |
| `8(a)` | Shall | ExternalEvidence | — | Verified | Cryptographic technology and products meet national cryptography authority requirements |  |
| `8(b)` | Shall | PartiallyDecidable | MH-016 | Verified | Encryption only for temporary files; decrypt before archiving or migration |  |
| `8(c)` | Shall | PartiallyDecidable | MH-013 | Verified | Signature protection covers all content except annotation and signature lists |  |
| `8(d)` | Shall | PartiallyDecidable | MH-014 | Verified | Signatures use national commercial cryptographic algorithms |  |
| `8(e)` | Shall | ExternalEvidence | — | Verified | Verification does not depend on specific medical software |  |
| `9.1(a)` | Capability | SoftwareRequirement | — | Verified | Generator creates OFD from data and files |  |
| `9.1(b)` | Capability | SoftwareRequirement | — | Verified | Generator converts multiple formats to OFD |  |
| `9.1(c)` | Capability | SoftwareRequirement | — | Verified | Generator processes OFD per GB/T 42133 |  |
| `9.1(d)` | Capability | SoftwareRequirement | — | Verified | Generation, conversion and processing follow this standard |  |
| `9.2(a)` | Shall | SoftwareRequirement | — | Verified | Reader renders, verifies seals, controls permissions, locates annotations, shows metadata |  |
| `9.2(b)` | Shall | SoftwareRequirement | — | Verified | Reader authorizes page editing and disables it for OFD-H files | MedicalEditPolicy 为调用方分类操作；阅读器行为不评估。 |
| `9.2(c)` | Shall | SoftwareRequirement | — | Verified | Format check follows long-term preservation rules and reports unprocessed items |  |
| `9.2(d)` | Should | SoftwareRequirement | — | Verified | Reader opens external attachments |  |
| `9.2(e)` | Shall | ExternalEvidence | — | Verified | Reader verified by a third party for trusted viewing and audit |  |

## 可执行规则

| 规则 | 条款 | 可执行条件 | 不能证明 |
| --- | --- | --- | --- |
| MH-001 | 7.1(a) | 根节点无限定 `DocType` 恰为 ASCII `OFD-H` | 文件内容是否为病历 |
| MH-002 | 6.1.1、6.1.2(a) | OFD.xml 可识别；只有一个 DocBody；DocRoot 解析为 `Doc_0/Document.xml` | 多文档拆分来源 |
| MH-003 | 6.2(c)、6.2(d) | `Version` 为 1.x 时所有 OFD XML 根命名空间为 `http://www.ofdspec.org/2016` 并在根上声明 `xmlns:ofd`（结构所指条目使用其他命名空间时为 Fail，仅标签载荷与附件豁免）；元素使用 `ofd` 前缀、属性无限定（宜） | 后续 GB/T 33190 版本的命名空间（Version 非 1.x 时为 Unknown） |
| MH-004 | 6.1.3(b) | 中央目录中压缩方法只为 0 或 8，且 XML 条目不以存储方式写入 | ZIP 6.2.0 规范的全部条款 |
| MH-005 | 6.1.3(a) | 本地文件头物理顺序：OFD.xml 第一，文档在页面前，页面按顺序 | 网络加载性能；ZIP 中央目录仍在文件尾 |
| MH-006 | 6.2(a)、7.1(b)、7.1(d)–(f) | 声明位置落在图 3 目录（`Doc_0`、`Pages/Page_N`、`Signs/Sign_N`、`Attachs`、`Tags`、`Annots`、`Res`、`Temps`）；FontFile/MediaFile 载荷在包内且位于 Res 目录 | 典型文件名；未声明的游离条目 |
| MH-007 | 6.2(b)、7.2(a) | Font/ColorSpace 所在资源表（含模板页 PageRes），见上文冲突说明 | 绘制参数和多页共用资源的放置 |
| MH-008 | 7.2(b) | 转述 GB/T 42133 的 AC-017/AC-018（已用字形与嵌入载荷） | 引用标准中的豁免条件；字体授权 |
| MH-009 | 7.2(c)、7.2(d) | Image 类型 MultiMedia（含模板页 PageRes）载荷在包内，且文件签名为 BMP/JPEG/JPEG2000/TIFF/PNG/JBIG2 | DICOM 转换是否保留诊疗关联信息 |
| MH-010 | 7.2(e)–(g)、7.4(c) | 音频为 MP3、视频为 MP4（其他格式 Unknown，因原文为“等”）；嵌入对象和内部附件不超过机构限值（默认 512 MiB） | 清晰度（7.2(h)）；外部附件的可达性 |
| MH-011 | 7.1(c)、7.3(a)、7.3(b) | DocInfo 含 Author、CreationDate、Keywords、DocID，以及调用方映射的患者、文件类型、密级 CustomData | 字段语义是否正确；未映射时为 Unknown |
| MH-012 | 7.5(a)、7.5(b) | 自定义标引载荷和 XML/JSON 内部附件为严格 UTF-8，XML 良构且声明编码为 UTF-8，JSON 至少为对象/数组外形 | 7.5(c) 头信息/元数据/标识的结构；JSON 完整语法 |
| MH-013 | 6.3(e)、6.3(g)、8(c)、7.3(d) | 最后一个签名的 References（FileRef 自包根解析，与 `OfdSignatureVerifier` 一致）覆盖全部页面，且覆盖除签名目录、签名列表、注释列表和注释文件外的全部条目；存在无法解析的引用时，缺页为 Unknown | 摘要值与签名值的密码学有效性；早期签名的范围 |
| MH-014 | 6.3(b)、8(d) | SignatureMethod 为 SM2、CheckMethod 为 SM3（角色互换或重复为 Fail）；标识为 GB/T 33560 OID `1.2.156.10197.1.501`/`.401` | 印章（GB/T 38540）内部算法；证书链 |
| MH-015 | 7.3(g) | Document/Permissions 未把 Annot 或 Signature 设为 false | 元数据著录和文档组合是否被其他机制禁止 |
| MH-016 | 8(b) | 存在 Encryptions.xml 时报告 Unknown（只允许临时使用） | 文件是否为临时文件；归档前是否已解密 |

## GB/T 42133 继承账本

修订号 `gbt48666-2026-archival-inheritance/1`，数据在 `MedicalArchivalInheritance.cs`。评估在同一字节快照上运行 `Ofdrw.Net.Archival`，再按本账本解释其 105 项结果：只有 `Inherited` 条款的 Fail 会使 OFD-H 的 `StandardOutcome` 为 Fail。

| GB/T 42133 条款 | 处置 | GB/T 48666 依据 | 理由 |
| --- | --- | --- | --- |
| 6.2.1(a) | Superseded | 7.1(a) | 声明 OFD-H 而非 OFD-A |
| 6.2.2(a) | Superseded | 5.3.2(c) | 交换文档需要权限控制，GB/T 42133 移除 Permissions |
| 6.2.2(c)、6.2.3(c)、6.2.5(a)、6.3.3(b)、6.9(a)、6.9(b) | Superseded | 7.3(f) | 关联对象需要内部跳转动作和外部链接，GB/T 42133 移除非 Goto 和外部动作 |
| 6.2.6(e) | Superseded | 7.2(d) | 7.2(d) 规定自己的图像格式清单 |
| 6.2.6(g)、6.7 | Superseded | 7.2(f)、7.2(g) | 允许嵌入 MP3/MP4 音视频 |
| 6.12(a)、6.12(b) | Superseded | 6.2(a) | 表 1 保留 Extensions.xml |
| 6.14 | Superseded | 7.6 | 保留历史版本或以内部附件存储修改内容 |
| 6.15.*、附录 A.1/1–13 | Superseded | 7.4 | 保留内部和外部附件，包括 DICOM 和音视频 |
| 6.16 | Superseded | 8(b) | 临时文件可加密，归档或迁移前解密 |
| 7.1(1)、7.1(2)、7.2(1)、7.2(2) | Superseded | 9.1、9.2、9.2(b) | 第 9 章规定医疗生成和阅读软件；9.2(b) 锁定 OFD-H 的页面编辑 |
| 5.1(a)–5.2(b) | ArchivalStageOnly | 5.1 | GB/T 42133 的归档范围与组织；GB/T 48666 第 5 章另行规定场景 |
| 6.2.1(d)、6.2.1(e)、6.2.2(b)、6.2.2(e)、6.13.1(c) | ArchivalStageOnly | 7.3(g) | 组装归档 OFD 时的处理（9.1(c) 届时按 GB/T 42133 处理） |
| 6.2.2(d) | ArchivalStageOnly | 7.4 | 附件取舍属于归档处理 |
| 6.10.1(a)(1)–(3)、6.10.1(b)、6.10.2 | ArchivalStageOnly | 6.3(f) | 已签名的 OFD-H 仍需接受注释；冻结注释属于归档处理 |
| 6.13.2(a)(1)–6.13.3(b) | ArchivalStageOnly | 8(c) | OFD-H 保留有效签名和印章；去技术化属于归档处理 |
| 其余全部条款 | Inherited（Unverified） | 6.1.2(b) | 视为“相关要求”；标准未列明，读法待确认 |
