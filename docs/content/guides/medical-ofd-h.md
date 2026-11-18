---
title: "电子病历 OFD-H（GB/T 48666）"
description: "可选包 Ofdrw.Net.Medical：按 GB/T 48666-2026 配置 OFD-H 写出，并对 OFD-H 文件做只读、分条款评估；说明能检查什么、不能证明什么。"
sidebar:
  order: 177
  label: "电子病历 OFD-H"
---

`Ofdrw.Net.Medical` 是自 0.5.0 起提供的可选包。它面向 GB/T 48666—2026《电子病历版式文档技术要求》（2027-04-01 实施）定义的电子病历版式文档，即根节点 `DocType="OFD-H"` 的 OFD，提供两部分能力：

- **写出 profile**：把 `OfdDocumentOptions` 配成 OFD-H 所需的声明、目录、命名空间、压缩、线性化顺序和元数据。
- **只读评估**：对一个 OFD 快照逐条款给出 Pass / Fail / Unknown / NotApplicable，并在同一字节上运行 [GB/T 42133 归档评估](https://whynpc9.github.io/ofdrw.net/guides/archival-assessment/)，按继承账本解释其结果。

先说清楚它**不**做什么：不签发符合性结论（整体结果只有 Fail 或 Unknown，永不 Pass）；不验证签名的密码学有效性；不评估阅读器、业务流程和引用标准；不嵌入字体、不签名、不生成双层页面。

## 安装

普通包和 Converter 元包都不会带入它，需要显式引用：

```sh
dotnet add package Ofdrw.Net.Medical --version 0.5.0
```

该包以精确的同批版本依赖带入 Archival、Packaging 和 Core；项目中的其他 `Ofdrw.Net.*` 直接引用也应一起升级到 0.5.0。

依赖方向为 `Medical → Archival → Packaging → Core`。

## 写出 OFD-H

```csharp
using Ofdrw.Net.Layout;
using Ofdrw.Net.Medical;
using Ofdrw.Net.Packaging;

var document = new FlowDocument();
document.Blocks.Add(new Paragraph("出院小结"));
var package = document.Render();

MedicalDocumentProfile.Apply(package.Options, new MedicalDocumentMetadata
{
    DocumentIdentifier = "EMR-2026-0001",   // DocInfo/DocID
    PatientId = "P-000123",
    ResponsibleParty = "主治医师",           // DocInfo/Author
    CreationTime = DateTimeOffset.Now,      // DocInfo/CreationDate
    Keywords = "出院小结",                   // DocInfo/Keywords
    FileType = "出院小结",
    SecretLevel = "内部"
});

await using var output = File.Create("discharge-summary.ofd");
await new OfdPackageWriter().WriteAsync(package, output);
```

`Apply` 设置：`DocType="OFD-H"`（7.1(a)）、`http://www.ofdspec.org/2016` 命名空间（6.2）、`Doc_0` 目录（6.1.2）、Deflate 压缩与线性化条目顺序（6.1.3），以及 7.3(a)(b) 的元数据。必填项为空、映射名重复或已有 CustomData 值冲突时抛出 `ArgumentException`，此时不会改动 DocType。

字体是否嵌入由生成方式决定，profile 不改变它；需要自包含字体时使用 DOCX 的字体自包含选项或为 Layout 提供字体字节，再用评估确认 7.2(b)。

### 患者、文件类型、密级的字段名

GB/T 48666 要求标识患者、文件类型和密级，但没有规定字段名或 schema。本包默认写入 `DocInfo/CustomDatas` 的 `EMR.PatientID`、`EMR.FileType`、`EMR.SecretLevel`，这只是本库约定。机构有自己的名称时，写出和评估都传同一个 `MedicalMetadataMapping`：

```csharp
var mapping = new MedicalMetadataMapping { PatientField = "HIS.PatientNo", FileTypeField = "EMR.DocClass", SecretLevelField = "EMR.Secret" };
MedicalDocumentProfile.Apply(package.Options, metadata, mapping);
var report = await new MedicalAssessmentService().AssessAsync(stream, new MedicalAssessmentOptions { Metadata = mapping });
```

评估时把 `Metadata` 设为 null，则 7.3(a)(b) 中这三项为 Unknown（`FieldMappingNotProvided`），不会被判 Fail。

### DOCX 来源

DOCX 转换入口仍然写出普通 OFD。先按 native 或默认模式转换，再读取、应用 profile、重新写出。转换时必须把 `OfdNamespace` 设为 `/2016`：读取得到的包保留来源命名空间，写出端为安全起见拒绝重映射（`NotSupportedException: Imported source namespace cannot be remapped safely`）。

```csharp
options.OfdNamespace = OfdConstants.StandardNamespace; // http://www.ofdspec.org/2016
using var ordinary = new MemoryStream();
await new DocxToOfdConverter(options).ConvertAsync(docx, ordinary);
ordinary.Position = 0;
var package = await new OfdReader().ReadAsync(ordinary);
MedicalDocumentProfile.Apply(package.Options, metadata);
await new OfdPackageWriter().WriteAsync(package, output);
```

这一步只改变包级声明、命名空间、条目顺序和元数据。页面内容不变，回归测试会比较页数和全部文字。

## 评估

```csharp
await using var input = File.OpenRead("discharge-summary.ofd");
var report = await new MedicalAssessmentService().AssessAsync(input);

Console.WriteLine(report.ExecutionState);          // Complete / BudgetExhausted / Error
Console.WriteLine(report.Declaration.Status);      // Medical / TypesetVariant / NotMedical / Missing ...
foreach (var clause in report.Clauses.Where(x => x.HasExecutableCheck))
    Console.WriteLine($"{clause.Clause} {clause.Outcome} {string.Join(",", clause.Observations.Select(o => o.ReasonCode))}");
Console.WriteLine(report.Document.StandardOutcome); // Fail 或 Unknown，永不 Pass
```

解读要点：

- **条款强度。**“应”条款的规则失败使条款 Fail；“宜/可”和“应支持”（如 6.1.3(a) 线性化、7.3(d) 签名）的失败只进入 `Advisories`，条款为 Unknown。没有可执行规则的条款保持 Unknown，`Document.UncheckedClauses` 列出它们。
- **部分可判定。**`Matrix.Category` 为 `PartiallyDecidable` 的条款即使 Pass，也只覆盖规则检查的部分。例如 8(c) 只检查最后一个签名的引用是否覆盖除签名、注释文件外的全部条目，摘要与签名值未验证。
- **字型资源表冲突。**标准 6.2 表 1 要求字型在 PublicRes，7.2(a) 要求在 DocumentRes。本库写出的字型在 PublicRes，因此 6.2(b) Pass、7.2(a) 为 Unknown（`ResourceTableConflict`）。这是标准自身的不一致，不是本库缺陷。
- **DocType 排版字符。**标准 PDF 把取值印成 `OFD⁃H`（U+2043）。从 PDF 复制的取值会报告为 `TypesetVariant` 并判 Fail；正确取值是 ASCII `OFD-H`。
- **GB/T 42133。**`report.Archival` 是未改动的归档评估。`InheritedArchivalClauses` 给出每项的处置：`Superseded`（GB/T 48666 另有规定，如 DocType、Permissions、加密、动作、音视频、附件）和 `ArchivalStageOnly`（归档时才处理，如注释冻结、签名去技术化）只作信息；`Inherited` 的 Fail 会使 `StandardOutcome` 为 Fail。哪些属于“相关要求”标准没有列明，`Inherited` 的读法标为 Unverified。

可执行规则 MH-001～MH-016、72 行条款矩阵与继承账本见[条款基线](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/baselines/gbt48666-2026-clause-matrix.md)。

## 编辑策略

GB/T 48666 9.2(b) 要求阅读软件对 OFD-H 文件自动禁用插页、调页等编辑，而 7.3(g) 要求不禁止元数据著录、注释、组合和签名。`MedicalEditPolicy` 为调用方分类操作：

| 操作 | OFD-H 下的决定 |
| --- | --- |
| InsertPage、RemovePage、AdjustPages、EditPageContent | `RequiresAuthorization` |
| AddAnnotation、Sign | `Allowed` |
| EditMetadata、Combine | `AllowedWithResign`（现有签名不再覆盖结果，需要移除后重签或追加签名） |

```csharp
MedicalEditPolicy.EnsureAllowed(declaredDocType, MedicalEditOperation.InsertPage, authorized: false); // 抛出 MedicalEditNotAuthorizedException
```

Ofdrw.Net 的通用编辑 API 不会自动调用这一策略。

## 已知限制

- 线性化只控制本库新写出包的条目顺序；ZIP 中央目录仍在文件尾，签名服务重写包时不保持该顺序。
- 只识别 MP3/MP4 的文件签名；其他音视频格式为 Unknown，因为标准原文是“MP3 等”。
- JSON 机读载荷只检查严格 UTF-8 与对象/数组外形，不做完整语法校验；7.5(c) 的头信息结构未评估。
- 按患者组合与重签、内外部附件和跳转动作、版本附件、扫描件双层页面，以及 OFD-H → OFD-A 归档派生，尚未实现。
