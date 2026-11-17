---
title: "OFD-H 专用包 Ofdrw.Net.Medical 设计记录"
---

# OFD-H 专用包 `Ofdrw.Net.Medical` 设计记录

日期：2026-11-17。工作包：[路线图](../engineering/plans/ofd-archival-roadmap.md) H0–H2。条款依据：[GB/T 48666-2026 条款矩阵与继承账本](../engineering/baselines/gbt48666-2026-clause-matrix.md)。

## 问题

GB/T 48666—2026《电子病历版式文档技术要求》定义了电子病历版式文档（根 `DocType="OFD-H"`）。它在 GB/T 33190 与 GB/T 42133 之上增加医疗应用要求，但在 DocType、权限、临时加密、动作、音视频和附件上与 GB/T 42133 的归档处理相反。通用包自 D0 起默认写出 `OFD`，专用默认值与校验不得回流通用包。因此需要一个可选包，既能按 OFD-H 配置写出，又能给出不冒充认证的只读评估。

## 决策

1. **独立 opt-in 包。**依赖 `Medical → Archival → Packaging → Core`。默认包、Converter 元包和 CLI 不引用它；`verify-package-artifacts.py` 的 `OPT_IN_PACKAGES` / `OPT_IN_DEPENDENCIES` 与包消费 E2E 校验依赖闭包。
2. **复用而非复制 GB/T 42133 评估。**`MedicalAssessmentService` 在同一字节快照上调用 `ArchivalAssessmentService`，原样保存其报告（`Archival` 属性），再经继承账本解释：`Inherited` 的 Fail 计入 OFD-H 整体结果，`Superseded` 与 `ArchivalStageOnly` 只作信息。字体嵌入（7.2(b)）直接转述 AC-017/AC-018 的观察，不另写字体解析。
3. **自有只读规则 MH-001～MH-016。**医疗包需要中央目录的压缩方法和本地文件头物理顺序，`System.IO.Compression` 不提供，因此包内有一个有界的中央目录解析器（拒绝 ZIP64 与越界记录）。XML 读取禁止 DTD，受调用方 `ArchivalPreflightOptions` 预算约束；预算耗尽或输入被拒是执行状态，不是条款结论。
4. **强度聚合。**Shall/Statement 条款：Fail > Unknown > Pass，全部 NotApplicable 时为 NotApplicable。Should/May/Capability：Fail 只进入 `Advisories`，条款为 Unknown。无规则条款保持 Unknown。整体 `StandardOutcome` 只有 Fail 或 Unknown；`CoverageStatus` 恒为 `Incomplete`，`Certification` 恒为 `NotAssessed`。
5. **不替标准消解冲突。**6.2 表 1 与 7.2(a) 对字型资源表的要求相互矛盾，规则对两条款分别给出 Pass/Unknown（`ResourceTableConflict`），只有字型落在 PageRes（两条都禁止）时 Fail。
6. **写出 profile 只配置选项。**`MedicalDocumentProfile.Apply` 修改 `OfdDocumentOptions`：`DocType=OFD-H`、`/2016` 命名空间、`Doc_0`、Deflate、线性化顺序、`DocID`/Author/CreationDate/Keywords 与映射的 CustomData。必填项缺失、映射名重复或已有 CustomData 值冲突时抛出 `ArgumentException`，且不修改 DocType。
7. **通用包中性扩展。**`OfdMetadata.DocId` 读写 GB/T 33190 `DocInfo/DocID`（读取保留、置 null 删除、合并不复制）；`OfdDocumentOptions.LinearizedEntryOrder` 默认 false，开启后写出顺序为 OFD.xml、文档、资源表、首页条目（Content.xml 在前）、文档 Res、其余页面、其他条目。默认顺序与现有字节保持不变。
8. **编辑策略只分类。**`MedicalEditPolicy` 按 9.2(b) 把插页、删页、调页和改页面内容归为 `RequiresAuthorization`，按 7.3(g)/6.3(f) 允许注释与签名，按 3.10/6.3(g) 把元数据编辑和组合归为 `AllowedWithResign`。通用编辑 API 不调用它，避免在没有真实消费者时改变 Layout 的兼容行为。

## 不做的事

- 不验证摘要、签名值、证书链或时间戳；MH-013/MH-014 只检查结构覆盖与算法标识。
- 不判断 7.5(c) 机读信息头或卫生行业 schema；标准未给出字段定义。
- 不检查阅读器（第 9 章）、流程（第 5 章）与引用标准（GB/T 33560、35275、38540、GB 18030）。
- 不改写、修复或签发 OFD-H；不把 OFD-H 转为 OFD-A（H4）。
- 不让 DOCX/PDF 转换入口直接产出 OFD-H；调用方以 `/2016` 命名空间转换、读取后应用 profile 再写出，见指南。读取得到的包不能改命名空间（`OfdSourceIdentity` 拒绝重映射），profile 不绕过这一保护。

## 验证

单元测试见 `tests/Ofdrw.Net.Medical.Tests`：矩阵、规则与基线文档逐行一致；继承账本覆盖 105 项；profile 产物通过所有包可判定的 Shall 条款；普通 OFD、排版连字符、存储 XML、非线性顺序、非规范目录、PageRes 字型、图像签名、音频与限值、元数据映射、权限、签名覆盖与算法、UTF-8 机读载荷、加密、预算和异常输入均有正反例。`e2e/Ofdrw.Net.Medical.E2E` 从打包产物消费。视觉复查记录见[验收记录](../validation/medical-ofd-h-h1-h2-20261117.md)。
