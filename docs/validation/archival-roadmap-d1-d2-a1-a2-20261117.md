# OFD-A 路线图 D1、D2、A1、A2：实施验收摘要

日期：2026-11-17。关联 [#145](https://github.com/whynpc9/ofdrw.net/issues/145) 与[分阶段计划](../engineering/plans/ofd-archival-roadmap.md)。代码基线 `03981b08`（计划修订），本记录对应的提交见 PR 分支。均属 main 开发版本，**没有发布 NuGet**。

## 范围

| 工作包 | 交付 | 不包含 |
| --- | --- | --- |
| D2 | 新文档的自定义标签经中性通道 `ofdrw-kv` / `urn:ofdrw-net:custom-tags:kv:1` 写出，载荷根位于所声明的命名空间；历史 EMR 通道可读、原样保存 | OFD-H 专用包；`SchemaLoc`（仓库内无可核对的 schema） |
| D1 | `DocxConversionOptions.NativeFontSelfContainment = RequireEmbedded` 与 CLI `--native-font-self-containment`：每个已用字体嵌入所声明族自身的字节，否则失败；按 OS/2 fsType 拒绝受限嵌入 | 默认行为（不变）；DualLayer、DOCX→PDF（显式拒绝）；字体授权判断 |
| A2 | 改变内容的 SDK 入口与 CLI 写命令拒绝声明 OFD-A 的输入（CLI 退出码 5）；`OfdOrdinaryDerivation` / `derive-ofd` 显式派生普通 OFD | 签名追加（未决定，行为不变）；第三方阅读器第 7.2 条行为 |
| A1a | 可选包 `Ofdrw.Net.Archival`，只读 `ArchivalAssessmentService`，分维度报告；GB/T 42133 条款强度与适用性矩阵（105 项，读自渲染页面） | 生成 OFD-A；整部标准 Pass |
| A1b | 预检选项 `RecognizeKnownCustomTags`（默认关闭，开启时使用独立 profile 修订）；AC-016 已知标签链闭合 | 未知 schema、Extensions、Annotations、CompositeGraphicUnits |
| A1c | Core 公共 `OfdFontPayloadInspector`；AC-017–AC-020 按包内字节与实际使用字形判断字体；字体豁免策略选项 | 批注外观与 CompositeObject 中的字体使用（保持 Unknown） |
| A1d | 基于 `generated-layout.docx` 与固定字体的确定性回归；显式引用 Archival 的隔离包消费者 | macOS/Windows 上的同一组只读检查 |

条款强度策略：标准正文未使用“应”，68 项为陈述式、24 项“宜”、9 项“可”、4 项附录 A“应”。陈述式与“应”条款上的既有规则保持 Fail 语义；“宜”“可”条款只给出建议性观察，不产生 Fail。

## 功能测试

环境：Linux 5.15（WSL2）x86_64、.NET SDK 10.0.401。最终状态全解决方案 13 个测试程序集：**5332 通过、19 跳过、1 失败**。

- 失败项 `BuiltInPdfItalicRendererTests.DifferentPageSizesKeepHeadersFieldsTablesAndImagesThroughDrawingSeam` 为第四位小数的舍入差异，在本机基线 `03981b08` 上同样失败，与本次改动无关；以 CI 结果为准。
- 跳过项：Converter.Pdf 17 项需要原始 OFL 字体夹具，Converter.Docx 与 Cli 各 1 项为平台相关用例，均非本次新增。
- 仓库存储检查（含 `--require-materialized e2e --require-materialized tests`）、公共 API 文档门禁、脚本单元测试通过。
- 本地包消费：`run-converter-package-e2e.sh 0.4.1-a1d.1` 打包并校验 13 个包，全部既有消费者及新增 Archival 消费者通过；普通 Converter 消费者的依赖闭包不含 Archival。
- 网站：`docs/site` 的 `npm ci`、`build`、`typecheck`、`check` 通过。Nimbus 不对 `.md` 正文做 authoring lint，因此不是完整的文档检查。

## 转换链路与视觉检查

D1、D2 改变转换输出的字节（标签文件、可选的字体嵌入），不改变排版代码。

- **自动渲染检查**：用最终代码与 main 构建的 CLI 分别转换 `generated-layout.docx`（SHA-256 `17bea68d…d7bb6`），覆盖显式 `native`、默认模式，以及 `require-embedded` + 固定 Noto Sans CJK SC 字体目录。三种配置的 2 页 PNG 均与 main 逐字节相同，提取文本相同；包内标签改为 `CustomTag_KeyValue.xml`。
- **PNG 视觉复查**：逐页查看 `require-embedded` 的 2 页：中文标题与正文、英文间距、斜体副标题、表头底色与边框、红色粗体局部样式、分页与右对齐日期均正常，无缺字、裁切、重叠或异常空白页。各子任务另有自己的逐页复查记录（含宋体改名字体变体）。
- **macOS Preview**：**未完成**。本环境为无 GUI 的 WSL2，无法打开 Preview。须在 macOS 上以 `DOCX → native OFD → PDF → Preview` 补验，补验前视觉验收保持未完成。

结论仅限上述样例、页面与配置，不推断任意 Word 文档完整保真，也不构成任何 GB/T 42133 符合性结论。

## 证据

运行产物在已忽略的 `artifacts/roadmap-impl/`（`final/`、`d1/`、`d2/`、`a1a-matrix/`），未上传。`final/sha256.txt` 的 SHA-256 为 `9ecc57884fefe2ba08a02a956543cd261bf966534ee25655516c8685fd0b4d7c`。条款矩阵所据标准副本 SHA-256 为 `7d76c086…1b82`，副本与页面图不入库。
