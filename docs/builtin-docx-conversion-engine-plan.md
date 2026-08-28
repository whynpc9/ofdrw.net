# Built-in DOCX Conversion Engine 实施计划

更新时间：2026-08-14

## 1. 背景与结论

实施前，`Ofdrw.Net.Converter.Docx` 提供 `Auto`、`MicrosoftWord` 和
`LibreOffice` 三种 `DocxConversionEngine`。`Auto` 在 macOS 上检测到
Microsoft Word 时选择 Word，否则选择 LibreOffice；选中的引擎不可用、启动失败、
超时或没有生成 PDF 时，转换直接失败。

本计划新增 `DocxConversionEngine.BuiltIn`，在当前 .NET 进程内完成
DOCX → PDF，并让 `Auto` 将其作为最后一级兜底。目标是确保部署环境没有
Microsoft Word 和 LibreOffice 时，受支持的 DOCX 仍可生成有效 PDF，并继续复用
现有 PDF → OFD 链路生成 OFD。

该能力定位为“无外部 Office 进程、可预测降级的可靠兜底”，不承诺与 Microsoft
Word 或 LibreOffice 像素级一致。损坏、加密、超过资源限制或完全由不支持对象组成的
文档必须给出明确错误或诊断，不能通过空白输出伪装成功。

### 当前实施状态

截至 2026-08-14，本计划首期已落地：公开 `BuiltIn` 引擎及转换结果诊断，
`Auto` 使用 Word → LibreOffice → BuiltIn 候选链；Open XML 读取、规范化模型和
MigraDoc/PDFsharp 渲染覆盖两页 A4 样例的文字、样式、分页、表格、内嵌图片、
页眉页脚及页码。输入/展开量/部件数/XML 元素/图片字节与像素限制、外部资源关系
拒绝、取消/超时和不支持特性策略均已接入。

本地 Release 双目标框架构建、全部 36 个测试以及独立 NuGet 消费端
DOCX → PDF → OFD 已通过；CI 已增加不安装 Office 应用的 Linux/Windows/macOS
BuiltIn 矩阵和 Linux `Auto` 实际探针。三平台结果仍以远端 CI 首次运行结果为准。

## 2. 目标

1. 新增显式可选的 `DocxConversionEngine.BuiltIn`。
2. 将 `Auto` 的候选顺序调整为：
   `MicrosoftWord → LibreOffice → BuiltIn`。
3. BuiltIn 转换过程不启动 Word、LibreOffice 或其他办公软件进程。
4. 保持现有 `IDocxToPdfConverter`、`IDocxToOfdConverter` 和流式 API 兼容。
5. 提供实际使用引擎、尝试记录、字体替换和不支持功能的结构化诊断。
6. 保证失败不会向调用方输出流写入半成品，也不会遗留临时文件。
7. 在不安装 Word 和 LibreOffice 的 CI/NuGet 消费环境中验证
   DOCX → PDF → OFD 全链路。

## 3. 非目标

首期不以重建完整 Microsoft Word 排版引擎为目标，也不承诺：

- Word/LibreOffice 级像素一致性；
- 浮动形状、VML、SmartArt、图表、文本框和艺术字的完整布局；
- 公式、OLE、宏和嵌入文件的完整渲染；
- 多栏、复杂文字环绕和 Word 专属精确断行；
- 域、目录和交叉引用的重新计算；
- 修订、批注、脚注和尾注的完整语义；
- 完全无原生库的 DOCX → OFD。现有 PDF → OFD 默认使用
  Docnet/Pdfium 原生运行时，但不要求安装独立应用。

## 4. 行为契约

### 4.1 显式引擎

显式选择 `MicrosoftWord`、`LibreOffice` 或 `BuiltIn` 时，只运行指定引擎。
指定引擎失败后不得静默切换，以便调用方获得确定行为。

### 4.2 Auto 降级

`Auto` 按平台和可用性建立候选链：

1. macOS 且 Word 预检通过时尝试 Microsoft Word；
2. LibreOffice 可执行文件预检通过时尝试 LibreOffice；
3. 最后尝试 BuiltIn。

以下错误允许尝试下一引擎：

- 引擎未安装或不可用；
- 进程无法启动或自动化权限被拒绝；
- 引擎执行超时；
- 进程非零退出或没有生成有效 PDF；
- 引擎自身报告可降级的渲染失败。

以下错误不得继续降级：

- 调用方取消操作；
- 输入流不可读或输出流不可写；
- 输出流提交失败；
- DOCX 损坏、加密或违反安全限制；
- BuiltIn 确认输入超出支持契约，且策略配置为 `Throw`。

每个候选引擎写入独立临时 PDF。只有候选输出通过 PDF 头、非空和可打开检查后，
才复制到调用方的 `pdfOutput`。最终失败应保留各候选引擎的摘要，但不得泄露输入
正文、临时路径或敏感文件名。

### 4.3 成功定义

对于未加密、结构有效、未超过资源限制且内容落在支持范围内的 DOCX，BuiltIn
必须生成可被仓库当前 PDF 读取链打开的有效 PDF。遇到可忽略的不支持元素时，必须按
配置执行 `BestEffort`、`Placeholder` 或 `Throw`，并产生结构化诊断，不能静默丢失。

## 5. 首期支持范围

### 5.1 页面与文档结构

- 页面尺寸、方向和页边距；
- 分节和显式分页；
- 基础页眉、页脚和页码；
- 文档默认值、主题字体和样式继承。

### 5.2 段落与文字

- 段落、Run、中英文混排；
- 字体、字号、粗体、斜体、下划线和颜色；
- 左/中/右/两端对齐；
- 缩进、段前段后、行距、换行和 Tab；
- 简单编号和项目符号；
- 超链接的显示文本和基础 PDF 链接。

### 5.3 表格与图片

- 表格列宽、边框、底色和单元格内段落；
- 简单横向/纵向单元格合并；
- 表头重复和表格自动跨页；
- 内嵌 PNG/JPEG 图片及显式尺寸。

### 5.4 字体

- 复用并扩展 `DocxConversionOptions.FontDirectories`；
- 从文档字体、主题字体、用户配置和系统字体建立确定的回退链；
- 输出字体替换诊断；
- 首期不捆绑 Microsoft Office 私有字体；
- 是否提供可选开源 CJK 字体包，在验证包大小和许可证后单独决策。

## 6. 技术方案

建议的数据流为：

```text
DOCX
  → Open XML 解析
  → Ofdrw.Net 内部规范化文档模型
  → MigraDocCore 流式排版
  → 现有 PdfSharpCore PDF 输出
  → 现有 PDF → OFD 链路
```

候选依赖：

- `DocumentFormat.OpenXml 3.5.1`：MIT，支持 `netstandard2.0`。它只负责
  OPC/Open XML 读取和强类型元素访问，不负责高级排版。
- `MigraDocCore.Rendering 1.3.67`：支持 `netstandard2.0`，并依赖仓库
  已使用的 `PdfSharpCore 1.3.67`，可降低 PDF 栈冲突风险。

不得将 Open XML 元素直接散落映射到 PDF 绘图命令。应先转换到内部规范化模型，
隔离 DOCX 解析、样式计算和渲染器类型，使后续可以替换布局实现或增加直接
DOCX → OFD 路径。

建议的内部组件：

```text
DocxToPdfConverter                 引擎编排和事务性输出
  ├─ MicrosoftWordDocxBackend      现有 Word 自动化
  ├─ LibreOfficeDocxBackend        现有 LibreOffice 进程
  └─ BuiltInDocxBackend
       ├─ DocxPackageReader        包、关系和资源安全读取
       ├─ DocxStyleResolver        默认值、主题和样式继承
       ├─ DocxDocumentNormalizer   规范化文档模型
       ├─ DocxFeatureAnalyzer      支持度与诊断
       ├─ DocxFontResolver         字体发现和替换
       └─ BuiltInPdfRenderer       MigraDocCore/PdfSharpCore 输出
```

引擎实现应通过内部 backend 接口解耦，并允许在测试中注入可用性和失败结果，避免
单元测试依赖真实 Word/LibreOffice 安装。

## 7. API 设计

保持现有接口不变：

```csharp
Task ConvertAsync(
    Stream docxInput,
    Stream pdfOutput,
    CancellationToken cancellationToken = default);
```

新增枚举值和选项：

```csharp
public enum DocxConversionEngine
{
    Auto,
    MicrosoftWord,
    LibreOffice,
    BuiltIn
}

public enum UnsupportedDocxFeatureBehavior
{
    BestEffort,
    Placeholder,
    Throw
}
```

建议在具体转换器上增加非破坏性结果 API，原有 `ConvertAsync` 调用它并丢弃结果：

```csharp
Task<DocxConversionResult> ConvertWithResultAsync(
    Stream docxInput,
    Stream pdfOutput,
    CancellationToken cancellationToken = default);
```

`DocxConversionResult` 至少包含：

- `ActualEngine`；
- `AttemptedEngines`；
- `Diagnostics`；
- `FontSubstitutions`；
- `UnsupportedFeatures`。

不得使用共享的 `LastResult` 属性，因为同一个转换器可能并发执行。

CLI 增加：

```text
--docx-engine auto|word|libreoffice|built-in
```

CLI 成功信息应显示实际使用引擎；发生降级或字体替换时输出简明警告。

## 8. 安全与资源限制

BuiltIn 直接解析不可信 ZIP/XML，必须在正式开放前具备以下限制：

- 输入 DOCX 最大字节数；
- ZIP 条目数量、单条目和累计展开大小；
- 压缩比和嵌套关系数量；
- XML 深度、节点数量和文本长度；
- 图片数量、单图字节数和解码后像素总量；
- 表格行列数、段落数、Run 数和最大估算页数；
- 转换总超时及取消检查；
- 禁止 DTD、外部实体和外部资源下载；
- 拒绝外部关系、加密包和无法安全解析的畸形 URI；
- 所有临时文件和目录在 `finally` 中清理。

限制应使用安全默认值，并允许服务端调用方在合理范围内收紧。超限错误必须能与
“不支持版式”及普通渲染失败区分。

## 9. 分阶段实施

### 阶段 0：技术验证和停止门

1. 在独立试验代码中引入候选依赖，不先修改公开 API。
2. 使用现有 `generated-layout.docx` 验证中文、两页 A4、表格、分页和颜色。
3. 增加一份含图片、页眉页脚、编号和单元格合并的虚构夹具。
4. 验证 `netstandard2.0/2.1` 编译、包依赖图和程序集无冲突。
5. 评估 CJK 字体解析、断行、内存和输出稳定性。

停止条件：

- MigraDocCore 无法稳定渲染 CJK、基础表格或分页；
- 与现有 PdfSharpCore/全局字体解析器产生不可隔离的冲突；
- 包依赖或许可证不满足发布要求；
- 同一固定字体环境下输出不可重复。

触发停止条件后，不进入公开 API 实现；应在“自研分页器”和“迁移到官方
PDFsharp/MigraDoc”之间重新评估。

### 阶段 1：引擎编排重构

1. 抽取 Word 和 LibreOffice backend，不改变现有显式行为。
2. 建立候选链、可用性探测和错误分类。
3. 实现一次输入落盘、多候选复用和事务性输出。
4. 增加 `BuiltIn` 枚举和 CLI 解析，但在渲染器完成前保持内部开关。
5. 为降级顺序、超时、取消和输出提交补齐单元测试。

### 阶段 2：BuiltIn MVP

1. 完成安全的 DOCX 包与关系读取。
2. 实现样式、主题、编号和分节解析。
3. 建立内部规范化文档模型。
4. 实现段落、分页、字体和基础页眉页脚。
5. 实现表格、图片及基础超链接。
6. 产生不支持功能和字体替换诊断。

### 阶段 3：安全与鲁棒性

1. 完成资源限制和畸形输入测试。
2. 覆盖加密、损坏、ZIP bomb、外部关系和超大图片。
3. 验证并发、取消、超时和临时文件清理。
4. 对完全不支持的内容验证 `BestEffort/Placeholder/Throw` 行为。

### 阶段 4：自动化与兼容性

1. 增加无 Word/LibreOffice 的 BuiltIn CI job。
2. 在 Linux、Windows、macOS 验证 BuiltIn DOCX → PDF。
3. 在支持 Docnet/Pdfium 的运行时验证 DOCX → PDF → OFD。
4. 保留现有 LibreOffice 集成测试和高保真路径回归。
5. 更新 NuGet consumer E2E，显式选择 BuiltIn 并验证页数、尺寸、文本和图片。
6. 对虚构语料执行语义断言；固定字体环境后再加入可维护的像素差异阈值。

### 阶段 5：文档与发布

1. 更新 README、CLI 帮助、包描述和 `docs/feature-parity.md`。
2. 更新 `THIRD-PARTY-NOTICES.md` 和依赖审计结果。
3. 明确 Word、LibreOffice 和 BuiltIn 的选择建议及保真度差异。
4. 检查 NuGet 包大小、依赖闭包和三平台消费验证。
5. 以新的预览版本发布，不覆盖或移动已有版本/标签。

## 10. 测试矩阵

| 类别 | 样例 | 核心断言 |
| --- | --- | --- |
| 引擎编排 | backend stub | Auto 顺序、显式不降级、错误分类、取消 |
| 基础文字 | 中英混排、样式、换行 | 文本可提取、字体/颜色、无空白页 |
| 分页 | A4、横向、分节、手工分页 | 页数、尺寸、方向、页边距 |
| 表格 | 边框、底色、合并、跨页 | 行列结构可见、无溢出、表头重复 |
| 图片 | PNG/JPEG、显式尺寸 | 图片存在、比例和页内位置合理 |
| 页眉页脚 | 页码、固定文字 | 各节继承和页码正确 |
| 字体 | CJK/Latin、缺失字体 | 确定回退、诊断完整、字符不丢失 |
| 不支持对象 | 图表、形状、公式、OLE | 占位/忽略/失败策略符合配置 |
| 安全 | 畸形 ZIP/XML、外部关系、超大图片 | 快速拒绝、错误分类、资源有界 |
| 全链路 | NuGet consumer | 无 Office 应用的 DOCX → PDF → OFD |

私有复杂文档可以继续用于本地视觉验证，但不得复制到仓库、测试夹具、NuGet 包或
Git 历史。仓库测试只使用虚构且可重复生成的样例。

## 11. 验收标准

以下条件全部满足后，才能宣称 BuiltIn 兜底可用：

1. 没有 Word 和 LibreOffice 时，默认 `Auto` 能对支持范围内 DOCX 生成有效 PDF。
2. `BuiltIn` 转换过程没有启动外部进程。
3. 现有两页虚构夹具生成两页 A4 PDF，并继续生成两页有效 OFD。
4. 支持范围内文本可提取，表格、颜色、分页和图片可见。
5. 不支持内容不会静默消失，调用方能获得结构化诊断。
6. 显式 Word/LibreOffice 的成功和失败语义保持兼容。
7. 失败不污染输出流，不遗留输入、PDF、profile 或字体临时文件。
8. 畸形、加密及超限 DOCX 被安全、可诊断地拒绝。
9. 从打包后的 NuGet 包运行通过，而不只是项目引用通过。
10. README 和 feature parity 明确 BuiltIn 是可靠兜底而非高保真 Word 替代品。

## 12. 预计改动范围

主要修改：

- `src/Ofdrw.Net.Converter.Docx/DocxConversionEngine.cs`
- `src/Ofdrw.Net.Converter.Docx/DocxConversionOptions.cs`
- `src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs`
- `src/Ofdrw.Net.Cli/Program.cs`
- `src/Ofdrw.Net.Converter.Docx/Ofdrw.Net.Converter.Docx.csproj`
- `tests/Ofdrw.Net.Converter.Docx.Tests/`
- `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/`
- `.github/workflows/ci.yml`
- `.github/workflows/publish-nuget.yml`
- `scripts/run-converter-package-e2e.sh`
- `README.md`
- `docs/feature-parity.md`
- `THIRD-PARTY-NOTICES.md`

预计新增：

- BuiltIn backend、DOCX reader、样式解析、规范化模型、字体解析和 PDF renderer；
- 转换结果与诊断类型；
- BuiltIn 专用单元、集成、安全和 NuGet 消费测试夹具。

## 13. 外部参考

- Open XML SDK：https://github.com/dotnet/Open-XML-SDK
- DocumentFormat.OpenXml 3.5.1：
  https://www.nuget.org/packages/DocumentFormat.OpenXml/3.5.1
- PdfSharpCore：https://github.com/ststeiger/PdfSharpCore
- MigraDocCore.Rendering 1.3.67：
  https://www.nuget.org/packages/MigraDocCore.Rendering/1.3.67
