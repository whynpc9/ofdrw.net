# Plan 028: 设计复用 FlowDocument 的纯文本入口

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "plans/028-plain-text-design.md"`；再运行 `git status --short`、`git diff -- "plans/028-plain-text-design.md"` 与 `git diff --cached -- "plans/028-plain-text-design.md"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE（仅设计）
- 来源：D01
- 优先级：P3；工作量：M；修复/实施风险：LOW–MED
- 类别：direction
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：本次执行结论：DONE（仅设计；独立26链接检查）；冻结交付 `e35fb82e`。源码基线、独立复核、实测项目、视觉范围、遗留门及产物路径见 [执行调度记录](execution-status.md) 和 [机器清单](execution-state.json)。 调查/设计完成只表示本计划指定交付完成，不代表附件修复或新功能已经实现。

## 为什么要做

路线图已提出纯文本入口；公开 FlowDocument 已有分页能力。先确定薄适配器契约，避免重建布局器。



## 当前代码与约定

[docs/capability-roadmap.md:108](/Users/wanghongyi/Projects/ofdrw.net/docs/capability-roadmap.md:108)

```
| P1-08 | 纯文本 → OFD | ofdrw `TextConverter` | 调用 P0-01 布局引擎 | 可设字号与页尺寸 |
| P1-09 | OFD → HTML | ofdrw `HTMLExporter` | 多页 SVG 包进单 HTML | 浏览器可预览；不承诺可选中全部文字 |
| P1-10 | 强类型对象补齐 | ofdrw-core | 最小强类型闭环已实现：大纲/书签、局部注释、GoTo、Path 裁剪、共享嵌套组合；其余继续 `Raw` / `Preserved*` | 最小范围读写往返、未知 XML、纯包消费及 8 页 Preview 验收通过；[记录与边界](evidence/issue09-typed-objects/README.md) |
| P1-11 | 多 `DocBody` | ofdrw-reader | 读包不再只取第一个文档 | 有多样本；写入策略有文档 |
| P1-12 | SkiaSharp 绘图层 | ofdrw-graphics2d | 可选适配，主要服务 P1-04 | 不替代 P1-01 的 OFD 原语 API |

issue #3 的完成定义就是 **P1-01 + P1-02**。API 形状以 issue 中的类型名为准，实现可以架在 Skia 或自绘路径上，但调用方应能只依赖 OFD 原语，而不必引用 Skia。

```

[src/Ofdrw.Net.Layout/FlowDocument.cs:21](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Layout/FlowDocument.cs:21)

```csharp
    public OfdDocumentPackage Render(CancellationToken cancellationToken = default) =>
        new FlowDocumentRenderer(Options, Blocks, cancellationToken).Render();
}

/// <summary>Base of the built-in flow blocks.</summary>
public abstract class FlowBlock
{
    internal FlowBlock() { }
```

其余核查入口：无额外入口。

**沿用范例**：[src/Ofdrw.Net.Layout/FlowDocument.cs:21](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Layout/FlowDocument.cs:21)

```csharp
    public OfdDocumentPackage Render(CancellationToken cancellationToken = default) =>
        new FlowDocumentRenderer(Options, Blocks, cancellationToken).Render();
}

/// <summary>Base of the built-in flow blocks.</summary>
public abstract class FlowBlock
{
    internal FlowBlock() { }
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `plans/028-plain-text-design.md`

此外可写本次任务的 `artifacts/plan-028/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。本计划不修改生产实现；临时探针只能写任务 artifacts。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-028-plain-text-ofd-design`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-028/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 建立范围与可验证样例

以 FlowDocument/Paragraph 为唯一布局基础，设计 API/CLI 的 UTF-8/BOM/换行、空输入、字体、字号/页尺寸/边距、取消和输入/页数额度；默认严格 UTF-8 解码，保留原始文本的验证口径。列清 API 归属与对 netstandard 的影响。

**验证**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "plans/028-plain-text-design.md"` → 已确认基线；在指定交付文件写齐样例、证据和预期，不依赖会话上下文。

### 2. 形成可审阅闭环

写出最小实施切片及测试/视觉矩阵：中英文比例字体、CRLF、连续/末尾换行、超长行、分页和原文抽取。提出候选参数示例、调用方场景与不支持项；只交付设计，不提前实现新包或第二套分页器。

**验证**：完成下列样例矩阵，并执行文档检查命令；缺失事实或未判断项逐项标明，不把设计完成写成实现通过。

### 3. 记录完成边界

状态更新只表示设计已交付；生产实现、功能测试、视觉验收和发布均未完成。可选方向不阻塞修复批次。

**验证**：`git diff --check` 退出 0，`git status --short` 只新增/修改本计划允许的文件（保留原有用户改动）。

## 回归/样例矩阵

- `ContractExamples`：正常/空/非法UTF8/额度超限示例都有预期
- `ReuseMap`：每个功能映射到现有 Flow API 或明确最小新增点
- `AcceptanceMatrix`：文本完整、页面视觉、取消不提交均有具体样例

每个场景必须记录预期、实际、结论及复跑入口；未知项不可改写成通过。


## 设计交付检查

交付文件必须有以下标题，各节填写实质内容而非占位：`## API 与边界`、`## 现有能力复用`、`## 样例与预期`、`## 验收命令`、`## 风险与未决项`、`## 后续实施切片`。

```bash
python3 - <<'PY'
from pathlib import Path
p = Path('plans/028-plain-text-design.md')
s = p.read_text()
for heading in ['## API 与边界','## 现有能力复用','## 样例与预期','## 验收命令','## 风险与未决项','## 后续实施切片']:
    assert heading in s, heading
print('design structure OK; still requires content review')
PY
```

这是文档结构检查，不能代替对契约与样例的人工审阅。设计里的未来验证命令必须标为未执行。


## Done criteria

- [ ] 设计具备 API/边界/样例/验收/切片，未把设计写成产品已实现。
- [ ] 所有本计划要求的命令都有实际结果；未执行/跳过项明确列出，无零测试伪通过。
- [ ] 不要求新增视觉产物的本计划没有凭静态/文本结果声称视觉通过。
- [ ] 未扩大到无关源代码/依赖/发布修改。
- [ ] `git diff --check` 退出 0，改动只在 Scope（原有用户改动保持），索引更新真实完成范围。

## STOP conditions

- 缺少真实调用方用例时保留设计选项，不把新 API 视为已批准产品；不替代 DOCX converter。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

后续实现先同步路线图并重新估算 API、CLI、消费和视觉门；本设计不是发布资格。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。
