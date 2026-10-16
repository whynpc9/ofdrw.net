# Plan 029: 设计有界多页 SVG HTML 预览

> 后续授权实施已完成并统一本地合并。实际交付见[029-followup-implementation.md](029-followup-implementation.md)，合并后验证见[最终验收](followup-integration-acceptance-2026-10-16.md)。下文保留最初调查/设计范围和当时状态。

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "plans/029-html-preview-design.md"`；再运行 `git status --short`、`git diff -- "plans/029-html-preview-design.md"` 与 `git diff --cached -- "plans/029-html-preview-design.md"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE（仅设计）
- 来源：D02
- 优先级：P3；工作量：M；修复/实施风险：MED
- 类别：direction
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：本次执行结论：DONE（仅设计；11链接复核）；冻结交付 `1ca3e287`。源码基线、独立复核、实测项目、视觉范围、遗留门及产物路径见 [执行调度记录](execution-status.md) 和 [机器清单](execution-state.json)。 调查/设计完成只表示本计划指定交付完成，不代表附件修复或新功能已经实现。

## 为什么要做

现有 SVG 导出是单页接口，路线图要求多页 HTML 预览。先解决读取复用、ID 隔离、转义及输出预算，再判断实施范围。



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

[src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs:37](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs:37)

```csharp
    public async Task ConvertAsync(
        Stream ofdInput,
        Stream svgOutput,
        int pageIndex = 0,
        CancellationToken cancellationToken = default)
    {
        if (ofdInput is null)
        {
```

其余核查入口：无额外入口。

**沿用范例**：[src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs:37](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs:37)

```csharp
    public async Task ConvertAsync(
        Stream ofdInput,
        Stream svgOutput,
        int pageIndex = 0,
        CancellationToken cancellationToken = default)
    {
        if (ofdInput is null)
        {
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `plans/029-html-preview-design.md`

此外可写本次任务的 `artifacts/plan-029/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。本计划不修改生产实现；临时探针只能写任务 artifacts。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-029-multipage-html-design`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-029/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 建立范围与可验证样例

设计一次读取 OFD、按选页顺序输出内联 SVG 的内部入口和公开 HTML API/CLI；规定零基 API/一基 CLI、重复页、标题转义、页间 ID/fragment 隔离、累计输出字节和页数额度。记录 F11/F12 修复前不能承诺的文本/路径保真。

**验证**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "plans/029-html-preview-design.md"` → 已确认基线；在指定交付文件写齐样例、证据和预期，不依赖会话上下文。

### 2. 形成可审阅闭环

列浏览器验证样例：相同资源 ID 多页、恶意标题/文本不执行脚本、SVG fragment 指向本页、多页/空选页/预算失败、取消与原子输出。确定 SVG 白名单/事件属性/外链策略后再写实现计划；只交设计，不承诺全选中文本或任意浏览器等价。

**验证**：完成下列样例矩阵，并执行文档检查命令；缺失事实或未判断项逐项标明，不把设计完成写成实现通过。

### 3. 记录完成边界

状态更新只表示设计已交付；生产实现、功能测试、视觉验收和发布均未完成。可选方向不阻塞修复批次。

**验证**：`git diff --check` 退出 0，`git status --short` 只新增/修改本计划允许的文件（保留原有用户改动）。

## 回归/样例矩阵

- `IdIsolation`：两页相同 clip/id 不互串
- `EscapingPolicy`：标题/内容不能变成 HTML 脚本或事件
- `OneReadBudget`：读包一次且累计输出有界
- `FidelityDependencies`：实施依赖 F11/F12 或明确可接受限制

每个场景必须记录预期、实际、结论及复跑入口；未知项不可改写成通过。


## 设计交付检查

交付文件必须有以下标题，各节填写实质内容而非占位：`## API 与边界`、`## 现有能力复用`、`## 样例与预期`、`## 验收命令`、`## 风险与未决项`、`## 后续实施切片`。

```bash
python3 - <<'PY'
from pathlib import Path
p = Path('plans/029-html-preview-design.md')
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

- 若必须执行源 SVG/script 或加载未知外部资源才能显示，停止；如需改变 Svg 公共接口先给兼容方案。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

浏览器与 SVG 转义/链接策略属于产品安全契约；可选方向不阻塞缺陷修复。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。
