# Plan 026: 同步当前路线图与已有实现状态

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "docs/capability-roadmap.md" "docs/feature-parity.md" ".scratch/capability-roadmap/issues/README.md" "docs/capability-roadmap-development-plan.md"`；再运行 `git status --short`、`git diff -- "docs/capability-roadmap.md" "docs/feature-parity.md" ".scratch/capability-roadmap/issues/README.md" "docs/capability-roadmap-development-plan.md"` 与 `git diff --cached -- "docs/capability-roadmap.md" "docs/feature-parity.md" ".scratch/capability-roadmap/issues/README.md" "docs/capability-roadmap-development-plan.md"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE
- 来源：F28
- 优先级：P2；工作量：S；修复/实施风险：LOW
- 类别：文档/方向
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：本次执行结论：DONE；冻结交付 `ec121b08`。源码基线、独立复核、实测项目、视觉范围、遗留门及产物路径见 [执行调度记录](execution-status.md) 和 [机器清单](execution-state.json)。

## 为什么要做

F28：当前 TODO 仍建议先做已完成的 flow/table/image/tools，纯文本 ticket 仍被 01 阻塞，而主分支已有公开 FlowDocument。另 feature-parity 的无已知漏洞断言与本次核实 advisory 不符。



## 当前代码与约定

[docs/capability-roadmap.md:5](/Users/wanghongyi/Projects/ofdrw.net/docs/capability-roadmap.md:5)

```
本文件是当前仓库的 **TODO 清单**。它把与
[ofdrw/ofdrw](https://github.com/ofdrw/ofdrw) 的能力差距、
[issue #3](https://github.com/whynpc9/ofdrw.net/issues/3)、
以及无密码厂商前提下的加密/签章边界收成一份可执行计划。

已实现能力仍以 [功能对照](feature-parity.md) 为准。对照表描述现状；本文件描述下一步做什么、做到哪、不做什么。

可开工的垂直切片在 [ticket 索引](../.scratch/capability-roadmap/issues/README.md)。按索引里的建议批次领票，不要按文件号把后面的 P1/P2 当成下一批。`Status: blocked` 的票要等阻塞项完成。
```

[docs/capability-roadmap.md:93](/Users/wanghongyi/Projects/ofdrw.net/docs/capability-roadmap.md:93)

```
建议批次：**P0-01/02 → P0-04/05 → P0-06/07/08/09**。布局是生成侧最大缺口；图片进出是最短路径；工具项可并行。P0-03 可紧跟表格之后。

## 5. P1：缩小与 ofdrw 的核心差距

含 [issue #3](https://github.com/whynpc9/ofdrw.net/issues/3) 点名的绘图 API 与字体体积。

| ID | 项 | 来源 | 落地 | 完成标准 |
| --- | --- | --- | --- | --- |
```

其余核查入口：[docs/feature-parity.md:25](/Users/wanghongyi/Projects/ofdrw.net/docs/feature-parity.md:25)；[.scratch/capability-roadmap/issues/README.md:33](/Users/wanghongyi/Projects/ofdrw.net/.scratch/capability-roadmap/issues/README.md:33)；[src/Ofdrw.Net.Layout/FlowDocument.cs:21](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Layout/FlowDocument.cs:21)。

**沿用范例**：[docs/feature-parity.md:25](/Users/wanghongyi/Projects/ofdrw.net/docs/feature-parity.md:25)

```
| 文档生成 | 布局、Canvas、Graphics2D、多种块元素 | 低层 builder、公开 FlowDocument/Paragraph/Span、Table/Row/Cell 与原生 OfdGraphics API | **部分支持**；中英段落/局部样式/分页、固定列宽/水平合并/整行分页与毫米路径/文字/状态/裁剪已实现；布局 Canvas 与复杂 Word 版式仍按各票范围交付 |
| 字体 | `ofdrw-font`，字体生成与字形处理 | 按全包用字生成静态 TrueType 子集、复合/GSUB/UVS 字形闭包、TTC 选面、内容身份载荷复用，缺字明确失败 | **部分支持**；TTF/glyf 子集已实现；CFF/变量/AAT/RTL 与保留未知 XML 的字体以全量和诊断保护，详见[字体子集契约](font-subset-and-reuse.md) |
| OFD → PDF | `ofdrw-converter` 导出能力 | 模板、图层、字体、文本游程、矢量路径、图片、变换、颜色/透明度和裁剪原点 | **已支持（预览）**；仍需更大票据/印章语料库做保真度回归 |
| PDF → OFD | `ofdrw-converter` 转换能力 | 默认双层；可选 Pdf.Vector 包经 19/04 输出受限 native PathObject/TextObject，校验原文与嵌入字体字形 | **部分支持**；[矢量模式](tutorials/18-pdf-vector-mode.md) 支持明确页框/路径/TrueType profile；压缩流、clip/图像/effects/复杂字体等明确整页回退或失败；默认 CLI 仍双层，扫描件无 OCR，不承诺任意 PDF 保真/阅读顺序标记 |
| DOCX → PDF / OFD | 上游不提供 Office 文档渲染 | 默认直接将 DOCX/OpenXML 原文写为原生 OFD `TextObject`，完全跳过 PDF；可选 `DualLayer` 使用 Word/LibreOffice/BuiltIn 页面图像，并继续以 OpenXML 原文作为语义层 | **已支持（预览）**；原文可直接抽取，Native 支持常见段落、表格、内嵌图片与按页页眉页脚；附属正文带标签追加，DualLayer 按实际页定位原文，不承诺复杂 Word 版式和浮动对象保真 |
| OFD → SVG | 上游推荐/生态具备 SVG 预览方案 | 自包含单页 SVG，支持模板、路径、文本、颜色、变换和内嵌图片 | **已支持（预览）** |
| OFD ↔ PNG/JPEG | 图片导出与一图一页导入 | `Converter.Pdf` 的 `OfdToImageConverter` / `ImageToOfdConverter` 及 CLI；指定零基页/ppm导出、PNG默认/JPEG、自然或固定页等比居中导入、有界暂存 | **已实现（本票源码）**；PDF/PDFium 预览保真边界，视觉门状态见 [Issue 02 验证记录](validation/issue02/README.md) |
| HTML 导出 | 转换模块覆盖 HTML 方向 | 无直接 HTML API | **未支持** |
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `docs/capability-roadmap.md`
- `docs/feature-parity.md`
- `.scratch/capability-roadmap/issues/README.md`
- `docs/capability-roadmap-development-plan.md`

此外可写本次任务的 `artifacts/plan-026/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-026-roadmap-state-reconciliation`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-026/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 建立范围与可验证样例

逐项对照公开 API/测试和当前基线，更新当前 TODO 与 ticket 索引的 ready/blocked/implemented 状态；标记旧 development plan 为历史设计或明确时点，不将历史不通过记录改为通过。只调整已可由本地证据确认的项。

**验证**：`rg -n 'TODO|ready-for-agent|blocked|已实现|未核实' docs/capability-roadmap.md docs/feature-parity.md .scratch/capability-roadmap/issues/README.md` → 每项状态有当前证据或明确未知。

### 2. 形成可审阅闭环

分别列源码已实现、已合并、视觉验收、已发布四种状态；未查询远端或没有当前产物证据的状态写未核实。纯文本的 Flow 前置已具备，HTML/归档仍为方向；依赖健康使用已核实 advisory 和检查日期，不宣称无漏洞。

**验证**：`git diff --check` → 退出 0；逐个核对改动本地链接存在，状态与未执行边界一致。

### 3. 记录完成边界

记录实际验证结果，未通过的门保持未完成。

**验证**：`git diff --check` 退出 0，`git status --short` 只新增/修改本计划允许的文件（保留原有用户改动）。

## 回归/样例矩阵

- `NoDuplicateWork`：flow/table/image/graphics/password v2 不再被当前索引当作从零待开发
- `NoFalseRelease`：本地源码能力不等于已发布 NuGet
- `HistoryPreserved`：旧验收与 not-reviewed 记录含义不变
- `LinksResolve`：改动的本地 Markdown 链接存在

每个场景必须记录预期、实际、结论及复跑入口；未知项不可改写成通过。



## Done criteria

- [ ] 当前状态有依据、本地链接有效、历史证据含义未改。
- [ ] 所有本计划要求的命令都有实际结果；未执行/跳过项明确列出，无零测试伪通过。
- [ ] 不要求新增视觉产物的本计划没有凭静态/文本结果声称视觉通过。
- [ ] 未扩大到无关源代码/依赖/发布修改。
- [ ] `git diff --check` 退出 0，改动只在 Scope（原有用户改动保持），索引更新真实完成范围。

## STOP conditions

- 发现源码与公开远端状态不一致时不要猜测；记录本地 baseline 并将远端状态标未核实，不自动发 issue/改 release。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

路线图更新与能力交付同 PR，当前索引只维护一处权威状态。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。
