# D09：Windows Audit025 分配异常有界处置

结论：**关闭本轮主动取证，保留原阈值、现有门禁与被动观察；历史 7,760-byte 异常根因 INCONCLUSIVE，未修复。** 本轮已读范围没有新增可比较的 Audit025 自然失败，不以随后通过或绿色 CI 解释历史失败。没有产品、测试、依赖、fixture 或发布变更。

## 基线与成本

干净工作树 fetch 后，从最新 `origin/main` `a1c0327e40cc378009b8528f70b99556ceb21180`（D08／PR104）建立 `codex/d09-audit025-passive-disposition`。开发协调 gpt-6.1-sol High；独立判别审阅 gpt-6-astra High；独立既有证据核对 gpt-6.1-terra High。

主动取证起点 `2026-11-02T15:52:39Z`（北京时间 23:52:39），包含委派核对的硬上限为 30 分钟；完成独立审阅及措辞复核后于 `15:57:47Z` 停止，实际窗口 **5 分 08 秒**，其后仅整理记录和执行 PR 门禁。只读必要源码、现有日志、小型测试 artifact 与运行元数据；不复制大归档。新增压力／复现循环、跨平台实验矩阵、CI 重跑均为 **0**。后续文档 PR 照常执行既有 exact-head CI，不作为根因实验或修复证据。

## 已读证据与可比较性

| 证据 | 判断 |
| --- | --- |
| [原失败 run 41258810433](https://github.com/whynpc9/ofdrw.net/actions/runs/41258810433)，main `f2ed06c6fb56b4b470b1425c5bafaddb6a02812d`，attempt 1；Windows job `124707605952` | 本轮读到原 job 日志：`Audit025_ElementOverloadNoDelta(delta: " \t\r\n")` 对 4,096 个 `A` 的调用记录 `7,760 bytes`，违反原 `<4096` 预算；不是把真实失败改称噪声。runner 标签 `windows-latest`，Windows Server 2025 `10.0.26100`，镜像 `windows-2025-vs2026`／`20261021.278.1`；setup-dotnet 记录安装 runtime `10.0.14`、SDK `10.0.403`，不等同于逐调用运行时身份。该版源码在同一测试进程 warmup 后测一次线程计数差，没有现版 no-GC 窗口与完整原始收据；不足以区分真实产品分配、运行时记账或测量干扰 |
| [PR85](https://github.com/whynpc9/ofdrw.net/pull/85) 与[分配边界](audit025-allocation-boundary.md) | 已交付隔离子进程、窗口有效性、原始计数及正／负控；没有产品修复。新旧协议不同，后来的 48-byte 结果不能倒推原 7,760 bytes 的成因 |
| 最近失败的 main run [41278534258](https://github.com/whynpc9/ofdrw.net/actions/runs/41278534258)，`b73e3ab9c52d9e1b9c3c65aef0fb7ec92bed066d`，attempt 1 | D08 已读取的 Linux `test-results` 唯一失败为 Audit024 启动取消时限，非 Audit025。本轮另读取现有 Windows `full-platform-tests-windows-latest` artifact（ID `13072068796`，API 压缩大小 1,218,299 bytes）：Audit025 通过，normal 的 96 个 measured target 均为 48 bytes，regression-control 按预期被拒绝，独立 pressure 留存 128 区间；这些不是新的自然分配失败，也不关闭历史根因 |
| 最近 main CI 元数据 | 最新失败仍是上述 41278534258；其后已完成的 D01–D07 主线运行成功，D08 基线运行 `41474314501` 查询时尚在运行。此范围未发现新增失败，不证明所有未读记录或未来运行均无异常 |

所读 Windows 收据自报 `.NET 10.0.14`、Windows `10.0.26100`、`win-x64`／X64、workstation GC，并留存实际 Core 路径／SHA／版本、fixture SHA、线程、原计数和窗口状态。它们用于识别该次测量环境，不是原失败的补造收据。

## 判别与停止条件

- **预算探针计量无效**：no-GC 预算样本原始计数为负／递减、窗口启动失败或失效、线程／GC 次数不符、结束失败，都按现有门禁失败；不能截零、取最小值或重试选绿。此项不是 pressure target／empty-spin 观察值的有效性门。
- **预算失败**：no-GC 预算探针有效窗口中任一 measured target `>=4096` 即失败；每区间 1 次 warmup 与 8 次 measured 全部保留，warmup 计数也须有效。空操作和正控／旧按字素负控须符合各自判据。不放宽阈值、不平均、不跳过测试。
- **归因未定**：有效超预算只证明该次违反预算，仍须可比较输入、源码／Core／fixture 身份、运行时／架构／GC、原始窗口和控制结果才能判断实际回归。不同计量窗口下的通过不能证明修复；pressure 中空操作的正增量仅支持测量干扰假设，不能把历史失败直接判为噪声。pressure 自身的完整性、时限和控制检查仍可使 CI 失败，其 observed target／empty deltas 不作预算重试。

没有新增判别证据即停止，不用满预算，也不制造失败。仅当新的自然失败带来输入与可比较测量证据，才先报告协调、提出明确判据和有限预算；不自动恢复压力循环或扩大矩阵。现有 CI 新失败仍须分类，不默认忽略。

## 独立核验

Astra 独立审阅后批准关闭理由；其要求将计数有效性规则明确限定为 no-GC 预算探针的意见已修改并复核。Terra 独立检查现有 12 份 TRX（零失败、21 项 Audit025 通过）、3 份 allocation 收据与原 Windows 日志，并复核该措辞修订。主线程重算 22 项源码／证据哈希均一致。核验收据保存在本工作树 ignored `artifacts/d09/terra-validation.json`，SHA-256 `18d86dc3d9ec2dc87fc4588eb489feb4b6f303be34aab6f788f8db1a5936ebd2`；其中 D09 文档哈希对应添加本节及结束时间之前的已审阅正文。run／attempt／artifact 归属由主线程通过 GitHub API 和下载命令取得，Terra 不宣称独立验证该远端来源。

本轮本地 build/test、包消费、转换与自动渲染均未运行；没有新页面，PNG 与 macOS Preview 均为 **0 页**，未启动 GUI。原文、重复文本和视觉验收不作新结论；不改写 preview.10 标签或发布内容。
