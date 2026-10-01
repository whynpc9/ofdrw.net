# Issue16 公开表格验收（2026-10-02）

## 基线与范围

- 独立 worktree：`/Users/wanghongyi/.codex/worktrees/0ee7/ofdrw.net`；分支 `codex/public-tables`。
- 依赖 PR #9 保持开放未合并，起点准确为 `5d71dc9e4fb3ccfa4dd028386436e58cd261bc39`。本会话读取 GitHub 核对 head、六项绿色检查和 22/22 resolved review 线程；依赖视觉证据在该提交的 `flow-layout-2026-09-29.md`。
- 用户授权 stacked PR base=`codex/public-flow-layout`；不合并、不发布、不推进下一票。
- Astra High 先做源码设计与修复复核；Sol Low 独立功能与包消费验证；主任务实现及本次页面验收。
- 首版有限范围：[公开契约](../public-tables.md)。RowSpan 明确拒绝，DOCX vMerge 诊断或 Throw。Native 对不完整网格行（含未适配的 gridBefore/gridAfter）严格失败，不扩大任意 Word 表格保真。

## 功能与自动渲染

- 第一次 Layout 72/72、DOCX 130/130；修复后 Sol Low 独立全套 **295/295**（7 程序集，0 失败/跳过；Layout74、DOCX132、PDF52，其他37），日志/TRX `artifacts/public-tables/independent/full-tests-after-fix.log` 与对应 `test-results-after-fix/`（实际目录以下清单为准）。
- 11 个本地包隔离消费、manifest、CLI Native/DualLayer 通过：`independent/package-e2e-after-fix.log`、`independent/packages/`、`independent/package-output/`。新增外部 consumer 从 nupkg 使用 Table/Row/Cell，生成 3 页水平合并/底色/对齐表，核对文字完整。未发布公共 NuGet。
- 真实缺陷回归：旧多 M 路径经 PDF 导出出现斜线，独立 console 同页内容将单边路径合并后产生 **1810** 个格内黑像素；修复后 **0**。证据 `independent/negative-repro.log` 及 `negative-repro/`。PDF 像素回归另检查两格填色及预期边框，防空白产物假通过；新增正向断言的最终结果待填。
- Astra 复核极薄空行越框、合并列浮点最右边越框、继承样式格内分页诊断三个 P2；修复及针对回归已闭合。不支持的格内分页从已解析格式检查，warning/Throw 均覆盖。
- SDK 10.0.401，macOS，按 AGENTS 设置 writable DOTNET_CLI_HOME、关闭首次体验/遥测、单节点及禁用 build server；NUGET_PACKAGES 显式复用现有缓存。VSTest 沙箱 SocketException 权限限制，关闭 build servers 后在沙箱外运行。初次 NuGet 网络不可达是环境失败，复用缓存后构建成功；未清全局缓存。

## 页面验收进展

样例 `e2e/Ofdrw.Net.Layout.E2E/PublicTableSamples.cs` 生成 public 表格及同文 DOCX Native/default；还重用 `generated-layout.docx` 显式 Native/default。所有 PDF 都由本次 OFD 导出；程序分别核对完整 OFD 原文与 PDF 无重复文字，没有使用直接 DOCX→PDF 替代。

首轮 `artifacts/public-tables/initial/` 生成五组共 11 页。实际在 macOS Preview 打开 `tables-public.pdf` 第一页发现斜线缺陷，不能通过；已关闭旧窗口。该轮产物和页面 PNG 保留作为失败证据。修复后必须重新生成、打开并逐页复验。

最终源码提交、产物哈希、PNG 目视与 Preview 页码将在修复后验收时补录。当前完整视觉门保持未完成。
