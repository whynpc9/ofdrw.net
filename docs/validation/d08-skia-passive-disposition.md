# D08：Skia 首失败证据有界处置

结论：**关闭本轮主动取证，保留被动观察；历史 native 字体拒绝根因 INCONCLUSIVE，未修复。** 本轮已读取范围没有新增能区分提前 Dispose、缓存串用或依赖缺陷的自然首失败，不能据此证明依赖无缺陷。没有产品、测试、依赖、fixture、字体或发布变更。

## 基线与成本边界

干净工作树 fetch 后，以 `origin/main` `c37ed5168a87a7bc5e00fbabde9d40f3cb24d587`（D07／PR103）建立 `codex/d08-passive-skia-evidence`，未使用旧 preview.10 分支。开发协调 gpt-6.1-sol High；独立判别审阅 gpt-6-astra High；独立既有证据核验 gpt-6.1-terra High。

主动取证起点 `2026-11-02T15:22:27Z`（北京时间 23:22:27），硬上限 30 分钟，包含委派核对；两位代理完成证据核对后于 `15:27:41Z` 停止，实际窗口 **5 分 14 秒**，其后仅整理文档与执行 PR 门禁。没有新建压力／复现循环、跨平台实验矩阵、CI 重跑或无新证据重跑；不要求用满预算。后续文档 PR 仍执行现有配置的 exact-head CI 门禁，不以其成功解释历史根因。

## 读到的首失败与最近自然失败

| 证据 | 本轮判断 |
| --- | --- |
| [历史首失败运行 41258810433](https://github.com/whynpc9/ofdrw.net/actions/runs/41258810433)，main `f2ed06c6fb56b4b470b1425c5bafaddb6a02812d` | [既有诊断说明](skia-font-read-evidence.md)记录 macOS `EqualIdsDoNotShareDifferentFontIdentityAndLateMismatchIsAtomic` 构造第二个原始 `wide.ttf` binding 时拒绝，构造期 SKData 仍在作用域内；没有原始 native 返回值，不能归因 Dispose、缓存或依赖。此处复用历史记录，不声称重新取得该轮所有原始日志 |
| 最近已失败的 main CI：[41278534258](https://github.com/whynpc9/ofdrw.net/actions/runs/41278534258)，`b73e3ab9c52d9e1b9c3c65aef0fb7ec92bed066d`，attempt 1，2026-11-01 05:51:40 UTC 创建 | 读取该运行现有 `test-results` artifact（ID `13071883977`，API 压缩大小 1,195,787 bytes），11 份 TRX 中唯一失败为 `Audit024ProtocolTests.Audit024_StartupProtocolCancellationIsBounded` 的 `Assert.True`；不是新的 Skia 有效字体失败，不在 D08 扩展调查 |
| 同一小型 artifact 的 11 份 Skia JSON | 7 份 controlled injection 初始 guard、1 份 controlled injection 装饰后传播事件；3 份 uninjected guard 均无 raw-style preflight。全部自报 `sourceSha` 为该运行 head（来自环境变量，不是独立源码认证）；零份 `uninjected-guard-after-raw-style-preflight` scope 记录（8 份受控注入本身的 preflight 标志为 true）。这些记录不构成有效字体自然复发证据；其中装饰后传播事件不是原始 first-chance |
| 最近主线 CI 元数据 | D01–D06 对应的 `2c49e0ef`、`e14a7ac3`、`e4c67ef6`、`259b7fbf`、`54719ef3`、`a9de8cae` 已完成成功；基线 `c37ed516` 的 CI `41469409447` 在首次查询时仍运行。通过不是根因修复证据；本轮不为等复发重跑 |

批量日志及单 job 日志端点各遇一次 EOF，随后改读上述约 1.2 MB 的现有测试 artifact 成功；未继续重试日志或下载大型页面／字体归档。只核对必要 TRX、JSON、元数据与源码，不复制历史源码树。

## 现有采集足够，不重复开发

[已交付诊断](skia-font-read-evidence.md)、`FontReadTrace`／`ValidatedFontBinding`、`SkiaDrawEvent`、测试宿主 `SkiaFontDiagnosticHost` 与 CI 上传配置已覆盖：不可变克隆长度／SHA、OS/2/head 原始结构与 style、generation、wrapper／handle、线程、阶段／ordinal／单调起止时间、首次 size／byte 实际结果，以及源／managed／OS／RID／已加载模块路径与磁盘哈希。首份记录独占创建、不覆盖；采集失败不替换原始异常，预算有限。不增加相同字段或诊断循环。

raw-style preflight 只检查限定结构／样式，不是完整字体有效性证明；即便出现 after-preflight guard，也须结合实际 fixture 和失败上下文分类。模块路径哈希不是内存哈希或逐调用地址认证。采集是 best-effort，缺 JSON 或本批零 uninjected after-preflight scope 记录不证明未发生复发；独占创建不等于完整 JSON 原子发布，写入失败可留部分文件。历史[复盘](pr78-and-stability-retrospective-2026-11-01.md)中的 2,400 次阴性及成功／注入回归不构成修复。

## 核验与后续触发

Astra 独立审阅关闭判据与现有代码／证据；已处理其“自然复发与 first-chance 分开判断”的意见。Terra 独立核对字段、11 份 TRX 与 11 份 JSON，主线程重算 22 项文件哈希均匹配；核验收据保存在本工作树 ignored `artifacts/d08/terra-validation.json`，SHA-256 `b168e1ae526b9d88ee3074925f38841b735937b7cbc515b80d0d149b7a99c2f7`。远端 run／attempt／artifact 归属由主线程通过 GitHub API／下载命令取得；Terra 不宣称独立核实该远端来源。

本轮本地产品 build/test、包消费、转换、自动渲染均未运行；无新页面，逐页 PNG 与 macOS Preview 范围均为 **0 页**，未启动 GUI。不把既有视觉报告或远端 CI 当成本轮 Preview。原文、重复文本、视觉门不作新结论；preview.10 标签及发布内容不改写。

只有新的自然首失败带来可区分生命周期／缓存身份／依赖机制的证据，才提出明确输入、判据、预算和停止条件的下一步；先报告协调，不自动恢复压力实验或长期研究。现有测试、有限采集及 CI 门禁保持。新 CI 失败仍须分类，不默认忽略；没有新判别证据即可结束本轮，不以“尚未复现”宣称修复。
