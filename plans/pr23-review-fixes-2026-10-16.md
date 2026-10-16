# PR #23 Cursor 审查修复验收

本记录对应已复核生产源码 `979a93e39dd8cec3fbb4bdd2121d301ea84b2e39`，PR 为 https://github.com/whynpc9/ofdrw.net/pull/23 。本次在 `codex/audit-integration` 合入两项有效审查修复；本记录提交仅增加验收证据，不改变已测试源码。远端新提交的 CI 和两家机器人复审须另行闭环，本地通过不代表远端已通过。PR 保持 OPEN，未合入 main、tag 或发布。

## 两项修复

P1（进程退出状态与清理）：原包在真实 Linux ARM64/musl 容器 PID 1 下，实际退出 0 的目标抛出 ECHILD。新实现使用独立原生 supervisor 作为目标的唯一父进程，通过私有协议返回真实状态；保留目标的 waitable 身份后才清理进程组，托管宿主不再按裸数字 PGID 发信号。不接管宿主全局 SIGCHLD、不在 CLR 中 fork、不将 ECHILD 伪装成退出 0。修复分支最终提交 `fc12ba7df1c7d9d3acd18a84cb13be86540ce85d`。协调者完整阅读 C、托管协议/提取、测试、构建/许可增量，并独立重跑 Core 21 项及原生协议 9 项；Windows 实现未改变。

四个不可变资产覆盖 Linux x64/ARM64（静态 musl）和 macOS x64/ARM64。源码、固定工具链、原生执行证明、完整 musl 许可及构建时哈希校验随源码交付，普通消费无需编译器或 Python。Core 和 CLI 均携带完整许可，包验证器对缺失和被改动许可均有回归。资产资格不等于发布许可。Linux 最低内核和老版本 macOS 未获实际执行认证；部署头目标不冒充运行证据。helper 被 SIGKILL/崩溃、内核不可中断状态、主动逃离进程组的目标仍有明确边界，不能安全证明时报告清理失败。详见 [交付方案](024-process-reaper-delivery-proposal.md)。

P2（持久化拆分附件）：原实现对 Doc_0、根目录和嵌套文档的附件索引/载荷分配 `.1` 副本。现在仅在保留同一文档条目身份时，将实际源附件映射到克隆并保留其来源所有权；普通新建冲突保护、显式删除和未知外链不抓取均保持。跨文档不透明引用继续拒绝。实现 `2acfaed0503b2d59f2a5ebc1161d804a4cf2fbf0`，最终分支 `50f3ccae73cf11f14d937084f2cf8d08b6a26524`。协调者阅读全部生产增量与 14 项新增回归，独立重跑附件/预检边界 133/133，核对 40 个清单文件及 69 个 ZIP 附件哈希。详见 [附件记录](027-pr23-split-attachments.md)。

## 当前源码功能与包验证

证据目录为 [artifacts/pr23-review-fixes](../artifacts/pr23-review-fixes/curated-evidence.json)，其中清单保存每个保留文件的 SHA256。所有 .NET 调用使用可写 CLI_HOME、禁止首次初始化/遥测和单节点禁用 build server 参数；先 restore。全套 build/test/docs/pack 持有共享 validation 锁。

| 门禁 | 实际结果 |
|---|---|
| API docs Rebuild / 完整方案 | 通过；11 个 TRX，3059/3059，零失败、零跳过 |
| 强制可信字体 | 65/65，禁止字体静默 skip |
| Python 脚本 | 31/31 |
| 新版本 `0.0.0-pr23.20261016.1` | 11 包真实消费通过；纯文本、HTML、只读预检新增 API/CLI 实际消费通过 |
| 新包进程消费 | Core 两个 TFM，28 组、252/252 断言；四个 RID、六个实际 ABI/libc 环境 |

进程矩阵覆盖 macOS ARM64 与 Rosetta x64；原生 Linux ARM64/x64 的 glibc/musl，分别运行普通宿主、真实 PID 1、显式忽略 SIGCHLD。断言实际架构/PID、真实退出 0/7、参数/空参数/Unicode/cwd/env、ENOENT/EACCES、超时、取消、包装父进程先退出后的后代清理、四资产嵌入与提取权限。物理包和实际加载 DLL 的哈希均记录。Ubuntu CI 新构建包还需在该次 CI 产物就绪后单独消费，不能用这组本地 macOS 打包结果代替。

保留首轮矩阵的真实失败：27/28 组通过，剩余一组仅即时存活探测失败。旧断言没有清理前的 starttime/PGID，也未保留原始 stat 和清理异常数据，无法据此区分内核退出过渡和仍在执行的后代。改进的观察器在清理前绑定身份，检查清理诊断为空，并要求在有限两秒内确认原身份消失或僵尸状态。产品源码未改。针对性 100/100 次跟踪均通过，最长观察 0.2255ms，首次样本都已消失；没有捕获到暂态非 Z，故不声称直接观测到了原失败原因，也不把它归为基础设施抖动。原失败、冻结旧 harness、原始结果和完整第二轮矩阵均保留。

与上一轮本地包比较，Core 压缩包 363421→453127 字节（+89706），CLI 25459050→25506876 字节（+47826），包含原生资产、协议代码和许可。字节身份与版本均不同，不能据此推断任意业务文档体积。包、完整缓存、运行时和宿主 Arial 字体保留在本地，不提交为审查证据；清单记录物理包哈希。测试的源差异和原始文本以 JSON 保真保存，避免改写原始空格。

## 当前产物视觉验收

采用确定性 `generated-layout.docx`。显式 native、默认、CLI native 均检查本次 DOCX→OFD→PDF→Preview（各 2 页）；源第 2 页选择（1 页）、纯文本 notice（2 页）、direct BuiltIn（2 页）、真实/模拟斜体边界（1 页）也完成逐页检查。另独立检查新 supervisor 启动实际 LibreOffice 的 DOCX→PDF→Preview（2 页），该结果不代替 native OFD 产物。

- 功能/自动渲染：8 个 PDF、14 个 PNG，生成与文本/转换检查通过。
- PNG 逐页目视：14/14；正文中文、英文间距、局部样式、表格填充/边框、对齐、分页和裁切均检查。
- macOS Preview：14/14，14 张原始 JPEG 及各页完整 URL/page AX 证据保留。PDF 查看前后 SHA256 相同；只关闭本任务窗口，gui 锁已释放。
- 实际浏览器：4 个 HTML、9 个页面出现项，检查宿主字体、混合媒体、裁切与 ID 作用域；无脚本、事件属性和外部请求，ID 唯一、物理页面比例误差小于 0.000006。HTML 前后 SHA256 相同，任务空间已结束，用户的登录空间未接管。

协调者复核全部 Preview URL/页码/AX/JPEG 格式及 SHA、PDF 前后身份、HTML 实际 DOM/请求/截图身份，并抽看 native 浏览器、LibreOffice 和斜体边界页面。详细范围和观察见 [视觉汇总](../artifacts/pr23-review-fixes/visual-acceptance-summary.json)、[Preview](../artifacts/pr23-review-fixes/preview/report.md) 与 [浏览器](../artifacts/pr23-review-fixes/browser/report.md)。

保留 direct BuiltIn 较轻的 CJK 字重及引擎布局差异，不声称跨引擎像素一致或补上粗体模拟；长 notice 的默认 HTML 节点预算拒绝保持，仅短 notice HTML 获实际验收。本组 PDF 不覆盖图片/页眉页脚/页码专项，不推断任意复杂 Word 均保真。原 004 Windows 链接权限跳过不改写为通过，既有 landscape PageSetup 未修；028 portable metrics、029 静态资源/预算/宿主字体回退、030 六条本地 APF 四态只读且不作认证/验签/解密/字体许可验证等边界继续保留。
