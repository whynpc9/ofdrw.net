# Plan 028 follow-up: implemented plain text to native OFD

2026-10-16。冻结起点 `3247c76f95c9a3312cc26144e24c4984767af26f`；独立分支 `codex/audit-followup-028`；工作目录 `/Users/wanghongyi/.codex/worktrees/audit-followup-028/ofdrw.net`。实现提交 `7e336785e91c835a86919966cd59908893281ecd`。本文件及证据由后续文档提交保存。未推送、未建 PR、未合并、未发布。父协调会话负责五任务集成；004 排除。

本轮已经完成首版实现及本地验收。原两份 028 设计的“仅设计”状态保留为历史；此次用户后续实施授权覆盖其禁止实现的旧限定。未修改设计文件、共享 plans 索引或其他任务状态。

## 实现与源码漂移

新增 Layout 包中的 `PlainTextOfdConverter` / `PlainTextOfdOptions`。构造验证、快照全部 Flow 配置，使用单 FlowDocument / Paragraph / Span 以及既有分页器，每次调用生成独立 package。两条公开入口为 `Render(string, token)` 和 `RenderUtf8Async(Stream, token)`；所有新增公共成员有 XML 文档，库保持 netstandard2.0 / netstandard2.1，无新包或新依赖。

UTF-8 以 4096 字节块严格解码、EOF flush，跨 read 字符保持 decoder 状态。只去掉流当前读取起点的一个 EF BB BF；后续 BOM 与 string 起始 U+FEFF 保留。字符额度按原输入 UTF-16 code units 计算，字节额度含 BOM，超字节边界最多多读一个字节。输入不 seek、不读 Length、不关闭；异常或取消允许位置前移。非法 UTF-8 保留 decoder inner exception；孤立代理及 XML-invalid 标量拒绝，FF 作为 Flow 消费的控制符例外。

CRLF/CR/LF、内部空行、FF、空输入与末尾空行均沿用 Flow；不 Split、Trim、统一字符正规化或新增段距。空输入/仅控制符为一空白页。tab 保留且为四空格 advance，无 tab stops。U+2028/U+2029 为普通内容。字符/页数/元素额度、不能容纳的字形或行高拒绝；窄 Flow 调用边界包装 InvalidOperationException，I/O 与取消原样传播。取消覆盖入口、读取、解码块/标量验证及 Render 前后。

显式 `text-to-ofd` 的 parser 位于独立 `Cli.PlainText.cs`；Program 仅分发/白名单/help 接线。位置或命名路径二选一；支持家族、字号、成对页面宽高、各边距和四个预算。Invariant 数值、NaN/Infinity、重复/未知选项、混合路径、多余参数及非 .ofd 输出在读取/staging 前拒绝。先得到完整 package，再复用 AtomicOutput 写入并提交；成功消息只在提交后输出。既有 convert 推断不变。

**窄设计校正（ADR-028-1）**：原设计将 EnsureDifferentPaths 描述为仅词法比较，并等待 004。冻结源码实际上已调用 `ConversionPathIdentity.RequireDifferentPaths`，能解析文件/父目录符号链接并限制循环。因此直接复用现有保护，没有另建路径身份实现、没有实施本轮排除的 004。新增命令的同路径、文件/父目录符号别名拒绝测试在最终全套通过。保留现有共享 Flow/字体/转换源码不变。

## 功能验证

所有 dotnet 均使用可写 CLI home、skip-first-experience、telemetry opt-out，显式原全局包缓存；不清缓存。构建使用 `--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false`。整套/API Rebuild/包消费各在独占 validation.lock 下完成，同 worktree 无 Rebuild 重叠；GUI 锁与 validation 锁未嵌套。

| 门 | 实际结果 | 证据 |
| --- | --- | --- |
| solution restore | 退出 0 | 本地 artifacts/plan-028/restore.log |
| Layout + Flow filter | 89 passed，0 failed，0 skipped | [TRX](../docs/evidence/plain-text-ofd/layout-tests/layout.trx) |
| CLI + OutputSafety filter | 34 passed，0 failed，0 skipped | [TRX](../docs/evidence/plain-text-ofd/cli-tests/cli.trx) |
| API XML 文档 Rebuild | 退出 0，0 新增未文档化 API；双 TFM 编译 | [日志](../docs/evidence/plain-text-ofd/api-docs.log) |
| solution 全套 | 2868/2868，0 failed/skipped | [汇总](../docs/evidence/plain-text-ofd/full-test-summary.json)、同目录 full-tests/*.trx |
| 强制可信字体门 | 65/65，0 skipped | [summary](../docs/evidence/plain-text-ofd/trusted-font-tests/summary.json) |
| 本地真实包消费 | 11 包 `0.0.0-audit028` 全通过；隔离 PackageReference 消费器与安装 CLI | [日志](../docs/evidence/plain-text-ofd/package-e2e.log)、[包 manifest](../docs/evidence/plain-text-ofd/package-manifest.json) |
| OFD 正文 | notice/control/empty 逐页 TextCode 拼接 ordinal 精确匹配，不删空白 | [输入预期](../docs/evidence/plain-text-ofd/visual/source-expectations.json)、[生成检查](../docs/evidence/plain-text-ofd/visual-commands.json) |
| PDF 无重复 | 纯文本及 DOCX native/default 均与各自源文本一致（PDF 比较允许排版空白变化） | [DOCX 检查](../docs/evidence/plain-text-ofd/docx-text-check.json) |

行为测试覆盖合法多字节逐字节分片、单/双 BOM、当前位置/不 seek/借用流、所有换行与内部/尾部 FF/空行、空输入、空格/tab/NBSP/组合标记/补充平面字符 ordinal 保存回读、U+2028/2029、长行与直接单 Paragraph Flow 几何等价、UTF-8 非法/过长/截断/代理编码、XML-invalid 标量、孤立 UTF-16 代理、字节/字符/页数/元素/不可容纳额度、配置快照与并发调用独立性、入口和 read 中取消、I/O 原样传播。CLI 验证两种语法、参数前置拒绝、别名拒绝、旧输出保留与取消 130；AtomicOutput 既有取消提交测试同时运行。

包 runner 新增安装后的 `text-to-ofd` / `extract-text` 检查；消费程序新增 string / UTF-8 两 API，保存回读完整空白与 Unicode，并拒绝非 TextObject 内容。包消费同时执行既有 native/default DOCX 两页链路、图像/图形/字体子集及视觉 corpus 门。NuGet 既有 ImageSharp 警告未被隐藏，未扩大到依赖升级。

## 自动渲染、PNG 复查、Preview 验收

以下三项分开完成；PNG 没有冒充 Preview。

| 样例 | 实际链路 | 页数 | 自动渲染 | PNG 人工复查 | macOS Preview |
| --- | --- | --- | --- | --- | --- |
| notice.txt（含单 BOM、CRLF、长行、内部/尾空行、中英文比例文字） | UTF-8 → 新 adapter → native OFD → PDF | 2 | 2/2 | 1–2 通过 | 1–2 通过 |
| control.txt `A\n\nB\n\n`，40×30 mm，字号10、边距5 | 同上 | 3 | 3/3 | A / 空白 / B；尾空行不增页 | 1–3 通过 |
| empty.txt（仅字节 BOM） | 同上 | 1 | 1/1 | 单空白页 | 1 通过 |
| generated-layout.docx，显式 native | DOCX → native OFD → PDF | 2 | 2/2 | 1–2 通过 | 1–2 通过 |
| generated-layout.docx，默认 | DOCX → default/native OFD → PDF | 2 | 2/2 | 1–2 通过；与显式 native PNG 一致 | 1–2 通过 |

逐页实际查看全部 10 页正文与页边界。检查中英缺字/乱码、比例英文 iii/WWW、长行折行、空行、基线、分页、裁切、重叠、重影、异常空尾页；DOCX 另查居中/右对齐、蓝标题、斜体副标题、局部红色粗体、浅蓝表头、内外边框及两页分页。所选 generated-layout 样例无图片/页眉页脚/自动页码，因此不对这些内容作新视觉结论。Control 第2页与 empty 页是预期空白。

GUI 独占租约下重新打开本次最终 PDF，逐页保存完整 file URL/页码 AX、原始 JPEG 字节及 SHA-256，未重压或伪改扩展名；[页证据索引](../docs/evidence/plain-text-ofd/preview/pages.json)。文件选择期间曾打开本任务附带的图像/临时预览窗口，均已关闭；只有索引中匹配当前正式 PDF URL 的页面计入验收。最后状态回到 Open 面板，用户窗口未关闭。[清理记录](../docs/evidence/plain-text-ofd/preview/cleanup-final.json)。

5 个正式 PDF 的查看前后 SHA 完全一致：[before](../docs/evidence/plain-text-ofd/preview/pdf-before.json)、[after](../docs/evidence/plain-text-ofd/preview/pdf-after.json)。最终产物与 PNG 位于 [visual](../docs/evidence/plain-text-ofd/visual/notice.pdf)；本地原始路径完整保留在页 AX 和命令清单中。

字体条件：纯文本 OFD 仍为 name-only SimSun，没有字体嵌入或实际家族度量。当前 macOS OFD-to-PDF 使用 `/System/Library/Fonts/Supplemental/Arial Unicode.ttf`，SHA-256 `876af2cd4854644e7f3e7feb2f688997fdb3343c6df6693611209c9dfb47ccec`，与 PDF 嵌入字体 Ofdrw-876af… 身份一致。PDF 含子集字体、ToUnicode；这仅证明所查看中英样例在该 fallback 下显示正常，不宣称 SimSun 真实测量、通用字形覆盖或跨设备同外观。

首次新增样例没有历史体积可比较：notice OFD 3008 bytes，control 1951 bytes，empty 1124 bytes；没有加入隐藏原文、栅格正文或完整字体造成放大。PDF/页面图/截图等各文件体积、哈希见 [manifest](../docs/evidence/plain-text-ofd/manifest.json)。

## 来源与复跑

[源码身份清单](../docs/evidence/plain-text-ofd/source.json) 将实际验证使用的工作区文件内容与实现提交关联。测试和候选包在本地提交前生成，源码文件随后原样提交为 `7e336785…`；包 RepositoryCommit 元数据可能仍指冻结起点。没有声称其元数据为后续证据提交 SHA。包 bytes、输入/输出与源文件分别留存哈希，最终文档提交不更改实现。原始 AX 的行末空格、样例的 CRLF/尾空行是证据内容，使用证据目录内 .gitattributes 保留字节并限定空白检查规则，没有修剪或重新编码这些文件。

完整本地产物保留在 `artifacts/plan-028/`，候选包位于 `artifacts/package-e2e/0.0.0-audit028/packages/`。可移交的样例、OFD/PDF/10页PNG、10页原始截图/AX、所有 TRX、命令/环境/包/源码身份随 docs/evidence/plain-text-ofd 提交；商业系统字体不复制进仓库。

实际命令由 [validation-commands.json](../docs/evidence/plain-text-ofd/validation-commands.json)、[visual-commands.json](../docs/evidence/plain-text-ofd/visual-commands.json) 与包日志完整记录。复跑先设置同上 dotnet 环境并 prepare-trusted-fonts；全套/文档/包消费须按父协调规则取得 validation.lock，Preview 另取 gui.lock。正常使用见 [用户文档](../docs/plain-text-ofd.md)。

## 完成范围与保留边界

首版纯文本入口、本地功能/文档/双TFM/11包消费/自动渲染/PNG/Preview 门均闭合，ready_to_merge=true。父会话仍须在五项全部完成后集成并验证共同 HEAD；本任务未自行合并。

不扩 Markdown、富文本、表格、图片、语法高亮、任意字体测量/嵌入、RTL/shaping 或第二分页器。抽取不承诺逐字节可逆、原换行形式或尾换行恢复；LF/FF 是布局控制。支持平台声明止于双 netstandard 编译及本机 net10 测试，不凭此声称旧 .NET Framework/Mono/runtime 验收。既有 003 BuiltIn DOCX→PDF 缺陷由另一任务处理，未用直接 DOCX→PDF 产物替代本次 native 链路验收。
