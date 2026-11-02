# 系统原生 PDF 对照的既有证据复核（2026-11-02）

**判定：既有系统生产者对照已复现真实 Preview Find 失败，本轮停止新增实验；N02 仍 OPEN，根因 INCONCLUSIVE。** CoreGraphics/CoreText 直接绘制、使用系统 Helvetica 的 `2026-01-01` 和 `Alpha`，在既有 Terra 实测中也出现页命中但不可操作的结果。这反驳“只有 OFDRW 生产者或原嵌入字体才能产生该现象”的排他解释，不证明所有 PDF 都有问题，也不证明产品没有独立缺陷。

本轮在专用分支 `codex/cn-en-preview-control`、基线 `1dcbaaa09f6debd611bb81c7ef97ecd493c3f89f` 只复核来源，不修改产品。新增授权为一个系统生产者对照问题、最多 90 分钟主动工作；不重置 [N02](cn-en-release-n02.md) 已结束的两个假设预算。找到既有实验后，协调者要求停止新增 PDF/GUI。该指令到达前曾生成一份临时两页中文 PDF；它只在 ignored artifacts 中，未由 Preview 验证、未纳入任何结论，生成脚本未提交。没有运行新的 .NET 构建、测试或 GUI，也没有重新执行旧生产者。

## 原始输入与真实 Preview 结果

以下两组资料均位于本机已有根目录，文件不是本轮重新生成：

`/Users/wanghongyi/.codex/worktrees/9c7f/ofdrw.net/artifacts/r22-cjk-filled-stroke/`

`native-font-word-controls/produce.swift` 使用 CoreGraphics PDF context 和整行 `CTLineDraw`，对比原 subset 字体与系统 Helvetica；可见黑色 fill，字号 10.001 pt，起点 `(472.1472745, 686)`，A4 页面。每份只有一个字符串，所以该字符串也是页末及文档末文本；没有 raster、隐藏哨兵、重复语义层或 PDF 后处理。系统 face 为 `/System/Library/Fonts/Helvetica.ttc`，原收据记录 SHA-256 `25eceb458d4baf628ee0b6a135a9f8ac5ec7b2826646720834bd2bf00dfc825a`，不是原 subset 文件身份。

| 旧控制文件（根目录内） | 字节 | SHA-256 |
| --- | ---: | --- |
| `native-font-word-controls/produce.swift` | 4,290 | `ee9443edd9d27210c151247ba50b46f9ced338b7ab8a42b3c3f5bd7541255502` |
| `native-font-word-controls/helvetica-date.pdf` | 6,608 | `2a93dab21e7213bb9b8977921fa0bd38b6b3d028440b70c94790138f0f44c5a3` |
| `native-font-word-controls/helvetica-word.pdf` | 6,710 | `1285418794fd32a9ee7dce68ce589bff4f3219dd93d191c7fa2f1f97995ca774` |
| `native-font-word-controls/preview-review/receipt.json` | 29,453 | `81afd5e5628f6902a81d06f4374ecfd6a632f6f1551074eae8227e7f86ba50ab` |
| `native-font-word-controls/root-closure-audit.json` | 3,993 | `9abf3546cbb0ca10bc9b569735219f57398d445283ff41a4cf034744d1d7cda7` |

真实验证者为 `gpt-6.1-terra` / High，原日期 2026-10-30，macOS 26.6.2 (25G83)。两份 Helvetica PDF 的查询各预期 1 处，均实际报告 page hit 1、`Page 1 Not found`、Next disabled、无可见数值命中计数；清空普通选择后，两次观察和一次结果行激活均无字形对齐的 Find 高亮。源控制中的 `Alpha` 是末词，不能与产品中的内部 `Alpha` 成功混同。完整四份原字体/Helvetica × 日期/Alpha 均同样失败；没有依据只归咎 subset 字体。

更早的 `native-coretext-search-controls-677205e1/` 也直接用 CoreGraphics/CoreText，但输入为原 subset 字体，查询分别为 `2026-01-01` 和 `量A`。原脚本 SHA-256 `733db3bcac69596aa74b80b8f974691dd63acf0f5875bdc930862263dfa5d895`；两 PDF 为 `date-same-font.pdf`（13,694 字节，`2e1a12421b2300a9269b18d4b501ce18811bae149c25113c92c0180f44d10d53`）、`mixed-same-font.pdf`（3,782 字节，`fdcef289b40a1e537d27d1fb3b00b0fd31870e76cc5da0810af07f31f7b5a462`）。其真实 Terra 收据 `preview-review/receipt.json` 为 `015758732ba57e9618cdb9f4e06e351fbccbe23573568f0c316374007ccdaa1e`；同样没有可操作的 Find 结果。该组原有蓝色全选背景不能算 Find 高亮。

旧 `native-control-xref-limitation-audit.json`（`43684db889afa1cfd5b60f5c4d2f2dc1e343536a30beedfc80cfa7055a9bfc43`）保留了原 subset 控制存在未引用 offset-0 xref 项的限制；Helvetica 两份无此项，也仍失败。不能将该限制归为产品缺陷或用它解释全部控制。所有结果只是指定主机、指定文件和原 UI 操作的观察，不是普遍 PDF 合规性结论。

## 本轮复核、产品身份与退出

开发者只读取脚本、输入、原真实 GUI 收据和 coordinator audit；递归核对 **146 个 path/expected-hash 记录**，全部可用、无哈希不符（同一路径可有无期望值两条记录，不是 146 个独立实验）。本轮记录：

`/Users/wanghongyi/.codex/worktrees/c52e/ofdrw-preview-control/artifacts/cn-en-preview-control-2026-11-02/existing-evidence-audit.json`

SHA-256 `03ab1e602a58b1c4b7ac7fb13906fe9eb8a0aa28973a5f0eb351faa415d3f0cf`。原收据所引用的页面、查询及激活截图文件也可用且哈希吻合；开发者本轮没有重新观看旧截图，不宣称新的 PNG/Preview 视觉验收。原窗口关闭状态沿用收据，本轮没有打开窗口。

对照产品仍是 [本地独立验证](cn-en-release-local-validation.md) 的 **ORIGINAL 配置**，不是 physical-bold 替代：`core-v2/cn-en-release-manifest.json` SHA-256 `b58358897e9a9c690e34027fc6f5d0fb370d88e40b1d1b451a3aeed9907f4564`，源码 `1e6f5eefb60b18ea9951769273b749a27ac9c174`。本轮重新核对该 manifest 及其原版 native/default DOCX/OFD/PDF 身份，文件未修改：

| 现有 `core-v2/` PDF | SHA-256 |
| --- | --- |
| `layout-original-native.pdf` | `bd19c3f3c47e64327d9a8665fae3c57acc5eabcaa118b4d739b382c498c4e282` |
| `layout-original-default.pdf` | `21c49b8eb8daa4cd62a4044d10ba8777cd87c7524a09520bce1a7d0472a6adbe` |

产品旧已知结果继续有效：original native/default 完整日期在实际 Preview 失败；基准 N02 的 `待复核` 也失败，内部 `Alpha` / `测试日期` / `2026` 成功、负控 `2026-01-02` 无结果。系统旧控制没有覆盖 `待复核`、中文日期前缀或负控；不能把它补写成这三个查询也被系统控制验证。其日期重现已足以回答本轮生产者排他性问题，因此不追加中文、分页或字体矩阵。

最小可共享复现为 [系统 Helvetica 控制脚本](../../scripts/validation/preview-native-terminal-control.swift)，从上述旧脚本提取两个公开虚构字符串、同一 face/字号/位置/整行绘制流程，移除了原 subset 输入依赖。**本轮未执行该提取脚本**；已实测 PDF 的身份和结果仅以上述旧原脚本及收据为准，新运行的字节也不得假定相同。未来如需向上游提交，可在单独授权的复现中运行，记录系统版本、输出哈希和实际 Preview 结果。

功能结论限于现有内容/位置预检和本次身份核验；自动渲染、PNG 复查及 macOS Preview **没有新增验收**。退出为 **producer-only / original-font-only 排他解释被反驳，具体机制仍 INCONCLUSIVE**；不修改产品、不豁免 Copy/完整 Find/原 Regular 视觉门禁、不关闭 N02、不发布、不推送。
