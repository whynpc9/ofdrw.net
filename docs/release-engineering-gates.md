# Q-01～Q-05 发布工程闸门

本文件描述可重复执行的检查及其边界。**2026-11-01 状态复核：**`docs/release-preview-acceptance.json` 保存的是 2026-10-18 的 `0.1.0-preview.9` 历史 `accepted` 记录，检查人是用户当时明确授权的 Codex AI reviewer，记录含 `review_exception`，不能称为人工签收。该版本已有 GitHub Release 和公开 NuGet 包；这些历史结果不构成本轮中英文候选的验收或发布证明。

当前源码与历史指纹不匹配，新版本也不匹配历史记录。本轮候选按 [2026-11-02 受限 profile 决定](next-steps-2026-11-02.md)冻结最终源码、版本及实际产物，重新完成全部门禁；不得复用旧指纹或将预检写为 `accepted`。实际预检及待办见 [N12 发布工程预检](validation/cn-en-release-n12-preflight.md)。

源码指纹按路径绑定普通跟踪文件的当前字节；LFS 路径按已暂存的合法 pointer 的 oid/size 规范化，并严格核对工作树 pointer 身份或已下载载荷的大小/SHA-256。同一 LFS 身份的指针/载荷状态不改变指纹；篡改载荷、未暂存的新 LFS 身份、非法 pointer 和冲突路径都会失败。新增 LFS 文件或属性变更先暂存，再生成最终指纹；普通源码的未暂存变动仍改变指纹。

| 闸门 | 自动化入口 | 判定与留存 |
| --- | --- | --- |
| Q-01 许可与第三方声明 | `python3 scripts/verify-package-artifacts.py --notices-only`；`scripts/run-converter-package-e2e.sh <version>` | 仓库 MIT 与 nuspec MIT 表达式一致；每个直接 NuGet 依赖的版本、许可、来源与还原的包元数据相符；当前受审清单的 40 个直接及传递依赖的许可元数据与受审查清单一致；12 个同批包内声明与仓库一致，消费前后包 SHA256 不变。 |
| Q-02 视觉语料 | 同批包消费 E2E 中的 `ValidateVisualCorpusAsync` | 票据、发票、模板样式、有效异常布局四个合成无隐私单页 OFD，均经 OFD→PDF→144 DPI PNG，与固定 golden 的全页及标题/正文区域分别比较 RMSE ≤ 0.035；保留实际 OFD/PDF/PNG、差图和 CSV 指标。此门不代替 Preview。 |
| Q-03 坏包与结构 | `dotnet test tests/Ofdrw.Net.Packaging.Tests/...` | 固定坏包覆盖 `..`/反斜杠/绝对路径、重复条目、条目数、总展开量、压缩比、无效 ZIP；恰好等于预算的有效对照必须接受。低预算模拟炸弹，不创建巨大载荷。 |
| Q-04 XML 文档 | `python3 scripts/check-public-api-docs.py` | 开启 CS1591，按源码路径和公开成员符号对比历史清单。新增缺文档成员使 CI 与 tag workflow 失败；旧债减少可直接通过。 |
| Q-05 Preview | `docs/preview-acceptance-template.md`；tag workflow 的 `scripts/verify-preview-record.py` | 候选需有实际逐页查看记录，分别包含本次显式 native 与默认模式 OFD→PDF→macOS Preview、哈希、检查人、页码、缺陷；检查人须如实具名，AI reviewer 不称人工签收。记录模板使用 `not-reviewed`；仓库现存记录为旧 preview.9 的 `accepted`，其版本/源码绑定使其不能放行新候选。脚本仅校验证据字段和源码绑定，不能证明人实际看过页面。 |

## preview.10 的 profile-qualified Q-05

用户 2026-11-02 选择限定真实 Regular/Bold，并明列 Preview 搜索兼容性与原始 Copy 未完成。新候选 `cn-en-physical-cjk-400-700-v1` 的准确字体/参数见[当前计划](next-steps-2026-11-02.md)。固定原 DOCX 不变；Q05 的实际验收 OFD/PDF 必须来自该 profile 的本次精确候选包，两模式各两页均逐页 PNG/Preview 通过。不得用旧物理配置 PDF、直接 DOCX→PDF 或 Regular-only 输出充当该记录。现有 Q02 golden、RMSE 阈值、完整回归及其它发布门保持原样。

在原核心 JSON 字段之外，最终真实新候选记录须补充以下声明（核心 validator 不变；附加字段由协调者/Astra 复核，脚本不证明其内容或查看行为）：

- `release_profile`：对象，包含 `id=cn-en-physical-cjk-400-700-v1`、明确 `scope`、实际 font-directory/manifest 及字体 SHA-256、`NativeEmbeddedBoldBehavior=Throw`、两模式及其独立产物证据位置。仅有配置名称不证明实际选中了目标 face，须链接输出资源身份收据。
- `known_limitations`：分别保留 `legacy-regular-synthetic-bold`（旧视觉 FAIL）、`macos-preview-find`（macOS 26.6.2/25G83 已观察 Find FAIL、根因未定）、`raw-unicode-copy-unverified`（原始 Unicode 剪贴板字节未验证）。每项含原结论/证据、影响边界及用户 2026-11-02 范围决定；不写成已通过。
- 原 `open_findings` 继续表示**该明确 profile 内的开放视觉缺陷**，只有全部实际必验页面通过才可为空。新候选的功能/原文完整性/重复文本和强制门仍必过；不得把新的 profile 内失败移到 `known_limitations` 来绕过门禁。

因此新记录中的 `status=accepted` 仅表示具名配置和实际页范围的 Q05 视觉签收，不能表述为 Regular-only、Preview 搜索或原始 Copy 已通过，也不表示所有发行工程门或公开发布已完成。旧 `release-preview-acceptance.json` 保持原 preview.9 记录，直到取得最终新候选的真实产物与验收结果；本次文档不提前改 JSON。历史 Regular-only FAIL、旧矩阵及此前“未达到可发布验收”结论不重标。

## Q-04 历史 CS1591 清理顺序

当前历史清单在 `docs/public-api-cs1591-baseline.json`，按成员符号而非行号记录，包含两个 SDK 目标框架去重后的 273 项。依次清理：

1. `Converter.Abstractions` 与 `Layout` 的公共入口；
2. `Converter.Docx`、`Converter.Pdf`、`Converter.Svg` 的选项和结果；
3. `Reader` 与 `Packaging` 的读取、资源预算和错误契约；
4. `Core` 模型和 `Signatures`；
5. CLI 中可见的公共成员。

历史成员补完文档后**不要重新生成基线**；检查器会逐项报告减少的旧债。最后一项旧债清理完毕时，须经审查将基线更新为空数组，避免“没有捕获 CS1591”与“全部文档已补齐”无法区分。仅在审查确实需要调整基线时运行 `python3 scripts/check-public-api-docs.py --update-baseline` 并审查差异。

## Q-01 依赖许可复核

`docs/third-party-dependency-baseline.json` 当前固定 40 个直接/传递包的版本、nuspec 许可类型与值、项目 URL，以及以文件声明许可的内容哈希。旧版仅提供 `licenseUrl` 的条目保留原 URL，不把它自动推断为 SPDX 结论。依赖图变化会阻断检查；审核新增包的真实许可和分发影响后，才运行 `python3 scripts/verify-package-artifacts.py --update-dependency-baseline` 并审查清单差异。CI 安装的字体与外部 LibreOffice/Word 不在 NuGet 发布载荷内，其授权与来源按 `THIRD-PARTY-NOTICES.md` 分别说明。

## 发布候选顺序

1. 从固定提交运行 Python 检查、.NET 全套测试和同批包消费 E2E。保存包 manifest、OFD、PDF、PNG、日志。
2. 使用 macOS Preview 分别打开**本次**显式 Native 与默认模式的 PDF，两种模式各逐页核对源 DOCX。既有消费目录通常命名为 `generated-docx-native.pdf` / `generated-docx-default.pdf`；preview.10 必须选择来自上述严格物理 profile 的实际产物并记录准确路径/哈希，不能因沿用文件名而误签 Regular-only 对照。检查后填写 `artifacts/<candidate>/preview-acceptance.md`，记录问题和体积变化。直接 `generated-docx.pdf` 不属于 Native 链路。
3. 只有明确 release profile 内两种模式均无遗留视觉问题且各两页均已检查，才把 `docs/release-preview-acceptance.json` 更新为 `accepted`，并填写候选 `package_version`；preview.10 必须同时附上上述 `release_profile`、`known_limitations`，而不是隐藏范围外失败或未完成的语义验收。`source_fingerprint` 用 `python3 scripts/verify-preview-record.py --print-fingerprint` 获取；五个 SHA256 分别来自样例 DOCX、显式 Native 与默认模式的 OFD 和各自导出的 PDF。记录只改自身，源码指纹不变；所有非验收 JSON 的跟踪文件先冻结，再构建和验收；改动其他跟踪文件或改用新版本后必须重做候选绑定及必要验证，不能沿用旧 fingerprint。PDF/OFD 哈希是本机实际查看的证据；因生成文件可含时间元数据，不要求另一台 CI runner 的重建 ZIP/PDF 与本机逐字节相同。
4. 对该记录提交打 tag。发布工作流先复测、打包、消费原包并核对哈希，再校验记录，最后推 NuGet。失败时保持上个已验证包/标签；修复后使用新提交、新版本重新验收，不覆盖同版本包。公开源安装回读与下游生产环境签收仍需另行记录。

`.NET` 命令遵守根 `AGENTS.md` 的 `DOTNET_CLI_HOME`、首次体验/遥测变量与单节点 flags。当前记录是历史 preview.9 的 `accepted`；本轮未完成新的 Preview 验收，不能把旧记录或本轮绿测写成新候选视觉通过。
