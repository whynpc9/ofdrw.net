# Q-01～Q-05 发布工程闸门

本文件描述可重复执行的检查及其边界。**2026-11-01 状态复核：**`docs/release-preview-acceptance.json` 保存的是 2026-10-18 的 `0.1.0-preview.9` 历史 `accepted` 记录，检查人是用户当时明确授权的 Codex AI reviewer，记录含 `review_exception`，不能称为人工签收。该版本已有 GitHub Release 和公开 NuGet 包；这些历史结果不构成本轮中英文候选的验收或发布证明。

当前源码与历史指纹不匹配，新版本也不匹配历史记录。本轮候选必须冻结最终源码、版本及实际产物，重新完成全部门禁；不得复用旧指纹或将预检写为 `accepted`。实际预检及待办见 [N12 发布工程预检](validation/cn-en-release-n12-preflight.md)。

源码指纹按路径绑定普通跟踪文件的当前字节；LFS 路径按已暂存的合法 pointer 的 oid/size 规范化，并严格核对工作树 pointer 身份或已下载载荷的大小/SHA-256。同一 LFS 身份的指针/载荷状态不改变指纹；篡改载荷、未暂存的新 LFS 身份、非法 pointer 和冲突路径都会失败。新增 LFS 文件或属性变更先暂存，再生成最终指纹；普通源码的未暂存变动仍改变指纹。

| 闸门 | 自动化入口 | 判定与留存 |
| --- | --- | --- |
| Q-01 许可与第三方声明 | `python3 scripts/verify-package-artifacts.py --notices-only`；`scripts/run-converter-package-e2e.sh <version>` | 仓库 MIT 与 nuspec MIT 表达式一致；每个直接 NuGet 依赖的版本、许可、来源与还原的包元数据相符；当前受审清单的 40 个直接及传递依赖的许可元数据与受审查清单一致；12 个同批包内声明与仓库一致，消费前后包 SHA256 不变。 |
| Q-02 视觉语料 | 同批包消费 E2E 中的 `ValidateVisualCorpusAsync` | 票据、发票、模板样式、有效异常布局四个合成无隐私单页 OFD，均经 OFD→PDF→144 DPI PNG，与固定 golden 的全页及标题/正文区域分别比较 RMSE ≤ 0.035；保留实际 OFD/PDF/PNG、差图和 CSV 指标。此门不代替 Preview。 |
| Q-03 坏包与结构 | `dotnet test tests/Ofdrw.Net.Packaging.Tests/...` | 固定坏包覆盖 `..`/反斜杠/绝对路径、重复条目、条目数、总展开量、压缩比、无效 ZIP；恰好等于预算的有效对照必须接受。低预算模拟炸弹，不创建巨大载荷。 |
| Q-04 XML 文档 | `python3 scripts/check-public-api-docs.py` | 开启 CS1591，按源码路径和公开成员符号对比历史清单。新增缺文档成员使 CI 与 tag workflow 失败；旧债减少可直接通过。 |
| Q-05 Preview | `docs/preview-acceptance-template.md`；tag workflow 的 `scripts/verify-preview-record.py` | 候选需有人工逐页记录，分别包含本次显式 native 与默认模式 OFD→PDF→macOS Preview、哈希、检查人、页码、缺陷。记录模板使用 `not-reviewed`；仓库现存记录为旧 preview.9 的 `accepted`，其版本/源码绑定使其不能放行新候选。脚本仅校验证据字段和源码绑定，不能证明人实际看过页面。 |

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
2. 使用 macOS Preview 分别打开**本次** `generated-docx-native.pdf` 和 `generated-docx-default.pdf`，两种模式各逐页核对源 DOCX。检查后填写 `artifacts/<candidate>/preview-acceptance.md`，记录问题和体积变化。直接 `generated-docx.pdf` 不属于 Native 链路。
3. 只有两种模式均无遗留问题且各两页均已检查，才把 `docs/release-preview-acceptance.json` 更新为 `accepted`，并填写候选 `package_version`。`source_fingerprint` 用 `python3 scripts/verify-preview-record.py --print-fingerprint` 获取；五个 SHA256 分别来自样例 DOCX、显式 Native 与默认模式的 OFD 和各自导出的 PDF。记录只改自身，源码指纹不变；改动其他跟踪文件或改用新版本后必须重做候选验收。PDF/OFD 哈希是本机实际查看的证据；因生成文件可含时间元数据，不要求另一台 CI runner 的重建 ZIP/PDF 与本机逐字节相同。
4. 对该记录提交打 tag。发布工作流先复测、打包、消费原包并核对哈希，再校验记录，最后推 NuGet。失败时保持上个已验证包/标签；修复后使用新提交、新版本重新验收，不覆盖同版本包。公开源安装回读与下游生产环境签收仍需另行记录。

`.NET` 命令遵守根 `AGENTS.md` 的 `DOTNET_CLI_HOME`、首次体验/遥测变量与单节点 flags。当前记录是历史 preview.9 的 `accepted`；本轮未完成新的 Preview 验收，不能把旧记录或本轮绿测写成新候选视觉通过。
