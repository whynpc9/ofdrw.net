# Q-01～Q-05 发布工程闸门

本文件描述可重复执行的检查及其边界。当前 PR 提供闸门和模板，**没有完成任何发布候选的人工 Preview 验收，也没有发布 NuGet**。

| 闸门 | 自动化入口 | 判定与留存 |
| --- | --- | --- |
| Q-01 许可与第三方声明 | `python3 scripts/verify-package-artifacts.py --notices-only`；`scripts/run-converter-package-e2e.sh <version>` | 仓库 MIT 与 nuspec MIT 表达式一致；每个直接 NuGet 依赖的版本、许可、来源与还原的包元数据相符；22 个直接及传递依赖的许可元数据与受审查清单一致；11 个同批包内声明与仓库一致，消费前后包 SHA256 不变。 |
| Q-02 视觉语料 | 同批包消费 E2E 中的 `ValidateVisualCorpusAsync` | 票据、发票、模板样式、有效异常布局四个合成无隐私单页 OFD，均经 OFD→PDF→144 DPI PNG，与固定 golden 的全页及标题/正文区域分别比较 RMSE ≤ 0.035；保留实际 OFD/PDF/PNG、差图和 CSV 指标。此门不代替 Preview。 |
| Q-03 坏包与结构 | `dotnet test tests/Ofdrw.Net.Packaging.Tests/...` | 固定坏包覆盖 `..`/反斜杠/绝对路径、重复条目、条目数、总展开量、压缩比、无效 ZIP；恰好等于预算的有效对照必须接受。低预算模拟炸弹，不创建巨大载荷。 |
| Q-04 XML 文档 | `python3 scripts/check-public-api-docs.py` | 开启 CS1591，按源码路径和公开成员符号对比历史清单。新增缺文档成员使 CI 与 tag workflow 失败；旧债减少可直接通过。 |
| Q-05 Preview | `docs/preview-acceptance-template.md`；tag workflow 的 `scripts/verify-preview-record.py` | 候选需有人工逐页记录，包含本次 native OFD→PDF→macOS Preview、哈希、检查人、页码、缺陷。仓库记录默认 `not-reviewed`，因此未验收的 tag 会在推包前失败。脚本仅校验证据字段和源码绑定，不能证明人实际看过页面。 |

## Q-04 历史 CS1591 清理顺序

当前历史清单在 `docs/public-api-cs1591-baseline.json`，按成员符号而非行号记录，包含两个 SDK 目标框架去重后的 273 项。依次清理：

1. `Converter.Abstractions` 与 `Layout` 的公共入口；
2. `Converter.Docx`、`Converter.Pdf`、`Converter.Svg` 的选项和结果；
3. `Reader` 与 `Packaging` 的读取、资源预算和错误契约；
4. `Core` 模型和 `Signatures`；
5. CLI 中可见的公共成员。

历史成员补完文档后**不要重新生成基线**；检查器会报告减少的旧债。仅在审查确实需要调整基线时运行 `python3 scripts/check-public-api-docs.py --update-baseline` 并审查差异。

## Q-01 依赖许可复核

`docs/third-party-dependency-baseline.json` 固定本轮还原图中 22 个直接/传递包的版本、nuspec 许可类型与值、项目 URL，以及以文件声明许可的内容哈希。旧版仅提供 `licenseUrl` 的条目保留原 URL，不把它自动推断为 SPDX 结论。依赖图变化会阻断检查；审核新增包的真实许可和分发影响后，才运行 `python3 scripts/verify-package-artifacts.py --update-dependency-baseline` 并审查清单差异。CI 安装的字体与外部 LibreOffice/Word 不在 NuGet 发布载荷内，其授权与来源按 `THIRD-PARTY-NOTICES.md` 分别说明。

## 发布候选顺序

1. 从固定提交运行 Python 检查、.NET 全套测试和同批包消费 E2E。保存包 manifest、OFD、PDF、PNG、日志。
2. 使用 macOS Preview 打开**本次** `generated-docx-native.pdf`，逐页核对源 DOCX；受默认模式影响时也打开 `generated-docx-default.pdf`。检查后填写 `artifacts/<candidate>/preview-acceptance.md`，记录问题和体积变化。直接 `generated-docx.pdf` 不属于 Native 链路。
3. 只有无遗留问题且两页均已检查，才把 `docs/release-preview-acceptance.json` 更新为 `accepted`，并填写候选 `package_version`。`source_fingerprint` 用 `python3 scripts/verify-preview-record.py --print-fingerprint` 获取；三个 SHA256 分别来自样例 DOCX、本次 Native OFD 和由它导出的 PDF。记录只改自身，源码指纹不变；改动其他跟踪文件或改用新版本后必须重做候选验收。PDF/OFD 哈希是本机实际查看的证据；因生成文件可含时间元数据，不要求另一台 CI runner 的重建 ZIP/PDF 与本机逐字节相同。
4. 对该记录提交打 tag。发布工作流先复测、打包、消费原包并核对哈希，再校验记录，最后推 NuGet。失败时保持上个已验证包/标签；修复后使用新提交、新版本重新验收，不覆盖同版本包。公开源安装回读与下游生产环境签收仍需另行记录。

`.NET` 命令遵守根 `AGENTS.md` 的 `DOTNET_CLI_HOME`、首次体验/遥测变量与单节点 flags。当前记录是 `not-reviewed`，所以不应把本票的绿测写成 Preview 完成。
