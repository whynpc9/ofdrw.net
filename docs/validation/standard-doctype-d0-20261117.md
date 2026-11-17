# 通用默认 OFD：D0 集成验收摘要

日期：2026-11-17。关联 [#145](https://github.com/whynpc9/ofdrw.net/issues/145) 与[后续计划](../engineering/plans/ofd-archival-roadmap.md)。

## 范围与源码身份

远端基线 `6b487e6d35ac1da59881a27e7dd069cdc1c13317`；代码提交 `aa444c371b99d171dff5a49edb14c4e45db1666c` 与 `54e41ddaa80c81355d568c930cd5c4a824b27d3c`。按用户指示在本地直接快进集成 main，连同计划和本报告提交推送，不另建 PR。本次集成没有新增生产代码、测试或 E2E 变更，复用下述精确代码身份的验证；没有重新生成产物或重新执行 Preview。

通用生成、DOCX 默认/Native/DualLayer、PDF 导入默认 DocType=OFD；读取保留显式声明，缺失属性回退为 OFD。保持现有命名空间选择。C# 中 OFD-H 字面量全部清理，显式非默认类型 roundtrip 以 OFD-A 测试；该测试不证明归档符合性。OFD-H 专用 package 和 OFD-A 归档实现均未在本次开发。

这些变化属于 main 开发版本，**没有发布 NuGet 版本**；0.4.0 发布说明与历史行为保持原义。

## 功能测试与包消费

环境：macOS 26.6.2 arm64、.NET SDK 10.0.401、Poppler 26.09.0。

| 轮次 | 实际验证 | 结果与边界 |
| --- | --- | --- |
| 默认类型修改，`aa444c37` | 全解决方案 11 个程序集；写入器历史 fixture 修正后重跑 Packaging | 综合最终 5107 通过、21 跳过、0 失败；并非初次全量运行无失败 |
| 同轮独立门禁 | 可信字体及 12 个 `0.4.1-doctype.1` 候选包隔离消费 | 字体 74/74，包消费通过 |
| C# 清理，`54e41dda` | Release 构建、受影响 Packaging/PDF/CLI 测试 | 277 + 40 + 7 = 324 通过；这一轮未重新跑完整解决方案测试 |
| 同轮包消费 | 12 个 `0.4.1-doctype.2` 候选包、隔离 SDK/CLI/font/text 消费 | PASS；receipt 的 sourceSha=`54e41dda…`、sourceChanges 为空 |
| D0 集成复核 | 读取上述 TRX/receipt，核对产物哈希、源码相同、C# 扫描与存储检查 | 复核既有证据，不冒充重新构建或重新执行测试 |

候选包仅用于本地验证，没有以这两个候选版本发布 NuGet。网站正文在默认值修改轮已通过构建、类型和链接检查，索引 50 页；Nimbus authoring lint 不覆盖 Markdown 正文。本次新增工程计划/验收摘要检查相对链接与差异格式，不将它们作为网站正文发布。

## 自动渲染、PNG 与 Preview

确定性样例为仓库 `generated-layout.docx`，第一轮使用可信字体归档内的 Noto Sans CJK SC；其 SHA-256 为 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`。

- **自动检查**：Native/default/DualLayer 原文保留；Native 选页 2、1、2 的文本和像素顺序正确。Native、DualLayer、PDF 导入各两页与 0.4.0 基线 PNG 字节一致；除时间/ID 外，包差异仅 OFD.xml 的 DocType，大小各减少 2 字节。
- **PNG 视觉复查**：第一轮查看 Native、DualLayer、PDF 导入共 6 页；默认与选页另有像素比较。清理轮默认/Native 共 4 页及通用生成 1 页与前一候选像素一致；历史 OFD-H 两页仍可读，导出像素与基线一致。
- **macOS Preview**：第一轮实际生成 OFD → PDF → Preview，共 4 份 PDF、8 页；DOCX 同时覆盖显式 Native 和默认模式。检查中文、英文、局部样式、表格与分页，没有发现该样例上的缺字、裁切、重叠、黑页或异常空白页。DualLayer 的不同换行和较软栅格外观与基线相同。验收后关闭相关窗口。
- 清理轮及本次集成没有再次打开 Preview。Preview 截图在原会话工具记录中，未另存截图；本地页面 PNG 是辅助证据，不冒称 Preview 截图。

以上只覆盖具名样例、页面和配置，不推断任意 Word 文档完全保真，不证明 OFD-A 标准符合性或独立 OFD 阅读器行为。

## 可追溯证据

现有证据保留在 `codex/standard-doctype` 隔离 worktree 的 ignored artifacts 中。本次未复制整套运行目录、字体或包缓存，也未将这些目录宣称为 GitHub 已上传附件。后续需要移交时按仓库存储规则打包必要输入、输出与页面证据。

| 本地相对路径 | 文件 SHA-256 |
| --- | --- |
| `artifacts/standard-doctype/visual/sha256.txt` | `9f6d84694ff8b92da9ccb2d8f47ea74feccfb50db4f450f6822746cbecbf1656` |
| `artifacts/standard-doctype/visual/preview-observations.md` | `40679207d0f30ffd32c7da4be9aab5940416f855273fdb81b444ab090195d25d` |
| `artifacts/doctype-cleanup/sha256.txt` | `422fe562343ef2e340fa15981de5c639cbbac5b07a9ef231e27cb23adb598126` |
| `artifacts/doctype-cleanup/package-e2e/package-e2e-run.json` | `099d4f23b8c673eaf668c54abfe17b8fd19026cd4cf68cedc4aa6f0084b1a23a` |

D0 集成时复核前后两份 sha256 清单中的 22 与 6 个文件，均与记录一致。受影响测试原始 TRX 在 `artifacts/doctype-cleanup/results/`。精确源码身份由上述提交保存；不改写原 receipt 的 SHA、时间或历史验收结论。
