# N02 中英文页末 Find 有界判别（2026-11-01）

**结论：OPEN，当前核心路径也存在真实 Preview 页末搜索失败；根因 INCONCLUSIVE，未证明局部修复。** 本项完成有界判别，不代表 Find 已修复或候选可发布。不得归入“仅 PR78 未采用路线”的非阻断范围。保留完整日期/普通中文搜索承诺，不以可复制、API 命中、部分前缀或侧栏页数豁免。

## 基线与证据复用

- 专用分支 `codex/cn-en-release-n02` 起点 `ba7682cfceafd6e28d526941e0eb039addfe400f`；调查时 `origin/main` 和 `ls-remote` 均为 `beb02832fff01e61538d90039f3e1224ed03d702`。
- 对 `src/`、`tests/`、`e2e/`、`scripts/`、`Directory.Build.props`、`Directory.Packages.props`、`global.json` 比较，分支起点、当前 main 与旧 main `b02162bef61abc17761349d891237fc2518d3989` **无差异**。main 漂移为 CI/文档。本次复用已封存的准确 b02162be 官方包消费产物并重新校验哈希；不是本次重新转换、重新打包或最终集成候选验收。
- 确定性输入 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`：SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`。
- 公开默认模式与显式 `native` 均走 OpenXML → native OFD；查看链路为 DOCX → native OFD → PDF → Preview。未使用直接 DOCX→PDF 冒充。

原 main 目录：

`/Users/wanghongyi/.codex/worktrees/r22-span-context-query/ofdrw.net/artifacts/r22-shared-position-consumer/main-b02162be/`

| 文件（上述目录内） | SHA-256 |
| --- | --- |
| `package-output/generated-docx-native.pdf`（2 页） | `5ec07c51aac0fc5c5152d4594681c1e3c37a79c14b8edb45cd86f49175d78051` |
| `package-output/generated-docx-default.pdf`（2 页） | `94c6a1ffce28ec016b59afcf589216fa6e487191b1c796aa6889f24c868fdadb` |
| `package-output/generated-docx-native.ofd` | `fe4bc425ad1cfb049bb5f8da24bb0a7f94a5b79acbb2563fb1d171af5b5db452` |
| `package-output/generated-docx-default.ofd` | `3dd6396152fb2efd3bc2b45f555a995ff1be7ddc251b1a1fce6b15bd83821880` |
| `seven-page-preview-input-manifest.json` | `b521015d7f7bfb68ea4e6a9efaa26035fd0421b93a11817fe544c27bac5f8539` |

原 PR78 根目录：

`/Users/wanghongyi/.codex/worktrees/9c7f/ofdrw.net/artifacts/r22-cjk-filled-stroke/`

复核 `date-find-end-boundary-2c4-v1/consolidated-forensic-report.json`（SHA-256 `6794a8361c7bc6006064903dfeafdfcc40962cac206e566c206ff95923a0b731`），保留 45 PDF/59 页、27 Copy、76 Find 中 66 通过/10 正例失败的历史结论：六份完整页末日期、native/default `量A` 与 `目`。本次没有重验全部旧 10 项，不能关闭它们。其 PR78 native PDF 哈希为 `223976fc6023d249dd7895f67bc38e76ccf4f1bc094e77e2e178b0e124f5162e`。

## 两个有界假设

1. **失败仅由未合并 PR78 绘制路线引起**：当前源码等价的 main 产物在实际 Preview 也失败，故该“仅 PR78”解释被反驳；不据此定位具体共享代码或阅读器内部缺陷。
2. **当前产物简单丢失页末原文或无法建立完整搜索位置**：native/default OFD 的第一页末尾都是 `待复核`，第二页末尾都是 `测试日期：2026-01-01`；PDFKit 完整日期与中文均可命中，位置面积为正，故简单文字丢失解释被反驳。API 与实际 UI 明显不同，API 不替代 UI。

最多使用上述 2 个假设，主动分析约 10 分钟，GUI 由专属 Terra 独立执行、等待时间单列，不用满 4 小时。旧不可变证据已覆盖并否定多种单因素解释（逐字 Tj/TJ、最后发出的 carrier、搜索模式、TL/T* 与单段 LF ActualText），本轮不重复这些尝试，也不以剩余额度盲试 PDF 技巧。未发现可证明因果的局部源代码修改，因此只交付判别与阻断记录，不改字体/GID/度量/原文，不引入哨兵、重复文字或 PDF 后处理。

## 功能、渲染与真实 UI 分别判定

本次本地记录目录（不随源码分发）：

`/Users/wanghongyi/.codex/worktrees/c52e/ofdrw-release-n02/artifacts/cn-en-release-n02/`

- **功能/内容检查 PASS（有界）**：读取原始 OFD XML；Poppler 原文核对；macOS 26.6.2 build 25G83 / PDFKit 1451.5.3 对 main native/default 与旧 PR78 native 共 3 PDF 执行 7 查询，均符合源预期并有非空命中位置。`pdfkit.json` SHA-256 `d6b6cc5a4da1466e5d94759a9cbdd367222afb7341f8ebd106ed94d0803aad37`。未运行 .NET 构建/测试，避免与其它专用任务竞争构建；没有产品改动需声称新回归通过。
- **自动渲染完成**：对 main native/default 全部 4 页以 Poppler 120 DPI 渲染。两种模式相应页面 PNG 哈希相同：p1 `0b452192799f06e7f74a35262679fbd1d43bb38f78f0f45e1347fd41e39a611f`；p2 `b549db27597b88b637b9b9c8e2ee6cc72291f47dc8525dfd3569eb21df7cbf2c`。
- **PNG 视觉复查完成（2 个独立页面/4 张等同图）**：目视确认中文/英文正文、表格与日期在页面上实际可见，未见本范围新的裁切/重叠/乱码；旧小表头“量”字可读性 FAIL 不由本次低分辨率页图关闭。此项不等于整份文档视觉接受。
- **真实 Preview Find FAIL（native/default 共 4 页）**：专属 `gpt-6.1-terra` / High 对已重新校验哈希的两份 main PDF 实测如下。N02 开发者没有操作 GUI，也没有把 API 结果当作 Preview 成功。收据 `/Users/wanghongyi/.codex/worktrees/c52e/ofdrw.net/artifacts/cn-en-release/local-validation/n02-baseline/preview-find.json` 已读并重新校验 SHA-256 `5d1d2bae2eacb64e058c7ec64c62726b0cfa03727b441b4bb9493e1976e1050b`。两份窗口已关闭、打开面板已撤销。**本轮 Copy 未重验**；原始截图与 AX 观察在 GUI 工具调用历史中，未另外保存截图文件，不冒称独立图片证据归档已完成。

| 查询 | 源预期 | Preview native/default 实测 |
| --- | --- | --- |
| `Alpha` | p1，1 处 | 可用高亮 |
| `测试日期` | p2，1 处 | 可用高亮 |
| `2026` | p2，1 处 | 可用高亮 |
| `2026-01-01` | p2，完整 10 ASCII，1 处 | 侧栏 Found on 1 page；结果 Not found、导航禁用、无实际高亮 |
| `待复核` | p1，3 个汉字，1 处 | 同上失败 |
| `01` | p2，2 处 | 两模式均只报告/高亮内部一处；native 下一项不能到终末一处 |
| `2026-01-02` | 0 处（负控） | No Results Found |

## 复查与退出

现有产物复查命令（不重新生成 PDF，不访问 GUI）：

```sh
swift artifacts/cn-en-release-n02/pdfkit-probe.swift artifacts/cn-en-release-n02/manifest.json artifacts/cn-en-release-n02/pdfkit-rerun.json
pdftotext -layout /absolute/path/generated-docx-native.pdf -
pdftoppm -png -r 120 /absolute/path/generated-docx-native.pdf artifacts/cn-en-release-n02/main-native
```

上述 Swift 与输入清单保留于本次 artifacts；清单包含绝对路径、大小、SHA-256、源码等价检查与旧证据指针。未来最终候选须基于准确 main/冻结字体重新生成 native/default，重复两个完整页末正例、内部控制和负控，在真实 Preview 记录实际高亮、可用导航与原始复制文本并关闭所开窗口。不能沿用本次旧产品等价产物作为最终集成产物验收。

退出决定：停止主动定位；核心页末 Find 仍为 **OPEN/发布阻断**，根因 **INCONCLUSIVE**。只关闭本次有界排查工作量，保留现有拒绝/回归和完整搜索门禁，不自动重开字体/渲染研究。
