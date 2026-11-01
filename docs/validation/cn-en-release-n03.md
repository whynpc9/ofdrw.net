# N03：极小正字号输出契约

源码基线：`ba7682cfceafd6e28d526941e0eb039addfe400f`，分支 `codex/cn-en-release-n03`。本项仅处理 [N03](../next-steps-2026-11-01.md#n03极小正字号的序列化契约)，不扩展 R22、字体或坐标精度。

## 决策与兼容性

保留 `0.###`，采用输出前明确拒绝。`OfdNumericFormat.WriterPositive` 检查实际格式化、读回后的数值，而不是用 `size/UPEM > 0` 或硬编码最小字号替代输出契约。

- `OfdPackageWriter` 创建新的 `TextObject` 时要求字号有限、严格大于零，且现有格式写出后仍大于零；否则抛出 `NotSupportedException`。直接页面、后续页面、共享复合对象和批注 appearance 都经过同一 primitive writer。全部 entries 构造成功后才建立目标 ZIP；字号拒绝发生在读取目标流能力/位置或写入任何字节之前。
- `TextPrelayoutOfdWriter.CreateSinglePage` 和 `CreateSinglePageWithPositionedGlyphs` 在创建内存 package 前以 `TextLayoutException` / `UnsupportedOutput` 拒绝不可写出的字号。查询/塑形/`TextPositionedRun` 的正数测量契约保持不变；查询可成功不表示可以输出。
- 公开方法签名、默认字号、正常字号、坐标/Delta 格式和页数策略不变。先前会生成无效 `Size="0"` 的新文本现在明确失败。现有 `SourceXml` 仍保留原始合法小数精度，不因本修复拒绝 `Size="0.0001"`；本项不对任意外来 XML 作新合法性保证。
- 实测 .NET 10 自定义格式会将 `0.0004999999999999999` 写成 `0.001`，因此该值可以输出；`0.0004999999` / `0.0004` 被拒绝，`0.0005` / `0.001` 写成 `0.001`。契约取决于实际 formatter，不承诺数学上精确的 `0.0005` 二分阈值或保留极小字号的原始度量。

## 验证

定向验证结果与产物见工作树的 `artifacts/cn-en-release-n03/`；未提交运行产物、字体副本或完整日志。测试和探针均设置专用 `DOTNET_CLI_HOME`、关闭首次体验/遥测，以 `-m:1 /nodeReuse:false /p:UseSharedCompilation=false --disable-build-servers` 顺序执行。

| 项目 | 结果/范围 |
| --- | --- |
| Packaging `TextSizeSerializationTests` | **18/18 通过，0 跳过**；`test-results/packaging-focused.trx` |
| Layout `TextPositionedConsumerTests`、`TextPrelayoutTests`、`TextPositionedRunTests` | **88/88 通过，0 跳过**；`test-results/layout-focused.trx` |
| 输出/读回探针 | **12 个 OFD 写出/读回成功；4 个 0.0004 mm 请求在 package 创建前拒绝**；`outputs/manifest.json` |
| 自动页面渲染 | 未执行，交接后续验证 |
| PNG 视觉复查 | 未执行，交接后续验证 |
| macOS Preview | 未执行，由协调者安排 Terra 独立验收；本报告不宣称视觉通过 |
| 完整回归与本地包消费 E2E | 本分项未执行；合并候选由协调者运行 |
| 仓库存储门禁 | 暂存后执行 `python3 scripts/check-repository-storage.py --require-materialized e2e --require-materialized tests` 通过；167 个 LFS 指针及测试/样例载荷已校验 |

回归覆盖：正/负/零/NaN/Infinity、`double.Epsilon`、格式边界内外、普通 4 mm 与既有三位小数、XML Size 和原文读回、坐标精度、目标流零触碰、两种 typed writer 入口，以及现有 DejaVuSans (UPEM 2048) / NotoSans (UPEM 1000)。复用既有字体，无新二进制 fixture。

测试环境为 macOS arm64、.NET SDK 10.0.401 / runtime 10.0.12。构建有既有 NuGet 审计与 nullable/analyzer 警告，无编译错误。初次边界测试曾误将 formatter 实际写为 `0.001` 的近阈值 double 列为拒绝，改正断言后纳入接受侧回归；未调整生产阈值。

复现命令（在本工作树根目录设置上文环境变量后执行；第二条以后复用 `NUGET_PACKAGES=/Users/wanghongyi/.nuget/packages`）：

```sh
dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj --filter FullyQualifiedName~TextSizeSerializationTests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj --filter 'FullyQualifiedName~TextPositionedConsumerTests|FullyQualifiedName~TextPrelayoutTests|FullyQualifiedName~TextPositionedRunTests' --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build artifacts/cn-en-release-n03/probe/N03.Probe.csproj --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet artifacts/cn-en-release-n03/probe/bin/Debug/net10.0/N03.Probe.dll "$PWD"
```

探针源码仅保留于本轮 artifacts；产品回归可独立通过已提交的测试重现，不依赖探针。

## 视觉验证交接

输入使用现有 `e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/DejaVuSans-R22.ttf` 和 `NotoSans-Regular.ttf`。探针以 `TextPrelayoutRequest → HarfBuzzTextPrelayout → CreateSinglePage[WithPositionedGlyphs] → OfdPackageWriter → OfdReader` 生成/读回；正常 4 mm 输入为 `AV\nTo`，边界输入为 `A`。成功 OFD 各 1 页；拒绝 0.0004 mm 时不生成 package 或文件。

后续需从本次成功 OFD 导出 PDF，再逐页 PNG / macOS Preview 检查；不得以重新创建的直接 PDF 替代。极小边界验证主要确认正 Size、原文和输出有效性，不声称其肉眼可读。常规文档的样式、表格、分页和 DOCX native/default 应在协调者的冻结核心样例矩阵补验，不能从本项短 ASCII 输入推断通过。

本项输出目录为 `artifacts/cn-en-release-n03/outputs/`，精确输入字体哈希、输出字节数/SHA-256、Size 字面值和原文读回由 `manifest.json` 记录。后续 PDF、PNG、Preview 路径和实际页范围由视觉验收者追加自己的记录。
