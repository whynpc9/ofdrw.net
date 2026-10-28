# R02 CGTransform 跨 TextCode 映射问题与验证现状

记录时间：2026-10-28（Asia/Shanghai）。代码基线为 `28473532aa4bc03c334556f2336081feb6df8c2d`，即 PR #66 的合并提交；其 Git tree `b4326765ec99a6710e095907f4a69bfb07124c18` 与已验收的 PR 头 `5f170b80e796a616cdbb489e0f96a32783b7a22a` 相同。本文汇总既有验收和本轮只读源码审计；文中的测试与视觉结果属于各自注明的候选版本，本轮没有重新构建或渲染。

**当前判断：**R02 原定 P1 范围内的显式 GID 绘制、Unicode 文本保留、指定的 1:N 与 M:1 示例、定位与缺字拒绝，已在受限输入范围内完成并通过对应验证。尚未解决的是**不等基数 `CGTransform` 涉及多个兄弟 `TextCode` 时，`CodePosition` 归属哪个文本段、如何计数，以及映射能否跨段**。标准现有条文和手头样本未提供足以决定这三个问题的输入及独立预期；不能由已完成的单段案例推定全体 OFD 的互操作性。

## 为什么需要同时保存字形与文本

OFD `TextObject` 中的 `TextCode` 保存源文字和定位信息；`CGTransform` 以 `CodePosition`、`CodeCount`、`GlyphCount`、`Glyphs` 指定源字符范围与原始字体 GID 的对应关系。一个源字符可以对应多个 GID，多个源字符也可以对应一个 GID。正确导出需要两条相互独立的结果：页面绘制原始 GID 的轮廓，文本提取或复制仍得到源 Unicode，而不是用字体 `cmap` 重新挑选可见字形，也不能添加第二份可见文字。

原始 R02 P1 要求包含一对多、多对一、跨 `TextCode` 偏移、`DeltaX/Y`、缺失字形，以及固定字体的视觉和文本双检；这是 OFD 读取与呈现范围。生成端的 typed/shaped glyph run 属于 R22，不能因为同样涉及 GID 就算作 R02 的已交付能力。原始范围记录于本地忽略目录 `artifacts/roadmap-v12-20261021/research-v1.2.txt` 第 23–29、84–91 行；该文件不是仓库交付物。

## 标准给出的边界

所核对的 [GB/T 33190—2016 原文](https://tsgdag.jhun.edu.cn/_upload/article/files/71/fc/4ed26d5841caab5bc999e19f8f14/02a71fcf-7b6e-4661-9991-eb309c7614c6.pdf) SHA-256 为 `d0232b4d75340b885a5e068df477ad2b9f77e4fd23ee3c7e6ddc51b147b27ae8`。表 48（印刷第 68 页、PDF 第 73 页）称 `CodePosition` 为 `TextCode` 中字符编码的从零开始位置，图 67–70 说明四种基数关系；但在一个 `TextObject` 含多个兄弟 `TextCode` 时，条文未明确给某个 `CGTransform` 指定目标兄弟，也未明确补充平面字符的索引单位。表 46（印刷第 65 页、PDF 第 70 页）说明没有某轴位移量时，该绘制点在该轴不移动。不能把字体自然字宽当成该条文明确规定的默认位移。

例如，第二个 `TextCode` 前出现 `CodePosition="0"` 的映射，至少要区分“第二段局部位置 0”与“整个 `TextObject` 合并文字的位置 0”；若 `CodeCount` 延伸到下一段，还要确定是否允许跨段。这个例子用于说明歧义，不是一个已经获独立确认的 OFD 样本。

## 已合并实现与分项验证

当前[支持范围说明](r02-cgtransform-profile.md)记录了 PR #57–#66 的递增边界。`OfdExplicitGlyphRun.Read` 仍以多个段拼接后的 BMP 字符位置处理已支持的 1:1 映射；不等基数的定位与组合分支要求一个 `TextCode`。遇到未建模的映射、缺失原始 GID 轮廓、字体嵌入不许可、重叠或不完整位移时，受限输出路径会在写出前拒绝，而不是把文件生成成功视为正确呈现。参见[当前解析入口](../../src/Ofdrw.Net.Core/Models/OfdExplicitGlyphRun.cs)。更详细的原范围对照表保存在本地忽略目录 `artifacts/r02-cgtransform/remaining-core-after66-20261028/remaining-matrix.md`，不随仓库提交。

| 验证维度 | 已有证据及适用范围 | 当前状态 |
| --- | --- | --- |
| 映射基数 | #59 的完整无定位 2–8:1、#60 的完整无定位 1:2、#62 的单段显式定位 1:N/N:M；#65–#66 的单段局部不等基数、普通 cmap 字符混排和至多两个有序映射。 | 指定小规模静态 TrueType `glyf` 输入已验收；不等基数跨段未实现。 |
| 多 `TextCode` | #58 的全局 1:1 映射及 #64 的普通 cmap 混排，保留各段的起点与完整位移。 | 受限全局 1:1 已验收；局部位置重置与跨段不等基数仍无确定契约。 |
| 功能与包消费者 | PR #66 源码头 `5f170b80`：11 个 Release 测试程序集共 4,329 通过、17 跳过、0 失败；候选 `.15` 的 12 个包通过验证与广义 E2E，隔离 `PackageReference` 消费者检查原始 Go 输入的 PDF/SVG/HTML 与读回。 | 对该精确源码头和候选包成立；本轮文档整理未重跑。 |
| 字形和文本 | #66 的 12 条包消费者 PDF 路径与 12 条 SVG 路径核对原始 GID；五条映射路径与独立 Go PDF 的坐标差小于 `0.00001 mm`。原始 `TextCode`、资源 ID、`SourceXml` 与逻辑文本另行核对。 | 指定样例的源字形与文本门禁已通过。 |
| PNG 与 macOS Preview | 父级验收检查了六张 PDF/SVG 白底 PNG；三份单页包消费者 PDF 逐页在 Preview 查看，没有观察到裁切、重叠或重复可见字形。字面复制分别得到 `กำข`、`AflBffiC`、`AflB`，之后关闭了验收窗口。 | 三个指定页面的视觉与字面复制通过；PNG 结果没有代替 Preview。 |
| Preview 查找 | 三页均出现“Not found”行与“Found on 1 page”横幅并存、导航不可用的矛盾反馈。 | 未判定为查找通过，也未据此判定 PDF 文本缺陷；需记录精确查询串和阅读器结果后单独归因。 |
| CI、审阅与合并 | PR #66 精确头的 10 个 CI 作业（含 `CI complete`）及独立 SES 成功；Codex 审阅完成且没有未解决意见。父级按常规路径合并，合并树与受测头一致。 | 该候选的合并门禁已关闭；不代表公开 NuGet 发布或全 R02 完成。 |

上述验收的本地可复查入口依次是 `artifacts/r02-cgtransform/single-textcode-composition-20261028/evidence-manifest.json`、同目录的 `preview/receipt.json`，以及 `artifacts/roadmap-v12-20261021/R02-parent-merge-decision-PR66-20261027T1912Z.json`。这些文件位于已忽略的本地 `artifacts/`，不会随本次文档提交，其他检出环境不能依赖它们存在。证据清单的 `visual.preview` 字段写于桌面验收前；Preview 结论应以之后的记录为准。

## 外部材料实际回答了什么

在十份外来 OFD 中，已统计 1,173 个带映射的 `TextObject`，**没有一个**同时含映射和多个 `TextCode`。继续扫描这同一批文件不能判定跨段规则。`z.ofd` 原件的 SHA-256 为 `ae942157e68402aced7c604d0d17b6421f276cadbab363fa781fea5f948b3596`；它含单段的不等基数映射，但映射引用的 Times New Roman、Tahoma、Latha 只有名称，嵌入的 STZJ 字体属于另一个 ID。用机器字体替换不能验证原件的 GID 外观。`doc1File.ofd` 原件的 SHA-256 为 `92a655b3966aae0b44dc9ca666849f68ee26028e12f537669e58489271f06f95`；其零 Layer ID 在现有身份校验中被拒绝，且嵌入 `glyf` OTF 缺少 PDF 受限配置所需的 OS/2。曾用修改 Layer ID 的本地副本完成局部 SVG 对比，但那不是原件端到端 PDF 验收。这两份材料均未给出多段不等基数的判例。

另一份独立实现 `xiaoqidun/ofdgo` 固定在 [`7e3f2c468aaa56d461f117cc045413c1b457e75c`](https://github.com/xiaoqidun/ofdgo/tree/7e3f2c468aaa56d461f117cc045413c1b457e75c)。其 [XML 解码](https://github.com/xiaoqidun/ofdgo/blob/7e3f2c468aaa56d461f117cc045413c1b457e75c/ofdgo_text.go#L393-L425)把紧邻某个 `TextCode` 之前的 `CGTransform` 解释为该段局部位置，再加上前面各段的 rune 数；[写出器](https://github.com/xiaoqidun/ofdgo/blob/7e3f2c468aaa56d461f117cc045413c1b457e75c/ofdgo_writer.go#L1359-L1387)逆向转换。[绘制循环](https://github.com/xiaoqidun/ofdgo/blob/7e3f2c468aaa56d461f117cc045413c1b457e75c/ofdgo_text.go#L153-L170)只在映射完全落入当前段时消费它，编辑器校验也拒绝跨段映射。其 PDF/SVG 调用路径在源码中可追到显式 GID 或轮廓，但该固定提交的 356 项目录树没有 `_test.go`、`.ofd` 或字体样例；本轮没有运行它、生成 PDF/SVG 或查看页面。因此它给出**可检验的生产者与读取器约定**，尚不是跨实现的实际输出判例，也不能单凭 rune 计数决定补充平面字符的标准索引单位。详细哈希留在本地忽略目录 `artifacts/r02-cgtransform/remaining-core-after66-20261028/independent-ofdgo-source-audit.md`，不随本次文档提交。

先前核对的 `zc310/ofd` v0.1.5 对局部位置重置和位移长度采用另一套启发式解释，不能用“两个 Go 项目都支持”概括为一致结论。ofdrw.net 当前受限的全局 1:1 规则、ofdgo 的 XML 局部绑定规则，以及 zc310 的启发式，分别是实现行为；标准条文本身没有明确裁定三者冲突的输入。

## 解锁跨段实现需要的证据

1. 取得**未经改写**、来源与 SHA-256 可追溯的 OFD，其中同一 `TextObject` 至少有两个 `TextCode`，后一个段存在 `CodeCount != GlyphCount` 的映射；若目标是验证局部重置，最好在后段使用 `CodePosition=0`。保留原始页面 XML、源文字、完整字体字节与可用于验证和嵌入的许可、预期 GID 序列及每个段的绘制原点。
2. 配套取得生产者对 `CGTransform` 与段归属的明确规则，或第二个独立阅读器在该原件上的页面和文本结果。证据必须能区分全局拼接位置与 XML 局部位置。若要支持 `CodeCount` 跨越段边界，还需独立的跨边界样例及预期结果；仅有段内不等映射不能推出它。
3. 在规则确定后，再设计源字符范围到绘制 GID 槽位的逐段投影与拒绝条件，检查原始 GID 路径、Unicode 提取、各段 X/Y 与完整 `DeltaX/Y`，并执行源码回归、精确包消费、PDF/SVG 页面 PNG、macOS Preview 和精确头 CI/审阅。缺失字体或不支持的结构仍应在产生可被误认为成功的输出前拒绝。

补充平面字符的 `CodePosition` 单位、无位移量时多 GID 的真实输出，以及 Preview Find 的矛盾状态各需单独的规范或原件证据。它们不能通过再制造一个单段 `fl`/`ffi` 字符串样例来代替跨段判例。当前没有证据支持新开一个扩大基数的代码 PR；R22 的生成端能力也仍由独立门禁决定。
