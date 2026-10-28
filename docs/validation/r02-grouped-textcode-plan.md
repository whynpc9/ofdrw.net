# R02 基于 XML 分组的多 TextCode 段内映射调整方案

日期：2026-10-28。方案基于文档提交 `2264914a84175a02d064ef5baa75bfb7a75cf4e1` 和其中的产品基线 `28473532aa4bc03c334556f2336081feb6df8c2d`，承接[问题与既有验证记录](r02-cross-textcode-context.md)。本文记录后续调整方向；策略名称与数据结构是设计草案，尚未实现。当前受限实现的既有测试和视觉结果不覆盖本方案。

下一轮目标是：**保留 XML 分组归属，使多个 `TextCode` 分别承载已有受限基数的段内映射，并让 PDF、SVG、HTML 使用一致的字形和文本投影。**优先核实并防止映射作用到错误文本段。真正越过段边界的映射、补充平面计数、缺省复杂定位和阅读器查找行为分别跟踪；R22 的生成端能力由独立工作线推进；本次已合入包含 PR #68 的 main `c321903e`，详见 [`text-shaped-output.md`](../text-shaped-output.md)。

## 规范依据与需要修正的判断

[GB/T 33190—2016](https://tsgdag.jhun.edu.cn/_upload/article/files/71/fc/4ed26d5841caab5bc999e19f8f14/02a71fcf-7b6e-4661-9991-eb309c7614c6.pdf)的**规范性附录 A.13 `Page.xsd`**在印刷第 119–120 页（PDF 第 124–125 页）定义 `CT_Text`。其中 `xs:sequence maxOccurs="unbounded"` 内先是允许零到多个的 `CGTransform`，随后是一个 `TextCode`，形成 `(CGTransform*, TextCode)+`。本轮重新提取该两页文本并查看 PDF 第 125 页的页面图，确认了这一嵌套结构。所用 PDF SHA-256 为 `d0232b4d75340b885a5e068df477ad2b9f77e4fd23ee3c7e6ddc51b147b27ae8`。

表 48 将 `CodePosition` 定义为 `TextCode` 内从零开始的字符编码位置。**将这一属性定义与重复分组结构结合，可为“本组映射作用于随后那个文本段，位置按该段局部计算”建立很强的规范依据。**这是对规范结构和属性的联合解释；Schema 的顺序约束本身不等于各阅读器已按该解释通过验收。

因此，原记录中“归属完全没有依据、须先取得外来跨段样本才能推进”的判断应收窄。可以先按规范结构开展分组模型和定向算法验证，外部互操作另设实物证据门禁。

| 问题 | 本轮依据与决策 |
| --- | --- |
| 映射归属哪个 `TextCode` | 按重复分组绑定本组随后出现的文本段；以段内位置建立新解析策略。 |
| 补充平面字符的索引单位 | 分组结构没有解决。第一轮继续拒绝非 BMP，不把 Go rune、C# `string.Length` 或字素簇数量直接当作标准。 |
| 映射是否可以越出本段 | 分组结构不构成支持依据。新策略要求源范围完整落在所属段内。 |
| 既有全局规则的兼容性 | 单独保留和回归；当前全局规则不应继续被描述为规范分组语义。 |

## 需要先确认的错绑风险

当前 [`OfdExplicitGlyphRun.Read`](https://github.com/whynpc9/ofdrw.net/blob/28473532aa4bc03c334556f2336081feb6df8c2d/src/Ofdrw.Net.Core/Models/OfdExplicitGlyphRun.cs#L74-L107)分别收集映射和文本段，再拼接文字解释位置；不等基数分支要求 `codes.Length == 1`，[`PathGlyphOrigins`](https://github.com/whynpc9/ofdrw.net/blob/28473532aa4bc03c334556f2336081feb6df8c2d/src/Ofdrw.Net.Core/Models/OfdExplicitGlyphRun.cs#L50-L69)也要求一个 run。当前投影没有保留映射与 XML 分组的对应关系。

下面是分析用片段，字体、对象属性和 GID 需要在后续测试中补齐，不是已经运行的回归输入：

```xml
<TextCode X="0" Y="6" DeltaX="8">AB</TextCode>
<CGTransform CodePosition="0" CodeCount="1" GlyphCount="1">
  <Glyphs>40</Glyphs>
</CGTransform>
<TextCode X="0" Y="18" DeltaX="8">CD</TextCode>
```

分组解释将映射作用于第二段的 `C`，当前全局位置解释会把位置 0 指向 `A`。从源码静态检查可知，现有范围校验不会仅因映射出现在第二段前而拒绝；若字体、GID、语义冲突及位移校验也满足，存在成功输出但错绑字符的风险。**实施时已用 `LegacyGlobalMappingIgnoresSecondTextCodeGroupPlacement` 复现：旧策略得到 GID `[40,37,38,39]`，映射源为 `A`；分组策略得到 `[36,37,40,39]`，映射源为 `C`。两种策略均能通过基本范围检查，不能自动择一。**

既有 [`MixedOneToOneText`](https://github.com/whynpc9/ofdrw.net/blob/28473532aa4bc03c334556f2336081feb6df8c2d/tests/Ofdrw.Net.Converter.Pdf.Tests/ExplicitGlyphTransformTests.cs#L97-L123)把全部映射放在全部文本段之前，验证的是既有全局规则。保留这些输入、预期与历史验收身份，另增交错分组的定向样例。仅删除单段限制，或对全部映射统一加累计偏移，都无法确定正确归属。

## 明确策略并保留兼容行为

建议在内部区分下列策略；名称暂定，公开配置入口、诊断形式和默认迁移时间在兼容性阶段决定。

| 策略草案 | 解释与范围 |
| --- | --- |
| `GroupedLocal` | 按 XML 子节点顺序保存分组，将待绑定映射归入随后 `TextCode`，按段内源字符位置解释。拒绝越段和尾部悬空映射。 |
| `LegacyConcatenatedGlobal` | 保留当前全局拼接规则及原有受限支持范围，作为明确的兼容路径。 |

当前小版本先明确选择新策略，不静默改变旧调用输出。旧策略可能错绑分组输入的风险必须随兼容说明公开，并由第一阶段的复现决定是否需要额外拒绝或诊断；保留兼容行为不构成正确性认证。

第一版不采用“局部失败就试全局”或“哪个能输出就选哪个”的自动回退。同一输入在两种解释下可能均满足范围检查，却对应不同的源字符或字形位置。可增加双策略分析诊断：比较所属段、源范围和绘制计划；若有差异，记录对象 ID、映射 XML 序号、所属段、局部范围与采用策略。诊断不隐式选择另一种输出。

## 统一投影与出口调整

处理流程为：保留源 XML，按子节点顺序读取分组，按明确策略解析映射归属，逐段构建源字符到字形槽位的投影，最后让 PDF、SVG、HTML 消费同一绘制计划。所有变换都发生在临时模型中，不拆分原始 `TextObject`，不改写 `SourceXml`、对象 ID、原始字体或原始文本来适配输出。

在 `GroupedLocal` 下，遇到 `CGTransform` 时加入待绑定集合，遇到 `TextCode` 时将集合绑定到该段、校验段内范围并清空；对象结束仍有待绑定项则拒绝。还需遵守已有对象结构校验，不能把未知或不合法子节点当成可忽略的分组分隔符。

在现有 `OfdExplicitGlyphSpan` 上逐步扩展，至少明确以下字段含义，避免一个全局下标同时表示字符和字形：

```text
TextObjectPlan
  Strategy
  Groups[]
    TextCodeIndex
    DecodedSourceText
    EffectiveOriginX / EffectiveOriginY
    SourceSpans[]
      LocalCodeStart / CodeCount
      LocalGlyphStart / GlyphCount
      LogicalText
    GlyphSlots[]
      OriginalGid
      X / Y
      SourceSpanIndex
```

源字符偏移按源字符数累计，输出字形偏移按字形槽位数累计。对不重叠映射、未映射字符按一对一处理的已准入输入，每段应满足 `输出字形数 = 源字符数 − Σ CodeCount + Σ GlyphCount`。这只是投影计数；显式映射仍使用原文件指定的 GID，普通字符才按既有规则查询同一嵌入字体的直接 cmap，不引入重新 shaping 或替代字体。

`PathGlyphOrigins` 应改为对指定分组计算原点：使用经验证的该段有效 X/Y，并在该段内部按输出槽位累计位移。不能用上一段的末字形位置代替下一段起点。新开放的多段不等基数范围先要求完整的字形槽位定位，沿用已验收的 `DeltaX/Y` 边界；省略 X/Y 的继承与省略轴向位移是不同规则，不能统一换成字体自然 advance。保留现有有限性、预算、取消及输出前预检约束。

PDF 的 [`DocumentFontContext`](../../src/Ofdrw.Net.Converter.Pdf/Internal/DocumentFontContext.cs)、[`OfdToPdfConverter`](../../src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs)以及 [`OfdToSvgConverter`](../../src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs)目前依赖对象级 `PathGlyphSourceSpan`、`SingleGlyphSourceSpan` 和 run 索引。Core 模型接纳新输入时，必须同时覆盖这些消费者及 HTML 预览路径。可见输出使用每段原始 GID 与坐标，逻辑文本在各出口恰好保留一次；沿用现有字形与语义分离路线，不给每个输出字形重复附加整段 Unicode，也不新增重复可见文字。

## 定向样例与验收矩阵

算法验证使用清楚标注的人工样例，外部互操作使用独立生产者或阅读器的实际结果。十份外来文件中 1,173 个映射对象均不含多 `TextCode` 的事实，限制的是该语料对互操作的证明能力，不阻止按附录结构设计下面的定向测试。

| 样例 | 必须区分的行为 |
| --- | --- |
| 第一段普通文字，只有第二段前有局部位置 0 的映射 | 无重复位置时也要绑定第二段，不能靠位置重置启发式换段。 |
| 两段各含局部位置 0，采用不同 GID 与原点 | 两个映射独立，不误判为同一源范围重叠。 |
| 第一段 2:1，第二段 1:2，均在各自段内 | 源字符累计与字形槽位累计互不混用。 |
| 第二段局部不等映射前后夹普通字符，给完整位移 | 普通与显式 GID 的槽位顺序、坐标、逻辑文本均正确。 |
| 源范围超出本段但仍落在拼接后的对象范围 | 分组模式在输出前拒绝越段。 |
| 局部和全局解释均通过范围检查但计划不同 | 明确策略选择或冲突诊断，不以输出成功自动选取语义。 |
| 尾部悬空映射、重叠范围、GID 数量不符、缺失轮廓、不完整位移 | 在向调用者提交输出前失败，保持原始数据。 |

第一轮继续拒绝非 BMP。后续可用 `A𠀀fl` 等定向文字分辨 Unicode 标量与 UTF-16 单元计数，但必须将转义解码、索引单位和字素簇计数分别验证。

| 验收层面 | 预期证据 |
| --- | --- |
| 结构与投影 | 每个映射的段归属、局部源范围、GID 序列、对象空间原点及逻辑 Unicode；旧全局 profile 回归保持原始身份。 |
| PDF 与 SVG | 分别核对原始 GID 轮廓和坐标；PDF 的文本提取无重复，SVG 的语义标签正确；HTML 实际使用同一计划。 |
| 源码与精确包 | 受影响测试、共享转换回归、隔离包消费者及实际载入 DLL 身份；不能用项目引用替代包消费。 |
| 页面与阅读器 | 本次生成页面逐页 PNG，macOS Preview 和 Windows 原生阅读器分别记录视觉、复制和查找；未完成项明确保留。WSL 的 Linux 命令行结果不替代 Windows 原生查看。 |
| CI 与审阅 | 以实施候选的精确头为准完成适用检查；旧 PR #66 的回执仅证明旧范围。 |

可在明确执行范围后，从已固定的 [`xiaoqidun/ofdgo` 提交](https://github.com/xiaoqidun/ofdgo/tree/7e3f2c468aaa56d461f117cc045413c1b457e75c)建立小型生产者验证程序，固定字体、参数、版本，生成第二段含局部不等映射的 OFD，再把同一未经改写的产物交给另一个独立阅读器和 ofdrw.net。此类输入应标注为“定向生成的互操作样例”，与自然收集业务原件区分。该项目自己的写出器和读入器只能提供内部一致性证据；源码中的[局部位置归一化](https://github.com/xiaoqidun/ofdgo/blob/7e3f2c468aaa56d461f117cc045413c1b457e75c/ofdgo_text.go#L393-L425)可供实现对照，不能替代页面验收。

## 四个实施阶段

| 阶段 | 交付范围 | 合入条件 |
| --- | --- | --- |
| PR A 语义依据与防错绑测试 | 补入附录 A.13，明确分组解释与历史全局行为，构造第二段局部位置 0 的最小输入。 | 运行确认当前行为并记录风险；旧行为的特征测试与新策略的目标预期清楚区分，不用修改旧输入或永久跳过断言制造通过。此阶段不宣称扩大支持。 |
| PR B 分组解析与投影模型 | 引入分组、局部源范围、字形槽位和明确策略，保留旧兼容范围。 | 原有全局 profile 回归不变；新计划和拒绝规则有确定性测试，尚未接入的出口不能接纳新范围。 |
| PR C 多段段内映射输出 | PDF/SVG/HTML 接入分组计划，只组合已支持的受限段内基数。 | 原始 GID、坐标、文本、输出前拒绝、精确包消费及计划中的跨平台页面门禁通过。 |
| PR D 互操作与兼容策略收口 | 固化独立输出、兼容矩阵、策略入口及默认迁移决策。 | 关键证据可复取，默认行为和未支持范围明确；不将局部通过写成完整 R02 或 R22 解锁。 |

这些是建议的 PR 边界，实际实现需避免合入会让新模型提前进入未适配出口的中间状态。下一步先完成 PR A 的规范记录与错绑复现，再据结果设计策略入口和模型改造。

## 证据保存与本轮状态

规范 URL、页码、哈希、源码基线、输入来源、运行命令、预期映射计划与验收结论应保存在可检出的简短记录中。确定性字体和输入复用现有 fixture，按 `.gitattributes` 走 LFS；大体积页面、工具输出和必要视觉记录放在可复取且有 SHA-256 清单的制品中，不能让关键结论长期仅依赖单机忽略目录。

方案记录之后，用户于 2026-10-28 授权实施。A/B/C 的语义依据、分组投影和三个消费者作为一个候选一起验证；D 的互操作证据及默认策略决策另列验收记录。使用方式与有界支持范围见[分组策略说明](r02-grouped-textcode.md)。先前暂停的 PR #67 审阅处理未恢复。

按本方案完成的分组策略候选与阶段结果见[实施说明](r02-grouped-textcode.md)及[验收报告](r02-grouped-textcode-validation.md)。产品源码已包含 PR #69 的完整组合字符源段规则；实现、精确包、独立输出和视觉证据分别列出，不转用历史 PR 的验收结论。
