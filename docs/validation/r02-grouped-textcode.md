# R02：按 XML 分组解释 CGTransform

`OfdGlyphMappingStrategy.GroupedLocal` 是显式选择的输入解释策略。每批 `CGTransform` 绑定 XML 顺序中随后出现的 `TextCode`，`CodePosition` 在该段解码后的 BMP 文本内计数。旧调用继续采用 `LegacyConcatenatedGlobal`，将位置解释为拼接文本的索引；本候选不迁移默认值，也不提供失败后自动换策略。

规范依据、历史问题和实施计划见[调整方案](r02-grouped-textcode-plan.md)。依据是 GB/T 33190—2016 规范性附录 A.13 中 `CT_Text` 的 `(CGTransform*, TextCode)+` 结构与表 48 的位置定义。分组解释不能推出补充平面索引单位，也不能推出一个映射允许跨出所属段。

## 调用方式

```csharp
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg;
using Ofdrw.Net.Converter.Svg.Converters;

var pdf = new OfdToPdfConverter(new OfdToPdfOptions
{
    GlyphMappingStrategy = OfdGlyphMappingStrategy.GroupedLocal
});
var svg = new OfdToSvgConverter(OfdGlyphMappingStrategy.GroupedLocal);
var html = new OfdToHtmlConverter(new OfdHtmlOptions
{
    GlyphMappingStrategy = OfdGlyphMappingStrategy.GroupedLocal
});
```

SVG 同时需要输入限额、字体解析预算或 name-only 绑定时，可使用四参数构造函数 `(loadOptions, maximumCmapParseWork, nameOnlyBindings, glyphMappingStrategy)`。映射字体仍必须来自原 OFD 的嵌入资源。三个转换器在构造时校验策略并保存选择；未知枚举值抛出 `ArgumentOutOfRangeException`。PDF 中嵌套签章文档沿用所选策略。

## 投影和支持边界

读取过程先校验原 XML 与渲染视图一致，再保留每组的 TextCode 序号、原映射序号、源字符偏移、输出字形偏移和段内 source spans。各组复用既有段内基数校验，不修改 `SourceXml`、原对象、ID 或字体。原始坐标的省略轴由既有 `OfdTextRenderProjection` 按前段起点解析，位移在组内累积。

每组可以组合[历史受限段内规则](r02-cgtransform-profile.md)已经支持的 1:1、2–8:1、1:2、完整定位不等基数以及受限混合映射。仍有以下限制：

- 一个对象最多 1,024 个源 BMP 字符和 1,024 个输出字形槽位。仍拒绝非 BMP、空白、控制/格式字符及已有限制之外的组合标记。
- 每个映射完整落入所属段；尾部悬空映射、重叠或越界、字形数量错误、不可用轮廓均失败。
- 多段不等基数组产生多个字形时，必须提供覆盖输出槽位全部间隔的 `DeltaX`；提供 `DeltaY` 时也须完整。只有一个输出字形不需要间隔。普通未映射字符使用同一嵌入字体的直接 cmap；多个普通字形需要已有规则要求的完整位移。
- 静态单色 TrueType `glyf`、原 GID、字体嵌入许可和轮廓点数上限继续生效。不重新塑形，不替换字体，不增加 CFF、变量字体或合成粗斜体能力。
- PDF 分组路线统一使用原 GID 轮廓与逐组一次的逻辑文本。语义字体需完整覆盖源字符，并通过已有语义 GID 唯一性检查；没有 Unicode cmap 的字体不在此分组 PDF 路线内。SVG/HTML 使用同一组计划生成路径并逐组附加可访问文本。

投影诊断中的 TextCode 和 CGTransform 序号从 0 开始。分组范围错误包含对象 ID、所属段、映射序号及策略。计划不提供两策略的自动冲突分析 API；测试显式运行两者，确认“均可接受但映射不同”的输入保持可区分。

## 定向用例

使用已有 OFL `NotoSans-Regular.ttf`，字号 6 mm，对象原点 `(10,10)`，两段起点分别为 `(3,6)`、`(3,22)`。以下坐标均为对象空间：

| 用例 | 源分组及映射 | 期望原 GID | 期望字形原点 |
|---|---|---|---|
| second-only | `AB`；`CD` 的位置 0 → 40 | 36,37 / 40,39 | (3,6),(11,6) / (3,22),(11,22) |
| both-local-zero | `AB` 的位置 0 → 40；`CD` 的位置 0 → 41 | 40,37 / 41,39 | 同上 |
| unequal-both | `fl` → 1653；`A` → 36,37 | 1653 / 36,37 | (3,6) / (3,22),(12,22) |
| mixed-second | `AB`；`CflD` 的位置 1、长度 2 → 1653 | 36,37 / 38,1653,39 | (3,6),(11,6) / (3,22),(12,22),(22,22) |

这些是定向设计输入。测试分别核对原 XML 保留、分组归属、字符/字形偏移、GID、字形原点、SVG/HTML 路径、PDF 文本提取及各出口写出前拒绝。测试文件：[`GroupedGlyphTransformTests.cs`](../../tests/Ofdrw.Net.Converter.Pdf.Tests/GroupedGlyphTransformTests.cs)。

## 验证状态

本候选正在验证。源码基线包含 main `c321903e`（PR #68）及 PR #48，初始实施父提交为 `8b0e1ff3`。91 项定向 CGTransform 测试通过，公开 API 文档门通过；首轮完整回归出现的 SVG 私有反射参数问题已修正，复跑及精确包/视觉验收待完成。

已用固定提交 `7e3f2c468aaa56d461f117cc045413c1b457e75c` 的 `xiaoqidun/ofdgo` 生成定向原件并转换同一输入的 PDF/SVG。公共写出器生成的文件仍须核对字体、XML、输出轮廓和文本，不把“写出成功”视为互操作验收。其 `fl → 一个字形` 写出会自动补 `DeltaX="0"`，这是源字符间隔数量，与本候选要求的输出槽位数量不同；这项差异单独保留，不能通过修改原件消除后宣称原件通过。

完整 R02、任意越段映射、补充平面索引、无定位复杂多字形、任意业务外来文件和跨阅读器查找一致性均未由本增量解决。R22 的 typed shaping 生成端有独立契约，见[生成端说明](../text-shaped-output.md)。
