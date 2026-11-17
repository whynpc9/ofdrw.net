---
title: "中文、英文与字体部署"
description: "为 DOCX Native 和 PDF 导出配置真实 Regular/Bold 字体、fallback 与严格粗体策略。"
sidebar:
  order: 25
  label: "字体部署"
---

中文转换需要调用方提供许可明确的字体文件。`FontFamily` 名称、实际 face 与 Bold/Italic 载荷是不同信息；Regular 加样式标记不能替代真实 Bold 文件。

> **可选最终文字排版（自 0.3.0 起）：** 设置 `FlowDocumentOptions.TextEngine` 或 `DocxConversionOptions.NativeParagraphLayoutEngine` 后，Flow、纯文本或 DOCX Native 的文本改用段落排版引擎，以真实字体塑形并嵌入字体子集（支持制表位、合成粗斜体、CFF/OTF 与增补汉字）。默认行为不变，见[启用最终文字排版引擎](https://whynpc9.github.io/ofdrw.net/guides/final-text-layout/)。

## 固定物理字体配置

固定配置使用静态 Noto Sans CJK SC 400 / 700 及 Noto Sans Regular。来源、SHA-256、face index 与具体适用范围见[发布字体配置](https://whynpc9.github.io/ofdrw.net/reference/release-profile-cn-en/)。部署时保留字体许可证；字体不加入 NuGet 包。

`DocxConversionOptions` 提供 FontDirectories、FontFallbackFamilies、UseInstalledMicrosoftOfficeFonts 与 NativeEmbeddedBoldBehavior；main 另有 NativeFontSelfContainment（见下文）。明确字体目录后，应确认输出实际嵌入的 face，而不是只检查配置名称。

- 真实粗体缺失时，严格配置使用 `NativeEmbeddedBoldBehavior=Throw`。
- 清空 fallback 后按所需顺序添加受控家族。
- `UseInstalledMicrosoftOfficeFonts=false` 仅关闭特定 Office 字体暂存，不表示禁用所有系统字体解析。
- 缺字、其他字体、物理 Italic 和复杂脚本需单独验证。

以下为严格物理字体配置。将 `fonts` 替换为已准备好 Regular/Bold 的目录：

```csharp
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;

var options = new DocxConversionOptions
{
    OfdMode = DocxToOfdMode.Native,
    NativeEmbeddedBoldBehavior = DocxEmbeddedBoldBehavior.Throw,
    UseInstalledMicrosoftOfficeFonts = false,
};
options.FontDirectories.Add(Path.GetFullPath("fonts"));
options.FontFallbackFamilies.Clear();
options.FontFallbackFamilies.Add("Noto Sans CJK SC");
options.FontFallbackFamilies.Add("Noto Sans");
var converter = new DocxToOfdConverter(options);
```

省略 `OfdMode` 即检查默认 Native。字体配置成功仍须核对输出页面及实际字体资源。

## 要求字体自包含（main，0.4.0 之后，尚未发布）

默认情况下，宋体、黑体、等线、楷体、仿宋、微软雅黑等系统中文族只写字体名称，不写 FontFile；未从字体目录加载到 `Noto Sans CJK*` 时也是如此。这是阅读器本机字体的兼容路径，默认值不变。

需要字体自包含的输出（例如归档）时，设置 `NativeFontSelfContainment = DocxNativeFontSelfContainment.RequireEmbedded`，CLI 对应 `--native-font-self-containment require-embedded`：

```csharp
var options = new DocxConversionOptions
{
    NativeFontSelfContainment = DocxNativeFontSelfContainment.RequireEmbedded,
    UseInstalledMicrosoftOfficeFonts = false,
};
options.FontDirectories.Add(Path.GetFullPath("fonts")); // 放入有授权的 SimSun 等字体文件
```

- 每个被文字使用的字体资源都嵌入该字体族自身的字节。字体文件自己的 name 表须记录所声明的族名或同族别名（如 SimSun/NSimSun/宋体/新宋体）；TTC 会按现有规则抽取。只按文件名对应、或只有其他字体族可作替代时，转换失败，不会把替代字节写在所声明的族名下。
- 只使用 `FontDirectories` 中的字体；未配置目录时使用平台字体目录中按现有规则发现的中文字体。不使用宿主解析器按名称找到的其他字节。
- 读取字体自身的 OS/2 `fsType`：限制许可嵌入（0x0002）和仅位图嵌入（0x0200）直接失败，异常指出字体族；禁止子集化（0x0100）时嵌入完整字体，同时启用 `NativeParagraphLayoutEngine`（只能写字形子集）则失败；没有 OS/2 表也失败。
- **许可位允许嵌入不等于取得授权**。转换成功会返回 `DOCX_FONT_LICENSE_NOT_VERIFIED` 提示，字体许可仍由调用方负责。
- 失败异常为 `InvalidDataException`，消息以 `DOCX_FONT_NOT_SELF_CONTAINED` 开头并给出字体族和原因。`NativeEmbeddedBoldBehavior` 仍独立生效：缺少真实粗体时默认的文字描边仍算已嵌入。
- 只适用于 Native DOCX→OFD（含默认模式）。DualLayer 与 DOCX→PDF 遇到该选项会抛出 `NotSupportedException`。

## 0.2.0 的严格粗体选项与渲染修复

`0.2.0` CLI 支持 `--native-embedded-bold-behavior throw`。本库的 OFD→PDF 和 OFD→Image 出口已修复固定 Regular-only 样例中“量”的字腔堵塞，详见[PR108 验收](https://github.com/whynpc9/ofdrw.net/blob/main/docs/validation/pr108-cross-platform-20261103.md)。固定样例的发布验收范围见[0.2.0 发布说明](https://whynpc9.github.io/ofdrw.net/releases/020/)。

部署故障先核对实际载荷、字体许可、缺字诊断和样式文件，再核对运行时原生依赖。字体子集和载荷复用见[字体参考](https://whynpc9.github.io/ofdrw.net/reference/font-subset-and-reuse/)。
