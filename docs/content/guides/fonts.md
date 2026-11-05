---
title: "中文、英文与字体部署"
description: "为 DOCX Native 和 PDF 导出配置真实 Regular/Bold 字体、fallback 与严格粗体策略。"
sidebar:
  order: 25
  label: "字体部署"
---

中文转换需要调用方提供许可明确的字体文件。`FontFamily` 名称、实际 face 与 Bold/Italic 载荷是不同信息；Regular 加样式标记不能替代真实 Bold 文件。

## 固定物理字体配置

固定配置使用静态 Noto Sans CJK SC 400 / 700 及 Noto Sans Regular。来源、SHA-256、face index 与具体适用范围见[发布字体配置](https://whynpc9.github.io/ofdrw.net/reference/release-profile-cn-en/)。部署时保留字体许可证；字体不加入 NuGet 包。

`DocxConversionOptions` 提供 FontDirectories、FontFallbackFamilies、UseInstalledMicrosoftOfficeFonts 与 NativeEmbeddedBoldBehavior。明确字体目录后，应确认输出实际嵌入的 face，而不是只检查配置名称。

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

## 0.2.0 的严格粗体选项与渲染修复

`0.2.0` CLI 支持 `--native-embedded-bold-behavior throw`。本库的 OFD→PDF 和 OFD→Image 出口已修复固定 Regular-only 样例中“量”的字腔堵塞，详见[PR108 验收](https://github.com/whynpc9/ofdrw.net/blob/main/docs/validation/pr108-cross-platform-20261103.md)。固定样例的发布验收范围见[0.2.0 发布说明](https://whynpc9.github.io/ofdrw.net/releases/020/)。

部署故障先核对实际载荷、字体许可、缺字诊断和样式文件，再核对运行时原生依赖。字体子集和载荷复用见[字体参考](https://whynpc9.github.io/ofdrw.net/reference/font-subset-and-reuse/)。
