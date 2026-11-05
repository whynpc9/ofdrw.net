---
title: "R10 受限图像与区域文字配置"
description: "导入 OFD 的受限 JBIG2、ImageMask 和基于原字体几何的区域文字提取。"
sidebar:
  order: 130
  label: "图像与区域文字"
---

此配置用于导入的 OFD 内容。Reader 保留原始 `SourceXml`、资源标识符和编码字节。导出器为渲染创建临时 PNG 像素；包写出器不会替换源媒体。

## 图像

- JBIG2 资源必须显式声明 `Format="JBIG2"`，并与独立文件签名匹配。准入子集仅包含一个尺寸已知的页面，以及按顺序出现的 `PageInformation`、一个覆盖整页的立即算术编码通用区域、`EndOfPage` 和 `EndOfFile`。不接受字典、全局段、引用、细化、MMR、未知段长度或尾随载荷。输入上限为 1 MiB，解码面积上限为 1,048,576 像素。解码器的尺寸与 RGB 计数必须与预检结果一致。
- `ImageMask` 在与图像相同的资源作用域中解析。自引用、无法解析或不匹配的声明以及非二值格式均会失败。准入掩码为一位灰度 PNG 或上述受限 JBIG2 配置，且尺寸必须与源图像相同。白色掩码像素保留源 alpha，黑色像素清除 alpha，随后由导出器应用对象 alpha。GB/T 33190-2016 规定了同尺寸二值掩码资源，但其 ImageObject 表未说明这种极性；在将此映射视为确定行为之前，仍需独立阅读器视觉验收。
- PDF、SVG 和 HTML 使用相同的规范化渲染投影。未知图像子节点、属性和资源引用会在公开输出前导致失败。格式错误的压缩数据仍可能在解码时失败；此结果不代表支持所有 JBIG2 变体。

托管解码器为 `JBig2Decoder.NETStandard` 1.5.2。原 JPedal 的 BSD-3-Clause 声明和 C# 移植版的 MIT 声明均保留在 `docs/licenses/JBIG2DECODER-NETSTANDARD-LICENSE.txt`，并随 Reader 包分发。测试位流由自创图案生成，并使用独立 `jbig2dec` 作为对照；详见 R10 测试样例 README。早期诊断使用的 ITU 衍生第三方样例未包含在仓库中。

## 区域文字

`OfdRegionTextExtractor` 接受一个相对于所选页面 PhysicalBox、以毫米为单位且具有有限正面积的矩形。它按背景模板、正文和前景模板的绘制顺序返回 Unicode 标量，不自行添加空格或换行。对于有墨迹的标量，只有其经变换的原字体字形四边形与矩形存在正面积交集时才予以纳入；仅接触边界不算选中。对于无墨迹标量，使用变换后的基线原点，左、上边界包含，右、下边界不包含。`OfdKeywordFragment` 值公开源 TextCode 偏移、字体身份、字形 ID、由 CTM/Delta 推导的基线以及精确的字体框几何。模板及其引用的 XML 会接受来源归属检查。批注外观被排除，并通过 `ExcludedAnnotationScope` 诊断报告。不支持的来源、变换或裁剪、字体几何问题以及预算耗尽都会使整个查询失败。

查询要求为实际嵌入字体显式提供 `IOfdGlyphGeometryProvider`。它不会替换为宿主机字体，也不宣称返回绘制像素的边界。
