---
title: "图片导入与导出"
---

`Converter.Pdf` 中的 `OfdToImageConverter` 和 `ImageToOfdConverter` 提供 OFD 与 PNG/JPEG 转换。

OFD 导出按所选页及 ppm（每毫米像素数）渲染；例如 144 DPI 约为 5.67 ppm。优先使用 PNG 验证文字、线框和透明度，JPEG 为有损格式。

图片导入生成居中、等比缩放的 OFD 页面，可使用自然尺寸或明确页面尺寸。不要把位图文字当作已经有可搜索的语义层。

CLI 使用 `ofd-to-image` 与 `image-to-ofd`，参数以安装版本的帮助为准。页号、像素、字节和临时输出预算见[转换约定](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/)。处理超预算输入会失败，不会无限分配内存。

图片出口经 PDF/PDFium 链路渲染，有独立的原生依赖。实际样例和页面范围见[图片验收](https://github.com/whynpc9/ofdrw.net/blob/main/docs/validation/issue02/README.md)。
