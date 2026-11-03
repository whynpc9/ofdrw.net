---
title: "可选 PDF 矢量扩展"
description: "了解独立 PDF Vector 源码扩展的原生 OFD 输出、页级失败与整页回退策略。"
sidebar:
  order: 180
  label: "PDF Vector 扩展"
---

源码扩展 `Ofdrw.Net.Converter.Pdf.Vector` 提供受限的 PDF→原生 OFD 路径和文字输出，依赖合作式 Skia 事件链路，不属于默认转换包发布集合。

使用时先选择明确的页级失败或回退策略。支持的页框、路径、TrueType 字体和原始 Unicode / GID 必须通过 profile 检查；不支持的压缩流、裁剪、效果、图片或复杂字体不可默默输出部分页面。

默认 Converter.Pdf 与 CLI 仍使用双层路径。启用源码扩展不代表所有 PDF 都可矢量化。详细行为见[矢量契约](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/pdf-vector-design-contract.md)，实际失败和视觉记录见[验收索引](https://github.com/whynpc9/ofdrw.net/blob/main/docs/evidence/pdf-vectors/README.md)。
