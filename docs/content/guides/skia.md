---
title: "可选 Skia 事件适配"
---

Skia 适配用于显式接入的合作 producer，将受支持的绘图事件转为 OFD 图元。它不会捕获任意现有 SKCanvas，也不等价于完整 Graphics2D。

默认 Layout 绘图不需要 Skia。只有应用已有可接入的生产端或使用 PDF vector 扩展时，才需要这一独立依赖。

先核对目标框架和 native-assets，再检查字体来源、事件顺序、矩阵及明确拒绝范围。设计见[Skia 契约](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/skia-adapter-design-contract.md)。间歇性字体读取问题仍按[已知问题](https://whynpc9.github.io/ofdrw.net/troubleshooting/known-issues/)被动观察。
