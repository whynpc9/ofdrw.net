# SkiaSharp 绘图层（可选适配）

**What to build:** 可选的 SkiaSharp 适配，把 Skia 绘制转到 OFD 原语。不替代 `OfdGraphics`；调用方仍应能只依赖 OFD API。主要为后续 PDF 矢量模式服务。

**Blocked by:** 04 类 Graphics2D 绘图 API（issue #3）

**Status:** ready-for-agent

- [ ] 可选包/适配能把基本 Skia 绘制落到 PathObject/TextObject
- [ ] 不替代 P1-01 的 OFD 原语 API；核心生成路径不强制引用 Skia
- [ ] 有样例说明何时才需要这个适配

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P1-12

## What to build

有人已经用 Skia 画图，希望落到 OFD。这张票提供可选适配层，而不是第二套官方绘图 API。`OfdGraphics` 仍然是调用方该依赖的表面。没有这层，04 必须已经能独立画完线框和字。

## Acceptance criteria

- [ ] 可选适配将线、矩形、路径、文字的 Skia 绘制转到与 04 相同的 OFD 对象
- [ ] 默认 Layout/Converter 包不因这张票而必须引用 SkiaSharp
- [ ] 文档写明：这是适配，不是 Graphics API 本身；issue #3 的完成定义仍是 04+05
- [ ] 有最小样例

## Blocked by

- 04 类 Graphics2D 绘图 API（issue #3）
