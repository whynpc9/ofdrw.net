# 类 Graphics2D 绘图 API（issue #3）

**What to build:** 调用方只用 OFD 原语 API（`OfdGraphics`、`OfdPen`、`OfdBrush`、`OfdFont`、`OfdGraphicsPath`、`OfdMatrix`，名称可微调）画线、矩形、路径、文字和变换，结果落到 `PathObject` / `TextObject`。不必引用 Skia。

**Blocked by:** None (can start immediately)

**Priority:** P1

**Status:** implemented-awaiting-review

- [x] 能画线、矩形、路径、文字并做变换，写入合法 OFD
- [x] 教程样例不依赖 Skia
- [x] 更新功能对照与 issue #3 实现对照

## Parent

[GitHub issue #3](https://github.com/whynpc9/ofdrw.net/issues/3)；[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P1-01

## What to build

issue #3 点名的绘图层。交付一套语义对齐上游 Graphics2D 的公开 API：画完一页票面框线或示意图后，包内是标准路径和文字对象，而不是位图。实现可以架在自绘路径上；这张票不得把 Skia 变成调用方的必需依赖。

## Acceptance criteria

- [x] 公开类型覆盖画笔、填充、字体、路径、矩阵；能画线、矩形、任意路径、文字，并施加变换
- [x] 生成的对象是 OFD `PathObject` / `TextObject`（及必要资源），用现有读写往返能读回来
- [x] 教程样例只依赖 Layout/Core 一类 OFD 包，不引用 SkiaSharp
- [x] Preview（OFD → PDF → Preview 或 SVG）能看出线宽、填充、文字基线和变换，无重影或裁切错误
- [x] 更新功能对照；不要关闭或改写 issue #3 的正文

## Blocked by

- None (can start immediately)

## 本次实现

- [设计/05 字体绑定契约](../../../docs/graphics-design-contract.md)
- [公开 API 教程](../../../docs/tutorials/16-native-graphics.md)
- [issue #3 实现对照](../../../docs/issue-3-implementation.md)
- [验收证据](../../../docs/evidence/graphics/README.md)

04 单票不新增 flow/table/Canvas，也不实现 05 字体子集或 19 Skia 适配。最新 head 评审与 Preview 尚未闭环时保持此状态。
