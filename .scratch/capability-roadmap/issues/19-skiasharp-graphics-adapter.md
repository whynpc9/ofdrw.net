# SkiaSharp 绘图层（可选适配）

**What to build:** 可选的 SkiaSharp 适配，把 Skia 绘制转到 OFD 原语。不替代 `OfdGraphics`；调用方仍应能只依赖 OFD API。主要为后续 PDF 矢量模式服务。

**Blocked by:** 04 类 Graphics2D 绘图 API（issue #3）

**Priority:** P1

**Status:** implemented-functional-validated-preview-pending

- [x] 可选包/适配能把基本 Skia 绘制落到 PathObject/TextObject
- [x] 不替代 P1-01 的 OFD 原语 API；核心生成路径不强制引用 Skia
- [x] 有样例说明何时才需要这个适配

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P1-12

## What to build

有人已经用 Skia 画图，希望落到 OFD。这张票提供可选适配层，而不是第二套官方绘图 API。`OfdGraphics` 仍然是调用方该依赖的表面。没有这层，04 必须已经能独立画完线框和字。

## Acceptance criteria

- [x] 可选适配将线、矩形、路径、文字的 Skia 绘制转到与 04 相同的 OFD 对象
- [x] 默认 Layout/Converter 包不因这张票而必须引用 SkiaSharp
- [x] 文档写明：这是适配，不是 Graphics API 本身；issue #3 的完成定义仍是 04+05
- [x] 有最小样例

## Blocked by

- 04 类 Graphics2D 绘图 API（issue #3）

## 本次明确受限实现

基线 `ab87bca`；Astra High 设计通过，协调方接受合作 producer 显式事件范围。
[设计契约](../../../docs/skia-adapter-design-contract.md)、[迁移教程](../../../docs/tutorials/17-skia-event-adapter.md)、[证据](../../../docs/evidence/skia/README.md)。

`SkiaDrawEvent` + `OfdSkiaAdapter.Append` 只接收 producer 明确提交的 Skia 类型事件，不拦截任意既有 SKCanvas/SKPicture/PDF 调用。原文由 producer 提供，字体唯一 ID 与实际载荷/face 元数据验证；不引入 05 服务。透明捕捉的真实失败探针保留。446/446 与主/Low 独立真实包消费通过，Preview 锁屏导致 0/14 未完成，最新 head review 尚待闭环；21 的任意 PDF 绘制来源仍需独立验证。

R3：首轮 Codex/Cursor 到齐后，取消问题以真实 nupkg 复现再修；枚举器获取/推进前取消检查、Dispose 前提交禁止，17回归与 Low 实际包消费证明原子性。当前运行时冻结 `e6e11ee`；14页 r3 新产物仍待 Preview，最新复审未闭环。
