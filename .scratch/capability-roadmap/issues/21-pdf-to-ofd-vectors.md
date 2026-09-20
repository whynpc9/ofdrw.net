# PDF → OFD 保留矢量（可选模式）

**What to build:** 可选的 PDF→OFD 矢量模式：用 PdfPig 取出的文本和路径写入 OFD（经 OfdGraphics / Skia 适配）。默认可仍为双层（页面图 + 透明文字）。矢量模式有样例对比。

**Blocked by:** 04 类 Graphics2D 绘图 API（issue #3）；19 SkiaSharp 绘图层（可选适配）

**Status:** ready-for-agent

- [ ] 可选矢量模式写出路径与文字对象；默认可仍为双层
- [ ] 有与双层模式的样例对比
- [ ] 扫描件/无路径页有文档化回退，不把空白当成功

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P1-04

## What to build

当前 PDF→OFD 以栅格页为视觉层、可抽取文字为透明层。这张票增加可选矢量模式：把 PDF 里的路径和文字写成 OFD 原语，页面图可以继续当稳定视觉层或按选项关掉。默认保持双层，避免阅读器突然换皮。矢量路径通过 04 的绘图 API 落地，并使用 19 的 Skia 适配承接 PdfPig/桥接绘制。

不承诺任意 PDF 特效、字体语义或阅读顺序标记一次做完。

## Acceptance criteria

- [ ] 显式矢量模式把样例 PDF 的可见路径和文字写入 OFD；用 SVG/PDF 导出能看出线框而不是整页位图（允许混合）
- [ ] 未开矢量模式时，行为与现网双层一致
- [ ] 有同一 PDF 的双层 vs 矢量对比样例；记录局限（字体、填充、图像）
- [ ] 无矢量内容的扫描页有回退（继续页面图或明确失败），不输出空白页冒充成功
- [ ] 更新功能对照；视觉验收基于本次 OFD 产物

## Blocked by

- 04 类 Graphics2D 绘图 API（issue #3）
- 19 SkiaSharp 绘图层（可选适配）
