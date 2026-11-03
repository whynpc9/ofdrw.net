---
title: "原生绘图"
description: "使用 OfdGraphics、画笔、路径和矩阵绘制原生 OFD 图元，无需 SkiaSharp。"
sidebar:
  order: 50
---

`Ofdrw.Net.Layout.Graphics` 提供 OfdGraphics、OfdPen、OfdBrush、OfdFont、OfdGraphicsPath 和 OfdMatrix，输出原生 OFD 图元，无需 SkiaSharp 或 System.Drawing。

在 Flow 中使用 CanvasBlock 可取得按块局部坐标工作的绘图上下文：

```csharp
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Graphics;

var document = new FlowDocument();
document.Blocks.Add(new CanvasBlock(100, 28, (graphics, box) =>
{
    graphics.FillRectangle(new OfdBrush(new OfdColor(225, 240, 225)),
        0, 0, box.WidthMillimeters, box.HeightMillimeters);
}));
```

坐标和尺寸使用毫米，Y 向下，文字 Y 为基线。Save/Restore 按栈恢复变换和裁剪；回调期间不要修改文档块集合。图片与 Canvas 块不可跨页拆分，尺寸超出可用正文区域时失败。

仿射矩阵、路径、裁剪和字体资源绑定的精确定义见[绘图契约](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/graphics-design-contract.md)。光栅渐变和图片 tile 不等于支持任意外来原生 gradient/pattern。
