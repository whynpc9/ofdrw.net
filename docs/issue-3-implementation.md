# issue #3 实现对照

本文件记录仓库增量，不修改或关闭 [GitHub issue #3](https://github.com/whynpc9/ofdrw.net/issues/3)。04+05 均闭环后才可判断 issue 的完成定义。

| 要求 | 当前实现 | 证据与状态 |
| --- | --- | --- |
| 公开 Graphics2D 类 OFD 原语 API | Layout.Graphics 的 Graphics/Pen/Brush/Font/Path/Matrix/Options/FillRule；线、矩形、曲线、文字、完整 affine、状态栈及 clip | [契约](graphics-design-contract.md)、[教程](tutorials/16-native-graphics.md)、[两页样例](../e2e/Ofdrw.Net.Graphics.Sample/Program.cs)；04 实现完成，本次 R5 10 页 Preview 已通过，429 回归与 11 包消费通过；`5f1a1e8` 双 bot/六检查及全部 16 线程已闭环；04 验证完成、PR 未合并，后续文档提交读回见 PR |
| 原生图元及可搜索文本 | PathObject/TextObject、普通 Unicode TextRun；沿用现有资源 ID 与 Writer/Reader | Layout 合同回归、PDF 栅格与 SVG 填充回归，实际 OFD 往返及本次 PDF/SVG 全页产物 |
| 无 Skia 必需依赖 | 生成样例只引用 Layout，传递 Core/Packaging；不引用 SkiaSharp/System.Drawing | 样例项目及生成产物；19 可选适配仍独立交付 |
| 字体子集化、复用与缺字回退 | 04 使用既有 FontResourceId/package.Fonts；没有另建字体服务 | [05 绑定边界](graphics-design-contract.md#与-05-的字体绑定边界)；05 并发票据负责，04 不宣称完成 |

功能测试、自动渲染、PNG 复查、Preview、CI 与最新 head 评审分别列在 [验收证据](evidence/graphics/README.md)。04 没有新增 flow/table/Canvas；01/16 分支只读参考，未并入本 stacked 增量。
