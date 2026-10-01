# 公开流式表格

`Ofdrw.Net.Layout` 的 `Table`、`Row`、`Cell` 可与 `Paragraph` 混排。描述可反复渲染，每次创建独立包；渲染期间不得并发修改描述。所有几何量以毫米计。首版提供固定列网格、水平合并、段落水平对齐、整格垂直对齐、底色和统一实线边框。

```csharp
var document = new FlowDocument();
document.Blocks.Add(new Paragraph("月度报表"));
var table = new Table { BorderWidthMillimeters = 0.2, SpaceAfterMillimeters = 4 };
table.ColumnWidthsMillimeters.Add(45);
table.ColumnWidthsMillimeters.Add(45);
table.ColumnWidthsMillimeters.Add(50);
var row = new Row { MinimumHeightMillimeters = 16 };
var heading = new Cell("合并标题 / Summary")
{
    ColumnSpan = 2,
    BackgroundColor = new OfdColor(210, 230, 250),
    VerticalAlignment = CellVerticalAlignment.Center
};
heading.Paragraphs[0].Alignment = ParagraphAlignment.Center;
row.Cells.Add(heading);
var amount = new Cell("128.50") { VerticalAlignment = CellVerticalAlignment.Bottom };
amount.Paragraphs[0].Alignment = ParagraphAlignment.Right;
row.Cells.Add(amount);
table.Rows.Add(row);
document.Blocks.Add(table);
document.Blocks.Add(new Paragraph("表后正文"));
var package = document.Render();
await new OfdPackageWriter().WriteAsync(package, outputStream);
```

## 网格、样式与容量

- 列宽为空时，根据第一行 `ColumnSpan` 总和等分正文宽；显式列宽必须为正有限值，总宽不能超过正文宽，保持给定毫米值，不自动缩放。较窄的表按 `Table.Alignment` 左/中/右对齐。
- 每行必须准确覆盖全部列。空表、空行、非正/越界跨度、缺格、多格及没有可用文字宽度的 padding/缩进会抛参数异常。空 `Cell.Paragraphs` 合法，仍占 padding 与行最小高度。
- `Cell.Paragraphs` 保留段落缩进、段前/后间距、换行、混合字号和局部 Span 粗体/斜体/颜色。各段 `Paragraph.Alignment` 控制水平对齐；`Cell.VerticalAlignment` 对齐整组段落。没有字号缩小或裁切来满足行高。
- `BackgroundColor` 只作用于该格。`Table.BorderColor`/`BorderWidthMillimeters` 是整表统一边框，宽度 0 关闭。合并格内不画竖线，公开表的相邻共享边只写一次，跨页片段各自封闭。外边向内偏移半笔宽，极薄行小于笔宽会失败。首版不提供逐边/逐格边框覆盖、虚线或冲突消解。
- 格内文本使用 `MaxCharacters`/`MaxTextElements`，空格也计入 `MaxTableCells`（默认 100000）；页数使用 `MaxPageCount`，支持取消。字符预算保留 UTF-16 计数语义。

## 分页与明确失败

完整测量一行后再落页。当剩余高度不足时，整行移至下一页，文字、底色和边框一起移动。新页仍不能放下时，公开 API 抛 `InvalidOperationException`，消息包含行号、行高与可用高度。`MinimumHeightMillimeters` 是最小高度，不是强制裁切高度。首行消费前段延期空行与段后间距；`Table.PageBreakBefore` 与段落采用相同的首块保护及清除延期空行规则。表后间距延至下一块，表末不会因尾间距产生空白页。

`RowSpan > 1` 抛 `NotSupportedException`；`RowSpan < 1` 是参数错误。格内 `Paragraph.PageBreakBefore` 和文字 `\f` 同样明确拒绝。首版不拆行、不重复表头，不支持跨页纵向合并、嵌套表和图片格。CJK 仍按 1 em 测量，公开 Flow 不嵌入字体，跨机器效果取决于阅读器字体；不承诺任意 Word 表格保真。

## DOCX Native 兼容边界

公开表与 Native 表格共用 `FlowTableLayout` 行宽/行高/垂直定位、`FlowParagraphLayout` 折行和 `FlowPagination` 不可拆块分页。Native 保留其字体、图片、字段、节页眉页脚、列宽归一化、0.8 mm padding 及既有格内段落策略；没有把 DOCX 简化为公开 Span。Native 既有逐格边框覆盖保持，公开表共享边去重保证不扩大到 DOCX。

Native 新增严格网格校验：非正 `gridSpan`、越界/不满网格的行或无可用格宽会抛 `InvalidDataException`，通过既有临时输出提交机制保持目标流不被部分写入。这会拒绝过去可能丢格或忽略余列的输入；包含 `gridBefore`/`gridAfter` 的省略网格表格不在首版适配范围。

DOCX `vMerge`（restart/continue/空元素）首版未实现：`UnsupportedFeatureBehavior.Throw` 拒绝，其余模式返回 `DOCX_VERTICAL_MERGE_DEGRADED`，带行/格位置，说明降为独立单元格。格内显式或继承样式的分页控制同样按策略拒绝或返回 `DOCX_CELL_PAGE_BREAK_DEGRADED` 后忽略。请用 `ConvertWithResultAsync` 读取诊断。`tblHeader` 仍不重复。DualLayer/直接 DOCX→PDF 的表格引擎没有迁移，不用于替代 Native 验收。

## 复现与验收

```sh
dotnet build e2e/Ofdrw.Net.Layout.E2E -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet run --project e2e/Ofdrw.Net.Layout.E2E -c Release --no-build --no-restore -- artifacts/public-tables/current e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx --tables-only
```

运行前按根 AGENTS 设置可写 `DOTNET_CLI_HOME`、关闭首次体验和遥测。样例生成中英比例字体、局部红色粗体/斜体、底色、合并、三种水平/垂直对齐及整行跨页。它核对输入→OFD 原文及 OFD→PDF 无重复文本，另外转换既有 `generated-layout.docx` 的显式 Native/default 基准。实际页面仍须逐页查看；[验收记录](validation/public-tables-2026-10-02.md)分别记录功能、自动渲染、PNG 与 macOS Preview。
