# D11：Regular-only TextStroke 横向填充强调（issue #107）

日期：2026-11-03（Asia/Shanghai）。基线 main `cf3f11a7719cb55d5088698e64304f73b5d3e5bf`。结论：**源码修复已提交；Linux 功能测试、自动渲染与逐页 PNG 复查通过；真实 macOS Preview 验收未完成**。按仓库规则，issue #107 在 Preview 补验前保持打开。

## 新增因果证据

D07 列出的重开条件是“定位到可达、独立且确定的产品绘制缺陷”。本轮固定 issue #107 的最小反例（`数量`，Size 3.704 mm，LineWidth 0.037 mm，固定 Noto Sans CJK SC Regular 子集），在公共 `OfdToImageConverter` 默认 144 DPI 下得到三组对照：

| 渲染 | “量”下／上字腔亮行均值 | 两字墨量 | 说明 |
| --- | --- | --- | --- |
| Regular 文字（Weight 400，无描边） | 202.9／213.7 | 84.0 | PDFium 对嵌入 TrueType 文字做网格对齐 |
| main：Weight 700 + 1% em 描边 | 27.8／26.0 | 146.8（1.75×） | 0.037 mm＝0.21 px 的描边被放大为 1 设备像素 hairline |
| 控制：同一原轮廓仅作路径填充（不加粗） | 81.0／29.1 | 122.2（1.46×） | 无网格对齐的原轮廓本身已使上字腔变暗 |

由此得到两条可复现的机制：

1. **hairline 下限**：写入端的 1% em 描边在普通尺寸下远小于 1 设备像素；PDFium、Poppler 均把它放大到约 1 px，实际加粗约为设计值的 5 倍，并同时在竖直方向吞掉 1.16 px 高的“日”字腔。这是确定的绘制缺陷，不是字体、GID 或坐标损坏。
2. **整字路径重绘会失去文字网格对齐**：PR78 各路线均以路径重绘整字（几何描边、并集、裁剪）。上表控制组表明，即使轮廓完全不变，路径化本身就使上字腔从 213.7 降到 29.1，因此任何“整字路径”方案都无法在 PDFium 144 DPI 下达到 >100 的原判据。

## 修复机制

仅改 OFD→PDF 渲染（因此也覆盖公共 OFD→Image）；OFD 写入、原文、GID、定位、advance、字体字节均不变，同一 DOCX 的修复前后 OFD 内容除时间戳外逐字节相同。

- 识别写入端自身的 Han 合成粗体描边（同色、Fill+Stroke、宽度＝1% em，且 TextObject 含 Han），**保留原 PDF 文字绘制**（继续享有阅读器的文字栅格化）。
- 不再绘制描边，改为额外填充一个“仅水平膨胀”的覆盖层：对每个字形的扁平化轮廓逐水平带计算 nonzero 区间，生成 `[L−h, L+h]`／`[R−h, R+h]` 梯形并在间距 ≤2h 时合并；h＝em/48（总宽 em/24，即 FreeType `FT_GlyphSlot_Embolden` 的常规强度，事先固定，未扫描）。
- 覆盖层只在原字形已有的竖直范围内横向生长，**不会进入堆叠横画之间的字腔**；以填充而非描边绘制，不受 hairline 下限影响；全部梯形在同一 nonzero 路径中填充，无相邻接缝；向内 h 的重叠避免与原字形边缘之间的抗锯齿浅缝。
- 旧文档的 2.5% em 描边、不同颜色／宽度的显式描边、纯描边文字、名称引用字体均保持原路径。

## 预登记与实验预算

实验前固定：输入与字体身份（下文）、144 DPI、原 PR78 字腔窗口及 >100 阈值、上横非空检查、OFD/PDF 原文一致。计划 ≤4 组候选，结果如下：A（整字 3 份水平平移并集、文字不可见）81/29 量级失败；B（再加 1% em 竖向）0/0 失败；R（不加粗路径控制）确认机制 2；S（本修复）通过。没有在看到结果后修改阈值或强度。

强调强度：S 在最小反例上的两字墨量为 Regular 的 **1.23×**（main 失败版为 1.75×）。历史 B/C 相对强调目标未能取回，因此本记录只声称“可见且明确强于 Regular”，不声称等同物理 Bold 或原 B/C 目标；回归测试以 1.15× 作为防止强调消失的下限，而非目标值。

## 固定输入

| 项目 | 身份 |
| --- | --- |
| DOCX | `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`，SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6` |
| CJK Regular | `Ofdrw-CI-NotoSansCJKsc-Regular.ttf`，21,659,544 字节，SHA-256 `3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa` |
| Latin Regular | `NotoSans-Regular.ttf`，SHA-256 `b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5` |
| 配置 | `FontDirectories` 仅含上述两 Regular；`UseInstalledMicrosoftOfficeFonts=false`；显式 `NativeEmbeddedBoldBehavior=TextStroke`；分别运行显式 `Native` 与默认模式 |
| 环境 | Linux（WSL2 5.15），.NET SDK 10.0.401，PDFium（Docnet，公共 OFD→Image）与 Poppler `pdftoppm -r 144` |

## 结果

### 功能测试

- 新回归 `DenseCjkSyntheticBoldTests`：最小反例在公共 144 DPI 导出中字腔 >100、上横有墨、墨量 ≥1.15× Regular，PDF 原文 `数量` 两字无重复；在未修改的 main 上同一测试失败（27.8／26.0）。另一测试固定覆盖层几何不进入堆叠横画间隙。
- `BoldAlignmentTests` 中 `O中`／0.08 mm 一行改为预期“不同于普通描边”，因为 Han 合成粗体现以填充覆盖层绘制；其余行不变。
- 按 CI 方式提供固定 CJK 字体（隔离 `HOME/.fonts`）与可信字体 fixture 后：`Ofdrw.Net.Converter.Pdf.Tests` 1212/1212 通过、0 跳过；`Ofdrw.Net.Converter.Docx.Tests` 308/309，唯一失败 `DifferentPageSizesKeepHeadersFieldsTablesAndImagesThroughDrawingSeam` 属于未改动的直接 DOCX→PDF 渲染器（0.0001 pt 舍入），在未修改的 main 上同样失败。
- 全解决方案（未提供宿主 CJK 字体）：Core、Packaging、Signatures、Cli、Layout、SesInterop、SkiaSharp、Pdf.Vector、Crypto.Password 全部通过；Pdf/Docx 的失败均为缺少宿主 CJK 字体（`Arial lacks U+91CD` 等），已在 main 上复现相同失败。
- 未运行本地包消费 E2E（`run-converter-package-e2e.sh`）；交由 PR 的 CI 执行。

### DOCX 链路（DOCX → native/default OFD → PDF → PNG）

两模式均 2 页；OFD 与 PDF（PdfPig）原文 189 字一致、无重复字母；7 个粗体对象（含“数量”）全部引用 Regular 字体面（Bold 标志为 false），OFD 中仍为 `Stroke=true`、1% em `LineWidth`。第 1 页“数量”表头按同一窗口测得（浅蓝表头底色）：

| 查看路径 | main | 本修复 |
| --- | --- | --- |
| 公共 OFD→Image（PDFium，144 DPI），native/default | 24.0／22.2 | 169.9／179.7 |
| PDF → Poppler `pdftoppm -r 144`，native/default | 0.0／0.0 | 200.6／207.2 |

### PNG 视觉复查

逐页查看了 native 与默认模式各 2 页（PDFium 144 DPI），并放大标题、二级标题、表头与第 2 页红色粗体：“量”字两个上部字腔在普通尺寸可辨；粗体仍明显重于相邻正文；无重影、错位、接缝、裁切、重叠、异常空白页；红色与蓝色局部样式未扩散；表格底色与边框不变。Poppler 页面同样检查了第 1 页表头（修复前“量”为实心块、“数”几乎未加粗；修复后两字一致加粗且字腔打开）。

### macOS Preview

**未完成**：本环境为无 GUI 的 Linux。macOS 不做竖向 hinting，原 Regular 字腔本身的观感取决于 Core Graphics；覆盖层不进入字腔，因此预期“量”与同字体 Regular 同样可辨，但这只是推断，不能替代 Preview 验收。补验时请用 `DOCX → native OFD → PDF → Preview` 与默认模式各查看 2 页，普通尺寸检查第 1 页表头与第 2 页红色粗体。

## 本次产物

位于本工作树 ignored `artifacts/i107/`（未提交）。显式 native 与默认模式的页面 PNG 逐字节相同。

| 文件 | SHA-256 |
| --- | --- |
| `docx-after/native.ofd` | `52241d36c0989f9d75b9ab3e0cc3d99454b270981631ad23385fa845264ce796` |
| `docx-after/default.ofd` | `4ab072a8ad5182c9b2cc0f5d0634ef449aec6e463a774a588bbc0b00e8d4081b` |
| `docx-after/native.pdf` | `d4fcebebd5df14967a52c23b4aff81281f74432d0e95b25f9c4780aac51f7b54` |
| `docx-after/default.pdf` | `2547cfb75f140c4c7d4e014ace8487594a1df6c9df60fd7a5a1e53792cf8efd7` |
| `*-image-p1.png`（PDFium 144 DPI） | `56e4e3cba74b158a16f64209c74125c87471a90ddbf3b848094d708832223bff` |
| `*-image-p2.png` | `f28520a917136c14708aa0850143e101ab573c5d25f1cf097a3bf4a8be67f835` |
| `*-poppler-1.png`（Poppler 144 DPI） | `c8a4cb56b165b670e55cecd3fa14d36d13c1ff03616efbdac3381709239310da` |
| `*-poppler-2.png` | `9c521fed7d1fb9201fa186762d82af6300f4dedbc21755bca797a1799065b2bc` |

## 遗留与范围

- 其他 OFD 阅读器直接渲染 OFD 中声明的描边，不受本修复影响；本修复只覆盖本库的 OFD→PDF／OFD→Image。
- 只对上述样例、字体与页面作出结论；不推断任意字体、字号或 Word 文档。
- 推荐配置仍为显式物理 Regular/Bold 与 `Throw`；本修复改善默认 TextStroke，不使其等同物理 Bold。
