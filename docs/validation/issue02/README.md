# Issue 02 图片进出验证记录

2026-10-02（Asia/Shanghai）。基线 `origin/main` = `be8b74b3cbd699732ec0638189d14037e052a770`。只交付 Issue 02，不依赖 01/16，不合并 PR 或发布 NuGet。

API 设计经 GPT-6 Astra High 子代理只读核查；实现由主代理完成。GPT-6 Sol Low 子代理独立验证发现并复验了 APNG、极低 ppm 和嵌套签章预算修复。

| 验证层 | 当前结果 | 实际范围 |
| --- | --- | --- |
| 功能回归 | 200/200 通过 | Core 5、Packaging 23、PDF/Image 89、Signatures 4、DOCX 49、CLI 30 |
| 本地包消费 | 11/11 通过 | 独立缓存消费 `0.1.0-issue02.review2`；新图片 API 和安装后的 CLI 两方向，加既有 DOCX/PDF/SVG/签章 E2E |
| 自动渲染 | 通过 | 新样例两页 text/image/path、PNG/JPEG选页；PNG/JPEG导入两页居中往返；Native/default基准文本完整、两页逐页渲染 |
| PNG/JPEG 目视复查 | 10/10 完成 | 新样例PNG第1–2页、JPEG第2页、导入往返第1–2页、Native/default各第1–2页，加重复嵌套外观一页 |
| macOS Preview | **未完成** | Computer Use 报告 Mac 锁定且自动解锁失败；已请求手动解锁。PNG 不代替 Preview |
| PR CI / Codex / Cursor | 待到齐 | PR 创建后补充最新 head、检查和线程状态 |

需要的 Preview 路径：本次 `source.ofd → source.pdf` 第 1–2 页、`input.png/input.jpg → imported.ofd → imported.pdf` 第 1–2 页；本次 `generated-layout.docx → 显式 Native/default OFD → OFD导出PDF` 各第 1–2 页。必须重新打开实际 PDF，检查后关闭相应窗口。无法访问 GUI 前，本票保持未完成。

## 功能边界

- API零基单页、CLI一基单个正整数；默认PNG，JPEG显式选择，格式不依赖后缀。非整数ppm、低于72DPI/高于300DPI、极小页面ceil尺寸有回归。
- 双页分别有文字、四象限图片和不同颜色路径，断言每项区域内容及白底，不以单页非空推断正确。
- PNG/JPEG一图一页、顺序、自然尺寸、固定页等比缩小/居中、不放大、原始字节/资源去重和导出后的四象限位置有回归。
- 坏格式、截断PNG、APNG、非法页号、NaN/Infinity、输入/累计/像素/工作缓冲/中间PDF/条目/输出超限、取消、已有目标保持有回归。APNG在解码前扫描编码容器拒绝。
- Stream最后发布I/O/取消无法回滚，CLI提供同文件系统原子替换。解码/native阶段仅前后检查取消。预算与保真边界见[教程](../../tutorials/15-image-conversion.md)。
- 图片导出签章只加载选中页外观，嵌套OFD继承调用方预算；原公开OFD→PDF默认行为保留。

## 可复查产物

完整实际产物在 [evidence.tar.xz](evidence.tar.xz)，索引/哈希/字节数在 [manifest.json](manifest.json)，均随PR提交。压缩包保留三个新OFD、对应PDF/逐页PNG/JPEG、两个本次DOCX显式Native/default OFD/PDF/逐页PNG、功能和包消费日志与11包manifest；读者无需访问被忽略的 `artifacts/`。

完整性检查：`python3 scripts/verify-image-io-evidence.py`。该命令核对归档本身和每个实际载荷的字节数、SHA-256、成员集合。解压：`mkdir -p artifacts/issue02/review && tar -xJf docs/validation/issue02/evidence.tar.xz -C artifacts/issue02/review`。随后在Preview打开解压目录里的五个PDF。小型[第1页PNG](export-1.png)、[第2页PNG](export-2.png)、[导入几何PNG](imported-1.png)也可直接在PR查看。

Native/default OFD各约15 MiB，主要是原有字体嵌入；它们的展开字体载荷各23,278,008 bytes。这是本次旧Native转换基准的实际体积，不是新增图片API膨胀。新source/imported OFD分别约2.9/3.3 KiB，新source/imported PDF约68/3.3 KiB；全部精确字节数见manifest。

源码、环境、哈希、字节数、样例、模式与检查范围由 manifest 记录。Native/default基准来自11包本地消费本次生成的OFD；PDF由这些OFD导出，未使用直接DOCX→PDF代替。

复现：先按根 AGENTS 设置 writable `DOTNET_CLI_HOME`、跳过首启/遥测、显式 `NUGET_PACKAGES`，运行全套单节点测试和 `scripts/run-converter-package-e2e.sh 0.1.0-issue02.review2`。图片样例生成测试入口：`OFDRW_IMAGE_EVIDENCE=<directory> dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests -c Release --filter FullyQualifiedName~SaveReviewEvidence`（附 AGENTS 构建参数）。CLI导出基准：`ofd-to-image generated-docx-{native|default}.ofd <page.png> --pages {1|2} --ppm 4`。

## 实际页面记录

- 新图片样例导出PNG第1–2页：已查看本次图片，中文“样例”、英文/粗体/局部紫色斜体、四象限图片和红/蓝路径分别可见；无全黑、乱码、重影、越界或异常空白。范围为80.3×60.4 mm两页。
- 图片导入再导出PNG第1–2页：已查看本次图片，PNG/JPEG四象限保持顺序与原始像素方向；40×20 mm图在60×50 mm页中央，左右10 mm、上下15 mm；背景和图形分界正常。
- JPEG第2页：已查看，本页中文/英文、四象限和蓝色路径可辨，未发现裁切/重影；JPEG有损边缘属约定编码行为。
- Native/default各第1–2页：均查看本次OFD经新API导出的PNG；第1页中英文标题/比例斜体、局部粗体、蓝灰表格底色和边框正常；第2页分页明确，红色粗体限制在对应文本，右对齐日期完整。未发现缺字、样式扩散、裁切/重叠、重影或多余空白页；既有的大段页内空白符合确定性样例显式分页。
- macOS Preview全部未完成；工具多次确认Mac锁屏，故没有逐页Preview结论或截图。

未据此推断任意复杂Word/OFD保真，也不宣称厂商阅读器互认。

## 首轮审阅修复

Cursor在`2d5d260`确认22载荷完整性与票据API/CLI契约，并提出以下增量；Codex在同一head无发现。

- Windows：中间PDF写句柄在PDFium按路径打开前明确关闭，临时文件在DocReader销毁之后才删除；文件句柄回归与图片API/CLI加入Linux/Windows/macOS矩阵。
- 固定页导入：只限制最终页/图像几何，按像素计算缩放因子，允许大于10000mm的自然尺寸和极低ppm；未指定页尺寸时仍拒绝超限自然页。
- CLI：省略`--output`时末个位置参数必须以`.ofd`结尾，防止误覆盖末张PNG/JPEG；混用`--input`和位置参数仍按出现顺序生成页。
- 选中页嵌套OFD签章超限或 malformed `InvalidDataException` 都失败；此严格语义在教程明确，原公开OFD→PDF保留跳过无效外观行为。

修复后全套200/200。产物与11包消费证据已在`ce23288`重新生成并通过；9张新PNG/JPEG逐页重新查看，几何与文字/样式/表格检查未见新缺陷。Preview仍受锁屏阻塞；尚不宣称完整视觉门或最新重审闭合。

## 第二轮复审修复

`2dd3c3e`的五项CI通过且Windows图片API/CLI实跑通过。两类bot的本轮结果均已读完：Codex提出重复嵌套签章展开与微小几何精度，Cursor提出非二进制整除浮点超出页轴；均纳入同轮。

- 图片严格路径按载荷内容缓存嵌套包与共享字体上下文，绘制也复用同一PDF form/bitmap。外包加唯一嵌套包累计占用展开载荷bytes、已物化非目录文件条目及页面预算；每次读前只给剩余额度。各ZIP原始条目数仍由loader逐包限制，不把物化条目误称累计ZIP目录项。单次转换最多1000个选中候选StampAnnot，无效边界/无载荷也计数，超限在载荷提取前失败。
- canonical包条目的载荷提取缓存成功与空结果，直接payload复用原始数组；ASN.1扫描保留一个最大候选，避免候选列表放大。提取副本、XML/native/font/PDF开销不属于展开载荷budget，不声明严格RSS上限。
- 每轴缩放结果clamp到已验证页轴，145×145／0.01ppm／10000mm页与39×39／0.001ppm／210mm页回归确认无ULP越界/负原点。
- 现有writer为0.001mm精度。最终导入页及图像每轴小于0.001mm明确拒绝；固定页、自然页、高ppm及缩小后的极薄图都覆盖，失败不生成零尺寸图元。

新增重复签章单份展开budget成功/单PDF form、不同载荷累积bytes/entries/pages失败、共享ASN.1载荷对象身份、无效候选限额与几何回归。主套200/200；Sol Low独立复核未发现新确定性缺陷。仅外观fixture不代表密码学签名有效。

第二轮实际产物均从`d8ff853`重新生成，11/11本地包消费已再次通过。10张当前PNG/JPEG逐页重新查看：两页文字/图片/路径、JPEG第2页、两页图片居中往返、显式Native/default各两页和共享嵌套外观一页均未见新缺陷。新外观样例6个红色10×10mm方形位于120×80mm页面同一行，5mm起点、18mm间距，全部可见且无裁切/异常叠加；该fixture只测外观，不含可验证密码学签名。Preview另外需打开本次`shared-seal.ofd → shared-seal.pdf`第1页。新bundle保留25个确切载荷及日志，源码/模式/字节数/SHA-256/待Preview页码见manifest。
