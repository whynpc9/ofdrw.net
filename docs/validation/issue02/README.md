# Issue 02 图片进出验证记录

2026-10-02（Asia/Shanghai）。基线 `origin/main` = `be8b74b3cbd699732ec0638189d14037e052a770`。只交付 Issue 02，不依赖 01/16，不合并 PR 或发布 NuGet。

API 设计经 GPT-6 Astra High 子代理只读核查；实现由主代理完成。GPT-6 Sol Low 子代理独立验证发现并复验了 APNG、极低 ppm 和嵌套签章预算修复。

| 验证层 | 当前结果 | 实际范围 |
| --- | --- | --- |
| 功能回归 | 269/269 通过 | Core 5、Packaging 23、PDF/Image 158、Signatures 4、DOCX 49、CLI 30 |
| 本地包消费 | 11/11 通过 | 独立缓存消费 `0.1.0-issue02.review9`；新图片 API 和安装后的 CLI 两方向，加既有 DOCX/PDF/SVG/签章 E2E |
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

复现：先按根 AGENTS 设置 writable `DOTNET_CLI_HOME`、跳过首启/遥测、显式 `NUGET_PACKAGES`，运行全套单节点测试和 `scripts/run-converter-package-e2e.sh 0.1.0-issue02.review9`。图片样例生成测试入口：`OFDRW_IMAGE_EVIDENCE=<directory> dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests -c Release --filter FullyQualifiedName~SaveReviewEvidence`（附 AGENTS 构建参数）。CLI导出基准：`ofd-to-image generated-docx-{native|default}.ofd <page.png> --pages {1|2} --ppm 4`。

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

修复后全套189/189。产物与11包消费证据已在`ce23288`重新生成并通过；9张新PNG/JPEG逐页重新查看，几何与文字/样式/表格检查未见新缺陷。Preview仍受锁屏阻塞；尚不宣称完整视觉门或最新重审闭合。

## 第二轮复审修复

`2dd3c3e`的五项CI通过且Windows图片API/CLI实跑通过。两类bot的本轮结果均已读完：Codex提出重复嵌套签章展开与微小几何精度，Cursor提出非二进制整除浮点超出页轴；均纳入同轮。

- 图片严格路径按载荷内容缓存嵌套包与共享字体上下文，绘制也复用同一PDF form/bitmap。外包加唯一嵌套包累计占用展开载荷bytes、已物化非目录文件条目及页面预算；每次读前只给剩余额度。各ZIP原始条目数仍由loader逐包限制，不把物化条目误称累计ZIP目录项。单次转换最多1000个选中候选StampAnnot，无效边界/无载荷也计数，超限在载荷提取前失败。
- canonical包条目的载荷提取缓存成功与空结果，直接payload复用原始数组；ASN.1扫描保留一个最大候选，避免候选列表放大。提取副本、XML/native/font/PDF开销不属于展开载荷budget，不声明严格RSS上限。
- 每轴缩放结果clamp到已验证页轴，145×145／0.01ppm／10000mm页与39×39／0.001ppm／210mm页回归确认无ULP越界/负原点。
- 现有writer为0.001mm精度。最终导入页及图像每轴小于0.001mm明确拒绝；固定页、自然页、高ppm及缩小后的极薄图都覆盖，失败不生成零尺寸图元。

新增重复签章单份展开budget成功/单PDF form、不同载荷累积bytes/entries/pages失败、共享ASN.1载荷对象身份、无效候选限额与几何回归。主套200/200；Sol Low独立复核未发现新确定性缺陷。仅外观fixture不代表密码学签名有效。

第二轮实际产物均从`d8ff853`重新生成，11/11本地包消费已再次通过。10张当前PNG/JPEG逐页重新查看：两页文字/图片/路径、JPEG第2页、两页图片居中往返、显式Native/default各两页和共享嵌套外观一页均未见新缺陷。新外观样例6个红色10×10mm方形位于120×80mm页面同一行，5mm起点、18mm间距，全部可见且无裁切/异常叠加；该fixture只测外观，不含可验证密码学签名。Preview另外需打开本次`shared-seal.ofd → shared-seal.pdf`第1页。新bundle保留25个确切载荷及日志，源码/模式/字节数/SHA-256/待Preview页码见manifest。

## 第三轮复审修复

`adaa7c8`两类bot结果均已读取。Codex提出不同位图在整页缓存中累积，Cursor提出ASN.1长度相加溢出；两项均修复。

- 位图缓存改为单槽，换载荷时先释放旧对象再创建下一份，只有相邻相同载荷复用；不会因数百个不同载荷而保留所有解码图片。资源所有权回归用500个不同载荷确认live/peak均不超过1。
- ASN.1标签/长度使用减法检查剩余切片，识别header和分配副本前校验`offset/length`确实在原数组内。`04 84 7F FF FF FF`及追加ZIPheader版本均抛`InvalidDataException`，输出保持；legacy公开PDF路径仍忽略不可读外观并导出正文。
- 全套203/203及Sol Low独立复核通过。新的产物与11包消费验证随后记录；Preview仍未完成。

第三轮25个实际产物均从`f7dcae8`重新生成，11/11本地包消费通过；当次10张PNG/JPEG逐页重新查看，检查范围与上述相同，未见新缺陷。最新源码、日志、字节数、哈希由当前manifest给出。Preview仍未完成。

## 第四轮复审修复

`8383723`两类bot结果均读取：畸形ASN.1会清空PDF已接受外观，以及`XImage.Dispose`不能释放底层像素，均已修复。

依据锁定[PdfSharpCore 1.3.67源码](https://github.com/ststeiger/PdfSharpCore/tree/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd)：`PdfImageTable`持有`XImage`，默认ImageSharp source持有解码图像；仅Dispose包装对象不足。严格图片路径改用只存编码字节/尺寸的`IImageSource`，PDF图像实现期间局部解码/编码并Dispose实际ImageSharp图像，文档表不再持有它的像素缓冲；不改全局image source或legacy PDF路径。真实PdfDocument/ImageTable回归各绘制12个不同PNG/JPEG，确认DrawImage后实际像素已ObjectDisposed、尺寸仍可用、最终PDF可保存。

ASN扫描区分strict与legacy：严格图片路径继续分配前拒绝坏长度；legacy遇到畸形候选返回并保留已经找到的合法候选，不抹掉其他签章。合法JPEG后另一坏签章和同ASNrecord先合法OCTET后坏trailer均有实际PDF红色像素回归。

全套207/207；Sol Low独立核查未见新确定性缺陷。单个多帧厂商签章位图的解码属于既有预览边界；工作缓冲估算不声明进程RSS硬上限。最新11包/产物记录随后更新；Preview仍未完成。

第四轮25个实际产物均从`4312bfb`重新生成，11/11本地包消费再次通过；当次10张PNG/JPEG逐页重新查看，范围与上述相同，无新缺陷。最新207项回归日志与全部载荷哈希在当前bundle/manifest。Preview仍未完成。

## 第五轮复审修复

`103ccd1`全部CI通过，两类bot均到齐。Cursor确认上一轮两项修复无新缺陷；Codex指出严格模式对非InvalidData解析/位图异常仍容错，两项均修复。

严格图片导出下，签章元数据、嵌套OFD的坏XML/缺失root/无效首页面，以及位图解码/绘制的非取消错误，都在发布前抛出`InvalidDataException`；取消保留原异常，内存耗尽不吞掉。legacy PDF仍容错。新增合法JPEG+坏第二外观的XML/缺失root/截断位图三例，严格输出保sentinel，legacy实际PDF红像素仍在。

全套210/210；最新包消费/页面产物随后记录。Preview仍未完成。

第五轮25个实际产物均从`1164218`重新生成，11/11本地包消费再次通过；当次10张PNG/JPEG逐页重新查看，检查范围同上，无新缺陷。最新210项日志/载荷哈希在当前bundle/manifest。Preview仍未完成。

## 第六轮复审修复

`47a9765`两类bot结果均已读。Codex/Cursor共同指出无可用载荷的早退，Cursor另指出嵌套首页面NaN/Infinity未被旧`<=0`挡住；同轮修复。

所选StampAnnot缺失/空/不支持载荷在strict路径直接InvalidData；嵌套首页面须有限且正，NaN/Infinity按strict失败、legacy跳过。五例新增回归分别检查严格输出保sentinel、legacy合法JPEG红像素仍在。全套215/215；新产物/包消费稍后记录，Preview仍未完成。

第六轮25个实际产物均从`65c40ed`重新生成，11/11本地包消费再次通过；当次10张PNG/JPEG逐页重新查看，范围同上，无新缺陷。最新215项日志/全部哈希在当前bundle/manifest。Preview仍未完成。

## 与03后续集成约定

共享触点：`src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs`（加载后渲染入口、签章准备/绘制）、`Internal/OfdSignatureAppearanceReader.cs`、`src/Ofdrw.Net.Cli/Program.cs`（partial类/命令路由/帮助）、`.github/workflows/ci.yml`、包消费Program/脚本。02不改变文字CTM、假斜体、注释多边形裁剪或PageBlock语义，不修改03分支。

内部接口：`ConvertPackageAsync(OfdDocumentPackage, Stream, IReadOnlyList<int>?, CancellationToken, bool strictAppearanceBudgets = false, int? maximumSignatureAppearances = null)`。公开PDF `ConvertAsync`保留default false；只有图片导出传true及签章候选上限。严格限制/缓存/错误传播应继续局限该入口，保留03的共享绘制实现及普通PDF兼容行为。`OfdSignatureAppearanceReader.Read`新增选中页ID、可空候选上限、CT；默认调用（包括SVG）保持容错候选语义。

## 第七轮复审修复

`542f798`两类bot结果均已读取：Codex指出选中StampAnnot的缺失/不可解析/零或负尺寸Boundary仍被静默跳过；Cursor指出Infinity和浮点溢出`1e309`还能进入PDF坐标。

严格图片导出现在对四个Boundary轴统一拒绝NaN/Infinity，并要求宽高为正；缺失或不可解析同样以`InvalidDataException`失败。legacy PDF跳过这些无效边界，保留其它合法外观。使用兼容netstandard2.0的NaN/Infinity检查。20个边界回归分别验证strict输出保持sentinel、legacy有效JPEG仍有红色像素；同记录两无效stamp的数量限制仍在边界解析前失败。全套235/235，Sol Low独立定向21/21通过；本轮新包消费和产物随后记录。Preview再次确认锁定，仍未完成。

第七轮25个实际产物均从`80c33c8`重新生成，11/11本地包消费通过；当次10张PNG/JPEG逐页重新查看，文字/样式/几何/表格/分页及重复外观检查范围同上，未见新缺陷。235项回归日志与全部载荷哈希在当前bundle/manifest。Preview仍未完成。

## 第八轮复审修复

`4840666`两类bot结果均已读取。Codex未发现主要问题；Cursor确认上一轮Boundary修复，并指出有限毫米值/页原点仍可在PDF转换中变成非有限操作数，以及正尺寸可被`0.####`写成零。

新增局部`PdfOperandGeometry`与锁定PdfSharpCore 1.3.67的实际操作数/格式保持一致：Boundary、相对页面原点的stamp位置、边缘加法须在`mm * 72 / 25.4`后有限；位图尺寸及嵌套form缩放因子须为可按`0.####`写出的正值。strict外页原点在DrawPage前检查；嵌套首页面在创建form前检查。strict失败保持caller sentinel；legacy跳过不合法stamp或nested外观。普通公开PDF的页正文逻辑不扩展。锁定源码依据为[DrawImage实现](https://github.com/ststeiger/PdfSharpCore/blob/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Drawing.Pdf/XGraphicsPdfRenderer.cs#L604)。

新增22项，覆盖有限`1e308`/超大边缘、极小正尺寸、NaN/Inf/溢出页原点、正常(2,3)mm页原点的实际红像素、相对位移减法溢出、嵌套form比例溢出/写零；legacy回归除有效JPEG红像素外，还解压PDF内容流确认无NaN/Infinity。全套257/257通过；本轮新包和产物随后记录。Preview仍未完成。

第八轮25个实际产物均从`eed7d6d`重新生成，11/11本地包消费及Sol Low独立48/48通过；当次10张PNG/JPEG逐页重新查看，检查范围同上，未见新缺陷。257项回归日志与全部哈希在当前bundle/manifest；Preview仍未完成。

## 第九轮复审修复

`46a0029`两类bot结果均读取。Codex指出已引用但缺失的签章列表/每签章XML仍早退；Cursor指出XForm BBox采用`0.###`，区别于cm操作数的`0.####`。两项修复：strict缺失元数据/空列表引用/缺失或空BaseLoc均InvalidData；nested第一页XUnit.Point宽高须按`0.###`写出正值，legacy继续跳过不合法外观。新增5种metadata和2种BBox宽高用例，strict保sentinel，legacy正文红path保持且无零BBox的Form字典。全套264/264通过；本轮包和产物稍后记录。

Cursor另质疑`-2e306`与`2e306`的相对溢出用例。保留该用例：实际函数按`mm * 72 / 25.4`逐步计算，`4e306 * 72`为Infinity，即使重排为`4e306 * (72 / 25.4)`会有限，也不能代替当前实现。Sol Low独立已编译CLI探针实测exit 1、placement错误和13-byte哨兵保留；264全套回归同样通过，旧head三OS图片CI均通过。本条以实际运行证据回应，未按误判修改样例。Preview仍未完成。

第九轮25个实际产物均从`c45e797`重新生成，11/11本地包消费和Sol Low独立55/55通过；当次10张PNG/JPEG逐页重新查看，检查范围同上，无新缺陷。264项回归日志/全部哈希在当前bundle/manifest；Preview仍未完成。

## 第十轮复审修复

`11152fe`两类bot到齐，六项CI全部通过。Cursor确认此前metadata/BBox修复且纠正原点算术质疑，无新缺陷；Codex提出缺失/空白PageRef在strict前被过滤。Astra补充设计复核也指出此项及已引用XML错误根节点导致零匹配；Sol Low的三份已构建CLI探针复现了错误列表根、错误Signature根及缺PageRef都会exit0覆盖哨兵。

strict现在按LocalName校验列表/Signature根（仍允许XML namespace），StampAnnot先校验非空PageRef再筛选明确非选中页，避免未知归属被静默省略。legacy不增加这些strict校验，正文及既有容错保持。新增错误根两种及PageRef缺失/空/空白三种，用实际caller sentinel及legacy正文红path断言。全套269/269通过；本轮新包和产物稍后记录。Preview仍未完成。
