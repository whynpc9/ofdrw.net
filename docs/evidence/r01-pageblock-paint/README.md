# R01 PageBlock、文字与颜色验收

日期：2026-10-23（Asia/Shanghai）；分支 `codex/r01-pageblock-text-color`；[PR #29](https://github.com/whynpc9/ofdrw.net/pull/29)。**产品源码** `3acd3f1aa093730ddf6bff39a431c82f76ede3ee`，已整合主线 `aac53f2875b3bdbe190364fbacdec5e0f0f59888`。环境：macOS Arm64、.NET SDK 10.0.401。本轮输入、输出、日志与 PNG 位于已忽略的 `artifacts/roadmap-v12/R01/`，文件 SHA-256 见 [hashes.txt](hashes.txt)。

## 已验证的行为边界

- Reader 在普通页及模板页按绘制顺序投影 PageBlock 的路径、图片、文字与复合资源，保留原始 XML 与逻辑文字；PDF、SVG、HTML 使用同一预检拒绝不支持的绘图，而不静默漏绘。官方 2009/2016 和无命名空间的受支持绘图可以在 PageBlock 内混用；普通顶层复合对象仍遵循原有命名空间契约。外来命名空间的 Page、Content、Layer、绘图与 Clips 均保持不透明或拒绝渲染。外来 Layer 的 0、1、多子元素在 Reader 中按整层 XML 与来源槽读取；带扩展属性、处理指令或非空文本的 Page/Content/Area/Layer 包络，以及 Content 中未建模的直接子项，在 PageBlock 入口按原始槽保留并拒绝渲染。无普通 Content、仅在外来 Page 包裹层中包含 PageBlock 的普通页和模板页也保留整页原始 XML，三个出口拒绝空白漏绘。Reader 已认证子项的 PageBlock 若只在包络上带外来属性、处理指令或非空文本，仍可提取子文字，三个出口凭绘制错误标记拒绝；来源被修改或不透明的 PageBlock 保持直接拒绝。Writer 继续拒绝无法安全重写的未知结构，包括 PageBlock 所在页面的额外或外来 Content 同名容器。
- 字体、媒体、复合资源、DrawParam 和颜色资源按正整数 ID 解析；`030` 与 `30` 数值等价，原 XML 拼写不变。模板字体和复合资源只从文档与模板自己的资源解析，不借用调用页 PageRes；本页 PageBlock 可解析本页 PageRes 的复合资源，另一页和普通顶层复合对象不借用这张只读映射；Split 克隆 PageBlock 时保留本页/模板只读复合资源闭包，Writer 保留原 PageRes XML。明确非 Image 的媒体不能供 ImageObject 使用，省略 Type/Format 的旧图片仍可用；显式未知 Format、与实际载荷不符的声明或结构损坏的 PNG/JPEG 在直接、复合 PageBlock 图片写出前统一拒绝。重复或歧义 ID、未知绘图、资源闭包无法证明及显式 CGTransform 字形替换在输出前拒绝。
- PageBlock 引用的字体若显式声明 `FontFile`，必须实际读到非空载荷；无 `FontFile` 的名称字体维持原有回退。`Fonts/Font/FontFile`、`MultiMedias/MultiMedia/MediaFile`、`CompositeGraphicUnits` 及 `DrawParams/ColorSpaces` 的容器和定义按完整结构验收；Res 根节点上的外来属性/子项、处理指令或非空文本同样不能被丢弃后继续解析 PageBlock 资源。同命名空间的既有 VendorExtension 保持回读兼容。媒体的严格来源判定覆盖直接与复合资源中的 PageBlock 图片；普通复合对象和批注的既有外来媒体属性回读行为仍由独立回归确认。
- PageBlock 图片的三出口共同准入范围是有界静态 PNG 与单扫描基线 JPEG；BMP/TIFF、交错或索引 PNG、带未准入元数据块的 PNG，以及渐进式 JPEG 不进入该共同范围。有效 BMP/TIFF 的定向正反例确认：PageBlock 的 PDF/SVG/HTML 在写出前拒绝，而普通顶层 PDF 仍成功解码；PDF 对普通顶层图元的既有图片解码路径不由此改变。声明中的 BMP/TIFF 仍被识别为既有资源格式，但 PageBlock 会在渲染前标为未准入；HTML 的静态预览本来也不接受这些格式。此处记录有限功能边界，不把某个出口能独自解码提升为三出口支持。
- 非 `Page` 的内容根节点在 Reader 中按一个不透明来源槽绑定，Writer 继续拒绝不安全重写；无 ID 的受支持 Layer 在写回时分配正整数 ID，两个同类型无 ID 图层的读写回读也通过。
- PageBlock 后续 TextCode 缺失的 X/Y 独立继承前一个 TextCode 的**起点**；首个缺轴则保留逻辑文字并拒绝渲染。直接、复合和嵌套复合资源中的首个 TextCode 缺 X/Y，以及模板中的相同情形，均保持源 XML 与可提取文字，三个导出在写出前拒绝。继承坐标下的 OFD 转义仍解码；首个 TextCode 的 X/Y 显式 NaN 或后续 TextCode 继承 NaN 时，逻辑提取也解码转义，PDF/SVG/HTML 继续在写出前拒绝。复合资源文字缺少 Font 或引用未声明字体时，Reader 保留 typed TextCode 和转义解码，三个导出在写出前拒绝。typed run 的实际编辑优先于旧源包络。普通顶层文字既有 typed 默认值保持不变。局部粗斜体、颜色、HScale、DrawParam 和字体选择由正反例覆盖。
- PublicRes 与 DocumentRes 指向同一物理 Res 时（包括 `./` 规范化别名），按解析后的包内路径只索引一次字体、媒体、画笔与复合资源；不同物理 Res 的同 ID 定义仍由来源身份检查拒绝。未引用媒体只做常量大小的格式识别；实际被 PageBlock 直接或复合图元引用的 PNG/JPEG 才执行有界结构准入，并按载荷缓存结果。压缩字节与解码像素工作量累计计费；未引用的损坏媒体不会阻断已引用图片的导出。PageBlock 文字含未知属性或子项时，原始 XML 和逻辑文字保持可用，导出仍由共享预检拒绝。有效文字旁边的 Path/Image 不支持样式、复合资源包络的外来属性/处理指令/非空文本、不透明复合子项与不可解析的复合放置也保留邻近 TextCode 提取，以绘制错误标记让三出口写出前拒绝；循环、预算、取消与来源不一致仍维持原拒绝边界。首个复合资源 TextCode 缺轴时也走逻辑投影，保留 OFD 转义解码。
- 路径和裁切检查命令完整性、弧线半径与字面标志位，并要求至少一个可绘线/曲线/弧段；只有 Move 或 Move+Close、开头/连续/尾随逗号的路径及裁切在三个渲染器中一致拒绝。合法数字间逗号仍可绘制。无命名空间 PageBlock 的普通 Clips 与两个官方命名空间按同一封闭契约处理；外来裁切继续拒绝。复合对象放置只准有一个同命名空间 `Clips` 容器；重复容器和扩展命名空间的同名节点均保留资源内逻辑文字，三个出口在写出前拒绝。复合对象放置或其文字叶节点的畸形裁切保留资源 TextCode 提取，但渲染预检拒绝。PageBlock 的文字、路径和图片叶节点若有两个同命名空间 `Clips` 容器，也在三个出口写出前拒绝，文字逻辑提取仍可用。颜色调色板所有条目先校验，合法 Value 优先于 Index。原始 XML 注释、空白和裁切布尔词元按受支持契约保留。

## 功能测试与本地包消费

- 源码 `3acd3f1` 的完整 Release 解决方案：**4022 passed、0 failed、0 skipped**，11 个测试项目；R01 PageBlock **412/412**。单节点构建并使用仓库独占锁；日志为 `p95-full-release.log`、`p95-focused.log`。
- 本地版本 `0.1.0-r01.p95` 的 **12 个基础 SDK 包和 1 个可选 SES 包**通过来源核对及独立消费者 E2E。`generated-layout.docx` 的显式 native 与默认 DOCX→OFD→PDF 各生成 V1/V4 测试签章，原文完整保留；签章前后逐页 PNG 一致。`p95-ses/files/functional.json` 与 `rendered-pages.json` 保存机器检查。Core、Packaging、Reader 三个精确本地包的 R18 附件清理消费者也通过，物理删除后回读为空。包未发布；默认密码学验证状态为 `Unsupported`，测试专用自签路径通过。
- 整合已实际合并的 R20 主线 `aac53f2` 后，精确本地 Core/Packaging 包的归档预检消费者完成 4 个公开 API 样例；`p95-preflight.log`、`p95-preflight/output/manifest.json` 与 `archival-structural.json` 保存运行及规则结果。R20 新增的 `AC-013`–`AC-015` 是只读、有界的违规检测；未发现违规不等于标准通过，缺乏完整来源、资源闭包与明文证明时保持 `Unknown`。本次验证不宣称档案合规认证。

## 自动渲染与 PNG 视觉复查

本轮从产品源码生成 **30 份 PageBlock PDF、31 页**，逐页导出 144 DPI PNG。24 张沿用样例与先前逐页目视检查的对应当前分支图像逐字节相同；另 4 张与 `p80` 对应 PNG 逐字节相同，`p80` 记录了与 `p77` 目视复查页的比对；两张同物理 Res 与规范化别名路径的红色方块和上轮直接目视页逐字节相同；本轮新增的复合对象单例裁切页直接目视检查为左半蓝色矩形、完整可见的黑色文字；省略 Format 的 JPEG 页与本轮显式 JPEG 页逐字节相同，符合样例预期，未见异常裁切、重叠、重影或空白页。实际检查范围如下：

逐页比对文件、页码、SHA-256、先前目视检查的对照路径、后四页的 `p80` 对照、新裁切页的直接目视记录与省略 Format 页的同轮对照见 `p95-png-qa.json`；[hashes.txt](hashes.txt) 所列的 186 个关键文件已由 `shasum -a 256 -c` 全部核对。

| 样例 | 页数与检查要点 |
| --- | --- |
| 13 个数值字体/媒体/复合/绘制资源、弧线与裁切基础样例 | 各一页：文字、红蓝方块、弧线、三角裁切及模板图片位置符合预期。 |
| 4 个混合 Layer/PageBlock 命名空间、2 个混合复合对象命名空间 | 各一页：`mixed block text`、`mixed composite text` 完整可见，无重复文字。 |
| `template-own-font` | 两页：共享模板文字在两个调用页均清晰可见，页面右下图形不遮挡正文。 |
| `textcode-inherited-start`、`textcode-inherited-escape` | 各一页：坐标继承样例故意使部分文字共起点；转义样例的 `escaped text`、`next line` 分行清楚，未显示字面 `\\0020`。 |
| `unqualified-pageblock-clip` | 一页：无命名空间 PageBlock 的蓝色三角裁切正确；外来 Clips 负例在输出前拒绝。 |
| `page-local-composite`、`template-local-composite` | 各一页：本页/本模板 PageRes 中的复合资源呈现相同蓝色三角形，边界完整。 |
| `jpeg-media-pageblock`、`omitted-format-jpeg-pageblock` | 各一页：显式与省略 Format 的 JPEG 均显示红色方块，逐页 PNG 一致；错配声明及损坏载荷由功能测试拒绝。 |
| `shared-res-drawparam`、`alias-res-colorspace` | 各一页：同一路径与 `./` 别名的资源表均画出预期红色方块；与上一轮直接目视页逐字节相同，未见裁切、重影或空白。不同物理 Res 的重复 ID 由来源身份检查拒绝。 |
| `composite-single-clips` | 一页：蓝色矩形只显示裁切范围的左半，黑色 `clip sentinel` 可见；同命名空间重复 Clips 和扩展命名空间同名节点在输出前被拒绝，逻辑文字仍可提取。 |

同一 `generated-layout.docx` 的显式 native OFD 和默认 OFD 再各自导出 PDF；V1/V4 四条本地消费链路各两页，共 8 张 96 DPI PNG，与签章前基线逐页一致。含两条基线的 12 张 PNG 与已目视检查的上一轮对应页逐字节相同。此前 `native-v1`、`default-v4` 各两页的目视记录覆盖中英文、斜体、蓝色表格底色和边框、第二页红色粗体、日期与分页；未见缺字、裁切、重叠、重影或异常空白页。查看链路为 **DOCX→OFD→PDF→PNG**；直接 DOCX→PDF 未充当 native 验收。

## macOS Preview 验收

已以产品源码 `3acd3f1` 的 32 份 PDF、35 页及其 SHA-256 提交 `gui-request-R01-20261023T1509Z.json` 独占 GUI 申请；**Preview 逐页验收尚未完成**，目前等待桌面租约安全释放。完成后记录实际查看链路、页面范围、结论与关窗结果。本报告只覆盖上述确定性样例；不推断任意复杂 DOCX 的完整保真、外部 OFD 阅读器互操作、V5/GBT 端到端互操作或签名法律效力。
