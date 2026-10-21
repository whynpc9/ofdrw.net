# R01 PageBlock、文字与颜色验收

日期：2026-10-21。实现基线 `35976dbee1f58eef5fe19f2750746a97d523cfec`，随后整合 R03 主线 `91e389f90e8368b13855af5320626209661d3e25`；分支 `codex/r01-pageblock-text-color`，.NET SDK 10.0.401，macOS。提交与 PR head 以 PR 记录为准。本目录仅保留报告和 [SHA-256 清单](hashes.txt)；本次生成的 OFD、PDF、SVG、HTML、PNG 和日志保存在已忽略的 `artifacts/roadmap-v12/R01/`。

## 行为与功能测试

- Reader 对普通页和模板的嵌套 `PageBlock` 建立只读投影，保留原始 XML 作为保存来源。投影按源顺序包含 Path、Text、Image；根节点有未知绘图语义或投影失效时，PDF/SVG/HTML 明确拒绝。
- 有界解析 DrawParam 相对继承、8-bit GRAY/RGB、无 ICC 的简单 CMYK 和调色板。缺失或歧义引用、渐变/图案及不支持的文字语义拒绝；文字 Fill、Stroke、HScale 在 PDF/SVG/HTML 中分别呈现。签章内嵌 OFD 的已识别颜色失败会向外传播，不再静默略去章面。PR 审查后补上缺失图片载荷拒绝，并确保颜色规范化的物化预算异常立即向外传播。
- `dotnet test Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false`：11 个项目共 **3047 passed、13 skipped、0 failed**。定向 R01 12/12，DOCX native/default 视觉控制 2/2 通过。使用共享独占构建锁及仓库要求的 `DOTNET_CLI_HOME`、首次运行/遥测环境变量。
- 本地打包版 `0.1.0-r01.review3`：11 个包校验及独立消费者 E2E 通过，包括 CLI、DOCX、PDF、图片、上游样本。该版本仅为本地消费验证，未发布。既有 ImageSharp 2.1.13 的 `NU1902/NU1903` 告警仍在，本票未改依赖。

## 整合 R03 主线后的复验

- R03 命名区域在 Layout 新增 API、持久化索引和填充；与 R01 修改文件没有 Git 冲突。联合源码全套 11 个项目 **3056 passed、13 skipped、0 failed**；R01 定向 12/12，DOCX native/default 控制 2/2。R03 命名区域 E2E 重新打开 OFD，确认填充状态及区域外文本坐标不变，并导出空白/填充各 2 页 PDF。
- 旧 PR head 的 CI `package-e2e` 在 document-tools 样本遇到 Reader 保留的注释索引元数据，误报 `Opaque OFD drawing 'Annotations'`。修复仅放行 Reader 生成的 `UnmodeledAnnotations` 元数据封装；`UnsupportedAnnotationAppearance` 和普通未知绘图仍拒绝。修复后的本地 document-tools API/CLI 矩阵验证 **36 个样本、55 页 PDF**，完整通过。
- 联合源码重新打包为仅限本地消费的 `0.1.0-r01.r03`：11 个包校验和独立消费者 E2E 均通过。R01 五个新生成 PDF 的 144 DPI PNG、SVG 和 HTML 源与修复前已逐页检查的页面逐字节一致；DOCX native/default 新生成 PDF 的各两页 PNG 也逐字节一致。
- 额外逐页查看命名区域空白/填充共 4 张本轮生成的 PNG：中英文、局部粗体/斜体/颜色、表格底色与边框、填充文字和双色图片可见；填充后区域外文字位置与分页稳定，未见缺字、重叠、裁切或异常空白页。对应产物路径及哈希列于清单。macOS Preview 对本轮组合产物的检查结果在下节单独记录。

## 后续审查修复与联合验证

- 无显式 `StrokeColor` 的轮廓文字按 OFD 默认黑色描边；未引用 ColorSpace 的注释灰度单分量按灰色处理，仍拒绝有未解析颜色空间的注释。另修复顶层路径保存重开：Writer 将已解析的颜色写成 RGB 三分量时删除旧 `ColorSpace`/`Index`，避免 GRAY/CMYK 引用把 RGB 数值误读。签章内嵌 OFD 若缺少图片载荷，已识别的拒绝向默认 PDF 转换传播，防止整章其他路径和文字静默消失。
- 回归分别覆盖纯描边文字的 PDF 像素和 SVG/HTML、注释灰度三出口及未解析颜色拒绝、顶层 GRAY/CMYK 路径保存重开后的三出口、带有兄弟路径/文字的缺图签章失败。最终联合源码的 11 个项目 **3061 passed、13 skipped、0 failed**；两条最后新增回归定向 2/2。文档工具 API/CLI 矩阵在修复后的联合源码验证 **36 个样本、55 页 PDF**。完整日志在 `artifacts/roadmap-v12/R01/`。
- 本次生成的普通/模板 PageBlock、调色板/DrawParam、文字描边/横向缩放、嵌套图片、纯描边文字和灰度注释页面已在 macOS Preview 检查；DOCX 显式 native 和默认模式各 2 页，以及 R03 命名区域空白/填充各 2 页也逐页检查。最后重生成的顶层 GRAY/CMYK 保存重开页在 Preview 第 1/1 页显示预期灰色与红色方块。所有对应文档窗口均已关闭。最终生成的既有样例页面 PNG 与这些 Preview 检查时的 PNG 逐字节一致。
- 本地版本 `0.1.0-r01.r03d` 的 11 个包校验和独立消费者 E2E 通过；CLI、DOCX、PDF、图片及上游视觉样本均由打包结果消费。仅用于本地验证，未发布。

## Split 克隆审查修复

- Codex 复审指出：Reader 建立的 `PageBlock` 只读子元素投影在 `OfdDocumentSplitter.Split` 克隆时丢失，导致返回包的文字提取拒绝该嵌套绘图。克隆现仅在投影仍与原始 XML 一致时递归复制子元素，保留克隆后 XML 的权威性和独立对象身份；原本已失效的投影继续失效。
- 新增普通页及模板页回归：Split 后提取文本、写出重读、PDF/SVG/HTML 三出口均须保留 `split nested`。定向测试 17/17；联合源码全套 11 个项目 **3063 passed、13 skipped、0 failed**。本地 `0.1.0-r01.r03e` 的 11 个包和独立消费 E2E 通过；文档工具 API/CLI 矩阵验证 **36 个样本、55 页 PDF**。
- 两份 Split 派生 OFD 各导出 1 页 PDF、PNG、SVG、HTML。本轮 PDF 在 macOS Preview 逐页检查：普通页和模板页均显示相同的 `split nested`，无裁切、重影或空白页；检查后已关闭窗口。两张本轮 PNG 字节一致，路径和哈希见清单。

## 隐藏未知绘图审查修复

- 后续 Codex 审查指出 Reader 的未知绘图以 `OfdRawElement.Xml` 保存，原可见性检查只读取已建模元素的 `SourceXml`，导致明确 `Visible="false"` 的非绘制对象也被预检拒绝。现对原始绘图读取其原始 XML 可见性；可见或未声明可见性的未知绘图仍明确拒绝。
- 回归覆盖普通页和模板 `PageBlock` 中隐藏未知绘图与可见兄弟文字：PDF/SVG/HTML 均保留 `visible sibling`；显式 `Visible="true"` 和未声明可见性的未知绘图仍被三出口拒绝。最终联合源码全套 11 个项目 **3066 passed、13 skipped、0 failed**；本地 `0.1.0-r01.r03f` 的 11 个包及独立消费 E2E、文档工具 **36 个样本、55 页 PDF** 通过。本轮各导出 1 页 PDF/PNG，两份 PDF 在 macOS Preview 逐页检查，均只有可见兄弟文字，无遮挡、裁切或异常空白页；窗口已关闭。

## 外来命名空间审查修复

- Reader 对 `TextObject`、`PathObject`、`ImageObject` 的类型分派现仅接受两种 OFD 标准命名空间和无命名空间兼容形式；同名扩展节点保留为原始未知绘图，不由 PDF/SVG 误当受支持对象呈现。普通页和模板的三类扩展节点共 6 个回归，逐一断言 Reader 保留原始类型，PDF/SVG/HTML 均明确拒绝。R01 定向测试 26/26。最初只接受与父层相同命名空间导致 PDF vector 回归失败；纳入历史兼容的另一个标准命名空间后，PDF vector 测试 315/315 通过。
- 最终联合源码全套 11 个项目 **3072 passed、13 skipped、0 failed**；本地 `0.1.0-r01.r03g` 的 11 个包及独立消费 E2E、文档工具 **36 个样本、55 页 PDF** 通过。此次源码重新生成普通页、模板页及隐藏未知绘图两种页的 PDF/PNG；各 1 页已在 macOS Preview 检查并关闭窗口。新 PNG 与先前检查过的对应页面逐字节一致：蓝色方块与 `nested text`，以及仅显示 `visible sibling` 的隐藏对象页面，均未见缺字、裁切、重叠或空白页。完整产物哈希见清单。

## 最后两项审查修复与 R12 主线整合

- PageBlock 内已识别的 Text、Path、Image 叶子现在也经过原始绘图语义准入；未知子节点及 `ImageObject Substitution` 在 PDF/SVG/HTML 三出口明确拒绝。普通页/模板页、文字/路径未知子节点与图片替代引用均有回归。此检查不改变 PageBlock 叶子的坐标。
- 带 `Stroke=true` 的文字若携带 `Cap`、`Join`、`MiterLimit`、`DashPattern`、`DashOffset`，现明确拒绝，避免把这些描边属性默认为标准实线；覆盖五种直接属性和 DrawParam 继承。R01 定向回归 **37/37**。
- 合入 R12 已并入主线的提交 `a50892cf9ecc0177d243b24a89419809057a8115`，本分支合并提交 `32a616e4261c95a599819b54f9c55090b2273aa3`。R12 可选 `Ofdrw.Net.Signatures.SesInterop` 保持独立：11 个基础 SDK 包，外加一个可选 SES 包；CLI 的默认依赖闭包不含 SES。签章的签发/验证经 `IOfdSignatureProvider`、`IOfdSignedValueVerifier` 接入 `OfdSignatureService`、`OfdSignatureVerifier`。
- 整合及上述修复后的全量解决方案 **3087 passed、13 skipped、0 failed**（11 个测试项目）。本地 `0.1.0-r01.r12a` 的 12 包来源清单、11 基础包消费和可选 SES 独立消费通过；显式 native/默认 DOCX→OFD 各生成 V1/V4 签章样本，四份签章产物的两页 PDF 像素均与各自签章前基线逐页一致。默认 CLI 对不支持的验证返回 2；本地样本不构成外部阅读器互操作或 V5/GBT 端到端验收。
- 本次签章消费生成的 `native-v1.pdf` 与 `default-v4.pdf` 在 macOS Preview 各检查两页：标题、英文斜体、中文正文、蓝色表格底色/边框、红色局部粗体与稳定分页可见，未见裁切、缺字、重影或空白页。其余两份签章 PDF 与基线在 96 DPI 两页 PNG 上逐字节一致。检查后的文档窗口已关闭；页面、清单、日志及哈希均保存在已忽略的 `artifacts/roadmap-v12/R01/r12-combined-ses/`。

## 签章 PageBlock 绘制失败标记

- 后续复审指出：签章内嵌 OFD 的 PageBlock 子路径若颜色空间解析失败，投影层会先以普通 `NotSupportedException` 拒绝原始颜色引用，签章呈现的容错分支因此可能吞掉错误并丢弃整个章面。投影现先检查 Reader 的 `RenderPaintError`，使用与绘制预检相同的失败标记抛出；未解析颜色空间的签章 PageBlock 新增端到端回归，要求 PDF 转换失败且不留下部分输出。
- 修复后 R01 定向测试 **38/38**；11 个项目全套 **3088 passed、13 skipped、0 failed**。本地 `0.1.0-r01.r12b` 的 12 包来源、11 基础包及可选 SES 独立消费、显式 native/默认 DOCX→OFD 与四份 V1/V4 签章样例均通过。六份基线/签章 PDF 的各两页 96 DPI PNG 与前一轮逐字节一致，故有效页面沿用前一轮的 Preview 逐页验收；新修复针对失败路径，未将这些样本扩展为外部阅读器验收。

## 最终绘图准入复审

- `PageBlock` 的 `TextCode xml:space="preserve"` 与普通层文字一样获准，普通页和模板页的 PDF/SVG/HTML 均可呈现双空格文本。签章内嵌 OFD 的可见未知绘图或被 PageBlock 准入拒绝的子对象，在默认 PDF 转换中现会明确失败，不会吞掉整个章面；严格签章路径仍把不支持的裁剪包装为原有的 `InvalidDataException`。第一次扩大异常传播范围时，该严格路径测试失败；缩小到默认模式后，关联测试 **44/44** 通过。
- 资源文件顶层 `Res` 只接纳两种官方 OFD 命名空间和无命名空间兼容形式；外来命名空间的 `ColorSpace`/`DrawParam` 不再被误当 OFD 绘制参数。索引调色板的 `CV` 逐项拒绝扩展属性、子节点和处理指令。各自独立回归及前述严格签章测试 **49/49** 通过。
- 最终联合源码 11 个项目 **3097 passed、13 skipped、0 failed**。本地 `0.1.0-r01.r12d` 的 11 基础包和可选 SES 包来源、独立消费及四份显式 native/默认 V1/V4 签章样例通过；六份基线/签章 PDF 的各两页 96 DPI PNG 互相一致。`resource-final-ses/files/native-v1.pdf` 和 `default-v4.pdf` 为**本轮重新生成**的产物，已在 macOS Preview 分别重新打开并检查第 1–2 页：中文、英文斜体、表格填充与边框、红色局部粗体、分页正常；未见缺字、裁切、重叠、重影或异常空白页。检查后关闭两份文档窗口。其他签章版本仅有本轮自动 PNG 等同性与功能测试结果，未逐个做 Preview 验收；外部 OFD 阅读器和 V5/GBT 端到端仍未验收。

## R21 主线整合及资源容器命名空间复审

- 合入 R21 主线 `ba2d14cc13967441c79a89a19b968c7a9782b756`，本分支合并提交 `35895878fce1c5c6996868cb679ffdf9573225bf`。R21 增加可选 `Ofdrw.Net.Text.HarfBuzz` 文本预布局包；本轮本地包消费共 12 个 SDK 包及一个可选 SES 包，CLI 默认依赖闭包仍不含这两个可选包。
- Codex 复审指出，标准 `Res` 根节点内的外来命名空间 `ColorSpaces` 或 `DrawParams` 容器可能被按本地名解析。Reader 现要求资源容器与根节点命名空间一致；回归覆盖外来根节点及外来容器中的 `ColorSpace`/`DrawParam`，避免它们提供绘制语义。相关定向测试与严格签章测试 **51/51**。
- 整合与修复后的解决方案 11 个项目 **3135 passed、13 skipped、0 failed**；本地 `0.1.0-r01.r21b` 的 13 包来源清单、独立消费、DOCX 显式 native/默认模式与 V1/V4 签章链路均通过。四份签章 PDF 与各自签章前基线的两页 96 DPI PNG 像素一致。R21 可选包的独立消费者另生成英文及中文预布局 OFD/PDF/PNG 与回执。本轮产物均在已忽略的 `artifacts/roadmap-v12/R01/r21-final-ses/`，日志与哈希见清单；包未发布。
- macOS Preview 重新打开本轮 `native-v1.pdf` 和 `default-v4.pdf`，各逐页检查第 1–2 页；中英文字、斜体、红色粗体、表格填充和边框、分页均可见，未见缺字、裁切、重叠、重影或异常空白页。另打开本轮 `r21-ascii.pdf` 与 `r21-han.pdf`，各检查第 1/1 页，英文宽窄字形及中文字符可见。查看路径为 DOCX → native/default OFD → 签章 → PDF → Preview，以及 R21 预布局 OFD → PDF → Preview；四份文档窗口检查后均关闭。其他签章版本只有自动 PNG 像素等同性与功能测试证据，不据此宣称外部 OFD 阅读器互操作或 V5/GBT 验收。

## 跨命名空间 DrawParam 颜色继承复审

- Cursor 复审指出：无命名空间或另一官方 OFD 命名空间的 `DrawParam` 颜色直接复制到绘图节点时，颜色节点仍保留资源命名空间，PageBlock 准入会拒绝；普通图层文字则可能找不到继承的描边色而回退黑色。现按绘图节点命名空间重建继承色，并按颜色本地名识别图元自身的显式覆盖色。测试分别覆盖这两种资源命名空间、PageBlock/普通图层、继承的绿色填充与红色文字描边，以及显式蓝色描边优先；PDF/SVG/HTML 均验证，定向 **4/4**，R01 类 **53/53**。
- 修复后完整解决方案 11 个项目 **3139 passed、13 skipped、0 failed**。本地 `0.1.0-r01.r21d` 的 13 包来源与独立消费、R21 文字预布局、DOCX native/默认模式及 V1/V4 签章链路通过；签章 PDF 与基线两页 PNG 像素一致。首次打包尝试因未提供本地受许可字体目录在 SES 消费阶段停止，设置 `OFDRW_SES_FONTS` 后以全新输出目录复验通过；失败属于验证环境，不是产品回归。
- 四个跨命名空间样本各生成 1 页 OFD → PDF → 144 DPI PNG，四张页面图逐字节一致：绿色填充方块、红色空心 `inherited red`、蓝色空心 `local blue` 均清晰且不重叠。最终源码又将四份样本重新生成到 `crossnamespace-pages-final/`，macOS Preview 对四份最终 PDF 的第 1/1 页逐份重新打开并检查。另重新打开本轮包的 `native-v1.pdf`、`default-v4.pdf` 各第 1–2 页，表格、中文/英文、红色粗体和分页正常；两份 PDF 的页面 PNG 也与上一轮 Preview 页面逐字节一致。全部窗口检查后已关闭。外部 OFD 阅读器及 V5/GBT 仍未验收。

## 跨资源文件 Relative 颜色覆盖复审

- Codex 复审指出：`DocumentRes` 与 `PageRes` 采用不同的受支持命名空间时，父子 `DrawParam` 的同名颜色原本按完整 XML 名称并存，父色可挡住子色。现沿 `Relative` 链按 `FillColor`/`StrokeColor` 的本地语义名替换，并归一到目标 DrawParam 命名空间；图元显式颜色仍优先。
- 新回归覆盖两种官方命名空间的双向组合、无命名空间资源的父/子边界，以及 PageBlock/普通图层，总计 **8/8**。三层继承中的子级蓝色描边覆盖父级红色，孙级黄色填充覆盖父级绿色，图元本地洋红色描边覆盖继承蓝色；PDF/SVG/HTML 均有断言，PDF 文本可提取。完整解决方案 11 个项目 **3147 passed、13 skipped、0 failed**。
- 八个样本各生成 1 页 OFD → PDF → 144 DPI PNG，页面图逐字节一致。macOS Preview 对八份 PDF 的第 1/1 页**逐份实际打开检查**：黄色方块和蓝色边框、蓝色 `child blue` 空心字、洋红色 `local magenta` 空心字均显示，未见缺字、重影、裁切、重叠或异常空白页；随后关闭全部窗口。
- 本地 `0.1.0-r01.r21e` 13 包来源与独立消费、R21 预布局、DOCX native/默认模式及四份 V1/V4 签章链路均通过；签章前后各两页 96 DPI PNG 像素一致。最新包的 `native-v1.pdf` 与 `default-v4.pdf` 又在 Preview 各实查第 1–2 页，表格、中英文、红色粗体、分页正常；本轮输出页面 PNG 与上轮逐字节一致。查看后关闭全部窗口。未作外部 OFD 阅读器和 V5/GBT 结论。

## 纯描边文字模拟粗体复审

- Codex 复审指出，`Stroke=true` 且 `Fill=false` 的文字在请求粗体但仅能解析到常规字形时，PDF 的显式轮廓仍采用常规笔画宽度。转换器现对原有描边加上模拟粗体宽度，保持字腔不填充。回归使用嵌入的 Noto Sans 常规字形生成普通和请求粗体的空心 `O`，断言粗体墨迹增加、字形边界稳定、内部仍透明，且 PDF 语义文本只包含一个 `O`。
- 定向回归 **1/1**；完整解决方案 11 个项目 **3148 passed、13 skipped、0 failed**。本地 `0.1.0-r01.r21f` 的 12 个基础 SDK 包及一个可选 SES 包来源、独立消费、R21 预布局、DOCX 显式 native/默认模式与四份 V1/V4 签章链路通过；签章前后对应页面 PNG 一致。包与测试仅作本地验证，未发布。
- 本轮普通与模拟粗体各生成 1 页 OFD → PDF → 288 DPI PNG，两份 PDF 已在 macOS Preview 分别打开对比：粗体轮廓明显加厚，`O` 的中空区域保持开放，没有重复字形或裁切。新包的 `native-v1.pdf` 与 `default-v4.pdf` 又在 Preview 各检查第 1–2 页：中文、英文斜体、表格填充和边框、红色局部粗体及分页符合样例。对应四张本轮页面 PNG 亦已逐页复查；所有窗口均已关闭。外部 OFD 阅读器和 V5/GBT 互操作仍未验收。

## 注释标记与调色板复审、R14 主线整合

- Codex 复审指出，扩展绘图借用 `UnmodeledAnnotationMetadata` 本地名时可能被当作内部注释元数据跳过。只有 Reader 生成且原始 XML 未改动的注释元数据现在携带内部来源标记；普通图层或 PageBlock 的同名绘图继续按未知绘图拒绝。另一处复审指出，引用的 `ColorSpace` 若含两个 `Palette`，原解析只选第一个。现拒绝重复调色板。三个定向用例覆盖普通图层、PageBlock 和重复调色板，PDF/SVG/HTML 均失败闭合；现有合法注释外观仍通过全套回归。
- 修复提交 `0a64c47` 后合入已并入主线的 R14 标准加密提交 `87d35cf11992fec57f14418f93fde85e98c1c48e`，整合提交 `71adec7f69fb87ece6c2c89db566d07e5bc31c06`。整合源码完整解决方案 11 个项目 **3159 passed、13 skipped、0 failed**。本地 `0.1.0-r01.r14g` 的 12 个基础 SDK 包和一个可选 SES 包来源、独立消费、R21 预布局及显式 native/默认 DOCX→OFD、四份 V1/V4 签章样本均通过；另用 R14 标准加密包的独立消费者验证口令/证书多收件人及原子失败契约。均未发布。
- 整合源码重新生成的 `native-v1.pdf` 和 `default-v4.pdf` 均在 macOS Preview 打开并检查第 1–2 页，正文中英文、斜体、红色局部粗体、蓝色表格填充与边框及分页符合样例；四张本轮 PNG 也逐页复查，与前轮验收页面图逐字节一致。全部 Preview 窗口已关闭。伪造标记和重复调色板输入是拒绝路径，没有可验收的成功页面；外部 OFD 阅读器与 V5/GBT 仍未验收。

## 注释索引包装与重复绘图颜色复审

- 继续复审发现，手工创建的 `Annotations` 原始绘图若伪装成 `<UnmodeledAnnotations>` 包装，原本会被当作 Reader 的注释索引元数据跳过。现仅放行 Reader 创建且 XML 未改动的包装；伪造包装在普通页面绘图和注释投影位置均被拒绝。更改 Reader 包装的 XML 后，来源标记也失效。
- 同一 TextObject/PathObject 中重复的 `FillColor` 或 `StrokeColor`，以及被引用的 DrawParam 中重复的颜色子节点，现在在颜色归一化与选择首项前拒绝。六个新增定向用例覆盖四种图元颜色组合、DrawParam 和注释包装；三个渲染出口的重复颜色输入均失败闭合。
- 源码修复提交 `30a3194de59e765d5c43a787239e790ffe3a2029` 的完整 Release 解决方案 11 个项目 **3165 passed、13 skipped、0 failed**。本地 `0.1.0-r01.review3b` 13 包来源与独立消费、R21 预布局、显式 native/默认 DOCX→OFD、四份 V1/V4 签章样例均通过；R14 标准加密的独立 3 包消费者也通过。第一次 SES 运行因字体目录缺少相邻存放的许可文件中止；使用同时包含受许可字体和许可文件的本地目录，以全新输出目录重跑通过。包未发布。
- 本次生成的 `review3-ses-retry/files/native-v1.pdf` 和 `default-v4.pdf` 是由本次 native/默认 OFD 导出。macOS Preview 对两份 PDF 各检查第 1–2 页：中英文标题/正文、灰色斜体、蓝色表格底色与边框、第二页红色粗体和分页均符合样例，无缺字、裁切、重叠、重影或异常空白页。四张本轮 PNG 与前一轮已验收的对应页面哈希逐页一致。两份 PDF 文档窗口均已关闭，并登记 `20261021T1141Z` GUI 代次释放。伪造包装和重复颜色属于拒绝路径，不存在成功页面；未验证外部 OFD 阅读器或 V5/GBT 互操作。

## 实体粗体与斜体轮廓复审

- Codex 复审指出，PDF 轮廓补绘解析到实体 Bold/Italic 字体时，缓存字体族却用默认 Regular 创建字体，可能抛错。现从实际嵌入字体字节读取物理字形样式，按对应样式创建轮廓字体；模拟粗体逻辑独立保留。两种各 1 页的合成字体回归覆盖实体粗体和斜体，PDF 语义文本各保留一个 `O`，页面确有墨迹。合成字体的轮廓预期分别为直立和倾斜的空心四边形，不将它们误读为真实字形的外观验收。
- 修复源码 `79e5d53e34530237fc1706e7eec31f9e939d02fc` 的定向测试 **3/3**，完整 Release 解决方案 11 个项目 **3167 passed、13 skipped、0 failed**。本地 `0.1.0-r01.review4` 的 12 基础 SDK 包和 1 个可选 SES 包来源、独立消费、R21 预布局、显式 native/默认 DOCX→OFD 及四份 V1/V4 签章链路均通过；未发布包。
- 本轮 `physical-face-review/outline-only-physical-bold.pdf` 与 `outline-only-physical-italic.pdf` 均由本轮 OFD 导出，在 macOS Preview 各查看第 1/1 页：实体粗体直立轮廓和斜体倾斜轮廓均可见，无空白、裁切或重复文字。包消费重新生成 `physical-face-ses/files/native-v1.pdf`、`default-v4.pdf`；在 Preview 各查看第 1–2 页，中英文字、灰色斜体、蓝色表格填充/边框、第二页红色粗体及分页符合样例。四张本轮 96 DPI 页面 PNG 与先前验收页面逐字节一致。四份文档窗口均关闭，GUI 代次 `20261021T1211Z` 已释放。其余签章 PDF 仅有自动页面 PNG 等同性和功能测试证据；未验证外部 OFD 阅读器或 V5/GBT 互操作。产物和日志在已忽略的 `artifacts/roadmap-v12/R01/`，哈希见清单。

## R13 主线整合与 PageBlock 绘制失败复审

- 合入已进入主线的 R13 签名范围提交 `070594e414fcb43b55a11a7927799afe797a3cfe`，本分支合并提交 `66f886727fa8b271054a87d7fe0cb005a101e726`。合并后 11 项目完整 Release 测试为 **3265 passed、0 skipped、0 failed**；R13 四包范围的本地包消费、独立 Go XML/SHA 核验通过。首次全套测试因本地受许可字体 fixture 未物化而失败，运行仓库的字体准备脚本后重新执行通过。
- Codex 复审发现：嵌套 PageBlock 的叶子绘制若有未解析颜色，投影层提前拒绝会阻断原始 `TextCode` 提取。`0ba5e6ab1d03da8ab13cb80a67112833126dcba3` 保留 Reader 的 `RenderPaintError`，使文本提取继续读取原文，而 PDF/SVG/HTML 绘制预检明确失败；普通页、模板页及签章拒绝路径定向 **3/3**。修复后 11 项目完整 Release 测试 **3267 passed、0 skipped、0 failed**。
- [GB/T 33190—2016 第 8.3.2 节表 26](https://tsgdag.jhun.edu.cn/_upload/article/files/71/fc/4ed26d5841caab5bc999e19f8f14/02a71fcf-7b6e-4661-9991-eb309c7614c6.pdf)规定：`Value` 出现时采用该值，仅缺少 `Value` 时才取调色板 `Index`。因此未采纳与此相反的复审意见；`34650a9fa7a7aa57e905c05b2a7172a6a022d3d2` 增加直接颜色和继承 `DrawParam` 冲突值回归，PDF/SVG/HTML 均断言红色 `Value` 优先于蓝色 `Index`，定向 **2/2**。
- 本地 `0.1.0-r01.r13paint` 的 **12 个基础 SDK 包与 1 个可选 SES 包**来源清单、独立消费、显式 native/默认 DOCX→OFD、四份 V1/V4 测试签章及页面 PNG 等同性通过；该本地包未发布。包消费针对 `0ba5e6a` 生产代码运行，后续 `34650a9` 仅增加测试。四份签章 PDF 与签章前基线各两页 PNG 像素一致；默认密码学验证仍返回 `Unsupported`，不据此作外部阅读器或签名法律有效性结论。
- 本轮重新生成的 `value-precedence-pages/palette-value-priority-direct.pdf` 与 `palette-value-priority-drawparam.pdf` 由各自 OFD 导出，在 macOS Preview 各查看第 1/1 页，方块均为红色；对应 PNG 也逐页检查。`final-paint-ses-e2e/files/native-v1.pdf`、`default-v4.pdf` 分别由本轮显式 native 和默认 OFD 导出，经 Preview 各检查第 1–2 页：中文、英文、蓝色表格填充与边框、局部红色粗体和分页可见，未见缺字、裁切、重叠、重影或异常空白页。查看后关闭四个文档窗口，并登记 `20261021T1231Z` GUI 代次释放。其余签章 PDF 只有功能和逐页 PNG 等同性证据；没有外部 OFD 阅读器或 V5/GBT 端到端验收。

## R15 主线整合与非法属性词元拒绝

- 合入已进主线的 R15 多 DocBody 提交 `2be7500f89ec6c4245ee7bbbaffe3945a202653f`，本分支合并提交 `69a5c44`。核对 Reader 的所选 DocBody 路径和 PDF 签章只取所选 DocBody 的清单；R01 的 PageBlock 绘制准入继续作用于所选文档内的元素。完整解决方案测试包含 R15 的多文档回归。
- Codex 复审指出，Reader 对已知属性的非法原始值可能使用默认值或夹限值，例如 ImageObject `Alpha`、PathObject `Fill`/`CTM`。投影现对**原始 XML** 的布尔、有限数值、边界、矩阵、Alpha 等已知标量做有界词元检查，再审查资源归一化后的绘图 XML。非法值在 PDF/SVG/HTML 的绘制预检中明确失败，不留下部分输出；错误留在投影叶子，原始 TextCode 仍可提取。裁剪 `Area Start` 仅接受当前可保持语义的原点，非原点明确拒绝。
- 新回归覆盖 Path/Image/Text 的非法布尔、矩阵、线宽、可见性、边界、Alpha、字号、斜体及字重 **11 个案例**，以及裁剪 Start 原点接受/非原点拒绝 **2 个案例**；R15 整合源码定向 **13/13**，完整 Release 解决方案 11 个项目 **3297 passed、0 skipped、0 failed**。其中非法 TextObject 仍可由文本提取器取得原文。Value 优先冲突回归另定向 **2/2**。
- 本地 `0.1.0-r01.r15scalar` 的 12 个基础 SDK 包加 1 个可选 SES 包来源及独立消费通过。显式 native/默认 DOCX→OFD、四份 V1/V4 测试签章与其签章前基线的两页 PNG 像素等同性通过；包未发布。默认密码学验证保持 `Unsupported`，不推出外部 OFD 阅读器或签名法律有效性结论。
- 最终源码新生成的 `r15-scalar-value-pages/palette-value-priority-direct.pdf` 与 `palette-value-priority-drawparam.pdf` 由 OFD 导出，macOS Preview 各检查第 1/1 页，方块均为红色；两份 144 DPI PNG 逐字节一致。`r15-scalar-final-ses/files/native-v1.pdf`、`default-v4.pdf` 分别来自本轮显式 native 和默认 OFD，经 Preview 各检查第 1–2 页：中英文、表格底色/边框、局部红色粗体及分页正常，未见缺字、裁切、重叠、重影或异常空白页。两条查看链路均为 DOCX → OFD → PDF → Preview；检查后关闭四个文档窗口并登记 `20261021T1301Z` GUI 代次释放。其余签章版本仅有功能和逐页 PNG 等同性证据。

## 最终词元与资源内容准入复审

- Codex 对 `5a75559` 的三项复审分别涉及字重合法域、`DeltaX/DeltaY` 尾部非法词元，以及被 `Elements()` 遍历忽略的 DrawParam 处理指令/正文。PageBlock 投影现按现有文字契约只接受 100–900 且步长为 100 的字重；对整个字距表达式逐词检查数字和 `g count value`，不按字形数提前停止，也不展开任意重复计数。词元扫描为 .NET Standard 2.0 兼容的有界分配，受物化预算及取消信号约束。引用的 DrawParam、FillColor/StrokeColor 与 ColorSpace 都拒绝处理指令和非空白正文；注释与空白仍可接受。Reader 的原文提取继续可用，PDF/SVG/HTML 在绘制前失败且不留下部分输出。
- 新回归覆盖非法字重 0/1000/150、合法端点 100/900、两轴字距的非法尾词元、不完整 `g` 重复表达式、合法 `g 1 3`、DrawParam 的处理指令/正文及合法注释/空白、关联颜色/颜色空间内容扩展。R01 类 **101/101**；最终源码完整 Release 解决方案 11 个项目 **3311 passed、0 skipped、0 failed**。Core 的 `netstandard2.0` 单独编译通过。第一次完整测试虽各项目测试显示通过，但该目标框架编译因 `ReadOnlySpan` 缺失而退出 1；改用兼容解析后，从头重跑的完整命令退出 0。日志在 `artifacts/roadmap-v12/R01/r15-review-*.log`。
- 本地 `0.1.0-r01.r15review` 的 12 个基础 SDK 包加 1 个可选 SES 包来源、独立消费者、显式 native/默认 DOCX→OFD、四份 V1/V4 测试签章及页面 PNG 等同性均通过；包未发布。另从最终测试源码生成 5 个单页正例 PDF：重复字距、100/900 字重、直接/继承的 `Value` 优先调色板。PDF 均为 1 页，144 DPI PNG 逐页复查：`A B` 间距清楚、字重端点分别呈常规/粗体、两份调色板页均为红色。最初测试正例的 1 mm 字距造成预期中的字形相贴；改为 3 mm 后重新生成并复查，最终产物无重叠。
- macOS Preview 对本轮上述 5 份 PDF 各检查第 1/1 页；对 `r15-review-final-ses/files/native-v1.pdf` 和 `default-v4.pdf` 分别检查第 1–2 页。查看链路为正例 OFD→PDF→Preview，以及由 `generated-layout.docx` 制作的本轮受许可字体样本→显式 native/默认 OFD→PDF→Preview，未使用 DOCX→PDF 代替。正文中英文、斜体、蓝色表格底色/边框、第二页红色局部粗体和分页符合样例；未见缺字、裁切、重影或异常空白页。四张本轮 DOCX 页面 PNG 与此前验收的对应页逐字节一致。检查后关闭全部 Preview 文档窗口，GUI 代次 `20261021T1341Z` 释放。非法输入是拒绝路径，没有成功页面；本次仍不作外部 OFD 阅读器或 V5/GBT 互操作结论。

## 注释空颜色默认值复审

- Codex 指出注释外观绕过 Reader 颜色规范化：无资源引用且没有 `Value` 的空 `FillColor`/`StrokeColor`，以及仅有 `Alpha` 的颜色，原应以黑色分量呈现，先前文字绘制校验会拒绝。Core 的注释绘制解析现给这种颜色补零分量并保留 Alpha；有 `ColorSpace`、`Index`、非法 `Value` 或未识别子内容的颜色继续拒绝。
- 两组实际注释输入分别覆盖空填充/描边色与 `Alpha=128` 的填充/描边色。PDF 原文可抽取；SVG/HTML 为黑色，SVG 同时保留填充与描边的不透明度。最终源码完整 Release 解决方案 **3313 passed、0 skipped、0 failed**，包括 Core `netstandard2.0` 编译；定向回归 2/2。测试及构建日志在 `artifacts/roadmap-v12/R01/annotation-default-*.log`。
- 从新生成的两份 OFD 各导出 1 页 PDF，并渲染 144 DPI PNG。macOS Preview 已各检查第 1/1 页：`Default` 为黑色、`Alpha` 为半透明灰色外观，字形完整，无裁切、重影或异常空白页；PNG 逐页复查结果一致。查看后已关闭这两份 Preview 文档窗口。
- 本地 `0.1.0-r01.annotation` 的 12 个基础 SDK 包及 1 个可选 SES 包来源清单、独立消费者、显式 native/默认 DOCX→OFD、四份 V1/V4 测试签章及逐页 PNG 等同性通过；包未发布。默认密码学验证仍为 `Unsupported`，不据此认定外部阅读器或签名法律有效性。
- 本轮 `annotation-default-final-ses/files/native-v1.pdf` 与 `default-v4.pdf` 均为各自 OFD 导出的新产物；在 macOS Preview 各检查第 1–2 页。查看链路分别为 DOCX → 显式 native OFD → PDF → Preview、DOCX → 默认 OFD → PDF → Preview。中文、英文、蓝色表格底色与边框、第二页红色粗体及分页与样例一致；四张本轮 PNG 与前轮已验收页面逐字节一致。检查后关闭全部文档窗口，并登记 `20261021T1422Z` GUI 代次释放。

## 注释路径与继承绘制参数复审

- Codex 对 `4f0ac41` 指出两处绘制准入缺口：未解析资源的注释路径可能使用类型化回退色，继承链上被子项覆盖的父级颜色可能在校验前丢失。渲染预检现在检查路径源 XML 的资源引用与直接填充/描边颜色；Reader 在合并每个 `DrawParam` 节点前校验其颜色，包含父项后来被覆盖的情形。注释路径仅声明 `Alpha` 时，Reader 用黑色默认分量保留透明度。与绘制无关的注释扩展元数据仍按原有路径保留。
- 负例覆盖缺失 `DrawParam`/`ColorSpace`、无颜色空间的 `Index`、非法颜色分量/Alpha，以及父级颜色的扩展属性、处理指令和缺失颜色空间；PDF/SVG/HTML 均在绘制前拒绝且输出为空。正例覆盖有效注释路径、仅有 Alpha 的半透明黑色路径、跨资源命名空间的相对 `DrawParam` 子项覆盖。修订后完整 Release 解决方案 **3322 passed、0 skipped、0 failed**；定向负例与 Alpha 回归 9/9，正例生成 13/13。Core `netstandard2.0` 由完整解决方案编译通过。
- 本地 `0.1.0-r01.pathpaint` 的 12 个基础包和 1 个可选 SES 包来源、独立消费者、显式 native/默认 DOCX→OFD→PDF、四份 V1/V4 测试签章与各自基线逐页 PNG 等同性通过；包未发布。默认密码学验证仍为 `Unsupported`。功能、生成链路与哈希见 `artifacts/roadmap-v12/R01/pathpaint-final-ses/` 及本目录清单。
- 从本轮 OFD 导出的三份单页 PDF 在 144 DPI PNG 与 macOS Preview 各检查第 1/1 页：半透明黑色路径呈灰色矩形，有效注释路径为红色，相对 `DrawParam` 页为黄色填充、蓝色描边，并保留蓝色 `child blue` 与洋红 `local magenta`。本轮 `native-v1.pdf`、`default-v4.pdf` 分别由显式 native 和默认 OFD 导出；两条 DOCX→OFD→PDF→Preview 链各检查第 1–2 页。中英文、斜体、蓝色表格底色/边框、第二页局部红色粗体及分页符合样例；未见缺字、裁切、重叠、重影或异常空白页。四张本轮 DOCX 页面 PNG 与既有已验收像素哈希一致；所有本轮 Preview 文档窗口已关闭，GUI 代次 `20261021T1502Z` 已释放。结论仅适用于这些实际检查的页面，不推出外部 OFD 阅读器互操作或任意 DOCX 保真。

## PageBlock 字体与绘图边界复审

- Codex 对 `15edd37` 指出 PageBlock `TextObject` 的缺失/未知 `Font` 会落入宿主字体回退，以及缺失或非正 `Boundary` 的图像可能被静默省略。Reader 现在给 PageBlock 未解析字体子项保留可提取的文本并标记渲染错误；投影在准入时要求 PageBlock 文本、路径、图像的原始 `Boundary` 存在且宽高为有限正数。普通层的既有解析行为不受这项 PageBlock 门禁影响。
- 新增 9 个负例覆盖普通/模板 PageBlock 的缺失与未知字体、图像缺失/零/负边界、路径负边界和文本缺失边界；PDF/SVG/HTML 均在输出前拒绝，文字仍可提取。定向测试 **9/9**；修订后完整 Release 解决方案 **3331 passed、0 skipped、0 failed**，Core `netstandard2.0` 通过。正例重新生成 **12/12**。本地 `0.1.0-r01.fontboundary` 的 12 个基础包及 1 个可选 SES 包来源、独立消费者、显式 native/默认 DOCX→OFD→PDF、V1/V4 测试签章和逐页 PNG 比对通过；默认密码学验证仍为 `Unsupported`，包未发布。最初两次消费尝试分别因字体目录缺失、许可文件缺失停止；使用已有完整字体与许可目录后全链路通过，属于测试资料前置条件。
- 本轮普通与模板 PageBlock 各 1/1 页的 144 DPI PNG 像素哈希相同，蓝色路径和 `nested text` 均可见；图像页 1/1 显示蓝色路径与红色图像；相对 DrawParam 页 1/1 显示黄色填充、蓝色边框/文本和洋红局部文本。上述四份从本轮 OFD 导出的 PDF 均在 macOS Preview 检查第 1 页。`font-boundary-final-ses-verified/files/native-v1.pdf` 和 `default-v4.pdf` 各自由本轮显式 native/默认 OFD 导出，两条 DOCX→OFD→PDF→Preview 链各检查第 1–2 页：中英文、斜体、蓝色表格底色/边框、第二页红色粗体与分页符合样例，未见缺字、裁切、重叠、重影或异常空白页。本轮四张 DOCX 页面 PNG 与既有验收哈希一致；全部 Preview 窗口关闭并登记 `20261021T1602Z` GUI 释放。结论仅覆盖实查页面，不推断外部 OFD 阅读器或任意 DOCX 的保真性。

## 本次页面视觉复查

| 样本/链路 | 实查范围 | 结果 |
| --- | --- | --- |
| `ordinary-block`、`template-block`：OFD → PDF/SVG/HTML/PNG | 各 1 页；PDF 与 SVG 的 PNG、HTML 浏览器截图、macOS Preview PDF | 蓝色方块与 `nested text` 顺序/位置一致；模板和普通页外观一致 |
| `palette-drawparam` | 1 页；同上 | 蓝色/青色描边方块、红色方块可见 |
| `stroke-hscale` | 1 页；同上 | 红色空心 `R01 text`，横向缩放可见；PDF 原文可抽取，无重复文字 |
| `image-block` | 1 页；同上 | 蓝色路径和红色图片方块按顺序可见 |
| `implicit-stroke`、`annotation-gray` | 各 1 页；本轮 PDF/PNG 与 macOS Preview，SVG/HTML 断言 | 未指定描边色的空心文字为黑色；未引用 ColorSpace 的注释单分量灰色文字可见 |
| `annotation-default`、`annotation-alpha`：注释空颜色和仅有 Alpha 的文字 | 各 1 页；本轮 OFD→PDF、PNG 与 macOS Preview，SVG/HTML 断言 | 默认黑色及半透明黑色均可见；填充与描边无缺失、裁切或重影 |
| `top-level-gray-cmyk`：顶层路径读入 → 保存 → 重开 → PDF/SVG/HTML | 1 页；本轮 PDF/PNG 与 macOS Preview | 灰色及红色方块正确，无丢色或空白页；三出口断言通过 |
| `split-ordinary-block`、`split-template-block`：Reader → Split → 写出 → PDF/SVG/HTML | 各 1 页；本轮 PDF/PNG 与 macOS Preview | `split nested` 均显示，位置一致；原始文本可提取 |
| `hidden-opaque-ordinary`、`hidden-opaque-template`：隐藏未知绘图与可见兄弟文字 | 各 1 页；本轮 PDF/PNG 与 macOS Preview | 仅 `visible sibling` 可见，模板与普通页一致；PDF/SVG/HTML 断言通过 |
| `generated-layout.docx` → 显式 native OFD → PDF → macOS Preview；默认模式单独重复 | 两条链各 2 页，另核对选第 2 页产物的 PNG | 中文、英文、斜体、红色局部粗体、表格底色/边框与分页符合样例；未见缺字、裁切、重影或异常空白页；选页 PNG 与完整文档第 2 页字节一致 |
| 原始 `upstream-ofdrw/999.ofd`；另制 render-only 副本 → PDF | 原始签名 SM3 引用完整性与无效 ColorSpace 拒绝；副本 PDF 5 页逐页 PNG | 原件的嵌入章将 `ColorSpace=4` 指向字体资源，转换明确拒绝。仅在**签名已失效的呈现副本**把两处引用改成正确 RGB ID 5 后，章面与 5 页内容可见；副本不作签名有效性或原件互操作结论。此副本未做 Preview 验收 |

PDF 和 SVG 的页面 PNG 均从本轮生成文件渲染，HTML 五页由浏览器截图复查。macOS Preview 已重新打开 R01 PDF，检查七个既有 1 页样本、顶层颜色保存重开样本、Split 普通/模板页及 DOCX native/default 各 1–2 页，检查后关闭 R01 窗口。最终源码重生成的七个既有 PDF 页面 PNG、DOCX native/default 四张页面 PNG 与此前 Preview 验收时的页面图分别逐字节一致。R03 命名区域空白/填充共四张最终 PNG 也与此前逐页检查的页面图逐字节一致。先前 `review3` 签章副本的五页 PNG 与此前逐页查看的 PNG 逐字节一致；它仅作包消费回归与 PNG 辅助检查。没有独立第三方 OFD 阅读器互操作结论，也不将这些有限样本推广到任意复杂文档。

## 产物身份

完整输入/输出哈希见 [hashes.txt](hashes.txt)。关键输入：`generated-layout.docx` 为 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`；`999.ofd` 为 `a2b9080ae35184135160c48ce45fb609c35d5ecfdb538ebc16132c48a2186a80`。本轮独立 native/default DOCX 产物及选页产物的哈希均列于清单。页面图/浏览器截图在同一已忽略的本地目录，可按清单对照重新复查。
