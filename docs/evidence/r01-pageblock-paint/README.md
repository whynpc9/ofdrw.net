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

## 本次页面视觉复查

| 样本/链路 | 实查范围 | 结果 |
| --- | --- | --- |
| `ordinary-block`、`template-block`：OFD → PDF/SVG/HTML/PNG | 各 1 页；PDF 与 SVG 的 PNG、HTML 浏览器截图、macOS Preview PDF | 蓝色方块与 `nested text` 顺序/位置一致；模板和普通页外观一致 |
| `palette-drawparam` | 1 页；同上 | 蓝色/青色描边方块、红色方块可见 |
| `stroke-hscale` | 1 页；同上 | 红色空心 `R01 text`，横向缩放可见；PDF 原文可抽取，无重复文字 |
| `image-block` | 1 页；同上 | 蓝色路径和红色图片方块按顺序可见 |
| `implicit-stroke`、`annotation-gray` | 各 1 页；本轮 PDF/PNG 与 macOS Preview，SVG/HTML 断言 | 未指定描边色的空心文字为黑色；未引用 ColorSpace 的注释单分量灰色文字可见 |
| `top-level-gray-cmyk`：顶层路径读入 → 保存 → 重开 → PDF/SVG/HTML | 1 页；本轮 PDF/PNG 与 macOS Preview | 灰色及红色方块正确，无丢色或空白页；三出口断言通过 |
| `split-ordinary-block`、`split-template-block`：Reader → Split → 写出 → PDF/SVG/HTML | 各 1 页；本轮 PDF/PNG 与 macOS Preview | `split nested` 均显示，位置一致；原始文本可提取 |
| `hidden-opaque-ordinary`、`hidden-opaque-template`：隐藏未知绘图与可见兄弟文字 | 各 1 页；本轮 PDF/PNG 与 macOS Preview | 仅 `visible sibling` 可见，模板与普通页一致；PDF/SVG/HTML 断言通过 |
| `generated-layout.docx` → 显式 native OFD → PDF → macOS Preview；默认模式单独重复 | 两条链各 2 页，另核对选第 2 页产物的 PNG | 中文、英文、斜体、红色局部粗体、表格底色/边框与分页符合样例；未见缺字、裁切、重影或异常空白页；选页 PNG 与完整文档第 2 页字节一致 |
| 原始 `upstream-ofdrw/999.ofd`；另制 render-only 副本 → PDF | 原始签名 SM3 引用完整性与无效 ColorSpace 拒绝；副本 PDF 5 页逐页 PNG | 原件的嵌入章将 `ColorSpace=4` 指向字体资源，转换明确拒绝。仅在**签名已失效的呈现副本**把两处引用改成正确 RGB ID 5 后，章面与 5 页内容可见；副本不作签名有效性或原件互操作结论。此副本未做 Preview 验收 |

PDF 和 SVG 的页面 PNG 均从本轮生成文件渲染，HTML 五页由浏览器截图复查。macOS Preview 已重新打开 R01 PDF，检查七个既有 1 页样本、顶层颜色保存重开样本、Split 普通/模板页及 DOCX native/default 各 1–2 页，检查后关闭 R01 窗口。最终源码重生成的七个既有 PDF 页面 PNG、DOCX native/default 四张页面 PNG 与此前 Preview 验收时的页面图分别逐字节一致。R03 命名区域空白/填充共四张最终 PNG 也与此前逐页检查的页面图逐字节一致。先前 `review3` 签章副本的五页 PNG 与此前逐页查看的 PNG 逐字节一致；它仅作包消费回归与 PNG 辅助检查。没有独立第三方 OFD 阅读器互操作结论，也不将这些有限样本推广到任意复杂文档。

## 产物身份

完整输入/输出哈希见 [hashes.txt](hashes.txt)。关键输入：`generated-layout.docx` 为 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`；`999.ofd` 为 `a2b9080ae35184135160c48ce45fb609c35d5ecfdb538ebc16132c48a2186a80`。本轮独立 native/default DOCX 产物及选页产物的哈希均列于清单。页面图/浏览器截图在同一已忽略的本地目录，可按清单对照重新复查。
