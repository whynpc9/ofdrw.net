# R01 PageBlock、文字与颜色验收

日期：2026-10-22（Asia/Shanghai）；分支 `codex/r01-pageblock-text-color`，[PR #29](https://github.com/whynpc9/ofdrw.net/pull/29)。R01 产品源码提交 `7abeb3163fc713ea9574b9f13cd0db81dc4263d9`，组合提交 `e6065ada574505d177204a2da21998a6fa748ec9` 已合入主线 `903a7c8a09e3e68cc1d1d9b6abfdccd64372db0a`（R17）。测试环境为 macOS Arm64、.NET SDK 10.0.401。输入、日志和输出在已忽略的 `artifacts/roadmap-v12/R01/`；[hashes.txt](hashes.txt) 记录本轮必要文件的 SHA-256。

## 范围与行为

- Reader 对普通页和模板中的 `PageBlock` 建立只读子元素投影；保存仍以原始 XML 为准，Split 后维持投影与文字提取。PDF、SVG、HTML 对未解析绘图、外来资源命名空间、非法原始标量及不支持的绘制参数明确拒绝。
- 有界颜色解析覆盖 DrawParam 相对继承、8-bit GRAY/RGB、无 ICC 的简单 CMYK 与调色板；文字填充、描边、横向缩放和支持的字体轮廓分别呈现。缺失或歧义引用、渐变/图案及不支持的文字语义继续拒绝。
- `PageBlock` 内的 `CompositeObject` 使用**原始**放置 XML 验证正宽高 `Boundary` 与包括 `CTM` 的标量词元。无效放置仍投影资源文字及同页、后续页文字，但 PDF/SVG/HTML 的预检均拒绝渲染；有效放置正常呈现。四种测试分别覆盖有效、缺失边界、零宽边界与非法矩阵。
- 字体资源中无效的 Bold/Italic 标志不会误满足 PageBlock 的字体解析；资源和文本仍可读取，三个渲染器拒绝输出。调色板 `CV` 内及条目之间的 XML 注释不改变颜色。方向和 Alpha 属性的等值整数写法按数值解析，仍只接受原有的 0/255 值。
- 支持的文字 Fill/Stroke 布尔词元允许首尾空白；`Visible=" false "` 的 PageBlock 子文字和路径保持隐藏，不会泄漏到文字提取、PDF、SVG 或 HTML；复合对象内隐藏裁切路径也按规范化值拒绝。
- 复合资源文本引用的字体声明存在但 Bold/Italic 标志异常时，Reader 保留该文本与同页兄弟文本用于逻辑提取，并标记渲染失败；PDF/SVG/HTML 均在写出前拒绝，不将异常字体当作可用资源。
- 数字布尔值 `Fill=" 1 "`、`Stroke=" 0 "` 和 `Visible=" 0 "` 在标量预检与 Reader 中统一按首尾空白规范化；普通 PageBlock 与复合路径保留蓝色填充、取消描边并隐藏红色路径。Reader 原先对无空白 `Stroke="0"` 的默认真值也已修正。
- R17 归档规则及副本 Handler 经原始 ZIP loader 处理 `Document.xml`，不调用 R01 的 Reader、Writer 或 PageBlock 投影；合并没有同文件冲突。以下组合测试与包消费验证了两个改动集同时存在时的行为，不能推断任意归档文件可渲染或符合完整标准。

## 功能测试与包消费

- 合并后的完整 Release 解决方案 11 个测试项目 **3422 passed、0 skipped、0 failed**；R01 PageBlock **143/143**，R17 `ArchivalCopyTests` **19/19**。Core `netstandard2.0` 包含在完整构建内。单节点构建使用仓库规定的 CLI 环境与共享独占锁。
- 本地版本 `0.1.0-r01.r17integration` 的 **12 个基础 SDK 包和 1 个可选 SES 包**通过来源验证及独立消费者 E2E；显式 native 和默认 DOCX→OFD 各生成 V1/V4 测试签章，再从各自 OFD 导出 PDF。原始文本完整，四份签章 PDF 与签章前基线逐页 PNG 一致。另以这个精确 Packaging 包的 DLL 独立运行 R17 四例消费者及归档副本，结果通过。默认密码学验证为 `Unsupported`；本地包未发布。

## 自动渲染、PNG 与 Preview 视觉验收

| 本轮产物与实际查看链路 | 实查范围 | 结果 |
| --- | --- | --- |
| `pageblock-composite-valid.ofd` → PDF → 144 DPI PNG；前轮 PDF → macOS Preview | 1/1 页，当前 PNG 与前轮一致 | 黑色方块位置与大小符合样例；无裁切、重影或空白页 |
| 普通/模板 PageBlock、图片、三种合法媒体资源与相对 DrawParam OFD → PDF → 144 DPI PNG | 七个样例各 1/1 页，逐页 PNG | 蓝色路径/文字、红色图片、黄底蓝框及局部洋红文字可见；普通/模板和三种媒体页面各自一致 |
| 两种调色板注释与三种等值整数写法 OFD → PDF → 144 DPI PNG；前轮各选一份 PDF → Preview | 五个样例各 1/1 页，当前 PNG 与前轮逐页检视结果哈希相同；前轮 Preview 两份各 1/1 页 | 注释位置不改变蓝色三角形；三种整数写法均显示相同 `numeric text`，没有重复文本 |
| 文字布尔空白、隐藏文字和隐藏路径 OFD → PDF → 144 DPI PNG；前轮 PDF → Preview | 三个样例各 1/1 页，当前 PNG 与前轮一致 | 文字正常绘出；隐藏样例仅显示 `visible sibling`，未见隐藏文字或蓝色路径 |
| 数字布尔空白文字、普通 PageBlock 路径及复合路径 OFD → PDF → 144 DPI PNG → macOS Preview | 三个合并后样例各 1/1 页，PNG 与前轮相同；当前 PDF 各在 Preview 查看 1/1 页 | 文字完整；两个路径样例均只显示蓝色实心方块与兄弟文字，未见红色隐藏路径或额外描边 |
| `generated-layout.docx` → 本地许可字体替换样本 → 显式 native OFD → PDF → macOS Preview；默认模式单独重复 | 合并后 `native-v1.pdf` 和 `default-v4.pdf` 各在 Preview 查看 1–2 页；四种 V1/V4 模式各两页 PNG 与前轮相同 | 中英文、英文斜体、蓝色表格底色/边框、第二页局部红色粗体与分页符合样例；本轮 Preview 未见缺字、裁切、重叠、重影或异常空白页 |

R17 合并后重新生成 19 个 PageBlock 样例各 1/1 页 PNG，与前轮逐页检视过的 PNG 哈希全部相同；四组 DOCX→OFD→PDF 的八张页图也与前轮相同。合并后五份当前 PDF 的 macOS Preview 实查路径、页码、文件与页面哈希记录在 `r17-integration-preview-release.json`；验收后已关闭所有 Preview 窗口及进程。先前 R01 源码的 Preview 检查路径、页码、哈希记录在 `gui-release-R01-20261021T1912Z.json`，更早复合、隐藏及 Palette/Alpha Preview 分别记录在 `gui-release-R01-20261021T1852Z.json` 和 `gui-release-R01-20261021T1812Z.json`。`rendered-pages.json` 的 `previewAccepted: false` 是自动化阶段的状态，独立 Preview 验收以 GUI 记录为准。签章前后 PNG 等同性属于自动渲染证据，不代表逐份签章 PDF 均已在 Preview 查看。`native-v1.ofd` 约 11.6 MB，对应 PDF 约 119 KB，未见明显体积变化。

结论仅覆盖上述样本与实际检查的页面。原始上游 `999.ofd` 的无效章内颜色空间引用仍明确拒绝；另制签名失效呈现副本的先前结果不作为原件互操作证明。外部 OFD 阅读器、V5/GBT 端到端互操作、签名法律效力及任意复杂 DOCX 的保真性均未验收。
