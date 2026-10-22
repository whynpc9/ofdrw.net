# R01 PageBlock、文字与颜色验收

日期：2026-10-22（Asia/Shanghai）；分支 `codex/r01-pageblock-text-color`；[PR #29](https://github.com/whynpc9/ofdrw.net/pull/29)。产品源码 `5b37b0f537ade4a47a186bfcd9896d00adbab9d7`，已整合主线 `62887ad6af62605411e1968da662be48ce1c091c`。macOS Arm64、.NET SDK 10.0.401。复查文件在已忽略的 `artifacts/roadmap-v12/R01/`；[hashes.txt](hashes.txt) 记录本轮输入、输出及验收记录。

## 行为边界

- Reader 将普通页和模板页的 `PageBlock` 绘图投影为临时只读子元素，保持原始 XML、父子关系和文字提取结果。PDF、SVG、HTML 按原顺序呈现支持的路径、图片、文字、DrawParam、复合对象及有限颜色空间。复合资源与文档级绘制资源在投影、渲染之间维持正确引用；投影不取得原 XML 的写入权。
- 字体、媒体、复合资源、颜色空间、DrawParam 及其引用按正整数 ID 解析；`030` 与 `30` 等值，原始 XML 拼写不变。字体选择以实际引用的数字 ID 为准。重复字体和媒体 ID 在源码身份捕获阶段拒绝，重复复合资源 ID 在 Reader 阶段拒绝。PageBlock 与复合资源中的有效文字及兄弟文字可提取、可渲染；异常字体声明保留逻辑文字，但阻止三个渲染器写出。
- 普通绘图及裁切路径的弧线要求正半径和字面 `0`/`1` 标志位。非法的原始裁切路径词元在 PDF、SVG、HTML 写出前统一拒绝；合法嵌套扩展不误判。空白、重复或语法不完整的路径数据也拒绝；XML 注释之间的路径空白按实际绘制文本保留。
- 颜色调色板的每个条目均经过校验，直接 `Value` 绘制不绕过资源结构检查；合法 `Value` 优先于 `Index`。未知绘图、资源歧义和不支持的样式在输出前拒绝。隐藏元素、布尔空白、复合对象放置与裁切、局部文字样式均由针对性测试覆盖。

## 功能测试与包消费

- 产品源码的完整 Release 解决方案：**3568 passed、0 failed、0 skipped**，11 个测试项目；R01 PageBlock **175/175**。使用单节点构建与仓库独占锁；日志分别为 `p11-full-release.log` 和 `p11-focused.log`。
- 本地版本 `0.1.0-r01.p8` 的 **12 个基础 SDK 包和 1 个可选 SES 包**通过来源验证及独立消费者 E2E。显式 `native` 与默认 DOCX→OFD→PDF 各生成 V1/V4 测试签章，原文完整保留；签章前后逐页 PNG 一致。`p8-ses/files/functional.json` 和 `rendered-pages.json` 保留机器检查结果。本地 Core、Packaging、Reader 精确包的 R18 附件清理消费者也通过，物理删除后回读为空。包未发布；默认密码学验证结果为 `Unsupported`，测试专用自签路径通过。

## 自动渲染、PNG 与 Preview

11 个本轮 PageBlock OFD 各导出一页 PDF，并渲染 144 DPI PNG。本轮逐页 PNG 复查与 macOS Preview 对应 PDF 的 11/11 页检查均通过：

| 页面样例 | 实查内容 |
| --- | --- |
| `numeric-font-ids`、`numeric-font-face-ids` | PageBlock 与复合资源文字、直接和嵌套字体选择、粗体均正确，无缺字或重影。 |
| `numeric-media-ids`、`numeric-media-ids-template` | 普通页与模板页的红色方块位置和尺寸一致；两张 PNG 哈希相同。 |
| `numeric-composite-ids`、`composite-paint-resources` | 复合文字完整，文档级红蓝绘制资源均正确呈现。 |
| `numeric-paint-ids`、`pageblock-arc`、`pageblock-clip-arc`、`pageblock-clip-arc-default`、`pageblock-comment-whitespace` | 红色方块、蓝色弧线与裁切区域、注释空白形成的三角形均符合样例；无异常裁切、重叠或重影。 |

`generated-layout.docx` 经本地许可字体样本分别生成显式 `native` OFD 和默认 OFD，再各自导出 PDF。V1/V4 自动生成链路各两页、共八张 96 DPI PNG，与对应签章前基线逐页哈希一致。本轮对 `native-v1.pdf`、`default-v4.pdf` 的两页分别在 PNG 和 macOS Preview 查看，链路均为 **DOCX→OFD→PDF→Preview**。中文及英文、英文斜体、蓝色表格底色和边框、第二页局部红色粗体与分页符合样例；未见缺字、裁切、重叠、重影或异常空白页。

Preview 本轮合计 **13 份 PDF、15/15 页通过**。检查后退出 Preview，中央 GUI 独占租约 `20261022T0208Z` 已释放；逐文件 PDF/PNG 哈希与页码记录于 `artifacts/roadmap-v12/R01/gui-release-R01-20261022T0208Z.json`（中央原件位于项目共用 `resource-leases/`）。结论仅覆盖上述样例和实查页面；外部 OFD 阅读器、V5/GBT 端到端互操作、签名法律效力及任意复杂 DOCX 的保真性未验收。
