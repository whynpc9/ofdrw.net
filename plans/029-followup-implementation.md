# 029 有界多页 HTML 预览实施与验收

本任务实施和规定范围的验证已完成，可供父协调会话集成。未推送、创建 PR、合并、打 tag 或发布。

- 冻结起点：`3247c76f95c9a3312cc26144e24c4984767af26f`。
- 分支：`codex/audit-followup-029`。
- 受管 worktree：`/Users/wanghongyi/.codex/worktrees/audit-followup-029/ofdrw.net`。
- 最终产品源码：`9a67e2658e94b45d695fb383b1f10f91cef178b3`。本报告的后续提交只修改验收文档。
- 本地实施提交：`aa5005df69f7576ca0940b5b97ce46e159d219dd`（主实现）、`1a21f90aca9b6f8e599739988005ec3ff3c00faa`（SVG 字体继承隔离）、`9a67e2658e94b45d695fb383b1f10f91cef178b3`（收窄并检查 PNG 灰度/RGB(A) 语法范围）。

## 实际实现与边界

在现有 SVG 程序集新增 `OfdToHtmlConverter` / `OfdHtmlOptions`，没有增加 NuGet 包或依赖。显式 `ofd-to-html` 的解析与文件写入位于独立 `Cli.Html.cs`；旧 `ofd-to-svg` 仍是单页，通用 `convert` 没有新增 `.html` 自动路由。

一次 Reader/ZIP 加载，首文档、排序页序号、首次 await 前有限复制选页、保留重复和出现顺序。每次调用独立计数/缓存；每个 occurrence 独立 clip ID 与字体 family。字体按内容及 collection face 缓存成功解析，同页相同内容只内嵌一次，跨页再次内嵌并计费。当前基线已有 014/015，HTML 保留受控 alpha/dash 输出，专项断言比较 HTML 与独立 SVG 的路径。

在投影、字形拆分、路径正规化、Base64/CSS/DOM 分配前保守预留生成字符/节点；流序列化精确计全局与 section UTF-8 字节。资源和重复页逐次计费。额外 `MaxDecodedImagePixels` 默认累计 100,000,000 像素，在 PNG IHDR/JPEG SOF 准入点、inflate/熵遍历之前收费；重复图片/页仍逐次计费。真实 PNG/JPEG 小预算测试验证两次 64×64 成功、第三次越过 8192 上限失败且公开输出未变，准入点取消保留 OCE；扫描行/MCU 循环另有取消检查。

采用有界私有分块内存 spool，代替设计提议的磁盘临时 spool，避免在共享临时目录写入文档明文。生成成功才复制到调用者流；流保留当前位置和所有权。最终复制/Flush 失败或取消可能留下前缀，明确抛错，不宣称通用回滚。CLI 复用 `AtomicOutput`，提交前失败保护已有目标。

准入边界是自己生成的受限 SVG DOM，非通用 SVG sanitizer。标题/正文按上下文转义；CSS 仅固定模板及受控 font-face，含防御性 CSP。未知可见 raw 绘图、不支持字形/注释/绘制属性、事件、外链、未知属性和不闭合片段均拒绝。只显示已有 SVG 绘制范围，不恢复交互/导航或签章信任 UI。

媒体没有调用任意格式通用解码器：非交错灰度/RGB(A) PNG 校验 chunk/CRC、静态元数据长度、zlib 行长/filter 字节和 Adler；JPEG 仅单扫描 8 位 baseline Huffman，检查段、量化/Huffman 表和熵块遍历。indexed/interlaced/animated PNG、文本/私有 PNG chunk、progressive/multiscan/arithmetic JPEG、BMP/TIFF/GIF、尾随载荷等明确拒绝。语法准入不证明浏览器解码器安全或任意色彩管理保真；输入/模型/font parse 限制仍独立，输出额度不是进程级内存/CPU/墙钟上限。

公开用法、预算、异常及限制见 [HTML preview](../docs/html-preview.md)。

## 最终功能验证

以下均针对产品源码 `9a67e265…`，不是旧设计检查或零命中筛选。

| 门 | 实际结果与证据 |
| --- | --- |
| 全仓 restore/build | 成功；包含 SVG `netstandard2.0`、`netstandard2.1`。[命令、退出码](../artifacts/plan-029-followup/final-gates/full-results.json) |
| 全套测试 | **2857/2857，0 failed，0 skipped**；11 份 TRX。HTML 库 29 个、HTML CLI 16 个实际命中。[逐 TRX 计数](../artifacts/plan-029-followup/final-gates/test-summary.json) |
| 公开 API XML 文档 | 通过，新增未文档化成员 0；272 个既存历史符号未扩张。[日志](../artifacts/plan-029-followup/final-gates/public-api.log) |
| 全部本地包消费 | **11 包**验证与原有真实消费者通过；仅本地版本 `0.1.0-preview.plan029`。[日志](../artifacts/plan-029-followup/final-gates/packages.log) |
| 新 HTML 包消费 | 独立纯 PackageReference API 消费者与新安装的本地 CLI 实际执行，顺序 `3,1,3`、转义和完整 HTML 断言通过。[日志](../artifacts/plan-029-followup/final-gates/html-package.log)、[可复跑 runner](../scripts/run-html-preview-package-e2e.sh) |
| 流/安全/失败控制 | 一次实际 Reader + 非 seek 流、快照/并发、重复页 ID/font、HTML 字面转义、局部引用闭包、精确 N/N−1 UTF-8、节点/字符/页/资源/像素小预算、预取消/读后取消/最终复制取消与写失败、坏后页不写公开输出、CLI 旧目标/同路径/rename 失败均有非零回归 |
| 原始文字与 PDF 重复控制 | 17 个 DOCX 文本 run 均在 native/default OFD 中保留，共 202 字符；标题、红色强调句在 PDF 提取结果均只出现一次。[记录](../artifacts/plan-029-followup/final-gates/text-retention.json) |

所有 dotnet 使用可写 CLI_HOME、禁首次体验/遥测和单节点/禁构建服务器参数。全回归、API Rebuild 与全包消费使用共享 validation 租约并在 finally 释放；未清理全局 NuGet。最终 `git diff --check` 和 shell 语法检查通过。

## 最终浏览器验收

最终 HTML 在字体继承修复和 PNG 收窄之后重新生成；[生成前哈希清单](../artifacts/plan-029-followup/final-gates/visual/manifest-before-browser.json) 记录完整源码 SHA。最终浏览器记录为 **Chrome 152.0.7977.54 / macOS，file URL**，实际 URL、字体加载、页面尺寸/顺序、请求及新截图 SHA 见 [最终 observations](../artifacts/plan-029-followup/final-gates/visual/browser-observations.json)。`artifacts/plan-029-followup/visual/` 中早先 aa500 的 HTML 截图仅属诊断历史，不用作新 HTML 的验收证据。

| 本次重新生成的 HTML | 实际查看范围 | 最终完整页面截图 |
| --- | --- | --- |
| [native.html](../artifacts/plan-029-followup/final-gates/visual/native.html) | 1、2 页 | [两页](../artifacts/plan-029-followup/final-gates/visual/browser-native-full.png) |
| [default.html](../artifacts/plan-029-followup/final-gates/visual/default.html) | 1、2 页 | [两页](../artifacts/plan-029-followup/final-gates/visual/browser-default-full.png) |
| [mixed.html](../artifacts/plan-029-followup/final-gates/visual/mixed.html) | 3、1、3，含重复 | [三次出现](../artifacts/plan-029-followup/final-gates/visual/browser-mixed-full.png) |
| [mixed-all.html](../artifacts/plan-029-followup/final-gates/visual/mixed-all.html) | 1、2、3，含异形宽页 | [三页](../artifacts/plan-029-followup/final-gates/visual/browser-mixed-all-full.png) |
| [duplicate-fonts.html](../artifacts/plan-029-followup/final-gates/visual/duplicate-fonts.html) | 3、1、3，三套内嵌 family | [三次出现](../artifacts/plan-029-followup/final-gates/visual/browser-duplicate-fonts-full.png) |
| [blank.html](../artifacts/plan-029-followup/final-gates/visual/blank.html) | 1 个明确空白源页 | [一页](../artifacts/plan-029-followup/final-gates/visual/browser-blank-full.png) |
| [injection-sentinel.html](../artifacts/plan-029-followup/final-gates/visual/injection-sentinel.html) | 3、1、3，标题含真实哨兵地址 | [三次出现](../artifacts/plan-029-followup/final-gates/visual/browser-injection-sentinel-full.png) |

合计 **7 文件、17 个页出现项**，全部完整页面截图均已目视复查。检查中文、英文比例间距、局部粗体/斜体/颜色、表格底色/边框、图像/裁剪、尺寸比例、页标签/重复顺序及异常空白/重叠；本次样例没有发现新增缺陷。`style-metrics.ttf` 中 ABC 的三个矩形 glyph 是测试字体的预期轮廓，不是缺字。420px 窄窗口下 SVG 为 388×548.7265625，与原 viewBox 比例一致，两页内容不重排：[尺寸](../artifacts/plan-029-followup/final-gates/visual/narrow-layout.json)、[截图](../artifacts/plan-029-followup/final-gates/visual/browser-native-narrow.png)。

SVG 根字体继承显式恢复初始值，实际 HTML 与独立 SVG 的 computed root/首文本 family 均为 `Times`：[独立 SVG 控制](../artifacts/plan-029-followup/final-gates/visual/standalone-font-control.json)。这只验证外壳继承隔离；未内嵌或不可用字体仍受浏览器、语言和宿主字体回退影响，不宣称其字形与 PDF 或任意宿主精确一致。

7 份实际浏览器打印 PDF 分别为 2、2、3、3、3、1、3 页，无额外打印页；其中 native 2 页与 mixed 的 3、1、3 共 5 张打印页 PNG 已逐张目视复查：[native 1](../artifacts/plan-029-followup/final-gates/visual/print-native-page-1.png)、[native 2](../artifacts/plan-029-followup/final-gates/visual/print-native-page-2.png)、[mixed 1](../artifacts/plan-029-followup/final-gates/visual/print-mixed-page-1.png)、[mixed 2](../artifacts/plan-029-followup/final-gates/visual/print-mixed-page-2.png)、[mixed 3](../artifacts/plan-029-followup/final-gates/visual/print-mixed-page-3.png)。仅覆盖本次 Chrome/A4 零边距设置，不推断所有打印设备或混合纸张设置。

### 可观察的零访问控制

在 `127.0.0.1:58460` 启动本任务独占监听哨兵，先打开独立可信 `/control`，脚本实际请求 `/positive`，CDP 与服务端均记录到了该请求。随后打开上述最终生成 HTML，注入标题含同一端口的 `/attack-script` fetch 与 `/attack-image` URL。结果：生成 HTML 的脚本节点 0、事件属性 0、外部请求 0；服务端 `/attack-*` **0 次**。不是只根据文本转义推断零访问。

证据：[正向 CDP events](../artifacts/plan-029-followup/final-gates/visual/positive-control-events.json)、[原始服务端请求日志](../artifacts/plan-029-followup/network-sentinel/requests.jsonl)、[最终汇总](../artifacts/plan-029-followup/final-gates/visual/network-summary.json)、[实际注入标题](../artifacts/plan-029-followup/final-gates/visual/injection-title.txt)。监听哨兵已停止。最终浏览器空间已关闭，保留原有用户空间；共享 GUI 租约已释放。

## PDF / PNG / macOS Preview 控制

实际链路为 **generated-layout.docx → 本次显式 native/default OFD → PDF → PNG / Preview**，不是直接 DOCX→PDF。另有 SVG 空白文字透明控制 OFD→PDF。PNG 自动渲染成功，native 1–2、default 1–2、blank-run-alpha 1 共 **5 张 PNG** 均已逐张目视复查；Preview 也对这 **5 页**实际打开、查看并关闭本人产物。

[Preview 原始验收记录](../artifacts/plan-029-followup/visual/preview-record.json) 每页记录完整真实 fileURL、页码 AX、原始 JPEG 路径/SHA 与 PDF 查看前后相同 SHA。截图按真实 JPEG 扩展保存，未重压。native/default PDF 均 183154 字节；OFD 均 170597 字节，最终 HTML 均 2424951 字节，主要增长来自 900780 字节字体逐页 Base64 内嵌，属于已记录的重复字体成本。

这些 OFD/PDF 控制在 `aa5005d…` 实施版本生成并完成 Preview；后续两提交仅修改 HTML shell/HTML PNG 准入与文档测试，未改独立 SVG/PDF/native 渲染路径。`OfdToSvgConverter.cs` 在 aa500 与最终源码的 blob 均为 `161ab52dfd1aaf7680cfaaacc1fd8fae3d3e7624`。它们是本任务本次生成并实际检查的控制产物，查看后字节一直未变；**不把它们描述为字体继承修复后重新生成的 HTML 截图**。最终新 HTML 的 17 个出现项另由上节独立新证据覆盖。

全部可复查产物/图像/TRX 的字节数与 SHA 见 [delivery manifest](../artifacts/plan-029-followup/delivery-manifest.json)。功能测试、自动 PNG 渲染、PNG 目视、macOS Preview 及最终 HTML 浏览器验收分别完成；仅对明确列出的样例和页面作结论。

## 剩余边界

本任务没有未闭合的必需实施/验收项。父协调会话仍负责五项共同评审、集成与合并后的联合验证；本任务不自行合并。产品范围保持首文档、受支持的静态绘制、有限预算与上述保守媒体 profile；不承诺任意复杂 Word/OFD、所有浏览器/打印机、原生可访问阅读顺序、外链交互或缺失字体的精确身份。默认预算可能因大字体或大量重复页拒绝输入，这是显式失败边界。
