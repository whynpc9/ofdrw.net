# R01 PageBlock、文字与颜色验收

日期：2026-10-21。实现基线 `35976dbee1f58eef5fe19f2750746a97d523cfec`，分支 `codex/r01-pageblock-text-color`，.NET SDK 10.0.401，macOS。提交与 PR head 以 PR 记录为准。本目录仅保留报告和 [SHA-256 清单](hashes.txt)；本次生成的 OFD、PDF、SVG、HTML、PNG 和日志保存在已忽略的 `artifacts/roadmap-v12/R01/`，最终本地包消费输出在 `artifacts/package-e2e/0.1.0-r01.review2/output/`。

## 行为与功能测试

- Reader 对普通页和模板的嵌套 `PageBlock` 建立只读投影，保留原始 XML 作为保存来源。投影按源顺序包含 Path、Text、Image；根节点有未知绘图语义或投影失效时，PDF/SVG/HTML 明确拒绝。
- 有界解析 DrawParam 相对继承、8-bit GRAY/RGB、无 ICC 的简单 CMYK 和调色板。缺失或歧义引用、渐变/图案及不支持的文字语义拒绝；文字 Fill、Stroke、HScale 在 PDF/SVG/HTML 中分别呈现。签章内嵌 OFD 的已识别颜色失败会向外传播，不再静默略去章面。
- `dotnet test Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false`：11 个项目共 **3045 passed、13 skipped、0 failed**。定向 R01 与依赖回归通过；DOCX native/default 视觉控制 2/2 通过。使用共享独占构建锁及仓库要求的 `DOTNET_CLI_HOME`、首次运行/遥测环境变量。
- 本地打包版 `0.1.0-r01.review2`：11 个包校验及独立消费者 E2E 通过，包括 CLI、DOCX、PDF、图片、上游样本。该版本仅为本地消费验证，未发布。既有 ImageSharp 2.1.13 的 `NU1902/NU1903` 告警仍在，本票未改依赖。

## 本次页面视觉复查

| 样本/链路 | 实查范围 | 结果 |
| --- | --- | --- |
| `ordinary-block`、`template-block`：OFD → PDF/SVG/HTML/PNG | 各 1 页；PDF 与 SVG 的 PNG、HTML 浏览器截图、macOS Preview PDF | 蓝色方块与 `nested text` 顺序/位置一致；模板和普通页外观一致 |
| `palette-drawparam` | 1 页；同上 | 蓝色/青色描边方块、红色方块可见 |
| `stroke-hscale` | 1 页；同上 | 红色空心 `R01 text`，横向缩放可见；PDF 原文可抽取，无重复文字 |
| `image-block` | 1 页；同上 | 蓝色路径和红色图片方块按顺序可见 |
| `generated-layout.docx` → 显式 native OFD → PDF → macOS Preview；默认模式单独重复 | 两条链各 2 页，另核对选第 2 页产物的 PNG | 中文、英文、斜体、红色局部粗体、表格底色/边框与分页符合样例；未见缺字、裁切、重影或异常空白页；选页 PNG 与完整文档第 2 页字节一致 |
| 原始 `upstream-ofdrw/999.ofd`；另制 render-only 副本 → PDF | 原始签名 SM3 引用完整性与无效 ColorSpace 拒绝；副本 PDF 5 页逐页 PNG | 原件的嵌入章将 `ColorSpace=4` 指向字体资源，转换明确拒绝。仅在**签名已失效的呈现副本**把两处引用改成正确 RGB ID 5 后，章面与 5 页内容可见；副本不作签名有效性或原件互操作结论。此副本未做 Preview 验收 |

PDF 和 SVG 的页面 PNG 均从本轮生成文件渲染，HTML 五页由浏览器截图复查。macOS Preview 已重新打开本轮 PDF，检查 R01 五个 1 页样本及 DOCX native/default 各 1–2 页，检查后关闭 R01 窗口。最终 `review2` 签章副本的五页 PNG 与先前逐页查看的 `review1` PNG 字节一致；它仅作包消费回归与 PNG 辅助检查。没有独立第三方 OFD 阅读器互操作结论，也不将这些有限样本推广到任意复杂文档。

## 产物身份

完整输入/输出哈希见 [hashes.txt](hashes.txt)。关键输入：`generated-layout.docx` 为 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`；`999.ofd` 为 `a2b9080ae35184135160c48ce45fb609c35d5ecfdb538ebc16132c48a2186a80`。本轮独立 native/default DOCX 产物及选页产物的哈希均列于清单。页面图/浏览器截图在同一已忽略的本地目录，可按清单对照重新复查。
