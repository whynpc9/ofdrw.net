# D03：SVG 使用声明字体族

日期：2026-11-02。范围为 [N01 的 C2](cn-en-release-n01.md#c2svg-css-使用声明字体族避免泄漏资源身份键)。基线为最新 `origin/main`：`e14a7ac3f28efd7e1d7be58a39abd4057b8c7a41`，含 D02/PR98。

## 缺陷与修复

`OfdToSvgConverter.AddText` 原先直接将 `text.FontName` 输出为 CSS 字体族。name-only 资源将字体身份键与声明字体族分开时，例如 `FontName="Fixture Resource|bold|italic"`、`FamilyName="Fixture Family"`，SVG 错误输出前者，导致宿主无法按声明字体族选脸。

现在优先非空白的 `resource.FamilyName`；缺失或空白时仍使用 `text.FontName`。嵌入字体/显式绑定随后继续覆盖为既有 hash 家族，粗体/斜体、绘制、定位及字形算法均未改动。HTML 复用相同 `BuildSvg` 路径，获得同样修正。

回归经公开 OFD writer→reader→SVG 出口检查有/无 FamilyName 两例，均保留 `Ab`、bold 和 italic。复用现有嵌入字体测试，增加冲突声明家族并断言 hash 家族仍绑定 `@font-face`。没有新增二进制字体或样例。修复前先运行四例（两种输入构造），两例声明字体失败、两例 fallback 通过；独立审阅确认 writer/reader 会把两种构造统一为 Run，最终精简为两例，避免重复覆盖。

## 验证身份

- 开发：gpt-6.1-sol High；独立静态审阅：gpt-6-astra High；本地功能与页面验证：gpt-6.1-terra High，独占 GUI。
- 最终产品/测试 `git diff --binary` SHA-256：`ec4b9a70956a8eae9588cf426c04304e2614b60b22b45fe473278d796775181c`。
- converter Git blob：`d6861a639ff8319ec76378cbb7d3a45e26e503bd`；测试 Git blob：`93a3d750911afc3c1d69c033aacdffb472eb08c8`。报告编辑不改变产品身份。
- .NET 验证通过共享 `run-exclusive.py` 串行，使用可写任务目录、单节点、禁用 build servers/shared compilation；没有清除全局缓存。

## 功能与页面结果

| 项目 | 本次结果 |
| --- | --- |
| 最终精简定向回归 | 2/2 通过；初次受影响 SVG/HTML/字体组 58/58 通过 |
| 完整 solution Release build/test | 11 个测试程序集；4625 通过、0 失败、17 个默认字体环境跳过；独立原字体门禁 74/74 通过、零跳过 |
| 本地包消费 | 未发布版本 `0.1.0-d03.1`，12 个包验证及完整消费通过；HTML API/CLI 精准本地包消费通过 |
| 目标样例 | 同一两页 OFD 实际导出 SVG/HTML；声明可用字体、声明缺失字体、无声明 fallback、嵌入 hash 与中文控制 |
| 宿主实际选脸 | 浏览器 CDP 确认可用族使用 `TimesNewRomanPS-BoldItalicMT`；不存在的 `Fixture Family` 在 SVG 使用宿主 `Times-BoldItalic`、HTML 使用 `PingFangSC-Semibold`，回退依上下文不同；fallback 使用 `TimesNewRomanPSMT`；嵌入资源为 custom hash 家族 |
| 目标页面 | 两页 SVG PNG 逐页检查，SVG→浏览器打印 PDF→macOS Preview 两页正文已查看并关闭；主线程另复核两页 PNG |
| 原始文本 | 目标 OFD/SVG/HTML 每页五个文本对象，逐项相符，无重复节点；打印 PDF 前两条 fill+stroke 粗斜行重复提取，普通/嵌入/中文行不重复。隔离基线编译后对同一输入生成的两页 SVG，仅前两条 font-family 属性改变；基线与当前经同一浏览器打印后的 `pdftotext -layout` 两页逐字节相同，重复在基线已可复现 |
| DOCX 控制 | 复用 `generated-layout.docx`，显式 native/default 各两页 OFD→SVG/HTML；源 DOCX 全部文字与两模式 OFD 一致，各页 OFD/SVG/HTML 文本序列相等（14、4），无重复节点；两个模式各两页 SVG/HTML PNG 和对应 SVG→浏览器打印 PDF→Preview 正文均逐页检查，窗口已关闭 |

目标样例 SVG 两页、HTML 两页，DOCX 两模式 SVG 四页、HTML 四页，共 12 个页面截图逐页检查；macOS Preview 共 6 页正文，链路为本次 OFD→SVG→浏览器打印 PDF→Preview，未用直接 OFD→PDF 代替。主线程另复核目标 SVG 两页及 native 控制两页 PNG，并独立比较基线/当前打印 PDF 文本。补验时用户明确授权替代浏览器空间；验收结束已关闭本次空间及 Preview 窗口。

两种 DOCX 模式的第一页表格填充/边框、中文/英文，第二页红色粗体与右对齐日期均已查看；普通字号表头“量”在 SVG/HTML PNG 与 Preview 中仍拥挤，保留视觉 **FAIL**。这是 D02 已记录的既有默认字体配置限制，本次不关闭该缺陷或宣称完整视觉通过。目标样例的 `narrow.ttf` 嵌入负控是合成矩形字形，色块为该字体设计，不能称可读正文。

独立 Astra High 最终审阅未发现阻断性 correctness 问题，测试冗余意见已处理并复核。功能、包消费及目标字体族选择验收通过；所需页面检查已完成，控制样例保留上述既有视觉 FAIL。静态审阅不能替代实际页面检查，本轮也不宣称任意文档整体保真。

## 边界

仅修复 CSS 字体族选择；不搬运 PR78 光学增强或扩展 SVG 粗体算法，不更改预览版 .10 标签、包或发布记录。name-only 字体实际可用性和字体替换由宿主决定；本补丁不承诺缺失字体时视觉保真。嵌入 hash 家族及现有严格绑定机制保持原契约。

本机为 macOS 26.6.2 arm64、.NET SDK 10.0.401。详细命令、结果及范围见 `artifacts/d03/terra-validation.md` / `.json`；输入、输出字节数与 SHA-256 见 `terra/hashes.json`、`terra/layout-hashes.json`、`terra/continuation-hashes.json`；独立基线/当前文本比较另见 `terra/baseline-print-text-comparison.json`。

本次运行产物和详细证据位于本工作树忽略的 `artifacts/d03/`；不提交运行日志、包缓存、构建目录或源码副本。
