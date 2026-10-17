# DOCX 字体与粗体跨平台修复验收 — 2026-10-17

本记录合并两轮修复和独立复测结果。原始基线为 `f90b904ee6d71029bb6d30d12c5fd7ea941ff644`；字体身份修复为 `77825d3107f767e334cd80d6db5e651e958b23bf`，粗体修复为 `efb56431e05ba34dab554e9965079938634974dc`。Windows 门禁提交仅涉及验收入口、参考图和 CI；之后的 PR 审查补丁见文末。未发布 NuGet 包。

## 修复行为

- Native/default DOCX 使用配置的真实 Noto CJK 字体嵌入 OFD，并以同一物理字体面的度量计算拉丁字符间距。不会把其他中文字体伪装成 Noto；宋体/黑体的名称字体兼容路径保留。
- 第一轮的 PDF 正常，但 Suwell 直接打开 OFD 时，常规嵌入字体上的粗体强调过细。第二轮保留真实字体和语义文字，以 0.025em 的标准文字描边补充强调；字体资源标志描述实际字体面，文本 Weight 描述请求的样式。
- 默认策略返回 `DOCX_EMBEDDED_BOLD_SIMULATED`。需要真实粗体时设置 `NativeEmbeddedBoldBehavior = DocxEmbeddedBoldBehavior.Throw` 并提供粗体文件；缺少实际粗体面时在写出前失败。真实粗体面不额外描边，导入文字的原有绘制属性不被补绘覆盖。用法见[转换约定](../conversion-contracts.md#字体与宿主程序)。

## 固定提交复测

使用同一真实 Noto 字体，SHA-256 为 `3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa`。复用 `generated-layout.docx` 及仅修改字体名的派生样例；每个样例分别运行显式 native 和默认模式。

| 验收项 | 实际范围 | 结果 |
| --- | --- | --- |
| 受影响功能测试 | macOS、原生 Windows、WSL 各 13 项，零跳过 | 39/39 PASS |
| 源码转换 | 三平台 × 两个 DOCX × 两种模式，12 份 OFD、24 页 | 原文完整；字体身份、规范化结构一致 |
| 统一环境渲染 | 三平台 OFD 和本地导出的 PDF，48 项页面比较及局部检查 | RMSE 0；无重复文字 |
| 同批包消费 | 同一组 10 个库包，三平台各 4 组转换 | PASS；实际 DLL 与包哈希一致 |
| macOS Preview | 12 份最终 PDF，逐份重新打开并检查两页 | 24 页 PASS |
| Windows Suwell 5.0.22.0831 | 源码及包产物共 24 份 OFD、48 页，另查 4 处 200% 放大 | PASS；粗体强调恢复 |

查看链路为 **DOCX → native/default OFD → PDF → Preview**；Suwell 直接打开 OFD。检查了中文、比例英文间距、局部粗斜体与颜色、表格底色和边框、对齐、显式分页。样例未覆盖图片、页眉页脚和自动页码，不据此推断任意 Word 文档保真。

修复工作区另完成 116 项 DOCX、85 项 PDF 和 1 项名称字体回归，以及本地完整 11 包消费 E2E。实际 Noto 子集约使本样例 OFD 从 4.3 KB 增至 291 KB，这是嵌入真实字体的预期变化。两轮原始产物、哈希、页面截图和运行记录保存在已忽略的 `artifacts/`；日常仓库不重复归档源码、包或完整日志。

Windows 两个 Audit023 符号链接用例按本轮约定排除，测试代码和系统安全设置均未修改。这不影响上述字体与页面结论。

## Windows 页面回归门禁

四个样例为 ticket、invoice、template、edge-layout。Windows 使用已有可信字体 fixture 的真实 Noto 字节，并用同一字体的实际度量明确写出逐字符 DeltaX。首次捕获缺少字距，PDF 正常但 Suwell 直接查看文字重叠，未作为参考图接受；修正的是合成样例的定位信息。输出 **OFD → PDF → 144 DPI PNG**，逐页审查后才加入 `e2e/visual-corpus/golden/windows/`。整页、标题和正文区域的阈值保持 0.035；Linux/macOS 参考图不变。

完整 Windows DOCX 包消费也显式提供可信 Noto：样例请求 Noto，但隔离字体目录不会自动成为系统字体目录。未传目录时粗体被宿主解析为 Agency FB，原文映射按预期失败。验收配置补齐实际字体，不恢复跨字体伪别名，也不放宽文字映射。捕获模式只生成待审查文件，不代表验收通过，也不写入参考图。

- 原生 Windows .NET 10.0.401、Poppler 26.09.0、ImageMagick 7.1.2-31、librsvg 2.62.4；完整 11 包消费 **16 步 PASS**，含 CLI、主 consumer、字体子集 consumer 和最终包哈希校验。
- 四个合成样例的 PNG、Suwell 直接 OFD、配对 PDF 和 macOS Preview 均逐页通过。正式比较 **12/12 PASS，RMSE 全为 0**。缺图、纯白图两个负例均退出 1，恢复后的四图哈希一致；最后仅补 DOCX 字体配置，视觉门禁函数未变。
- 补齐字体后的 DOCX native/default/CLI 共 6 页，在 PNG、Preview 和 Suwell 检查通过；三组逐页 PNG 字节相同。BuiltIn PDF 另查 2 页，中文原文恢复。

**遗留视觉问题：BuiltIn DualLayer 的 2 页粗体强调偏细。** Suwell 直接查看和对应 BuiltIn PDF 均可见此差异，记录为视觉 FAIL。本次两轮产品修复限定于 Native 字体身份和强调，不宣称 DOCX 所有模式的整体视觉验收通过。这个问题不由四个合成样例的回归门禁覆盖，也未通过更新参考图或放宽断言接受。

精简原始记录位于本轮忽略目录 `artifacts/xplat/20261017-windows-visual-gate/`。四张 Windows 参考图通过 LFS 保存（合计 41,811 字节）；完整日志、字体副本、候选包和屏幕截图不加入日常源码树。CI 新增 Windows job，消费 Linux job 产出的同批 11 包字节并保留清单和页面；本地 Windows 候选版本为未发布的 `0.1.0-xplat-windows-gate.1`。


## PR 审查补验

- 已复现并修复组合样式选择：提供常规和粗体、缺少粗斜体时，粗斜体文字优先使用实际粗体字节，保留斜体请求；严格粗体策略不再错误拒绝已有粗体。复用小型样式 fixture 验证字节选择、资源标志及不叠加粗体描边，不将其冒充真实字体的视觉样例。
- Mac 完整 DOCX **294/294**、Windows 字体回归 **9/9** 和固定 Noto **2/2**、WSL 字体回归 **9/9** 通过，零跳过。Windows 四组规范结构和 8 张页面 PNG 与此前已审查产物相同；Mac 重新构建并消费 11 包通过，6 张 native/default/CLI 页面 PNG 也与已审查产物字节相同。既有 BuiltIn DualLayer 视觉 FAIL 保留。
- Windows CI 的可信字体清单原始哈希检查暴露了 CRLF 检出差异；通过 `.gitattributes` 固定该清单为 LF，不放宽精确字节校验。字体准备失败时停止后续测试。
- macOS CI 另暴露既有 KDF 取消测试的观察任务被线程池延迟。观察改用预先启动的独立线程，保留取消异常、后端释放、原文件未覆盖及无临时文件断言；密码产品实现未改，完整密码测试 98/98 通过。


后续审查补充了未配置 Noto 的边界：所有 `Noto Sans CJK*` 家族都必须实际注册后才嵌入；未注册时保留名称字体及 Arial 拉丁兼容度量。先复现 JP/TC/KR 的误嵌入及独立的窄字距失败，再验证修复。Mac 完整 DOCX **298/298**、可信字体 **68/68**、重新构建并消费 11 包通过；Windows 字体回归 **13/13** 加固定 Noto **2/2**、WSL **13/13** 通过，零跳过。Mac 6 张、Windows 8 张既有页面 PNG 与已审查版本相同，未改参考图。新增宿主替代字体测试不宣称缺少原字体的文档能跨平台保真。
