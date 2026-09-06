# .NET builds in the Codex sandbox

- Before running `dotnet`, set `DOTNET_CLI_HOME` to a writable task/temp directory and set `DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1` and `DOTNET_CLI_TELEMETRY_OPTOUT=1`.
- Run sandboxed builds/tests single-node with `-m:1 /nodeReuse:false /p:UseSharedCompilation=false`; add `--disable-build-servers` when the installed SDK supports it (.NET 6 does not).
- If local IPC reports `SocketException (13): Permission denied`, treat it as a sandbox constraint, shut down build servers, and retry with the flags above.
- Do not routinely clear the global NuGet package cache; it is not the cause and may make restricted-network restores fail.

# 转换与渲染的视觉验证

- 测试或修改 DOCX/OFD/PDF 转换、排版、字体、表格或渲染功能时，必须验证实际生成的页面。自动化测试通过、文件生成成功、页数正确或文本可提取，均不能代替视觉验收。
- 使用包含中英文、比例字体、局部粗体/斜体/颜色、表格底色与边框、对齐及分页的确定性样例；新增功能或修复应补充对应样例。优先复用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`。
- DOCX→OFD 必须重点验证显式 `native` 和默认模式，确认原始文本完整保留，并检查分页选择。视觉验收必须基于本次生成的 OFD；如通过 PDF 查看，应采用 `DOCX → native OFD → PDF → Preview`，不得用直接 DOCX→PDF 的结果代替 native 产物。
- 使用 macOS Preview App 打开最终产物并逐页查看正文；OFD 经 PDF 导出后查看时，报告中明确说明查看链路。重新生成文件后，应重新载入或关闭后再打开，避免验收旧的缓存页面。
- 对照源文档或明确的样例预期，检查中文缺字/乱码、英文间距与断词、局部样式是否扩散、字号与基线、表格填充/边框、对齐、分页、裁切、重叠、重影和异常空白页。检查样例涉及的图片、页眉页脚及页码。必要时放大关键区域，并用页面 PNG 辅助复查。
- 修复视觉缺陷后，重新生成受影响产物并重复视觉检查；同时确认 OFD 原始文本完整，PDF 导出的样式补绘没有引入重复文本。为实际缺陷增加有意义的回归断言，运行受影响测试；涉及共享排版、字体或转换链路时，运行全套回归及本地包消费 E2E。
- 保存可复查的产物、页面图和验收记录（可放在 `artifacts/` 下），记录源码基线、样例、转换模式、实际查看路径、测试结果、已检查页面、遗留问题及明显的产物体积变化。只对实际检查的样例和页面作出结论，不推断任意复杂 Word 文档均能完整保真。
- 最终报告分别说明功能测试与视觉验收结果。发现视觉缺陷时不得宣称整体通过；若 Preview 或所需查看工具不可用，明确记录未完成的验收项及原因，不以自动化测试或替代渲染结果冒充 Preview 验收。
