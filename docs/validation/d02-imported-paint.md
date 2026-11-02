# D02：导入 OFD 重写保留 paint/CTM

日期：2026-11-02。范围为 [N01 的 C1](cn-en-release-n01.md#c1导入的-name-only-文本重写保留原-paintctm)。基线为最新 `origin/main`：`2c49e0eff115bc159df1f240e133938eb3f8f990`。

## 缺陷与修复

合法导入 `TextObject` 保留 `SourceXml` 时，writer 仍对 name-only 字体应用新建文本的 faux bold/italic。实际复现 `回A`、Size=4、Weight=700、LineWidth=0.023：线宽变为 0.18，并强制 Stroke/Fill；省略 CTM 的斜体还会增加矩阵和扩展 Boundary。九种既有反例在未修复的基线上全部失败（9/9）。

`ApplyEmphasis` 现在对非空 `SourceXml` 提前返回，让原 XML 保有填充、描边、颜色、透明度、线宽、既有 CTM 及省略默认值的语义。ID/Font 重绑定、现有模型 Weight/Italic 写入及裁切处理保持既有契约。新建文本仍执行原来的默认强调策略，嵌入字体仍遵守显式 opt-in。

复用独立的九种导入反例（thin、different-color、alpha、outline-only、no-stroke、italic-ctm、sim-sun-explicit、omitted-paint、italic-no-ctm），比较两遍写入/读回后的完整 XML。另补两例新建 name-only 粗体/粗斜体，断言原 0.18 线宽、颜色/alpha、CTM 及重写稳定性。没有新增字体或文档二进制样例。

## 验证身份

- 开发：gpt-6.1-sol High；独立静态审阅：gpt-6-astra High；本地功能和页面验证：gpt-6.1-terra High，独占 GUI。
- 审阅及本地验证绑定的产品/测试补丁 SHA-256：`43e64c56b0715ee9aecd0820d249b78055dc33373fa9362cc043eb3648d86b49`（基线上 `git diff --binary`，仅 writer 与 `EmbeddedBoldCompatibilityTests.cs`）。
- writer Git blob：`25ee06cb98d96bf41103ac33338f9da125d26d44`；测试 Git blob：`882a9dc208ec34aa52d8ae7a6ead81d701c7a52c`。后续仅验收报告编辑不改变产品身份。
- 所有 .NET 构建/测试通过共享 `run-exclusive.py` 串行执行，使用独立可写 DOTNET_CLI_HOME、禁用 telemetry/首次体验、单节点和禁用 build servers/shared compilation；复用全局包缓存。

## 功能与页面结果

| 项目 | 本次结果 |
| --- | --- |
| 定向 `EmbeddedBoldCompatibilityTests` | 16/16 通过；修复前九例全失败 |
| 完整 solution Release build/test | 构建通过；11 个测试项目，4623 通过、0 失败、17 个默认环境字体跳过 |
| 独立原字体门禁 | `run-trusted-font-tests.py --no-build`：74/74 通过，零跳过；覆盖默认全套的字体环境缺口 |
| 本地包消费 E2E | 独立未发布版本 `0.1.0-d02.local.1`，12 个包验证及消费通过 |
| 导入原始/两次重写 | 原始 TextObject 直接注入合法 OFD 容器后经 reader 读取；三轮均与原 XML 比较通过。三个 PDF 各 1 页，抽取文本均精确匹配九行、178 个非空白字符，无丢失/重复 |
| 自动渲染 | 三个导入 PDF 的 144 DPI PNG 字节完全相同，SHA-256 均为 `05f76e95cf1b4c08bdbea2ad9833d9f9f0d9079e7655839618b647a8b6847565` |
| DOCX native/default 控制 | 本次包消费生成的 OFD→PDF 各 2 页；OFD 原文与 PDF 均匹配源 DOCX 的 189 个非空白字符。DOCX→OFD 选第 2 页、OFD→PDF 选第 2 页分别验证；两模式 DOCX 选页 PNG 与全量第 2 页字节相同 |
| PNG 视觉复查 | 导入原始/两次重写 3 页，native/default 完整样例 4 页，DOCX 选页 2 页，共 9 页逐页复查 |
| macOS Preview | 导入原始/最终各 1 页、两模式完整样例各 2 页、DOCX 选页各 1 页，共 8 页逐页检查并关闭窗口；查看链路是本次 OFD→PDF→Preview，DOCX 源先经 native/default OFD |

导入九例使用独立名称资源（含 SimSun），通过显式 name-only binding 绑定仓库 `LXGWWenKai-Regular.ttf` 的 ASCII/回字 regular TrueType 子集；不代表实际 SimSun 选脸或真实粗体。源字体 SHA-256 `39ad71264b588165b469e35e6afb162a378dacd1f95348160240ba9038ac3009`，探针子集 `fbae2215b10bb98d5911fd509697d8d4cc43fef1c330327bde68a2ac0d38acd7`。只对指定九例、所用绑定及上述页面作结论。

DOCX 两模式采用 macOS 默认字体解析，PublicRes 的 PingFang SC regular/italic/bold 资源指向同一 regular 内嵌子集（SHA-256 `a7245b7b10db0d13ea023768e24ad3703481e374a1814f8b5d8ae402fa62a337`），另有 name-only `Noto Sans CJK SC|bold`。这不是 [preview.10 的真实 Bold 配置](../release-profile-cn-en.md)。本次仍观察到普通字号表头“量”拥挤，保留视觉 **FAIL**；不能将此轮控制页称为全部视觉通过。该限制已属于 preview.10 接受的边界，D02 不关闭它，也不据此展开光学算法；九例导入原始与两次写回的页面一致性独立成立。

本机为 macOS 26.6.2、.NET SDK 10.0.401。完整路径、产物字节数和 SHA-256 见 `artifacts/d02/terra-validation.json`，测试/页面/Preview 范围另见同目录 `terra-validation.md`。未进行任意复杂文档、原生 OFD 阅读器或 Copy/Find 的新验收。

## 边界

本次只修复导入 writer 的绘制语义覆盖；不搬运 PR78 的光学增强、producer、planner 或新绘制算法，不更改预览版 .10 标签或发布记录。原字体普通字号“量”拥挤、历史 Find 和通用复杂文档保真边界不由此补丁关闭。独立审阅未发现有效 correctness 意见，属于静态结论，不能替代实际功能/页面检查。

运行资料保留在本工作树忽略的 `artifacts/d02/`：修复前复现、定向测试、Astra 审阅记录、Terra 功能日志、原始/两遍重写 OFD/PDF、PNG、Preview 记录与哈希清单。没有提交 artifacts、构建目录、包缓存、源码副本或重复字体载荷。
