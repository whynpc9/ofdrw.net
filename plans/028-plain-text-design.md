# Plan 028：复用 FlowDocument 的纯文本 OFD 入口设计

状态：**设计 DONE；生产实现、功能测试、包消费和视觉验收均未执行**。所有新 API、命令、测试类型都是推荐提案，不是当前可调用能力，也不代表真实调用方已批准这些产品选择。

源码核对基线：`ec121b087e095db4b4325cd4d61e61cbb4372142`，2026-10-15（包含 Plan 026 的路线图状态核对）；`FlowDocument` 相对原审查基线 `d0f88c7321a57443a4cb0ec92631664f66ed8cc0` 未变。[路线图 P1-08](../docs/capability-roadmap.md) 明确写着“Flow 前置具备，适配器未实现”。任务原计划在操作者根 checkout 的 `plans/028-plain-text-ofd-design.md`，不在本 frozen base 的 Git tree 中；本次只新增本文，索引由协调者维护。不修改生产源码、依赖、目标框架、包或发布状态。

## API 与边界

### 推荐 API 归属（未实现）

在现有 **Ofdrw.Net.Layout** 包和程序集、`Ofdrw.Net.Layout` 命名空间增加以下类型；不新建 Converter.Text 包，不新增通用转换框架。适配器最终创建一个 `FlowDocument`、一个 `Paragraph`、一个 `Span` 并调用现有 `FlowDocument.Render(token)`。已有 rich Flow 调用方仍直接使用 Paragraph/Span。

```csharp
// 提案签名，实施时补 XML 文档；当前不可调用。
public sealed class PlainTextOfdOptions
{
    public FlowDocumentOptions Layout { get; } = new FlowDocumentOptions();
    public long MaxInputBytes { get; set; } = 4_000_000;
}
public sealed class PlainTextOfdConverter
{
    public PlainTextOfdConverter();
    public PlainTextOfdConverter(PlainTextOfdOptions options);
    public OfdDocumentPackage Render(
        string text, CancellationToken cancellationToken = default);
    public Task<OfdDocumentPackage> RenderUtf8Async(
        Stream input, CancellationToken cancellationToken = default);
}
```

返回独立内存 package；调用方使用现有 [OfdPackageWriter.WriteAsync](../src/Ofdrw.Net.Packaging/OfdPackageWriter.cs) 保存。首版不提供目标文件/流、`TextReader`、可选编码、回调或页选择重载，避免薄入口自行建立文件事务。写任意流时 writer 仍可能留下部分字节，调用方负责 staging；CLI 使用既有原子文件提交。

构造时验证并复制全部 `FlowDocumentOptions` 属性值及 `MaxInputBytes`（DefaultColor 是现有不可变 OfdColor，可安全保留引用），外部后续修改不影响 converter。每次调用创建独立 Flow/package，不缓存输入、页面或字体资源，使同一实例的调用不共享可变渲染状态。null 参数抛 `ArgumentNullException`，不可读流/非法配置抛 `ArgumentException` 或更具体的 `ArgumentOutOfRangeException`。运行入口先检查取消；选项构造验证先于输入读取。

### UTF-8、BOM、换行及空输入

- 字节入口固定严格 UTF-8，采用 `new UTF8Encoding(false, true).GetDecoder()`，分块保留 decoder 状态，EOF 时 flush。非法续字节、过长编码、截断序列和编码的代理码点全部转为带 inner exception 的 `InvalidDataException`；不插入 U+FFFD。合法字符跨 read 边界必须成功。
- 只剥离本次流读取起始的一个 `EF BB BF`。不自动识别 UTF-16/32，不支持 GBK；调用方可自行明确解码后使用 string API。第二个 BOM、正文 U+FEFF 以及 string API 起始 U+FEFF 都作为内容保留。
- 从流当前 Position 读到 EOF，支持 non-seekable 流，不访问 Length、不复位、不关闭调用方流。失败/取消允许输入已前移，不承诺重试和同流并发读取。
- CRLF、CR、LF 由 Flow 规范化成一个逻辑换行。适配器不 Trim、不按行 Split 成多个 Paragraph、不额外插段距、不改变 Unicode 组合形式。连续换行保留正文间空行；末尾换行沿用 Flow 延迟空行规则，不增空尾页。
- 空流、BOM-only、空字符串生成一张空白页，不抛无内容错误。仅换行/仅 form feed 不强制物化空尾页；空格和 tab 是内容，不属于空输入。
- tab U+0009 保留原字符，沿用 Flow 固定四个空格 advance，**没有 tab stop 对齐**。U+000C form feed 是分页控制：有后续正文才物化分页，连续 FF 可形成中间空页，尾部 FF 不增空尾页。U+2028/U+2029 作为普通文本标量，不另行解释为换行。
- string API 拒绝孤立代理；两入口都拒绝 XML 1.0 无法表示的标量（如 NUL、U+000B、U+FFFE/U+FFFF），已由 Flow 消费的 FF 除外。支持 TAB/LF/CR 和有效 XML 标量。拒绝用 `InvalidDataException`，不删字符或修补代理，避免返回 package 后才由 XML writer 意外报错。

### 限额、取消及失败

建议沿用 [FlowDocumentOptions](../src/Ofdrw.Net.Layout/FlowDocument.cs) 默认 `MaxCharacters=1_000_000`、`MaxPageCount=10_000`、`MaxTextElements=1_000_000`，允许调用方调低。`MaxTableCells` 保留在 Layout options 中但入口不创建表格，不另建表格预算。

推荐 `MaxInputBytes=4_000_000`，只用于字节入口，按实际读取计数且包括 BOM。字符按剥离字节 BOM 后、换行规范化前的 UTF-16 code unit 数计（CRLF 和补充平面字符各计 2）；string 入口不虚构编码字节额度。先限制读取与解码，再调用 Flow，而不是全量读取后才依赖 Flow。用固定小块、long/checked 累计；到字节边界最多多读 1 字节区分恰好 EOF 与超限，不按声明长度分配大数组。字符超过额度立即中止。

输入字节/字符/页数/TextElement 超限，以及单字形、不换行单元或行高无法容纳时，入口抛 `InvalidDataException`。只在调用 Flow 的窄边界包装用于限额/无法布局的 `InvalidOperationException` 并保留 inner；Flow 自身现有异常契约不改。参数和 I/O 异常不统一吞成输入错误。

取消检查覆盖入口、每次 ReadAsync、解码/标量验证块、构造 Flow 前后、Render 前后；令牌传入 Flow。`OperationCanceledException` 原样传播，不重试、不返回部分 package。不承诺硬实时终止：流需配合取消，Flow 当前会物化段落/glyph 列表，额度限制工作量而非保证常量内存。不引入第二个 streaming paginator。

### 页面、字体与兼容性

几何及字号均为毫米，默认 A4 `210×297`、四边 `25.4`、字号 `10.5 * 25.4 / 72`、`SimSun`、黑色、左对齐，直接来自 Flow。宽高/字号必须有限且大于 0，边距有限且非负，边距和小于对应页面边长。几何合法但行高/字形放不下仍是输入失败，不裁切降级。首版不开放段落对齐、最小行高、Markdown、富文本、图片、表格、页眉页脚或语法着色。

字体只配置 `Layout.DefaultFontFamily`。`FlowDocumentRenderer.GetFont` 当前创建 name-only `OfdFontResource`，`FlowLatinMetrics` 使用 portable metrics；家族名不会触发真实字体测量、字体文件嵌入或字形覆盖验证。首版不增加字体扫描、`--font-directory` 或字体文件参数，不承诺任意字体精确 advance 或跨设备相同外观。调用方如自行向 package 字体资源填入 Data，仍走现有 [OfdFontEmbeddingOptions](../src/Ofdrw.Net.Core/Models/OfdFontEmbeddingOptions.cs) 和 writer；不建立第二种字体通道。实际 CJK 缺字需由真实页面验收暴露。

沿用 [Layout 工程](../src/Ofdrw.Net.Layout/Ofdrw.Net.Layout.csproj) 的 `netstandard2.0;netstandard2.1`。使用数组版 `Stream.ReadAsync(byte[], int, int, CancellationToken)`、`Decoder.Convert`，异步 `ConfigureAwait(false)`；避免新 TFM-only overload。[global.json](../global.json) SDK 10.0.100/latestFeature、CLI/tests 的 net10 不构成 .NET Framework、Mono 或旧 runtime 兼容证明。

### 推荐 CLI（未实现）

```text
ofdrw text-to-ofd <input.txt> <output.ofd>
ofdrw text-to-ofd --input <input.txt> --output <output.ofd>
  [--font-family NAME] [--font-size-mm NUMBER]
  [--page-width NUMBER --page-height NUMBER]
  [--margin-left NUMBER] [--margin-right NUMBER]
  [--margin-top NUMBER] [--margin-bottom NUMBER]
  [--max-input-bytes INTEGER] [--max-characters INTEGER]
  [--max-pages INTEGER] [--max-text-elements INTEGER]
```

新增显式 `text-to-ofd`，首版不扩展 `convert` 推断，不支持 stdin/stdout/`-`。输入扩展名不限但仍必须严格 UTF-8，输出要求 `.ofd`。宽高配对，未指定的边距各自使用默认。数值按 InvariantCulture，边距可为 0，其他正数；NaN/Infinity、重复选项、混用位置与命名路径、未知选项、多余参数在打开输入/staging 前拒绝。`--pages`、DOCX engine/mode 和 `--font-directory` 均拒绝，不静默忽略。

参考 [Cli.Images.cs](../src/Ofdrw.Net.Cli/Cli.Images.cs) 的独立命令 parser；[Program.cs](../src/Ofdrw.Net.Cli/Program.cs) 增加分发/help，接入 Plan 004 完成后的路径身份保护，不复制当前仅 full-path 比较的 `EnsureDifferentPaths`。读输入、生成完整 package、写 [AtomicOutput](../src/Ofdrw.Net.Cli/AtomicOutput.cs) 临时流、`CommitAsync(token)`，之后才发成功消息。失败或提交前取消保留旧目的文件、清理 staging；提交后到来的取消不能撤销文件替换。沿用 Cli.RunAsync 退出码：0 成功、1 错误、130 令牌取消。

## 现有能力复用

| 能力 | 已核对的真实符号/文件 | 最小新增点 |
| --- | --- | --- |
| 原生文本生成 | [FlowDocument.cs](../src/Ofdrw.Net.Layout/FlowDocument.cs)：`FlowDocument.Render`、`Paragraph(string)`、`Span.Text`、`FlowDocumentRenderer.DrawLine` | 一个完整字符串送单 Paragraph，不栅格化、不隐藏叠加第二份正文 |
| 换行/单词/grapheme | [FlowParagraphLayout.cs](../src/Ofdrw.Net.Layout/Internal/Flow/FlowParagraphLayout.cs)：`FlushText`、`Layout`、`IsNonbreakingSpace` | 不重写布局算法，不 Split/Trim |
| 分页/额度/空尾页 | `FlowDocumentRenderer.ValidateOptions`、`Place`、`StartPage`、`FlowPagination.NeedsNewPage` | 读取前额度、options 快照，保留 Flow 实际页数语义 |
| 比例英文/tab | [FlowLatinMetrics.cs](../src/Ofdrw.Net.Layout/Internal/Flow/FlowLatinMetrics.cs)：`AdvanceMillimeters`、`FlowTextMetrics.NormalizeLatinForMeasurement` | 不把 measurement 的 NFC 当正文改写 |
| 基线行为测试 | [FlowDocumentTests.cs](../tests/Ofdrw.Net.Layout.Tests/FlowDocumentTests.cs)：`TrailingControlSequence_DoesNotCreateBlankTailPages`、`CrossSpanCrLfAndGrapheme_UseLogicalParagraphStream`、`PublicFlow_MultipageStyledText_RoundTripsWithoutCoordinates` | 新入口验证这些映射，不再造测试分页器 |
| 保存/回读 | [OfdPackageWriter.cs](../src/Ofdrw.Net.Packaging/OfdPackageWriter.cs)：`WriteAsync`；[OfdReader.cs](../src/Ofdrw.Net.Reader/Readers/OfdReader.cs)：`ReadAsync` | 使用实际 OFD 往返检查 TextObject |
| 阅读抽取/PDF | [OfdTextExtractor.cs](../src/Ofdrw.Net.Reader/Extraction/OfdTextExtractor.cs)：`ExtractPages`/`Extract`；[OfdToPdfConverter.cs](../src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs)：`ConvertAsync` | 阅读文本和原文口径分开；由实际 OFD 导出 PDF |
| CLI 事务/取消 | `Cli.RunAsync`、`AtomicOutput.CommitAsync`；[OutputSafetyTests.cs](../tests/Ofdrw.Net.Cli.Tests/OutputSafetyTests.cs)：`CanceledCommit_ShouldDiscardPartialOutput`、`CanceledCommand_ShouldReturnCancellationExitCode` | 补新入口错误/中途取消，不建新提交机制 |

“原始文本完整保留”明确如下：输入 BOM、CRLF 原字节形式不进入 OFD；LF/FF 作为布局控制，不写可见 TextCode。定义 N 为去掉字节 BOM 并将 CRLF/CR 规范化为 LF 的字符串。按生成的页/元素顺序拼接 `OfdTextElement.Text`，必须 ordinal 等于 N **只去掉 LF/FF** 的结果；空格、tab、NBSP、组合标记、补充平面字符全部保留，保存回读后再比较。不能用现有测试的 `Compact` 或 `char.IsWhiteSpace` 丢空白后声称全保留。LF/FF 单独由坐标/页数验收，不能只靠去控制符比较。

`OfdTextExtractor` 按可见元素坐标重建阅读行，以 Environment.NewLine/默认 FF 页分隔；自动折行增加新换行、无墨迹空行可消失、末尾换行不可恢复。因此 `extract-text` 仅验证已知样例阅读顺序，不是逐字节恢复器。PDF 抽取另查无重复正文。首版不承诺原文件字节、换行种类/尾换行回读；如真实需求要求，应另立源文件保存契约，不偷偷添加隐藏文本层。

## 样例与预期

### 候选调用场景与调用示例

候选场景：UTF-8 日志/通知正文转原生 OFD；应用把自己拥有的 Unicode string 生成为可分页 OFD。尚无实际调用方的规模、字体和分页承诺，不据此推出 DOCX 替代功能。以下均为未来示例，当前新类型不能编译调用：

```csharp
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;
var options = new PlainTextOfdOptions { MaxInputBytes = 100_000 };
options.Layout.PageWidthMillimeters = 210;
options.Layout.PageHeightMillimeters = 297;
options.Layout.DefaultFontSizeMillimeters = 12 * 25.4 / 72;
options.Layout.DefaultFontFamily = "SimSun";
options.Layout.MarginLeftMillimeters = options.Layout.MarginRightMillimeters = 15;
options.Layout.MaxCharacters = 50_000;
options.Layout.MaxPageCount = 100;
var converter = new PlainTextOfdConverter(options);
var package = converter.Render("通知 Alpha iii WWW\r\n第二行 e\u0301\n\n尾行", token);
// 或 await converter.RenderUtf8Async(utf8Input, token);
// token、utf8Input、stagingOutput 由调用方提供；库不提交文件。
await new OfdPackageWriter().WriteAsync(package, stagingOutput, token);
```

CLI 提案：`ofdrw text-to-ofd notice.txt notice.ofd --font-family SimSun --font-size-mm 4.233333 --page-width 210 --page-height 297 --margin-left 15 --margin-right 15 --max-pages 100`。省略参数即默认；locale 不改变小数格式，`4,2` 拒绝，`4.2` 接受。`--pages 1` 失败且不替换既有 OFD。

### ContractExamples / AcceptanceMatrix

以下全部是**实施预期，实际未执行**。默认 A4 用于内容样例；几何 G 固定宽 40、高 30、四边距 5、字号 10 mm，内容区高 20 mm，Flow 行高 13 mm，每页恰能放一行。精确页数用可计算小样例，不猜综合样例页数。

| 输入/场景 | 预期文本、页数或失败 | 拟新增验证入口；实际 |
| --- | --- | --- |
| 无 BOM/单 BOM `中英 Alpha`，合法字符每字节分片 | 文本相同、1 页；BOM 不入正文；跨 read 字符成功 | `PlainTextOfdTests.Utf8BomAndChunkBoundaries`；未执行 |
| 空流/BOM-only/空字符串/仅 LF、CRLF、FF | 1 空白页、零文本元素；高数量空行仍受输入额度 | `EmptyInputUsesOneBlankPage`；未执行 |
| G：`A\r\nB` / `A\rB` / `A\nB` | 拼接 AB，均 2 页，A/B 均在对应页 top=5；阅读抽取可为 A FF B | `NewlineVariantsShareFlowPagination`；未执行 |
| G：`A\n\nB` / `A\n` / `A\n\n` | 分别 3/1/1 页，第一例中间页空白，尾部换行不增页 | `InteriorAndTerminalBreaks`；未执行 |
| G：`A\fB` / `A\f\fB` / `A\f` | 分别 2/3/1 页；中间空页与尾 FF 不物化 | 同上；未执行 |
| `  中英 iii WWW\tX\u00A0Y e\u0301  ` | 所有空白/组合文本 ordinal 保留，默认 1 页；视觉检查比例/tab advance | `WhitespaceAndGraphemesRoundTrip` + PNG/Preview；未执行 |
| emoji/补充汉字 | 原生 Unicode 字符 ordinal 保留；不保证 emoji 实际字体覆盖 | 同上；未执行 |
| 无空格 200 个 W、多段中英 | 窄页多行/多页，不裁字不丢字；与同 options 单 Paragraph 直接 Flow 一致 | `LongLineMatchesFlow`；未执行 |
| G，MaxPageCount=1，`A\nB`；小 TextElement 上限 | InvalidDataException，inner 保留 Flow 错误，无 package；CLI 1、不提交 | `OutputLimitsFail` + CLI safety；未执行 |
| C3 28、EOF E4 B8、ED A0 80 | 非法 UTF-8 严格拒绝，无 U+FFFD | `InvalidUtf8FailsWithoutReplacement`；未执行 |
| FF FE BOM、NUL、VT、孤立代理 string | InvalidDataException，不换编码/删字符 | `UnsupportedEncodingAndControls`；未执行 |
| MaxInputBytes=3：ABC/ABCD（含 non-seekable） | 恰好成功/超限，不依赖 Length，流保持打开 | `InputBudgetBoundaries`；未执行 |
| MaxCharacters=2：😀 / 😀A / `A\r\n` | 成功/拒绝/拒绝，按规范化前 UTF-16 长度 | 同上；未执行 |
| 宽 10、左右边距 5，负数/NaN/Infinity | 参数错误，读取之前失败；合法几何但行太高则输入失败 | `InvalidOptionsBeforeRead`；未执行 |
| 预取消/读中取消/Render 中取消/提交前取消 | OCE / CLI 130，无 package、不覆盖旧文件、无 staging 残留；受控流而非睡眠竞态 | `CancellationDoesNotPublish` + existing commit test；未执行 |
| CLI 同路径/别名、重复选项、--pages、宽高不配对 | 错误 1、不覆盖输入；接入完成后的路径保护 | `PlainTextCommandTests` + OutputSafetyTests；未执行 |

边界均用小额度和短输入；超宽 NBSP 单元也用小样例，不实际创建千万页/百万对象作为“证明”。新增样例记录期望、实际、结论和复跑入口；当前实际统一未执行。

## 验收命令

### 本次文档检查

已通读任务计划、根 [AGENTS.md](../AGENTS.md)、本文引用的符号/测试/runner；目标文档在旧基线到当前 HEAD 无历史差异，工作区/暂存区无该文件的既有改动。基线 Git tree 没有 plans 目录，因此只创建本文需要的目录，不复制其他计划。完成后实际执行下列结构/链接/空白检查，结果交协调者；结构通过不代替逐节内容核对。

```bash
python3 - <<'PY'
from pathlib import Path
import re
p = Path('plans/028-plain-text-design.md')
s = p.read_text()
for h in ['## API 与边界', '## 现有能力复用', '## 样例与预期',
          '## 验收命令', '## 风险与未决项', '## 后续实施切片']:
    assert h in s, h
for target in re.findall(r'\]\(([^)]+)\)', s):
    if '://' not in target and not target.startswith('#'):
        assert (p.parent / target.split('#')[0]).resolve().is_file(), target
print('design structure and local links OK; content reviewed separately')
PY
git diff --check
git status --short
```

本次实际结果：6 个必需标题及 26 个相对本地文件链接全部通过；已逐节人工复核契约与源码，`git diff --check` 和暂存差异空白检查退出 0，差异只有本文。未运行 dotnet，未生成 OFD/PDF/PNG；没有把设计、历史产物或源码核对当视觉验收。

### 未来功能命令（全部未执行）

从后续实施 worktree 根执行，使用 global.json 对应 SDK，按 AGENTS.md 设置可写 CLI home、单节点和关闭 build servers。先 restore；不清全局 NuGet cache。以下 filter 包含尚不存在的新测试，必须发现非零用例并记录实际数量，零命中不算通过。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-028/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1
mkdir -p "$DOTNET_CLI_HOME"
dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build src/Ofdrw.Net.Layout/Ofdrw.Net.Layout.csproj -c Release -f netstandard2.0 --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build src/Ofdrw.Net.Layout/Ofdrw.Net.Layout.csproj -c Release -f netstandard2.1 --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --no-restore --filter 'FullyQualifiedName~PlainTextOfdTests|FullyQualifiedName~FlowDocumentTests' --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --no-restore --filter 'FullyQualifiedName~PlainTextCommandTests|FullyQualifiedName~OutputSafetyTests' --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
```

全套回归保留，因为入口进入共享布局/转换链路。IPC SocketException(13) 按 AGENTS 关 build servers 后重试；缺 restore assets 不是产品新回归。netstandard 编译和 net10 测试不证明旧 runtime；如需支持声明，另在明确目标环境跑真实包消费者。

### 未来产物、包消费及视觉门（全部未执行）

后续 fixture 在 artifacts/plan-028 写入确定性 notice.txt 和 G 样例；新命令 build 完成后才执行下列命令。这里查看的必须是新适配器产物，不能用直接 text-to-PDF 代替。

```bash
dotnet run --project src/Ofdrw.Net.Cli/Ofdrw.Net.Cli.csproj -c Release --no-build -- text-to-ofd artifacts/plan-028/notice.txt artifacts/plan-028/notice.ofd --font-size-mm 4.233333
dotnet run --project src/Ofdrw.Net.Cli/Ofdrw.Net.Cli.csproj -c Release --no-build -- extract-text artifacts/plan-028/notice.ofd artifacts/plan-028/notice-extracted.txt
dotnet run --project src/Ofdrw.Net.Cli/Ofdrw.Net.Cli.csproj -c Release --no-build -- ofd-to-pdf artifacts/plan-028/notice.ofd artifacts/plan-028/notice.pdf
pdftoppm -png -r 144 artifacts/plan-028/notice.pdf artifacts/plan-028/notice-page
pdftotext -layout artifacts/plan-028/notice.pdf artifacts/plan-028/notice-pdf-text.txt
open -a Preview artifacts/plan-028/notice.pdf
```

逐页 PNG 检查字形/英文间距/基线/空行/分页/裁切/重叠/异常空页，Preview 逐页检查本次 PDF。重新生成后关闭重开，验收后关闭窗口。链路记录 **UTF-8 → adapter → native OFD → OFD-to-PDF → macOS Preview**；OFD 文本按 ordinal 口径验证，PDF 抽取另查无重复正文。记录实际字体环境，缺字就是视觉失败，不以家族名存在推断字体可用。

共享布局回归还须复跑现有 [generated-layout.docx](../e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx) 的显式 `--docx-ofd-mode native` 与默认 `docx-to-ofd`，从本次 OFD 分别导出 PDF/逐页 PNG/Preview，涵盖粗斜体、颜色、表格底色边框、对齐、分页，以及样例中的图片/页眉页脚；原文和页面选择沿用既有 DOCX E2E 门。不能拿直接 DOCX-to-PDF 替代。纯文本不新增 rich styles，却保留这条共享回归门。

包消费在 [Converter.Pdf.E2E/Program.cs](../e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs) 和 [run-converter-package-e2e.sh](../scripts/run-converter-package-e2e.sh) 添加纯文本新 API/CLI 消费后运行。当前 runner 的 public Flow/Native DOCX 不能自动证明新入口：

```bash
# 仅未来实施阶段生成本地候选包；示例私有版本，按真实候选替换。
bash scripts/run-converter-package-e2e.sh 0.0.0-plan028
# 已有本次候选包时只消费；不上传/发布。
bash scripts/run-converter-package-e2e.sh 0.0.0-plan028 --consume-only --packages-dir artifacts/package-e2e/0.0.0-plan028/packages --output-dir artifacts/plan-028/package-output
```

本次没有 pack 或运行 E2E。后续记录源码基线、样例、模式、环境/字体、实际链路、文件哈希、测试数量、看过的页、遗留问题和明显体积变化，复用任务 artifacts，不建新证据框架。Linux 无 GUI 时完成可执行测试/逐页 PNG 复查，Preview 标未完成；完整产品视觉门保持打开。功能测试、自动渲染、PNG 视觉、Preview 四项分别报告。

## 风险与未决项

| 项目 | 推荐方案/代价 | 实施前待确认事项 |
| --- | --- | --- |
| 实际调用方与规模 | 日志/通知为候选，沿用 Flow 限额；单段物化，不建 streaming paginator | 没有真实样本、内存 SLA/时延/输入规模，需取得样例并确定是否调低默认值 |
| 原文含义 | 保存全部非 LF/FF 正文字符和空白，控制符靠页坐标验证 | 如要求 CRLF/尾换行/原字节回读，当前方案不足，另定源文件保存契约 |
| 字体 | Flow name-only 与 portable metrics，不扫描/嵌入新字体 | 部署字体、许可、缺字行为、精确任意字体度量均未确认；若必需，先解决共享字体契约并重估范围 |
| FF/tab/U+2028 | FF 分页、tab 固定 advance、U+2028/2029 作为内容 | 属于推荐兼容规则，不代表调用方已要求；tab stop 等应另立小契约 |
| API 归属 | Layout 无新包，依赖与 Flow 一致 | 公共命名/归属仍需后续 API review；不把设计提案当批准的新 API |
| CLI 路径身份 | 串行整合 Plan 004 共享保护 | 当前 EnsureDifferentPaths 仅词法 full-path，不能声称别名防护已完成；实施时复核实际结果 |
| runtime | 两个 netstandard 目标使用共同 API | 未知目标 .NET Framework/Mono/旧 .NET 列表，不以 SDK roll-forward 宣称支持 |
| 复杂脚本/emoji | 字符保留与页面视觉分开 | RTL/shaping、彩色 emoji、任意阅读器保真未承诺，实际样例有缺陷须报告 |

无硬前置不等于授权绕过产品决策。可选方向不阻塞修复批次；本设计不关闭历史页面门，不授权产品实现、合并或发布。后续触及共享 Flow/字体/converter 先重读当前源码与相应回归范围。

## 后续实施切片

以下全部未实现，新增路径是候选位置，不是现有文件。按可独立审阅单元推进，不修改原计划虚构完成状态。

1. **输入契约/API**：在 [Layout 工程](../src/Ofdrw.Net.Layout/Ofdrw.Net.Layout.csproj) 新增 `PlainTextOfdOptions.cs`、`PlainTextOfdConverter.cs`，必要时私有 `Internal/PlainTextUtf8Input.cs`。实现 options 快照、严格分块 UTF-8/BOM/标量校验、额度/取消，唯一渲染调用为 FlowDocument.Render。新增 `tests/Ofdrw.Net.Layout.Tests/PlainTextOfdTests.cs` 落实 ContractExamples，跑两个 TFM build 和 Layout filter。不要把新 decoder 塞入共享 Flow；退出条件为小样例与直接 Flow 等价、拒绝分支实际可测。
2. **保存回读/原生样例**：在新测试中使用 writer/reader 做 ordinal 空白文本往返、控制符坐标/页数、原生 TextObject 且无额外图像/重复正文断言，产生 G/notice 输入和产物。复用 [FlowDocumentTests](../tests/Ofdrw.Net.Layout.Tests/FlowDocumentTests.cs) 现有回归，不复制测试 paginator。发现 Flow 现有缺陷先报告并分开修复，不为适配器私改分页算法。
3. **CLI**：在完成路径修复的基线上串行修改 [Program.cs](../src/Ofdrw.Net.Cli/Program.cs) 的分发/白名单/help，新增 `src/Ofdrw.Net.Cli/Cli.PlainText.cs` 局部 parser，按工程约定需要时显式引用 Layout，复用 AtomicOutput。新增 `tests/Ofdrw.Net.Cli.Tests/PlainTextCommandTests.cs`，补 OutputSafetyTests 的失败/取消/别名用例；跑 CLI filter，确认既有 convert 行为不变。
4. **包消费/完整验收**：修改 [Converter.Pdf.E2E/Program.cs](../e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs) 添加 installed Layout 的 string/UTF-8 API 与新 CLI，必要时小改 runner。先全套回归，再消费本次本地候选包。执行上面的 OFD 文本、PDF、逐页 PNG/Preview 门，并复跑 generated-layout native/default 链路。仅实际检查的样例/页可报通过；缺字、重叠或 Preview 未完成不能称产品整体通过。
5. **文档/产品状态**：实施完成后才同步路线图 P1-08 和使用说明、字体/原文边界。版本/发布由单独授权管理；本次只交付设计状态给协调者，不标 P1-08 为源码已实现。
