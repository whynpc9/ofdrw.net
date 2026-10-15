# 029 有界多页 SVG HTML 预览设计

状态：**设计已交付；生产实现、功能测试、浏览器验收、PNG 复查、macOS Preview 验收及发布均未执行**。本方向为可选 D02，不阻塞缺陷修复批次。本文所有新增类型、选项、内部入口、测试名和 HTML 命令均为建议，当前不存在；示例不是产品输出。

设计源码基线：`ec121b087e095db4b4325cd4d61e61cbb4372142`（2026-10-15）。原计划编写基线为 `d0f88c7321a57443a4cb0ec92631664f66ed8cc0`；指定交付文件及本文核查的 SVG、Reader、CLI、AtomicOutput 在两基线间没有已提交差异。按协调者要求，仅新增本文及 `artifacts/plan-029/`，总索引由协调者维护。输入计划 029 和关联计划 014/015 从主 checkout 只读读取；它们在所选基线的隔离 worktree 中不属于已跟踪文件。

## API 与边界

### 公开 API 建议

建议复用现有 `Ofdrw.Net.Converter.Svg` 程序集，新增 `OfdToHtmlConverter` 与 `OfdHtmlOptions`，不新增 NuGet 包、不改变目标框架。以下是建议签名，**未实现、未编译**：

```csharp
namespace Ofdrw.Net.Converter.Svg.Converters;

public sealed class OfdToHtmlConverter
{
    public OfdToHtmlConverter();
    public OfdToHtmlConverter(OfdHtmlOptions options, OfdPackageLoadOptions? loadOptions = null);
    public Task ConvertAsync(Stream ofdInput, Stream htmlOutput,
        IReadOnlyList<int>? pageIndexes = null,
        CancellationToken cancellationToken = default);
}
```

公开成员必须补 XML 文档。`OfdHtmlOptions` 建议属性及有限默认值如下；这些是首版产品预算提案，尚未由真实语料校准，不是当前可配置项。

| 属性 | 建议默认值 | 契约 |
| --- | --- | --- |
| `Title` | `"OFD preview"` | 只当纯文本使用；不自动采用不可信 OFD 元数据；允许空标题 |
| `MaxTitleCharacters` | 1,024 | UTF-16 code unit 上限；无效代理项、XML/HTML 不允许的控制字符拒绝 |
| `MaxPages` | 200 | 输出页出现次数上限，重复选页逐次计费 |
| `MaxOutputBytes` | 64 MiB | 从 `<!doctype html>` 至末尾的累计 UTF-8 字节，含 shell、标题、字体、图片、CSS、所有 SVG 与重复页；不含 BOM，因为不输出 BOM |
| `MaxPageBytes` | 16 MiB | 单个页 section 的 UTF-8 字节，包含它的 SVG、页标签及内嵌资源 |
| `MaxGeneratedNodes` | 100,000 | 全次调用生成的元素、属性、文本节点总数；在新增节点前收费 |
| `MaxGeneratedCharacters` | 32 Mi | 全次调用新生成字符串的保守容量预算，含属性、路径、文本、转义和 Base64；独立于输出字节，实际重建/复制也收费 |
| `MaxEmbeddedResourceBytes` | 32 MiB | 本次输出将引用的原始图片/字体字节总量；相同资源每次内嵌都计费，编码前收费 |

MiB = `1024 * 1024` 字节；Mi 字符 = `1024 * 1024` 个 UTF-16 code unit。所有额度必须为正、可安全表示，不提供零表示无限的捷径；`MaxPageBytes <= MaxOutputBytes`。构造时和每次转换开始时均校验并复制选项，调用期间只使用不可变快照。`pageIndexes` 在首次异步等待前以先查 Count、再有限复制的方式取得快照；不保留调用者可变列表。调用者不得在读取快照期间并发修改它。相同转换器的并发调用不共享计数、DOM、临时文件或渲染上下文。

输入压缩/解压/条目/原文 XML/模型容量仍受 [OfdPackageLoadOptions](../src/Ofdrw.Net.Packaging/Archive/OfdPackageLoadOptions.cs) 与 [OfdModelLimits](../src/Ofdrw.Net.Core/Models/OfdModelLimits.cs) 的现有独立限制；新额度不替代它们。HTML 输出上限不能被描述为整个进程的内存上限，Reader 当前会物化整个包。

### 页选择、所有权及失败

- API 索引为零基：将 `package.Pages.OrderBy(p => p.Index)` 排序后的序号作为索引，保持现有 SVG 单页含义；它不是 XML 中的 Page ID。`null` 表示所有页的排序顺序。显式空列表拒绝，空文档也拒绝，不生成成功的空 HTML。
- 保留选择顺序和重复，例如 `[2, 0, 2]` 输出第 3、1、3 页；每次出现独立隔离 ID、计页数及字节。范围错误、负索引在开始生成正文前拒绝；`null` 的全选结果也必须先验证 `MaxPages`。不静默去重、排序或截断。
- 流必须非 null、输入可读、输出可写且非同一对象。输入从当前位置读到结束，不要求 seek、不主动归零、不关闭调用者流；输出从当前位置写，不清空已有数据，不关闭流，成功后 Flush。不同包装器别名同一底层流无法可靠识别，调用者有责任避免；CLI 则检查同路径。
- 参数错误：`ArgumentNullException` / `ArgumentException` / 页索引 `ArgumentOutOfRangeException`；损坏 OFD、累计额度超限或无效生成值：`InvalidDataException`；超出明确支持的 SVG/媒体/绘制语义：`NotSupportedException`；调用者取消原样传播 `OperationCanceledException`；I/O 错误原样传播，不转成成功或空页。禁止捕获预算/取消后换回宽松路径。
- `Title` 超长、含无效字符或预算配置不合法属于调用参数错误，读包之前用 `ArgumentException` 拒绝；有效正文遇到损坏源字符/不合法数值时则按输入数据错误处理，不能静默换成空文本。
- 生成与策略验证在私有有界临时 spool 完成，成功后才向公开输出流复制；生成/选页/额度/策略失败时公开输出保持原字节。spool 包含可能敏感的文档，使用私有随机 CreateNew 临时文件、完成后删除，不能使用调用者输入中的路径或固定文件名。
- 对任意 `Stream` 不承诺原子复制：最终复制或 Flush 期间取消、写入失败会留下完整输出的前缀；不尝试不可靠的通用回滚。调用者不得使用抛异常后的输出。调用前已经取消时，不读输入、不写输出。
- CLI 使用已有 `AtomicOutput` 在目标目录暂存；成功复制并 Flush 后调用 `CommitAsync(token)`，失败/取消仅清理临时文件。已有目标在提交前保持原字节。最终 rename 是提交点；提交点之后到达的取消不能撤销已提交文件，也不能报告目标未变。不承诺断电持久性或任意网络文件系统的原子语义。

### 建议 HTML CLI

新增显式 `ofd-to-html`，保持 `ofd-to-svg` 单页行为及现有命令参数兼容。首版不把通用 `convert` 的 `.html` 扩展名自动路由扩大为新行为。建议命令如下，**未存在、未执行**：

```text
ofdrw ofd-to-html --input document.ofd --output preview.html \
  --pages 3,1,3 --title "OFD preview" \
  --max-pages 200 --max-output-bytes 67108864 --max-page-bytes 16777216
```

CLI 页号是一基，`3,1,3` 对应 API `[2,0,2]`；未提供 `--pages` 才表示全选。`--pages ""`、全空白、空片段 `1,,2`、0、负数、降序范围、越界和整数溢出一律拒绝。正序范围 `1-3,2` 按 `[0,1,2,1]` 展开。新命令先解析正数 `--max-pages`，在展开前用 `long` 计算并累计范围长度，超额立即失败，不能先分配巨大列表再检查。输入字符长度也要有有限上限（建议 8,192），解析过程中检查取消。不修改其他命令历史解析语义。额度其余属性使用上述默认值，首版无需全部开放 CLI 开关。

沿用 CLI 退出码：成功 0、参数/数据/策略/预算/I/O 失败 1、调用者取消 130；只在 commit 完成后输出成功消息。必须先 `EnsureDifferentPaths`，拒绝目录/符号链接/硬链接等别名的能力边界需由既有输出安全测试确认，不能仅凭字符串比较宣称所有别名均安全。

## 现有能力复用

### 已核实的源码事实

下表为本次静态阅读证据；是现状，不是已运行测试结论。

| 入口 | 当前行为 | 设计影响 |
| --- | --- | --- |
| [OfdToSvgConverter](../src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs) `ConvertAsync`（37–71） | 每次新建 Reader、读完整 OFD、排序页、验证 composite/annotation、私有 `BuildSvg`、整串序列化与 UTF-8 分配、写出 | 多页不能循环调用此公开入口；输出额度必须在 DOM/编码分配之前接入 |
| 同文件 `BuildSvg` / `AddImage`（74–236） | 用 XElement 构造 SVG；局部 clip ID 为 `object-0-clip-0`、`image-0-boundary` 等，引用为 `url(#...)`；图片 `data:{MediaType};base64,...` | 多 SVG 同 DOM 会重复 ID；需要按页出现序号的分配器和媒体校验 |
| 同文件 `AddEmbeddedFonts` / `AddText`（172–337） | 字体 family 为 `ofd-font-` 加数据 hash，`@font-face` 内嵌 TTF/OTF；正文用 XElement 文本，font-family 为属性；空白有透明处理 | 复用字体覆盖、几何与空白语义；HTML CSS 必须由可信模板构建，不能把源字符串送入 style raw text |
| 同文件 `EnumerateElements`（368–383） | 背景模板、页元素、前景模板、注释投影，使用 composite 展开；明确拒绝部分字形/注释外观 | 内部共享入口必须保持顺序与拒绝行为，不通过 HTML 回退忽略异常 |
| [OfdReader](../src/Ofdrw.Net.Reader/Readers/OfdReader.cs) `ReadCoreAsync` / `ReadArchive`（48–145） | 一个 loader 读包，再物化模型；选择首个 DocBody 的 DocRoot；保留条目 | 一次读取指一次 Reader/loader 调用，不是每页重开 ZIP；首版仍只预览 Reader 选中的首文档，不声称覆盖所有 DocBody |
| 同文件 `ToMediaType`（1262–1271） | JPG/JPEG、BMP、TIFF 映射，其余默认 PNG；没有以实际图片内容证明格式 | HTML 不得只信 MediaType 或 Format 字符串，必须做内容与允许 MIME 一致性校验 |
| [CLI Program](../src/Ofdrw.Net.Cli/Program.cs)（104–118、261–278、542–592） | `ofd-to-svg` 至多一页；CLI ParsePages 一基转零基、允许范围，无 HTML 命令；空白值返回 null | 新命令需独立严格有界解析，不能把空白当全选或原样复用无限展开 |
| [AtomicOutput](../src/Ofdrw.Net.Cli/AtomicOutput.cs) | 同目录随机临时文件，Flush/关闭/检查取消后 `File.Move(..., overwrite:true)`，未提交则清理 | 可复用文件提交机制；库的 Stream 语义仍需单独说明 |
| [SVG 项目](../src/Ofdrw.Net.Converter.Svg/Ofdrw.Net.Converter.Svg.csproj) / [SVG 几何测试](../tests/Ofdrw.Net.Converter.Pdf.Tests/SvgGeometryTests.cs) | 库为 netstandard2.0/2.1；SVG 回归目前在 Pdf.Tests 的 net10.0 工程 | 内部重构兼容两个 TFM；现有测试是待重跑的兼容控制，不是 HTML 验收 |

### 一次读包与内部渲染入口建议

建议在 Svg 程序集内新增内部 `SvgRenderSession`（名称待实施审阅），由已读的 `OfdDocumentPackage`、选页快照、独立 materialization、取消 token 与累计预算组成；不是新的公开可变 package 重载。

1. 做纯参数/选项快照和已取消检查；调用一次 `OfdReader.ReadAsync(input, loadSnapshot, token)`。Reader 全包读取/验证仍发生在选页之前，选择少量页不能绕过恶意未选页的输入额度。
2. 排序一次，验证所有页索引、总输出页数、有限且正的页面尺寸与 viewBox；选择列表不在后续循环重新枚举。模型不做克隆或写入，不改 SourceXml。
3. 对同一 package 做一次 composite/annotation 验证。建立本次调用共享的字体解析缓存，按内容身份及 collection face 识别、缓存成功结果；缓存大小受原始字体字节与生成字符预算约束。既有字体解析的未覆盖 CPU 预算由计划 011 处理，不能声称 HTML 输出预算已经补齐。
4. 顺序处理每个页出现项，不并行组装所有页。内部入口从 `BuildSvg` 提取，接受 prefix/预算/字体上下文，在节点、路径/文本拆分、字素展开、Base64 及 CSS 构造前计费；页完成后验证白名单与本页引用闭包、序列化到全局计数 spool，再释放该页 DOM。
5. 全局计数器包含 shell、CSS、字体重复和所有页的转义输出。输出完成并验证结束标签后，从 spool 有界分块 Copy 到调用者输出，再 Flush；所有阶段检查同一 token，包含图遍历、编码、序列化、复制和 CLI commit 前。

现有单页 `ConvertAsync` 签名及默认第一页不变，可走提取后的相同内部渲染器；不要为 HTML 新增基于公用 SVG 字符串的拼接或正则清洗接口。XML 声明只属于独立 SVG，内联 SVG 不含 XML 声明/DOCTYPE/处理指令。DOM 与输出 serializer 使用有限块 UTF-8、无 BOM、无 DTD，无整份 `ToString` 再 `GetBytes` 的新路径。

累计字节必须用 `length > limit - used` 的先检查后累加或 checked long，超限不写入下一块。Base64 长度按 `4 * ((n + 2) / 3)` 用 checked long 预估并收费；包含 `data:` 前缀、CSS/属性/XML 转义的容量。输出 Writer 的字节计数才是 `MaxOutputBytes` 最终事实，预估不是替代。不能先 `Convert.ToBase64String`、`XElement` 或拆成巨量 tspan 后才查预算。单页 DOM 有限、字符/节点累计有限，仍不等于硬性的进程内存/CPU/墙钟上限；需要操作系统级隔离的宿主应另设边界。

### 页 ID / fragment / 字体作用域

HTML shell 使用固定 `ofd-preview-root`、`ofd-preview-page-{occurrence}`；0 起始的 occurrence 与源页号分开保存，只含 ASCII 数字。每个 SVG 内部 ID 前缀为 `ofd-preview-svg-{occurrence}-`，即使同一源页重复输出也不复用前缀。源 Page ID、标题、资源 ID 都不拼进 HTML ID。

优先由渲染器分配完整 ID；如内部已有局部树，先建立完整本页旧 ID → 新 ID 映射，拒绝重复旧 ID，再同时改写定义及唯一允许的 `clip-path="url(#id)"`。每个引用必须解析为本页 map 的唯一目标，目标为 clipPath；不允许跨页、空 fragment、外部 URI 或未定义 ID。禁止全文替换字符串，以免改变正文或 data URI。其他 fragment 型属性当前不在白名单内，新增 gradient/filter/use 时必须先扩展元素、属性、类型和引用闭包契约。

字体不是 DOM ID，仍要隔离 CSS 名称。family 建议 `ofd-preview-font-{occurrence}-{contentHash}`，每页从其自身字体映射生成，系统字体名仅放在编码后的 `font-family` presentation 属性内，不插进 CSS。相同字体可复用解析结果，但首版在每页内嵌同份受控 @font-face 并逐次计费；不引入跨页资源索引。后续去重是独立优化，不能在此设计中宣称已避免字体重复体积。

### HTML / SVG 白名单与事件、外链策略

安全边界是 **OFD 模型 → 自己生成的受限 SVG DOM → 验证 → HTML serializer**，不是承诺能清洗任意 SVG。不会执行源 SVG/script，也不读取未知外部资源来补显示。未知元素/属性/URI/不支持媒体整次拒绝，不静默删除后称为保真输出。

此白名单约束的是生成树。源模型还需要单独的绘制准入检查：被投影遍历到的可见 raw 绘图、无法建模的字形替换/注释外观必须明确拒绝，不能继承旧循环可能跳过未知 raw element 的行为后宣称完整预览。非绘图未知元数据保留在原 OFD，不注入 HTML；只展示静态外观的 Actions/GoTo 限制需在支持说明中写明。这项严格行为只属于新增 HTML API，旧单 SVG 的兼容控制独立。

| 允许元素（SVG 命名空间） | 允许属性/值类别 |
| --- | --- |
| `svg` | 固定 version、有限正 `width`/`height` 的 mm 数、四个有限 viewBox 数；命名空间声明仅固定 SVG、必要的 xml，不允许重绑定/额外外部命名空间 |
| `defs` | 无源属性 |
| `g` | `transform`（仅渲染器生成的 translate/matrix 有限数）、`opacity` 0..1、`clip-path`（仅当前页 `url(#generated-id)`） |
| `clipPath` | 生成 `id`、固定 `clipPathUnits="userSpaceOnUse"` |
| `path` | `d`（渲染器支持语法的有限数/命令）、`transform`、`fill`/`stroke`（固定 none 或生成 rgb）、`fill-rule`/`clip-rule`（nonzero/evenodd）、`stroke-width` 非负有限、`stroke-linecap`（butt/round/square）、`stroke-linejoin`（miter/round/bevel）、`stroke-miterlimit`（有限且 >=1）、`fill-opacity`/`stroke-opacity`（0..1） |
| `rect` | 非负有限 width/height，仅裁剪边界 |
| `text` | 有限 `x`/`y`、字体 family presentation 字符串、受控 normal/bold 与 normal/italic、正有限 font-size、生成 fill/opacity/transform、`xml:space="preserve"`、唯一内联样式常量 `white-space:pre`；正文只接受文本节点或受控 tspan |
| `tspan` | 有限 `x`/`y`、0..1 `fill-opacity`，正文为文本节点 |
| `image` | 有限正 width/height、固定 preserveAspectRatio、有限 transform、smooth/crisp-edges、`href` 仅经内容验证的 PNG/JPEG Base64 data URI |
| `style` | 唯一可选 `type="text/css"`；内容仅可信生成的 @font-face，family 为上面的 ASCII 生成名，src 仅通过 OpenType 校验的 TTF/OTF data URI，format/weight/style 取封闭枚举 |

仅允许上表按元素归属的属性；不能把某元素允许的属性任意用于另一元素。任意大小写/命名空间的 `on*` 事件属性、源 `style`、CSS `@import`/选择器/自定义属性、`script`、`foreignObject`、`a`、`use`、`iframe`、SMIL 动画、`filter`、`xlink:href`、`xml:base` 均拒绝。未来 `stroke-dasharray`/`stroke-dashoffset`/对象 opacity 需随 F12 的受控输出扩展允许值语法和白名单，不能先宽松接受任意字符串。源 OFD Actions/GoTo/链接不转为可点击行为；HTML 首版仅展示已支持的绘制投影，不能声称交互完整。

`image` 只首发支持 PNG/JPEG：以现有可复用的解码/格式验证能力确定实际 MIME（实施前核查可用依赖），拒绝字节不匹配、`image/svg+xml`、HTML/XML 伪装、BMP/TIFF、GIF、未知或混合恶意载荷；不根据 Reader 默认 PNG 推断真实格式。若现有依赖无法验证，先补最小实现方案，不用仅检查几个 magic bytes 宣称全部载荷安全。禁止 HTTP(S)、file、blob、协议相对地址、javascript、vbscript 及任何非上述 image/font data URI；不触发网络字体下载。字体使用现有 OpenType face/coverage 约束，不接收任意来源 CSS/data MIME；这不证明浏览器字体解码器不存在缺陷。

HTML shell 只由固定模板生成：doctype、html lang、head meta charset/viewport、title、可信固定 style、body/main/section/h2 和 SVG。所有用户 title、页标签与正文在对应文本上下文编码；不把已编码文本再按 raw HTML 插入。title 包括 `</title><script>...` 时 DOM 中仍是一个纯文本 title。SVG 文本 `</text><script>`、引号、`&`、`<` 必须由同一受控 serializer 编码；不得把 XML 转义当成 CSS raw text 转义。@font-face 不含任何源名称/自由文本，杜绝 `</style>` 终止标签来源。

shell 的屏幕布局为单列 section，标题在 SVG 外，不占源页面坐标；每个 SVG 保持原始 mm 尺寸/viewBox，使用固定 `max-width:100%;height:auto` 在窄窗口等比缩小，不移动或重排页内对象。打印样式仅由可信模板给 section 加分页断点、末节避免额外分页，隐藏仅属预览的 h2；混合纸张尺寸和浏览器打印可否与源一一匹配是待实测项，不承诺自动配置所有打印设备。无需脚本、页面懒加载或外链样式。页标签显示源一基页号与当前出现序号，空白源页也保留一节。

建议 shell 在样式前给出防御性 CSP：`default-src 'none'; script-src 'none'; connect-src 'none'; img-src data:; font-src data:; style-src 'unsafe-inline'; object-src 'none'; base-uri 'none'; form-action 'none'`。内联 CSS 是可信模板及受控字体所需；不加载 JS，不依赖 CSP 来补救策略遗漏。CSP 的 file/http 加载差异、meta 支持与 data 字体行为仍须目标浏览器实测；本文不声称所有指令/浏览器等价或 meta 能提供 frame-ancestors 等 HTTP 头能力。接入网站的宿主应按自身部署设置响应头，不能把生成文件的 meta 当站点安全配置。

## 样例与预期

所有下列测试名、fixture 和复跑筛选名均为**后续建议**。本次仅逐项审阅设计契约，实际产品运行结果全部为“未执行”，不能把表格行数当成通过用例数。建议未来在现有 Pdf.Tests 中新增 `HtmlPreviewTests`（可继续复用 SVG 工程引用）、CLI.Tests 中新增 `HtmlPreviewCliTests`，独立 fixture 放任务测试目录；不需要另建测试项目。

| 建议场景/输入 | 预期事实 | 当前实际 / 结论 | 未来复跑入口 |
| --- | --- | --- | --- |
| `IdIsolation`：3 页不同颜色/不同裁剪、每页局部 `image-0-boundary`/`object-0-clip-0`；选 `[2,0,2]` | 输出 3 section，ID 全局唯一，每个 clip 指向当前 section 的 clipPath；重复第3页两次视觉相同且无串页 | 未执行；隔离契约已设计 | `HtmlPreviewTests.IdIsolation` + 浏览器逐页 |
| `EscapingPolicy`：标题 `</title><script>alert(1)</script>&"`；正文中英文 `</text><script>`、`&`、引号；注入 `onload`/源 style | 标题/正文保持字面文本，document 无 script/on*；内部策略探针带事件或外链时拒绝；打开、点击、打印均无脚本/网络请求 | 未执行；编码、拒绝和浏览器 gate 已列 | `HtmlPreviewTests.EscapingPolicy` + 浏览器监测 |
| `FragmentClosure`：两页同局部ID、同页重复ID、悬空片段、跨页 clip | 合法局部引用重写；后三者明确失败，不回退到同名其他页定义 | 未执行；闭包契约已设计 | `HtmlPreviewTests.FragmentClosure` |
| `OneReadBudget`：NonSeekableInput + Reader/loader 内部计数探针，多页/重复页 | 读包/ReadAsync 调用恰好一次，不 seek；顺序及重复保留；读到 EOF 后不再次请求重新读取；所有 output 页均计费 | 未执行；一次读取需探针证明，单凭成功不算 | `HtmlPreviewTests.OneReadBudget` |
| `Selection`：null/空/负/越界/全选超额/文档无页；CLI `3,1,3` / `1-3,2` / `1,,2` / 超长及大范围 | 合法列表按约定；空及错误失败、无半份公开输出，CLI 不先展开无界范围；0 返回码只属完整产物 | 未执行；输入语义已设计 | `HtmlPreviewTests.Selection` / `HtmlPreviewCliTests.Selection` |
| `OutputLimits`：输出上限恰好 N 与 N-1；很小页/字符/节点/资源额度；多字节中文、转义符、重复字体页 | N 成功、N-1 失败；shell/转义/Base64/重复页累计；任意额度失败均在相应大分配前；无溢出绕过 | 未执行；预算路径需分配探针核查 | `HtmlPreviewTests.OutputLimits` |
| `Cancellation`：预取消、读包中、页面生成、Base64、最终复制中、CLI commit 前 | OCE 保留；预取消零输入读取；复制前公开输出不变，复制中允许前缀；CLI commit 前目标原字节不变且临时文件清理 | 未执行；分阶段故障注入待实现 | `HtmlPreviewTests.Cancellation` / `HtmlPreviewCliTests.Cancellation` |
| `AtomicOutputFailure`：已有目标、后页不支持/超额、写入抛异常、同输入输出、受控 rename 失败 | CLI 目标保持原字节、无成功消息；流复制失败契约独立；成功替换后文件完整 | 未执行；不由 Stream 测试替代 CLI 文件测试 | `HtmlPreviewCliTests.AtomicOutputFailure` |
| `ExternalResourcePolicy`：HTTP/font URL、SVG data image、伪 PNG、BMP/TIFF、合法 PNG/JPEG/TTF/OTF | 仅允许受控合法媒体；非法项拒绝，网络捕获零请求；拒绝时不靠浏览器尝试显示 | 未执行；格式验证能力待核实 | `HtmlPreviewTests.ExternalResourcePolicy` + 浏览器监测 |
| `UnsupportedDrawing`：可见 raw 绘图、不支持的 glyph/注释外观、仅有未知非绘图元数据 | 前三者拒绝而非漏画后成功；非绘图元数据不注入 HTML，输入原 OFD 字节/模型不改 | 未执行；生成白名单与源绘图准入分别验证 | `HtmlPreviewTests.UnsupportedDrawing` |
| `FidelityDependencies`：源 TextCode 继承轴/反斜杠四位编码、对象 Alpha/dash；混合中文/英文比例字体、局部样式、表格与分页 | F11/F12 修复并集成后按相同有效语义比较单 SVG/HTML；修复前明确这些文件不在保真承诺内，不标视觉通过 | 未执行；依赖记录，不把 HTML 封装当修复 | 计划014/015对应回归 + `HtmlPreviewTests.FidelityDependencies` |
| `BrowserLayout`：至少3页、异形页尺寸、嵌入字体/PNG、空白页、重复页 | section 顺序、页号、尺寸比例正确，无额外空页/裁切/重叠；缩放不改变内容；打印每个 section 单独分页作为单独实测项 | 未执行；浏览器布局非静态 DOM 可证明 | 浏览器打开本次 HTML，逐页截图/打印检查 |

建议小样例以 [SvgGeometryTests](../tests/Ofdrw.Net.Converter.Pdf.Tests/SvgGeometryTests.cs) 的 Writer→Reader→SVG 公开链构造，字体取现有测试字体；记录实际字体文件身份。继承/编码场景必须使用带 SourceXml 的真实读取样例，不能用新建模型绕过 F11。综合样例复用 [generated-layout.docx](../e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx)，先生成本次显式 native 和默认 OFD，再由这些 OFD 分别生成 HTML 与 PDF。固定标题和选页顺序，保存每个输入、OFD、HTML、PDF、页 PNG 的身份和源码基线；体积变化单独解释字体重复成本。

以下为建议 HTML 结构示意，**不是已生成样例或可直接运行的完整 HTML**；省略 head/CSP/可信样式和真实 SVG 内容：

```html
<main id="ofd-preview-root">
  <section id="ofd-preview-page-0" data-source-page="3">
    <h2>第 3 页（预览 1）</h2>
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
      <defs><clipPath id="ofd-preview-svg-0-image-0-boundary" clipPathUnits="userSpaceOnUse"><rect width="20" height="20" /></clipPath></defs>
      <g clip-path="url(#ofd-preview-svg-0-image-0-boundary)"></g>
    </svg>
  </section>
</main>
```

## 验收命令

### 本次实际执行的设计检查

在隔离 worktree 执行的漂移/范围检查：

```bash
git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- plans/029-html-preview-design.md
git status --short
git diff -- plans/029-html-preview-design.md
git diff --cached -- plans/029-html-preview-design.md
```

写入前结果：指定交付文件无已提交差异、无未提交/暂存差异，worktree 干净。另核查 SVG、Reader、CLI、AtomicOutput 在 d0f88c7→ec121b0 没有提交差异。主 checkout 的原有未跟踪计划和技能目录未修改。文档完成后运行 `python3 artifacts/plan-029/check-design.py` 检查必需标题、真实本地链接、引用源码符号及契约清单，再执行 `git diff --check` 和 `git diff --cached --check`。实际输出及范围见 [checks.txt](../artifacts/plan-029/checks.txt)；检查器只证明文档结构/来源路径和人工核对项目存在，不证明产品实现或设计安全性已由运行验证。

人工内容复核覆盖：流所有权和复制失败、一次读包、选页/重复/空集、所有计费位置、单页/累计输出、页ID/字体作用域、白名单按元素归属、CSS raw text、外链和媒体拒绝、取消、文件提交、F11/F12、样例与未来门、兼容和切片。核对结果记录于 checks.txt，不能把此人工审阅称为功能测试。

### 后续实施时才运行，以下全部未执行

下列命令以本仓库现有工程/runner 为基础；`HtmlPreviewTests`、`HtmlPreviewCliTests` 和 `ofd-to-html` 是尚待实施的建议，实施者必须先创建并确认非零测试，不能现在复制运行后把零命中称为通过。所有 dotnet 命令先使用可写 CLI_HOME 与单节点参数；库同时覆盖 netstandard2.0/2.1，测试宿主 .NET 10 不代表库可使用新 API。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-029/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --no-build --filter 'FullyQualifiedName~HtmlPreviewTests|FullyQualifiedName~SvgGeometryTests' --logger trx --results-directory artifacts/plan-029/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --no-build --filter 'FullyQualifiedName~HtmlPreviewCliTests|FullyQualifiedName~OutputSafetyTests' --logger trx --results-directory artifacts/plan-029/cli-tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test Ofdrw.Net.sln -c Release --no-build --logger trx --results-directory artifacts/plan-029/full-tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan029 --output-dir artifacts/plan-029/package-output
python3 scripts/check-public-api-docs.py
```

必须分别确认 HTML 测试类非零发现，不能以 SvgGeometryTests/OutputSafetyTests 的命中数掩盖 HTML 类零命中。包 E2E 的现有 runner 不自动覆盖新 HTML API；实施时追加纯 NuGet 消费者调用及已安装 CLI 的 HTML 实例，使用本次本地包，不宣称本地发布版本号已存在于公共源。测试失败先排 restore/assets，出现 SocketException(13) 按 AGENTS.md 处理，不清全局缓存；预算极值仅用小额度/受控进程，不在主 test host 跑无界范围。

建议未来 CLI 样例（必须使用上一步真实安装的本地包工具；下面 `ofdrw` 是该工具的逻辑名称，不是已经证实的 PATH 命令）：

```text
ofdrw docx-to-ofd --input e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx --output artifacts/plan-029/native.ofd --docx-ofd-mode native
ofdrw docx-to-ofd --input e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx --output artifacts/plan-029/default.ofd
ofdrw ofd-to-html --input artifacts/plan-029/native.ofd --output artifacts/plan-029/native.html --pages 3,1,3 --title "OFD preview"
ofdrw ofd-to-html --input artifacts/plan-029/default.ofd --output artifacts/plan-029/default.html
ofdrw ofd-to-pdf --input artifacts/plan-029/native.ofd --output artifacts/plan-029/native.pdf
ofdrw ofd-to-pdf --input artifacts/plan-029/default.ofd --output artifacts/plan-029/default.pdf
```

`3,1,3` 仅适用于先确认至少3页的确定性样例；若综合 DOCX 不满足，使用上表专门三页 fixture，不默默变更验证目标。浏览器打开本次绝对路径 HTML，保存实际浏览器版本、加载方式（file / 本地 HTTP）、网络零请求证据与逐页截图；至少确定一个目标浏览器版本再承诺支持。检查页间 clip、字体、标题注入和混合样式；如果要支持其他浏览器，逐一完成同矩阵，不推断等价。

对本次 OFD 导出的实际 PDF，未来逐页执行 PNG 渲染及目视复查，并在 macOS 用 Preview 重新加载最终文件逐页查看、结束后关闭窗口。建议命令形状如下（路径是实施时记录的实际绝对路径占位，**未执行**）：

```text
pdftoppm -png -r 144 <本次native.pdf绝对路径> <本次产物目录/native-page>
open -a Preview <本次native.pdf绝对路径>
```

必须报告四项：功能测试、自动 PNG 渲染、PNG 视觉复查、macOS Preview 验收。HTML 本身另列浏览器逐页验收；PDF/Preview 仅验证 `DOCX → native OFD → PDF → Preview` 链，不替代 SVG/HTML 浏览器验收，直接 DOCX→PDF 也不替代 native OFD。检查中文缺字、英文间距、局部样式扩散、字号基线、表格填充边框、图片、页眉页脚页码、裁切重叠、重影及异常空白页。重新生成后全部重新加载，不能验收缓存旧页面。无 GUI 时按 AGENTS.md 保留产物/环境/哈希与逐页 PNG 复查范围，Preview 和需要的浏览器门保持未完成。

本次没有创建可渲染产物，不运行上述 dotnet、包 E2E、转换、浏览器、PNG 或 Preview 命令；静态设计不承担这些尚不存在产品的验收。

## 风险与未决项

1. **F11 / F12 尚未由本任务修复。** 基线 Reader raw TextCode 省略轴采用独立默认值、源反斜杠编码未在 renderer 统一解码；SVG 只额外处理 cap/join/miter，未覆盖有效 dash/offset/对象 Alpha。计划 014/015 为关联修复入口。设计可独立交付；HTML 产品广泛保真验收须等它们集成，或在独立后续授权中明示受限样例范围，不能默认宣传正确显示这些输入。
2. 计划011字体解析预算、现有 path 正规化支持及不支持的 glyph/annotation 外观不是本计划解决范围。首版只继承明确支持的投影；对于无法证明的绘制语义拒绝或在产品支持范围中明示限制，不能用整页光栅化隐藏差异。
3. 内嵌 SVG 的 XML→HTML 解析差异、大小写和 namespace 行为、可信 @font-face 在浏览器中的兼容性、CSP/file 语义待实现时验证。本文没有引用外部标准，也不编造合规认证；精确 HTML/SVG/CSP 规范条款与浏览器差异如实施需要应另行核实。
4. PNG/JPEG 完整载荷验证所能复用的现有依赖尚需实施前核实；仅记录了 Reader 的 MIME 推导行为。若没有合适验证入口，先审阅依赖/最小验证方案，不在本任务中新增图片库。BMP/TIFF 等现有单 SVG 可尝试输出的媒体在 HTML 首版明确拒绝，是新增 API 的保守边界，不改旧单页 API。
5. 字体每页重复内嵌可显著增大产物；默认额度可能拒绝大量页面/大字体，应以真实语料记录并校准，不通过默认无限制修补体验。是否后来移到全局受控字体表留给独立优化审阅；首版不实现懒加载/虚拟滚动/脚本按钮。
6. 此预览不承诺全选中文本、原生可访问阅读顺序、搜索提取完整性、任意复杂 Word 保真或所有浏览器等价。正文仍是 SVG 绘制投影，源 OFD 不变。跨文档/跳转/外链交互、签章可信展示、任意源 SVG 导入、脚本扩展和标准认证均不在首版。
7. 单次 Reader 仍物化全包；有限输出与生成预算只覆盖所定义资源。OS 内存/CPU/墙钟、临时目录安全与可用空间、最终复制失败和网络文件系统提交边界应由宿主明确控制。基线 `EnsureDifferentPaths` 对所有文件别名的保障尚待实际安全测试，不能扩大结论。

## 后续实施切片

以下各片均需另行授权实施，本次仅交接设计；建议源码/测试位置不存在不代表缺失交付。避免同时改 SVG renderer 与计划011、014、015 的共享函数，集成后重读真实符号再实施。

| 片 | 建议改动面 | 可审阅交付与进入下一片条件 |
| --- | --- | --- |
| A：契约与内部复用 | `src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs` + 建议 `Internal/SvgRenderSession.cs`；Pdf.Tests 的现有 SVG 控制 | 提取 package/page 内部入口、调用内缓存与 token；保留公开单页签名和异常；现有单页回归重跑且逐页 SVG 浏览器复查，未引入 HTML public API |
| B：受限生成与预算 | 建议 `Internal/HtmlPreviewBudget.cs`、`SvgPreviewPolicy.cs`、`SvgIdScope.cs`；建议 HtmlPreviewTests | 在所有大分配前计费、局部ID闭包/字体命名、白名单/CSS/media；小额度、恶意载荷、外链、重复页探针非零运行；先证明拒绝再写 shell |
| C：HTML 公共流 API | 建议 `Converters/OfdToHtmlConverter.cs`、`OfdHtmlOptions.cs` + bounded spool | 一次读包计数证明、参数快照、顺序/重复/空页、UTF-8 精确额度、分阶段取消与 Stream 前缀语义；两个 netstandard TFM 编译及公开 XML 文档检查 |
| D：CLI 与文件提交 | 建议 `src/Ofdrw.Net.Cli/Cli.Html.cs`，Program 分发/帮助 + 复用 AtomicOutput；建议 HtmlPreviewCliTests | 严格有界一基解析、新命令 opt-in、取消130、旧文件保持及 rename 故障注入；旧 SVG 命令行为回归，不顺带改其他命令解析 |
| E：消费者与真实视觉 | 现有测试/包 runner 按独立授权扩充，保存 `artifacts/plan-029/` | 本次本地包 API/已安装CLI消费，全套回归，明确F11/F12集成状态，浏览器逐页/网络证据、PNG与Preview分项记录；存在视觉缺陷时产品门保持未完成 |

各片复核实际差异、未来 API 和异常契约，不将 A 的单页兼容通过当成 C/D/E 通过；预算或白名单失败不能削弱断言或自动回退到不受限单页串接。设计完成不授权推送、PR、合并、发布或修改上述生产文件。
