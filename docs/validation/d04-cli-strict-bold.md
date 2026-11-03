# D04：CLI 严格嵌入粗体策略

日期：2026-11-02。基线为最新 `origin/main`：`e4c67ef6786ce63e17d044ab1243dce881ac1f13`，含 D03/PR99。

## 行为与范围

CLI 新增 `--native-embedded-bold-behavior text-stroke|throw`，对应公开 API 的 `DocxConversionOptions.NativeEmbeddedBoldBehavior`。省略时保留 `TextStroke` 默认值；`throw` 在 Native 实际使用的嵌入粗体缺少真实 Bold 时拒绝。`docx-to-ofd` 与推断 DOCX→OFD 的 `convert` 共用路由；不改变直接 DOCX→PDF、dual-layer 或 name-only 字体行为，不新增字体引擎或关闭所有系统字体解析。

帮助和[字体教程](../content/guides/fonts.md)说明合法值、默认值、严格正负例与限定范围。新开关属于后续源码，不包含在已发布的 preview.10 包中；本次不改原标签或重复发布。

## 复现、回归与审阅

新回归复用既有确定性 `narrow.ttf`，动态构建仅供参数路由的 DOCX，不增加字体或文档载荷。覆盖默认/显式 Native、`convert` 推断、严格拒绝的真实错误、已有目标与临时文件、默认/显式合成粗体诊断及 OFD paint、大小写/空白、非法值、缺值和帮助。

修复前新增 15 例中 13 例失败、2 例默认兼容通过；修复后 CLI 全集 229/229 通过，新增 15 例均通过、零跳过。严格失败验证的是转换器报告 `actual bold face`，不是仅解析成功或返回任意错误。该原始度量字体用于语义回归，不用于页面可读性验收。

开发为 gpt-6.1-sol High；gpt-6-astra High 独立审阅无可操作发现。独立本地功能与页面验证由 gpt-6.1-terra High 执行，GUI 独占。所有 .NET 构建/测试/打包经共享 `run-exclusive.py`，使用任务目录、单节点并禁用 build servers/shared compilation；没有清除全局包缓存。

产品 `Program.cs` SHA-256：`f17c47938fd63f7aa940eb6dd22790cb2edc3ac749274b90dd1693c4a66ca6fb`；新增回归 SHA-256：`ab969b9de8b1aaba4d8ba902dca7343800cf9336329106233d20b15da0449964`。

本次日志、包、生成页面与详细审阅材料位于本工作树忽略的 `artifacts/d04/`，不提交构建目录、运行日志、包缓存或字体复制。

## 本地功能与页面验收

| 项目 | 本次结果 |
| --- | --- |
| 全 solution Release build/test | 11 个测试程序集，4640 通过、0 失败；默认字体环境 17 项跳过。独立原字体 mandatory 门禁 74/74 通过，零跳过。 |
| 本地包消费 | 未发布 `0.1.0-d04.1`，12 包验证及标准 32 步完整消费通过；单独安装 CLI tool，公开 API 应用精确引用同一 feed。CLI/Docx/Pdf 消费 DLL 与 nupkg 内字节一致。 |
| CLI 严格配置 | 固定真实 Regular/Bold 目录，显式 native/default 均成功；独立 Regular-only 目录下，两入口及 `convert` 均退出 1 并报告 `actual bold face`。已有目标保留；原本无目标的负例不发布文件。 |
| 默认兼容与页选择 | 省略策略及显式 `text-stroke` 均成功并保留 `DOCX_EMBEDDED_BOLD_SIMULATED`；`--pages 2` 只输出一页，文字等于完整产物第 2 页。 |
| API 严格配置 | 教程公开 API 配置的 native/default 正例通过；缺 Bold 均抛 `InvalidDataException`，写入零字节。示例自行创建的空文件由调用方负责，与 CLI 原子发布保证分开。 |
| 原文与物理字体 | 四份 CLI/API OFD 均保留源文档 189 个非空白字符且同序；两份 CLI PDF 提取同样逐字相等、无重复。四份 OFD 的两页 Content.xml 和 PublicRes.xml 哈希相同；七个粗体对象包含“数量”，引用 weight 700、Bold 标志有效的真实面子集，无 Stroke 文本对象。 |
| 自动渲染检查 | 本次 CLI 两模式各两页 PDF，经 Poppler 渲染为 110 DPI、910×1287 的四张 PNG；未用直接 DOCX→PDF 代替。 |
| PNG 视觉复查 | Terra 逐页检查四张 PNG；主线程另行逐页复核同四张。中英文、比例英文、局部粗斜体/颜色、表格底色边框、对齐和分页未见新问题；常规字号“数量”可辨，无缺字、裁切、重影或异常空白页。 |
| macOS Preview | Terra 实际打开本次 `strict/native.pdf` 与 `strict/default.pdf`，核对文件身份并逐页查看共四页正文，验收后关闭本次窗口。链路为 DOCX→native/default OFD→PDF→Preview。 |

输入复用未修改 F01 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`，SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`。字体为[既有固定 profile](../content/reference/release-profile-cn-en.md)的真实 CJK 400/700 加 Latin Regular，先验证源文件哈希再复用，不安装系统字体、不新增字体载荷。实际三个 OFD 字体资源均为 Noto Sans CJK SC：Bold、Regular、Regular/italic 请求；后两者使用相同 Regular 子集，本轮不声称物理 Italic。

本机 macOS 26.6.2 arm64、.NET 10.0.401。结构/字体/文本/包身份原始记录为 `artifacts/d04/terra/inspection.json`；CLI 命令与退出码为 `terra/strict/results.json`；API 为 `terra/api-output/results.json`；包清单及 32 步记录为 `terra/package-output/package-e2e-run.json`。Preview 截图未单独保存；保留的视觉记录是四张 PNG，不把它们称为 Preview 截图。详细独立验收记录及文件大小/哈希清单保存在 `artifacts/d04/terra-validation.md`、`.json` 与 `terra/hashes.json`。

### 关键输出身份

| 本次 CLI 输出 | 字节数 | SHA-256 |
| --- | ---: | --- |
| `native.ofd` | 551453 | `4eae2459f9fe253c40aca0ecd17d19bac28a84a1a0df9f908e3508ba338817d4` |
| `native.pdf` | 144787 | `96bf24a6bd81ceaf6d9ab6ed8ce2efb4c6a841e148181bd010f553004063d0da` |
| `default.ofd` | 551453 | `bb9c562bc8783c0fb77fd89a77f6d38e7293e8036a18d9091aa252f94fe7c0eb` |
| `default.pdf` | 144787 | `e6023a723784f8f23700ab3ec2a39140915acd6c3bdb6d527684861dfbb1cc23` |

两模式输出体积基本相同，ZIP/元数据导致整体文件哈希不同不表示页面或选面不同。子集 SHA 不等同于完整源字体 SHA。

## 验收边界与遗留限制

本次功能与页面通过仅覆盖上述具名严格真实字体配置、未修改 F01 和两模式各两页。样例没有图片、页眉页脚或页码，不由本轮推断这些功能或任意复杂 Word 文档均保真。

旧 Regular-only 合成粗体“量”字视觉 FAIL 保留；本轮未重测或关闭该项。Preview Find 与原始 Unicode Copy 也未重测，不宣称修复或通过。严格策略不是全局缺字体检测，name-only 字体仍依赖宿主解析；直接 PDF/dual-layer 不受该策略约束。
