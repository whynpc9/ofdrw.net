# D10：ImageSharp 告警与兼容升级评估

结论：**截至 2026-11-03，没有保持现有 netstandard2.0/netstandard2.1 的官方安全升级；本轮以评估和后续决策方案结束，依赖漏洞未修复，NU1902/NU1903 继续披露。** 不改包引用、目标框架、许可声明、图像实现或审计配置；不发布新版本，不改写已发布 preview.10 的内容。

源码基线为最新 `origin/main` 的 `6ce353b71704942e750bd07d8eef7cdccc66f062`（含 D09/PR105），独立分支 `codex/d10-imagesharp-security-assessment`。本轮使用 `codex-security:triage-finding` 做限定静态核对，不是全仓安全审计或漏洞利用验证。未找到适用的 `SECURITY.md`；库/CLI 处理调用方文档，是否接受不可信输入取决于宿主部署，不能假定所有输入可信或所有部署均可远程利用。

## 1. 直接与传递依赖

三个直接引用均为 `SixLabors.ImageSharp 2.1.13`，三个项目均保留 `netstandard2.0;netstandard2.1`：

| 项目 | 图与实际用途 |
| --- | --- |
| `src/Ofdrw.Net.Reader/Ofdrw.Net.Reader.csproj` | 受限 JBIG2 转 PNG、ImageMask 合成，`Readers/OfdImageRenderProjection.cs` |
| `src/Ofdrw.Net.Converter.Pdf/Ofdrw.Net.Converter.Pdf.csproj` | 图片导入/导出、OFD→PDF 的像素识别、裁切、透明度与 PDF bitmap source；另直接引用 PdfSharpCore 1.3.67 |
| `src/Ofdrw.Net.Converter.Docx/Ofdrw.Net.Converter.Docx.csproj` | 内嵌图像识别和 built-in renderer；另引用 MigraDocCore.Rendering 1.3.67，并引用 Converter.Pdf/Layout |

官方发布的 [PdfSharpCore 1.3.67](https://www.nuget.org/packages/PdfSharpCore/1.3.67) nuspec 声明 `SixLabors.ImageSharp >=1.0.4`，不是锁定 1.0.4；本仓库直接引用选择 2.1.13。[MigraDocCore.Rendering 1.3.67](https://www.nuget.org/packages/MigraDocCore.Rendering/1.3.67) 经 PdfSharpCore、PdfSharpCore.Charting 和 MigraDocCore.DocumentObjectModel 传递到该图。Layout→Reader，Converter.Svg→Reader，Converter 元包、可选 Vector 和 CLI 等消费者也会继承相关依赖；只改一处引用或删直接引用不能移除该图，降到 1.0.4 也不是安全修复。

依赖来源以当前 csproj 和官方 nupkg/nuspec 为准；现行发布盘点见 [40 项依赖基线](../third-party-dependency-baseline.json)，许可见 [THIRD-PARTY-NOTICES](../../THIRD-PARTY-NOTICES.md)。外部版本索引与公告可随时间变化，后续实施前须重查。

## 2. 官方版本、修复与许可对照

本轮浏览官方 NuGet 与维护者公告，并读取 NuGet flat-container 版本索引及实际 nupkg 的 `lib/`、nuspec 和 LICENSE；不是按当前 SDK 能运行就推定库的 TFM 兼容。

| 版本 | 实际包 TFM | 五条当前告警 | 许可/决策 |
| --- | --- | --- | --- |
| [2.1.13](https://www.nuget.org/packages/SixLabors.ImageSharp/2.1.13)（最新稳定 2.x） | net472、netcoreapp2.1、netcoreapp3.1、netstandard2.0、netstandard2.1 | 均在受影响范围；无已发布安全 2.x | nuspec 为 Apache-2.0；[v2.1.13 LICENSE](https://github.com/SixLabors/ImageSharp/blob/v2.1.13/LICENSE) 同为 Apache-2.0 |
| [3.2.0](https://www.nuget.org/packages/SixLabors.ImageSharp/3.2.0)（最新稳定 3.x） | **仅 net6.0** | 五条均标注此版本修复 | 包内 Six Labors Split License 1.0；不能用于当前两个 netstandard 目标的直接升级 |
| [4.1.2](https://www.nuget.org/packages/SixLabors.ImageSharp/4.1.2) | **仅 net8.0** | 五条均标注此版本修复 | Split License 1.0；同样不兼容当前目标 |
| [4.1.3](https://www.nuget.org/packages/SixLabors.ImageSharp/4.1.3)（当前最新稳定） | **仅 net8.0** | 高于这五条的修复下限 | Split License 1.0；不是 netstandard 兼容候选 |

这里的“修复”限于下表五条，不承诺这些版本没有其他或未来漏洞。[3.2.0 发布说明](https://github.com/SixLabors/ImageSharp/releases/tag/v3.2.0) 明列安全修复回移及 license enforcement；它不是 3.1.12 的无条件替换。[3.2.0](https://github.com/SixLabors/ImageSharp/blob/v3.2.0/LICENSE) 和 [4.1.3](https://github.com/SixLabors/ImageSharp/blob/v4.1.3/LICENSE) 的许可按使用情形区分授权，包括开源/源码可用、传递消费及其他条件；不能将旧版 Apache-2.0 行机械沿用，也不能断言所有消费者必然需要付费。后续选型须核对实际消费方式、包内许可及构建期授权要求，不在本轮替用户选择许可或绕过检查。NuGet 2.1.13 页的 README 文字与其版本专属许可元数据不同；固定版本判断采用实际 nuspec 和源 LICENSE。

## 3. 五条具体公告与可达性

官方维护者公告的当前结构化 affected/patched 字段如下。部分聚合公告正文保留旧的连续版本范围，故不把 `<4.1.2` 简化成所有 3.x 都安全，也不按 NU1902/NU1903 字符串合并风险。

| 公告 / 严重度 | 官方受影响范围；修复版本 | 本仓库静态核对与判定 |
| --- | --- | --- |
| [GHSA-jjfr-hcj7-qf5w](https://github.com/SixLabors/ImageSharp/security/advisories/GHSA-jjfr-hcj7-qf5w) / High；CVE-2026-106115 | `>=2.1.0,<3.2.0` 和 `>=4.0.0,<4.1.2`；3.2.0 / 4.1.2 | T6/Group4 TIFF **编码**越界。产品未调用 TiffEncoder/SaveAsTiff，PDF source 选择 JPEG/BMP，其他输出显式 PNG/JPEG。对已检查产品路径为 `not_actionable`，不是包版本无漏洞 |
| [GHSA-wmxv-xphr-5c9g](https://github.com/SixLabors/ImageSharp/security/advisories/GHSA-wmxv-xphr-5c9g) / Moderate；CVE-2026-106116 | `>=2.0.0,<3.2.0` 和 `>=4.0.0,<4.1.2`；3.2.0 / 4.1.2 | BigTIFF IFD 非前进循环，可在 Identify 阶段发生。已有顶层格式拒绝，见下文；依赖内部/嵌套元数据路径未穷尽，`needs_review`，队列 1 |
| [GHSA-j9gm-c75j-xc9q](https://github.com/SixLabors/ImageSharp/security/advisories/GHSA-j9gm-c75j-xc9q) / High；CVE-2026-106110 | `>=2.0.0,<3.2.0` 和 `>=4.0.0,<4.1.2`；3.2.0 / 4.1.2 | T4/Group3 TIFF **编码**越界，需明确选择该编码器或继承 TIFF 元数据后重新编码 TIFF。产品无该输出路径，`not_actionable`；解码 classic TIFF 本身不是此公告的编码 sink |
| [GHSA-j3p4-wp97-rph4](https://github.com/SixLabors/ImageSharp/security/advisories/GHSA-j3p4-wp97-rph4) / High；CVE-2026-106113 | `>=2.0.0,<3.2.0` 和 `>=4.0.0,<4.1.2`；3.2.0 / 4.1.2 | 浮点 TIFF→HalfVector4→HistogramEqualization 的不安全索引；产品未使用 HistogramEqualization/HalfVector4，`not_actionable` |
| [GHSA-gwg2-r3hj-4w44](https://github.com/SixLabors/ImageSharp/security/advisories/GHSA-gwg2-r3hj-4w44) / Moderate；CVE-2026-106114 | `>=1.0.0-beta0001,<3.2.0` 和 `>=4.0.0,<4.1.2`；3.2.0 / 4.1.2 | 2.1.13 的触发方式是 lazy `IccProfile.Entries`；不同于 v4 的显式 ICC conversion。产品无直接 Entries/ICC conversion 调用，但未穷尽依赖内部所有 metadata 处理，`needs_review`，队列 2；不把一般 Image.Load 直接称为 2.x 的已证实 ICC exploit |

所有五条均确认 **包版本受影响**；没有一条在本轮被确认为应用级可利用漏洞，没有运行恶意样例或建立远程攻击结论。`not_actionable` 只针对公告指定的产品触发操作，不覆盖消费者自行调用 ImageSharp 的场景。测试生成 TIFF fixture 与产品 TIFF 编码不是同一个入口；没有为了表面消警删除 fixture。

具体控制和残余证明缺口：

- `Core/IO/ImageFormatAdmission.cs:RejectBigTiff` 识别两种字节序的顶层 BigTIFF magic，在 `DocxModelReader.ReadImagePayload`、`OfdToPdfConverter.ValidateImage` 和 `EncodedBitmapImageSource` 的 Identify/Load 之前执行。图片导入 `ImageToOfdConverter` 在 Identify 前要求 PNG/JPEG；遮罩投影先核对实际编码格式与资源声明。PNG/JPEG 的字节/像素预算不能代替依赖解析器修复，也不能据此证明所有嵌套 metadata 不可达。
- `Reader/Readers/OfdImageRenderProjection.cs` 的普通图片返回原始 bytes，受限遮罩才解码；`Converter.Pdf/Converters/PdfCropImage.cs`、透明度路径和 `ImageIO/EncodedBitmapImageSource.cs` 用 RGBA32 与显式 PNG/JPEG/BMP 编码器。Vector 的 RGB 输出和 Docnet 栅格输出从像素缓冲构造图像，不能与任意编码字节解码混同。
- built-in DOCX 经 `DocxModelReader`→`BuiltInPdfRenderer`→MigraDoc/PdfSharpCore 的 `ImageSharpImageSource<Rgba32>`；签章 legacy bitmap 在 `ValidateImage` 后进入 XImage。已检查 [PdfSharpCore 固定源码](https://github.com/ststeiger/PdfSharpCore/blob/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Utils/ImageSharpImageSource.cs)：只解码后写 JPEG/BMP，没有 TIFF 编码或直方图处理。下一步最小风险核对是沿 2.1.13 的 EXIF/ICC metadata 解码及上述重编码内部，明确是否会进入 ReadValues64 / IccProfile.Entries；未完成前保留 `needs_review`。

## 4. 可执行方案与代价

本轮选择保持现状并披露阻断；这是有限评估的关闭，**不是安全修复完成或安全发布批准**。不抑制审计，不伪造安全版本，也不以现有测试/准入控制代替修复。

| 后续候选 | 实施前必须决定的事项与成本 |
| --- | --- |
| 现代目标的独立升级方案 | 明确授权新增/调整 TFM 及 Split License 审查后，在独立分支原型验证 3.2.0 或当时稳定 4.x。处理 ImageSharp API/程序集变化、PdfSharpCore/MigraDoc 对旧 API 的二进制调用；不能仅提升 CLI 的 TFM 就让 netstandard 库引用 net6/net8。若保留旧目标，其告警仍单独披露 |
| 保留 netstandard 的安全回移 | 先向维护者确认可用的兼容修复及明确授权；若自维护 Apache-2.0 衍生分支，需逐项独立修补/审查、独立包身份、长期安全更新与公告映射。不能未经许可复制 Split License 新源码，也不能把自建包改名为官方安全版本来规避审计 |
| 替换图像后端及传递依赖路径 | 先冻结 PNG/JPEG/经典 TIFF/掩码/透明度/ICC/方向与 DPI 行为、跨平台部署和许可证。复用可选 Skia 也会把 native 部署带入当前纯托管路径；还须处理 PdfSharpCore/MigraDoc 的默认 ImageSharp source，单删三个直接引用不够 |

任一实际实施都须：固定最小反例与有意义回归；检查完整直接/传递图和双 TFM 包消费；测试 PNG/JPEG 读取、经典 TIFF 等既有输入、透明度/裁切/遮罩与畸形拒绝；运行受影响完整回归与隔离包消费；本次重新生成 native/default DOCX→OFD→PDF、逐页 PNG、macOS Preview；分别核对原文、重复文本、图像视觉与许可清单。不能以一次编译、无 NU190x 或旧页面作为验收。实施授权不在本轮自动扩大。

## 5. 本轮验证范围

本轮仅修改评估文档和路线图链接，不运行本地全套构建、图片回归、包消费或页面生成。功能测试/自动渲染/PNG 复查/Preview 均为 **本轮未执行（无产品改动）**；不引用旧产物作为新修复的验收。Terra High 独立执行一次限定 Converter.Docx 的 restore：Reader/Pdf/Docx 各两个 TFM 均解析为 2.1.13，各有相同五条 NU1902/NU1903；并独立核对官方包 TFM、许可和公告范围。`verify-package-artifacts.py --notices-only` 因本 worktree 未还原 Converter.Svg 的 assets 而未完成，没有为文档扩大还原范围。PR 仍须完成现行 exact-head CI 和独立 Astra High 审阅，再实际合并。

官方公告 JSON、版本索引、实际包提取的少量 nuspec/LICENSE 和固定上游源码放在 ignored `artifacts/d10/`，不提交包、源码树副本或完整日志。复查入口为上文官方链接、当前 csproj 与固定源码 SHA；独立核验记录及存储检查结果在 PR 中记录。
