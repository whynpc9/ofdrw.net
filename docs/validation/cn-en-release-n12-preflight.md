# 中英文候选 N12 发布工程预检（2026-11-01）

> 本文为 2026-11-01 预检及当时阻断状态。2026-11-02 的候选版本核对、限定配置及已知限制决定见[当前计划](../next-steps-2026-11-02.md)。旧记录不改为通过；最终候选仍须新构建、原包消费、实际页面和准确源码绑定。

本轮只做发布工程预检、文档纠错、LFS 指纹阻断修复及轻量检查，**未形成可发布候选**。没有运行 .NET 构建/完整测试/转换，没有操作 GUI，没有更新候选验收 JSON，没有 tag、GitHub Release 或 NuGet 推送。最终出口仍等待集成修复后的准确 main、同批候选包、Terra 实际页面验收及 Astra 独立复核。协调者本轮转交的 N02 状态为：真实 Preview 的核心日期/页末“待复核”Find 失败仍 OPEN，API 查询通过不能代替 GUI；本预检未独立重验该结果，候选保持阻断。

## 源码与在线核对

- 预检工作树：`codex/cn-en-release-n12`，起点 `ba7682cfceafd6e28d526941e0eb039addfe400f`。该 SHA 是预检文档基线，不是最终发布 SHA。
- 本日 `git ls-remote origin refs/heads/main` 得到 `beb02832fff01e61538d90039f3e1224ed03d702`；读取了该提交的实际 CI 与 `docs/ci-performance.md`。本轮集成候选另行冻结，不能用此前规划基线或本文 SHA 代替。
- 已通过 `gh api repos/whynpc9/ofdrw.net/tags --paginate`、`gh api repos/whynpc9/ofdrw.net/releases --paginate` 核对：最新现有 tag 为 `v0.1.0-preview.9`；对应 [GitHub Release](https://github.com/whynpc9/ofdrw.net/releases/tag/v0.1.0-preview.9) 于 `2026-10-18T01:48:07Z` 公开，`prerelease=true`、`draft=false`。
- 官方 NuGet V3 `https://api.nuget.org/v3-flatcontainer/<lowercase-package-id>/index.json` 本日逐 ID 查询：默认清单中既有的 11 个 ID 均返回成功，最新版本均为 `0.1.0-preview.9`，均无 `0.1.0-preview.10`。例如 [Core 官方索引](https://api.nuget.org/v3-flatcontainer/ofdrw.net.core/index.json)、[Converter 官方索引](https://api.nuget.org/v3-flatcontainer/ofdrw.net.converter/index.json)、[CLI 官方索引](https://api.nuget.org/v3-flatcontainer/ofdrw.net.cli/index.json)。
- `Ofdrw.Net.Text.HarfBuzz` 官方索引返回 HTTP 404；四个独立可选扩展 `Ofdrw.Net.Graphics.SkiaSharp`、`Ofdrw.Net.Converter.Pdf.Vector`、`Ofdrw.Net.Signatures.SesInterop`、`Ofdrw.Net.Crypto.Password` 也均为 404。这表示本次未查到公开版本，不能证明 ID 所有权或 API key 的创建/推送权限。

**版本建议：**`0.1.0-preview.10` 是本日检查未发现被占用的下一个 preview 版本，可供冻结时选择；尚未预留。发布前重新核对全部最终产品 ID 的 NuGet 索引和 tag，核对新 HarfBuzz ID 的账号/推送权限，不检查或输出密钥明文。不要将 CI 内部 `preview.8`、当前 props 的 `preview.9` 或未发布的本地候选当作公开新版。

## 实际 tag 工作流与产品边界

[publish-nuget.yml](../../.github/workflows/publish-nuget.yml) 由 `v*` tag push 触发，不由普通 main push 或手动运行触发。它严格解析 `v<SemVer>`，使用 .NET SDK 10，并依次执行：指定 LFS fixture 获取、restore、许可/还原依赖图检查、带准确版本的公开 API Release Rebuild、mandatory trusted-font 零跳过测试、完整 solution 测试、逐产品 pack、原包隔离消费 E2E、候选绑定 Preview 记录检查、包 manifest 复核、NuGet push。工作流没有创建 GitHub Release 的步骤。NuGet 推送不能作为事务；中途失败可能留下部分已公开产品，必须逐 ID 回读，不能把 `--skip-duplicate` 当作完整发布成功证明。

当前唯一默认发布清单由 `scripts/verify-package-artifacts.py:PACKAGES` 定义，实际为以下 12 个产品：

```text
Ofdrw.Net.Core
Ofdrw.Net.Packaging
Ofdrw.Net.Layout
Ofdrw.Net.Text.HarfBuzz
Ofdrw.Net.Reader
Ofdrw.Net.Converter.Abstractions
Ofdrw.Net.Converter.Docx
Ofdrw.Net.Converter.Pdf
Ofdrw.Net.Converter.Svg
Ofdrw.Net.Signatures
Ofdrw.Net.Converter
Ofdrw.Net.Cli
```

普通 SDK 包需含 `netstandard2.0`/`netstandard2.1` DLL 和 XML；CLI 需含 `tools/net10.0/any/Ofdrw.Net.Cli.dll`；Converter 是无运行时 DLL 的 metapackage。校验器要求准确产品集合、版本、SDK 内部依赖版本、MIT 表达式、随包声明/使用说明、Core/CLI 的 musl notice 及 SHA-256 manifest。四个可选扩展不在 tag 发布循环中；现有独立消费门禁仍须通过，但不得将其本地成功写成公开发布。若要同时公开可选扩展，先明确产品范围、发布及原包验证流程；本预检不擅自扩大发布集。

## 就绪与待办

| 项目 | 本次证据与状态 | 最终候选仍需完成 |
| --- | --- | --- |
| 发布入口 | 已读实际 tag workflow；无工作流削弱 | 准确 main、候选版本与最终产品清单冻结 |
| 许可声明 | 根 MIT、props MIT、直接包声明和 40 项受审依赖清单已静态核对；旧文档的 38/22 项已纠正 | 当前工作树 16 个 src 项目均无 `obj/project.assets.json`；`--notices-only` 按预期因未 restore 失败。最终 restore 后必须执行真实还原图/缓存 nuspec 复核，不更新 baseline 来绕过变化 |
| 原生资产 | `verify-process-supervisor-assets.py` 通过全部四个既有 RID 的字节/哈希、头部、provenance、license、历史 execution receipt 和生成文件一致性检查 | 这是静态资产检查，不是本机重跑四个 ABI 的执行验收；最终仍运行 mandatory 字体、源码/消费测试及实际 native 身份记录 |
| API 文档与指纹 | 受审 CS1591 清单 273 个符号；`test_release_gates.py` 10/10 通过，包括 pointer/已下载内容等价、大小/哈希/错 pointer 拒绝、普通未暂存改动、冲突拒绝、新暂存 LFS 身份改指纹及提交前后一致 | 用候选版本执行 checker 的 Release Rebuild；禁止重新生成 baseline 隐藏新增缺文档成员 |
| 历史 Preview | JSON 实际是 2026-10-18 preview.9 的 `accepted`，含用户当时授权 AI reviewer 的例外；不是人工签收 | 对 preview.10 检查因版本不匹配退出 1；对 preview.9 因源码不匹配退出 1，均为正确阻断。没有复用旧 fingerprint |
| 当前源码/包/页面 | 未运行完整回归、pack、消费或转换；没有新页面验收结论 | 集成 N03/N13 及有界 N02 结论后，准确 main 的完整三平台适用测试、同批原包消费、PNG、真实 Preview、Copy/Find、Astra 复核和有效审阅/CI 结论 |
| 发布说明 | props/README/打包项仍指向 preview.9 历史发布说明 | 新版本确定后更新版本/新 release notes 打包路径、中英文入门与字体配置、支持范围及兼容性；同步范围/能力文档，再冻结指纹 |
| 新产品权限 | HarfBuzz 包此前无可查询公开版本 | 发布前核对新 ID 和推送账号/API key scope；当前未验证凭据能力，无须为预检暴露密钥 |

这些待办使当前状态保持 **preflight complete / candidate pending**，不能标为 release-ready 或 accepted。N04/N05 没有新自然失败时不追加根因研究；N01/PR78、完整 R02/R20/R22、范围外脚本和新架构不是本轮发布前置。现行必验页面上的缺字、不可读、裁切/重叠或未关闭视觉缺陷仍阻断候选。

## 最终验证命令：准确源码、准确版本、原包消费

以下是待执行清单，不是本次已通过记录。在干净、已集成全部代码及文档的最终候选工作树执行；`.NET SDK 10` 和各平台工作流规定的 LibreOffice、Poppler、ImageMagick/librsvg、zstd、固定字体和独立 R22 reader 环境须先准备。沿用现有 CI 的工具版本及固定字体字节，不自动接受主机系统字体。

主 Linux 构建/原包生产命令（bash；版本在最终冻结时再次核对）：

```bash
export OFDRW_N12_VERSION=0.1.0-preview.10
export OFDRW_N12_ROOT="$PWD/artifacts/cn-en-release/$OFDRW_N12_VERSION"
export DOTNET_CLI_HOME="$OFDRW_N12_ROOT/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
mkdir -p "$DOTNET_CLI_HOME"
git status --short
git rev-parse HEAD HEAD^{tree}
git lfs pull --include="e2e/**,tests/**" --exclude=""
python3 scripts/check-repository-storage.py --require-materialized e2e --require-materialized tests
dotnet --info
dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
python3 scripts/verify-process-supervisor-assets.py
python3 scripts/verify-package-artifacts.py --notices-only
python3 -m unittest discover -s scripts/tests -v
git lfs pull --include="docs/validation/public-tables-evidence/**" --exclude=""
python3 scripts/verify-public-table-evidence.py
python3 scripts/check-public-api-docs.py --package-version "$OFDRW_N12_VERSION"
python3 scripts/run-audit025-diagnostic.py
dotnet test Ofdrw.Net.sln -c Release --no-build --no-restore --verbosity minimal --logger trx --results-directory "$OFDRW_N12_ROOT/tests" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
python3 scripts/run-trusted-font-tests.py --no-build --results-directory "$OFDRW_N12_ROOT/trusted-font-tests"
./scripts/run-converter-package-e2e.sh "$OFDRW_N12_VERSION" --pack-only --no-build --packages-dir "$OFDRW_N12_ROOT/packages" --output-dir "$OFDRW_N12_ROOT/pack-log"
python3 scripts/verify-package-artifacts.py "$OFDRW_N12_ROOT/packages" "$OFDRW_N12_VERSION" --verify-manifest
./scripts/run-converter-package-e2e.sh "$OFDRW_N12_VERSION" --consume-only --packages-dir "$OFDRW_N12_ROOT/packages" --output-dir "$OFDRW_N12_ROOT/linux-consumer"
"$OFDRW_R22_PYTHON" scripts/validation/r22-mixed/check-export-batch.py "$OFDRW_N12_ROOT/linux-consumer/text-prelayout" --output "$OFDRW_N12_ROOT/linux-consumer/r22-independent"
python3 scripts/verify-package-artifacts.py "$OFDRW_N12_ROOT/packages" "$OFDRW_N12_VERSION" --verify-manifest
```

pack-only 要求全新空 feed；不得重建并覆盖原 feed。主构建的 API checker 已完成带版本的完整 Rebuild，不需在其前面重复普通 build。不要直接沿用普通 CI 目前硬编码的 `0.1.0-preview.8` 作为候选版本证明；普通 CI 门禁与最终准确版本的包验收分别记录。

Windows/macOS 上从同一最终源码、同一个 Linux 生产的原始 `packages` 目录运行下列等价命令。先按相应 shell 设置上面的 `.NET` 环境变量、`OFDRW_N12_VERSION`、新输出目录和已下载 feed 绝对路径；Windows 使用 PowerShell 的 `$env:`/变量语法，不能直接粘贴 bash 的 `$VAR`。完整源码测试在这两个平台各自 build 一次；所有测试和适用跳过理由都保留。

```text
dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
python scripts/run-audit025-diagnostic.py
dotnet test Ofdrw.Net.sln -c Release --no-build --no-restore --logger trx --results-directory <fresh-platform-tests> --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
python scripts/run-trusted-font-tests.py --no-build --results-directory <fresh-platform-trusted-font-tests>
python scripts/run-converter-package-e2e.py <frozen-version> --consume-only --packages-dir <same-linux-feed> --output-dir <fresh-platform-consumer>
<OFDRW_R22_PYTHON> scripts/validation/r22-mixed/check-export-batch.py <fresh-platform-consumer>/text-prelayout --output <fresh-platform-consumer>/r22-independent
```

最后一条在 macOS 加 `--pdfkit`，Windows 不加。Python 消费脚本会重新 restore 源依赖元数据供许可检查，但不重建生产包；它校验主消费及 HarfBuzz 消费的实际选中 DLL 与 nupkg、隔离缓存和输出副本。记录实际运行 TFM/DLL/native 身份及前后包哈希，不以源码测试替代消费结果。

还须完整执行准确 main 的 `CI complete` 全依赖与独立 `SES dev interop`；不得只跑上面的主命令块来冒充全部 CI。现行 CI 保留的 BuiltIn DOCX/加密等平台 filters、文档工具、graphics、可选 Skia/vector、password/standard crypto 和独立 SES 均不得删除。可选 Skia/vector 的准确版本入口为：

```bash
./scripts/run-skia-package-e2e.sh "$OFDRW_N12_VERSION" "$OFDRW_N12_ROOT/packages" "$OFDRW_N12_ROOT/skia"
./scripts/run-pdf-vector-package-e2e.sh "$OFDRW_N12_VERSION" "$OFDRW_N12_ROOT/packages" "$OFDRW_N12_ROOT/pdf-vector"
```

它们各自新建扩展 feed，不能把扩展包混入默认 12 包 feed。独立 SES 的现行脚本自行生产其测试 feed，不是这次唯一默认 feed 的消费证明；两类证据应分别保留。

## 最后封存与发布前门禁

1. 完成中英文使用说明、版本 props/新 release notes 及所有跟踪文档变更，再冻结最终源码/树和包 manifest。保留准确 main 的 CI complete、独立 SES、有效审阅处理结论。Terra 与 Astra 的输入必须是该准确候选。
2. 从本次原包消费结果分别检查 `generated-docx-native.ofd → generated-docx-native.pdf` 与 `generated-docx-default.ofd → generated-docx-default.pdf`。逐页 PNG、真实 macOS Preview、文本完整/重复、Copy/Find 各自记录；新增核心样例/范围独立记录。现有 gate 针对固定 DOCX 的两页/两模式，不代表新增核心矩阵也已验证。
3. 关闭并重开本次 PDF 排除缓存，验收后关闭相应窗口。依 [模板](../preview-acceptance-template.md)保留源/字体/输出哈希、实际工具版本、页码、检查人、时间、截图、体积及遗留问题。实际页面 FAIL 或未验部分不得写 accepted。
4. 本次修复实际发布阻断：旧 `source_fingerprint` 对全部跟踪工作树原始字节哈希；桌面已下载历史 docs 证据而 tag workflow 只下载 e2e/tests 时，同一提交也会不同。现在仍遍历除验收 JSON 外的所有跟踪路径：普通文件读取当前字节；LFS 文件读取 index 的合法标准 pointer，并将已验证的 pointer 或载荷统一为相同 oid/size 身份。已下载内容必须严格符合大小与 SHA-256，错 pointer、原始载荷误入 index、未暂存的新 LFS 身份和冲突均拒绝。新增 LFS 路径/属性先暂存，再生成指纹；普通未暂存源码仍改变指纹。没有关闭校验、跳过证据路径或修改 tag workflow。历史验收指纹不复用。
5. 仅在准确包/页面/审阅均通过后，在最终树运行 `python3 scripts/verify-preview-record.py --print-fingerprint`，据真实新产物更新 `docs/release-preview-acceptance.json`；仅该记录自身变动不改变指纹。同一提交已验证的 LFS 指针/载荷状态现在等价。随后执行 `python3 scripts/verify-preview-record.py <frozen-version>` 和 `python3 scripts/verify-package-artifacts.py <same-feed> <frozen-version> --verify-manifest`。其他跟踪内容再变动，须重验准确候选；不得复制旧 fingerprint。
6. 暂存后、提交前运行 `python3 scripts/check-repository-storage.py`；fixture 变动加两个 `--require-materialized` 参数。保存必要可复取证据，按需 LFS 归档，不 force-add artifacts、缓存或完整日志。
7. 本预检不执行 tag/push。批准实际发布后，按未削弱的 tag 工作流发布，再逐产品核对官方 NuGet 版本、公开包下载/安装、CLI/最小消费者回读及实际 GitHub Release 状态。“本地可发布”“tag 工作流成功”“公开安装已回读”分别报告。
