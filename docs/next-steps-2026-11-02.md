# 中英文受限预览候选计划（2026-11-02）

用户本日选择：**“限定已验证真实 Regular/Bold 配置，将 Preview 搜索兼容性与未完成原始 Copy 明列为预览版限制（便于尽快发布）。”** 本文取代 [2026-11-01 计划](next-steps-2026-11-01.md)中新候选的范围/阻断解释，不改写历史运行结论。目标为 `0.1.0-preview.10` 的可审查、可发布准备候选；当前未发布，不创建 tag、GitHub Release 或推 NuGet。新候选完成下列门禁前也不宣称 release-ready。

## 1. 新候选固定配置

Profile ID：`cn-en-physical-cjk-400-700-v1`。DOCX → 显式 native/default OFD → PDF，使用[字体教程](tutorials/cn-en-font-deployment.md)的 API 配置：

- 隔离的显式 `FontDirectories`；仅放下列三份字体及其许可证，不依赖系统安装替代面。
- `NativeEmbeddedBoldBehavior=DocxEmbeddedBoldBehavior.Throw`；`UseInstalledMicrosoftOfficeFonts=false`；清空 fallback 后依次设置 `Noto Sans CJK SC`、`Noto Sans`。
- 两次独立转换：一次显式 `OfdMode=Native`；一次不设置 `OfdMode`，验证当前默认 Native。源 DOCX、原文、样式、分页不修改。

| 字体 / face index | 源字节 / SHA-256 | 限定用途 |
| --- | --- | --- |
| `Noto Sans CJK SC` Regular / 0，静态 weight 400 | 21,659,544 / `3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa` | 中文与已覆盖的混排常规面 |
| `Noto Sans CJK SC` Bold / 0，静态 weight 700 | 21,638,668 / `647a6e375c5a0d8958111173d7b7cec5237e0784cc4490663e43834dfff491be` | 同族真实粗体；不得替换成 Regular 加样式标志或描边 |
| `Noto Sans` Regular / 0 | 569,208 / `b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5` | 比例英文常规面 |

CJK 来源 revision `f8d157532fbfaeda587e826d4cd5b21a49186f7c`、源 VF SHA-256 `990c807e79c25662a5a9ecf7f971baeb2bf2eab9a559e5ecf15cdfdb8561d21f`；fontTools `4.59.2` 离线实例化 400/700，运行时不是变量字体。来源、固定参数和 OFL 许可证详见教程。字体不加入 NuGet，不增加字体载荷。实际输出的嵌入子集身份另验，不能要求其哈希等于源字体。

此配置没有 Noto Sans Bold 或物理 Italic，不承诺任意家族/face、系统字体、缺字 fallback、合成字重的视觉质量。需要其他真实样式须提供对应合法字体并另验；既有斜体路径可模拟，不能称为物理 Italic。`Throw` 只约束选中的 Native 嵌入粗体；Office 开关只关闭特定暂存行为，两者均不是全局字体隔离。CLI 未提供严格粗体策略开关，故严格 profile 使用 API；CLI 默认行为仍保留并回归。

## 2. 原证据与新候选分别判定

[N00](validation/cn-en-release-n00.md) C01–C08 和 34 页历史清单保持可追溯，不删失败样例、不重写测试/golden/阈值。C03 物理配置是 DOCX 核心视觉承诺；C02 Regular-only 默认 TextStroke 作为原回归及已知限制独立保留。C01、C04–C08 的功能、原文、分页、样式、图像及实际选中 face 仍验，不把 DOCX profile 外推为 Flow 自动嵌入/新布局能力。

[此前 Terra 验证](validation/cn-en-release-local-validation.md)的 34/34 PNG 与 Preview 检查、原 Regular 量字 FAIL、物理配置量字清晰、Find FAIL、Copy 未完成均绑定旧源码/旧产物。其“未达到可发布验收”结论保留；不能把用户新范围决定当作旧 FAIL 已修复。`d447a0d5` 是本文修改前的文档基线，其祖先 `5d83fb38` 合并 N03/N12；最终源码 SHA/tree/指纹应取本批文档冻结后的实际提交。

新候选仍阻断：profile 内原文丢失/重复/乱码/缺字、不可读正文、严重样式串染、裁切重叠/错误分页、无效输出、候选包不可消费、字体来源/许可不明，或任何现行强制测试/CI/Q01–Q05 失败。范围外原 Regular-only 视觉 FAIL 不单独阻断本 profile，仍须原样披露；新发现的 profile 内失败不得移入 known limitations 来放行。

## 3. 明示的预览版限制

| ID | 原始结论 / 版本范围 | 新候选承诺边界 |
| --- | --- | --- |
| `legacy-regular-synthetic-bold` | 原 Regular-only 的 native/default 首页面“量”在 144 DPI 字腔堵塞，历史视觉 FAIL 保留 | 发布建议仅使用上述真实 Regular/Bold 配置；不声称修复默认 TextStroke 的光学问题 |
| `macos-preview-find` | macOS 26.6.2 (25G83) Preview 对原生成 PDF 的完整页末日期/部分整词显示 Found on 1 page 与 Not found，导航禁用/无高亮；根因未定 | 具名兼容性限制，不承诺本版在 Preview 中通用搜索成功；新产物仍记录实际 UI 观察，不用 API/AX/PDFKit 代替 |
| `raw-unicode-copy-unverified` | 原始剪贴板 Unicode 字节未验证；粘贴进搜索框可改写空格，不能证明原始值 | 明确未完成，不称 Copy PASS；不承诺原始复制保真。其他阅读器 Copy/Find 也未验证 |

用户决定使以上已披露限制不再单独阻断这个新预览 profile；不豁免原文完整性、PDF 重复文本检查、逐页视觉或工程门。保留实际失败/缺证据，不以“限制”字样改写测量结果。

## 4. 版本与发布集合冻结

于 **2026-11-02 01:13:57 UTC** 逐 ID 请求官方 NuGet V3 flat-container 索引，并执行 `git ls-remote --tags origin`：默认 12 个产品中 11 个既有 ID 均为 HTTP 200、最新 `0.1.0-preview.9`、无 preview.10；`Ofdrw.Net.Text.HarfBuzz` 为 HTTP 404。远端 tag 最新 `v0.1.0-preview.9`，没有 `v0.1.0-preview.10`。本轮固定候选版本为 **`0.1.0-preview.10`**，不是版本预留或公开发布。

| 官方索引 / 默认产品 | HTTP / 最新公开版本 |
| --- | --- |
| [Core](https://api.nuget.org/v3-flatcontainer/ofdrw.net.core/index.json) | 200 / preview.9 |
| [Packaging](https://api.nuget.org/v3-flatcontainer/ofdrw.net.packaging/index.json) | 200 / preview.9 |
| [Layout](https://api.nuget.org/v3-flatcontainer/ofdrw.net.layout/index.json) | 200 / preview.9 |
| [Text.HarfBuzz](https://api.nuget.org/v3-flatcontainer/ofdrw.net.text.harfbuzz/index.json) | 404 / 本次未查到公开版本；不证明 ID 所有权/推送权限 |
| [Reader](https://api.nuget.org/v3-flatcontainer/ofdrw.net.reader/index.json) | 200 / preview.9 |
| [Converter.Abstractions](https://api.nuget.org/v3-flatcontainer/ofdrw.net.converter.abstractions/index.json) | 200 / preview.9 |
| [Converter.Docx](https://api.nuget.org/v3-flatcontainer/ofdrw.net.converter.docx/index.json) | 200 / preview.9 |
| [Converter.Pdf](https://api.nuget.org/v3-flatcontainer/ofdrw.net.converter.pdf/index.json) | 200 / preview.9 |
| [Converter.Svg](https://api.nuget.org/v3-flatcontainer/ofdrw.net.converter.svg/index.json) | 200 / preview.9 |
| [Signatures](https://api.nuget.org/v3-flatcontainer/ofdrw.net.signatures/index.json) | 200 / preview.9 |
| [Converter](https://api.nuget.org/v3-flatcontainer/ofdrw.net.converter/index.json) | 200 / preview.9 |
| [Cli](https://api.nuget.org/v3-flatcontainer/ofdrw.net.cli/index.json) | 200 / preview.9 |

表中简称均加 `Ofdrw.Net.` 前缀。原始响应/标签记录保存在 ignored `artifacts/cn-en-release/preview10-version-check.json`。实际 tag 推送前还须重新检查版本占用，并核实新 HarfBuzz ID 的发布权限；不读取或披露密钥明文。`Directory.Build.props` 候选默认版本和随包说明改为 preview.10；README 的已发布版本/安装命令保留 preview.9，直到真实公开源发布与安装回读。可选 Skia/vector/SES/password 扩展不加入默认 12 包发布集合。

## 5. 冻结、验证与交接顺序

1. 完成全部非验收 JSON 的跟踪内容：范围/教程/能力索引、候选 release notes、版本及打包路径。先提交并通过仓库存储检查，再冻结准确 SHA/tree；不要在候选构建和收据之后补文档导致指纹漂移。
2. 协调者从该准确源码生产一次全新不可变 12 包 feed/manifest，完整回归、mandatory trusted-font、适用三平台完整源码测试/原包消费、Q01–Q04、CI complete 和独立 SES 按[门禁](release-engineering-gates.md)/[实际 tag workflow](../.github/workflows/publish-nuget.yml)执行。旧绿测不转移。保留任何自然失败，不删门、不改 baseline 来过关。
3. Terra High 领取独占 GUI，依据相同候选原包和固定 profile 重新生成并逐页检查实际 PNG/Preview。Q05 的固定 F01 输入 SHA-256 为 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`；两模式各两页均须通过 profile 内视觉验收。不得用旧物理 PDF 或直接 DOCX→PDF 替代；扩大矩阵另留记录。Find/Copy 和遗留 Regular 对照保留原始观察及边界。
4. Astra 独立复核准确源码/包、profile 实际选面、证据及已知限制，协调者审核原始页面记录与结果。实际 native 资产身份/平台适用性按收据记录，不由包文件存在推断已加载。
5. 只有 profile 内必验页面无开放视觉问题且其余门已通过，才以真实产物填写 `release-preview-acceptance.json`，附 `release_profile`、`known_limitations`，保留核心字段与 `open_findings=[]` 的原含义。用最终树生成/校验 fingerprint；仅该 JSON 自身的更新不改指纹。不得提前改旧 accepted 记录。缺页/未检查/新增 profile 内缺陷一律不能 accepted。
6. 最终报告分别给功能、自动渲染、PNG、实际 SVG、Preview、Copy/Find、CI/审阅及公开发布结果。此阶段无 tag/push NuGet；可发布候选、发布工作流成功、公开包/签名/独立消费回读是不同结论。完整 R02/R20/R22、PR78 光学优化及新部署范围仍非前置。
