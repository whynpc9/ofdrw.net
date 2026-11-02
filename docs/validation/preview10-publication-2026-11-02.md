# preview.10 最终发布记录（2026-11-02，D01）

`0.1.0-preview.10` **已发布并完成公开回读**。[GitHub Release](https://github.com/whynpc9/ofdrw.net/releases/tag/v0.1.0-preview.10) 为非草稿预览版（`draft=false`、`prerelease=true`），发布时间 **2026-11-02 05:27:19 UTC / 北京时间 13:27:19**。12 个主 NuGet 包已公开，新增 `Ofdrw.Net.Text.HarfBuzz`；11 个符号包成功推送，`Ofdrw.Net.Converter` 元包无 snupkg。安装与支持边界见 [README](../../README.md)、[发布说明](../release-notes-0.1.0-preview.10.md)及[字体教程](../tutorials/cn-en-font-deployment.md)。

## 1. 身份与发布路径

| 对象 | 固定身份 / 可复查入口 |
| --- | --- |
| 产品标签 | [`v0.1.0-preview.10`](https://github.com/whynpc9/ofdrw.net/tree/v0.1.0-preview.10)；annotated tag object `956837fb05cd38863b38818c5852b892bcdb632c`，解引用为下列产品源码；恢复发布时未移动标签 |
| 产品源码 / 公开包 repository commit | [`8f27f92b51b70a9b01082da35ff1cfe989d46123`](https://github.com/whynpc9/ofdrw.net/tree/8f27f92b51b70a9b01082da35ff1cfe989d46123)；全部 12 个公开主包均绑定此提交 |
| 发布控制器 main | [`3568c00bb52408f1151095823a1337cb97cb25dc`](https://github.com/whynpc9/ofdrw.net/commit/3568c00bb52408f1151095823a1337cb97cb25dc)；仅为运行发布流程的身份，不能当成包源码 |
| 成功发布 | [Publish NuGet 41391145165](https://github.com/whynpc9/ofdrw.net/actions/runs/41391145165)，`workflow_dispatch`、`completed/success`；[发布 job](https://github.com/whynpc9/ofdrw.net/actions/runs/41391145165/job/125102594218) |
| 发布前 Q05 | [标签中的 accepted JSON](https://github.com/whynpc9/ofdrw.net/blob/8f27f92b51b70a9b01082da35ff1cfe989d46123/docs/release-preview-acceptance.json)，生产者 `cb6e71df50cac46bfb3f40abee4df6b091d0d7ee`、profile-qualified accepted；[该生产者 CI](https://github.com/whynpc9/ofdrw.net/actions/runs/41376054001) / [独立 SES](https://github.com/whynpc9/ofdrw.net/actions/runs/41376053806) 成功。它不是公开包源码/消费凭证 |

[首次发布 41382154515](https://github.com/whynpc9/ofdrw.net/actions/runs/41382154515) 在首次上传因过期 API key 返回 403，未发布任何包。[PR96](https://github.com/whynpc9/ofdrw.net/pull/96) 接入 NuGet Trusted Publishing，[PR97](https://github.com/whynpc9/ofdrw.net/pull/97) 增加 main 上安全恢复原不可变标签的手动入口，均已合并。成功流程从原产品 SHA 重新构建、测试、消费、校验 Q05 后，以 OIDC 短期凭据上传；未改产品、版本、原标签或验收规则。

## 2. 远端证据与归档边界

以下 Release 附件提供稳定下载地址，包含包身份/内部依赖/字节校验值，以及 OIDC、12+11 次上传和发布 job 的关键证据。D01 从公开地址下载两份小型 JSON，字节与既有本地证据一致，重算 SHA-256 与 GitHub asset digest 一致；未下载全部公开包或重做签名/产品测试。

| Release 附件 | 字节 | SHA-256 |
| --- | ---: | --- |
| [public-packages.json](https://github.com/whynpc9/ofdrw.net/releases/download/v0.1.0-preview.10/public-packages.json) | 11,215 | `5d96bb1455f06b83a7d9c367b64a913fa8eff764f52d2442dc0b94d863617488` |
| [trusted-publishing-proof.json](https://github.com/whynpc9/ofdrw.net/releases/download/v0.1.0-preview.10/trusted-publishing-proof.json) | 7,616 | `856412e2ec75dd9ff9332f26c893bb861a978f2fe282e7bdfc710a89522c6615` |

成功发布运行的 [release-validation artifact 13117986465](https://github.com/whynpc9/ofdrw.net/actions/runs/41391145165/artifacts/13117986465) 保存发布流程的包 manifest、消费页面及测试资料。GitHub 元数据报告大小 **79,504,168 字节**、digest `54a5308f2042518cd92b433857a83e1c92ad79f174637fce570989a67e073158`，当前到期时间 **2027-01-31 04:56:05 UTC**。这是有保留期限的补充资料，不能冒称永久归档；本机未完整下载复算该 ZIP。该 artifact 生成在上传前，不是上传后公开源消费记录。

公开源回读与页面结果的长期摘要为 [Release body](https://github.com/whynpc9/ofdrw.net/releases/tag/v0.1.0-preview.10)及本文。下列收据/产物哈希从既有 `publish-preview10` 实测记录摘录；本次只读核对收据，未重新运行产品或 GUI。原始收据及页面仍在 ignored artifacts，可按哈希复取；不将个人电脑路径作为结论的唯一入口，也不在源码仓库重复提交包、字体、页面或完整日志。

## 3. 公开源消费与页面验收

原公开验证完成于 **2026-11-02 05:22:06 UTC**，验证者 `gpt-6.1-terra` High（用户授权的 AI reviewer，不是人工签字）；协调者另复查四张 PNG，最终 Astra High 审核通过。公开验证源为 nuget.org，使用全新缓存及 PackageReference 消费，结果与后续公开 Release 摘要一致。

| 项目 | 实际结果 / 限定范围 |
| --- | --- |
| 包与签名 | 12/12 主包版本、内部 SDK 依赖及产品源码 `8f27f92` 一致，NuGet 仓库签名验证完成；符号包成功上传不等于独立符号服务器消费验收 |
| API | 独立消费已运行，7 个实际加载的 SDK DLL 与公开 nupkg/cache/bin 哈希匹配 |
| HarfBuzz | 独立消费的 8 个 SDK DLL 匹配，版本/塑形调用运行，报告 `HarfBuzzSharp/14.2.1.301;libHarfBuzzSharp/14.2.1`；原生库实际加载路径未独立观察 |
| CLI | 从公开源安装，67 个安装资产匹配，help 与文本往返完成；提取丢终端 LF，不能称字节完全一致；CLI 无严格粗体开关 |
| 固定输入 / 配置 | F01 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`，SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`；`cn-en-physical-cjk-400-700-v1`，字体身份见[范围计划](../next-steps-2026-11-02.md#1-新候选固定配置) |
| 严格 API 参数 | 固定 Noto Sans CJK SC 静态 Regular 400/Bold 700 + Noto Sans Regular；`NativeEmbeddedBoldBehavior=Throw`、显式 fallback、`UseInstalledMicrosoftOfficeFonts=false`；该 Office 开关不代表禁用全部系统字体解析 |
| 功能 / 原文 | 显式 native 与 unset/default Native 各两页；两模式 OFD/PDF 保持 189 个非空白源字符序列，PDF 无重复文本；7 个粗体对象、3 个蓝色填充；抽查嵌入子集的“量”700 / “文”400 轮廓匹配固定源字体 |
| 自动渲染 | 两模式 OFD→PDF→PNG 完成，各两页，共四页 |
| PNG 视觉复查 | 四页通过；“数量”字腔清晰，中英文/英文斜体、局部红色粗体、表格底色/边框、分页和右对齐日期可见，无明显裁切、重叠或样式扩散 |
| macOS Preview | **公开 NuGet 消费 DOCX → 显式 native/default OFD → PDF → macOS Preview 26.6.2 (25G83)**，两模式第 1、2 页均实际查看通过，共四页；任务文档窗口已关闭 |
| Copy / Find | 本次未重测 Preview Find；原始 Unicode 剪贴板 Copy 未完成验收；其他阅读器 Copy/Find 未验证，均不称 PASS |

公开消费产物的身份如下，区别于发布前 Q05 JSON 中的旧候选包产物；OFD/PDF 可含时间元数据，不要求跨次生成字节相同。

| 模式 / 产物 | 字节 | SHA-256 |
| --- | ---: | --- |
| native OFD | 551,452 | `34218b82e1cac7c8bc27316bfe10c74f2953d665d91bec2abeb1bde9932c3688` |
| native PDF | 144,787 | `ddb79f0e87527f2e672235120c63dde4b8e1f7eb04f65b286adfdeb6bff5308a` |
| default OFD | 551,452 | `70e3eaa141b52d996bc639e3de0b5c483f9b31a0cccf3d67d77e98382f7eb3bf` |
| default PDF | 144,787 | `bf2584567e592a6ff17543fa6298bfc7c70abd88b87b5a09551051e7977a04dc` |
| 两模式第 1 页 PNG（各一张） | 各 61,523 | `389b6f4c75e89a02a42905927908dcfeb44ba83c90793a577e41bb0b69deb1ff` |
| 两模式第 2 页 PNG（各一张） | 各 31,337 | `7d9ebcc94bd2d1378cc13116c73abb8c3b87e6c3f0041f4d415ac974d1f666d4` |

关键公开消费收据 SHA-256：`publication-validation.json` 为 `31e484fdc72e166c04b586482da865380422d4c6b97799edbba9a2c0d7e0a8c4`；`consumer-identities.json` 为 `f734192ed99f91c1f92cb549bdd64215d80e7a023c453972f5a2b72de2337d34`；`profile-check.json` 为 `e09fce62ad637d61ec45ae35794c5e07f1eff06d40370ddbde0b24b6197f8b93`。哈希用于识别原记录，不表示它们已作为 Release 附件上传。

## 4. 保留限制与 D01 核查

- Regular-only 默认 TextStroke 的表头“量”字字腔视觉 **FAIL 未修复**，不能因物理 Bold 配置通过而重标旧结果。
- macOS Preview 页末日期/部分整词 Find 异常的根因未定；原始 Unicode Copy 缺少验收，不能称已证实的 Copy 缺陷。见[历史独立验证](cn-en-release-local-validation.md)与[系统对照](cn-en-preview-control-2026-11-02.md)。
- 仅覆盖上述配置、F01 和四个实际页面，不保证任意 Word、任意字体/物理斜体或跨机像素一致；严格粗体通过 API，CLI 不提供同等开关。
- [PR78](https://github.com/whynpc9/ofdrw.net/pull/78) 已关闭未合并，保留原分支/证据；[N01 C1/C2](cn-en-release-n01.md#3-两个有限正确性候选) 未开工。N04/N05 被动观察、N09 取消，RTL/复杂文字/高级字体和完整 R22 等仍为未来议题；现有 ImageSharp NU1902/NU1903 告警继续披露。

D01 于 **2026-11-02** 轻量复核 GitHub Release 状态/附件、成功与失败工作流、远端 main/tag、PR78/96/97 状态；NuGet 全部 12 个 ID 的 preview.10 索引及两份 Release 附件下载校验于 **06:59 UTC** 完成，均与上述结论一致。既有消费/页面收据只读核对，不重跑整套构建、视觉或 GUI。本次仅同步文档入口与证据，产品、CI、验收 JSON 和规则均不改，未重新发布或移动标签，不开始 D02。
