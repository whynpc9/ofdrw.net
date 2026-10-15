# 2026-10-15 能力路线图首轮集成

## 范围与状态

人类授权梳理已完成票并执行首轮 merge。保留所有已审查提交历史；不发布 NuGet，不创建 tag 或 dispatch，不修改主 checkout 的未提交文件。

下表是本轮开始时逐票已完成实现、测试与视觉验收的候选。Issue01、16、03 已合入 main，其合并后 CI 全绿；其余八票在独立 `codex/roadmap-merge-without-password` 集成。原 PR21 因口令 KDF P1 被[PR22 排除版汇总](https://github.com/whynpc9/ofdrw.net/pull/22)取代；人类明确选择先排除 Issue11，再合入其余已完成任务。本文件是合并前验收快照；最终合并状态以 PR21 和 main 提交历史为准。

| Issue | 能力 | PR | 接纳的原 head |
| --- | --- | --- | --- |
| 15 | 发布工程门禁 | [#8](https://github.com/whynpc9/ofdrw.net/pull/8) | `c1d8ab20a760a635a562183e29856feed7a7ae17` |
| 01 | 公开段落布局 | [#9](https://github.com/whynpc9/ofdrw.net/pull/9) | `5d71dc9e4fb3ccfa4dd028386436e58cd261bc39` |
| 16 | 公开表格 | [#10](https://github.com/whynpc9/ofdrw.net/pull/10) | `17d0d0626d957cd58a87d5157172b054185328c6` |
| 03 | 水印、拆分、混合、清签 | [#12](https://github.com/whynpc9/ofdrw.net/pull/12) | `df71c7f20e0c45e9cba9cb78f2d2a91d61413ec9` |
| 04 | 原生绘图 API | [#13](https://github.com/whynpc9/ofdrw.net/pull/13) | `ab87bcade2ec5df6fa1f7ea0b86c590d13225653` |
| 05 | 字体子集与复用 | [#14](https://github.com/whynpc9/ofdrw.net/pull/14) | `d8dfda455b44ef5857c3a87088ed08cf80d810dc` |
| 12 | 可选 SES 开发互操作 | [#16](https://github.com/whynpc9/ofdrw.net/pull/16) | `02f516ea61f5f73cdf3a6a015d4e02d6b46ddd7e` |
| 19 | 合作式 Skia 事件适配 | [#17](https://github.com/whynpc9/ofdrw.net/pull/17) | `d878aeda57c1da79917e11d0a35bcc028bc4c1fd` |
| 21 | 受限 PDF 原生矢量转换 | [#18](https://github.com/whynpc9/ofdrw.net/pull/18) | `0a97c980fa0846c931e199caf26e88f437472744` |
| 06 | 关键词定位与受控替换 | [#19](https://github.com/whynpc9/ofdrw.net/pull/19) | `7caa8b54f2b28e0479d051e6ddf9b6d3227b5c3d` |
| 09 | 强类型对象、裁剪与组合 | [#20](https://github.com/whynpc9/ofdrw.net/pull/20) | `03b5f29484a623a00ca8d92725bd7f1f8036491a` |

## 集成边界

- 默认包清单仍为精确 11 包。SES、Skia 和 PDF.Vector 保持可选；Password 本轮不纳入；核心能力开关与默认验签限制不变。
- 普通生成保留 05 的字体子集策略；06 的受控替换保留已接纳原字体字节。选面、字体配置克隆和回调变更检测同时保留。
- 01/16 流式布局、表格消费者与 04 绘图消费者同时保留；统一重复 solution 项目标识。
- 保留矩阵精度、Latin 组合测量、非绘制空白、typed fill rule、多子路径 AddMove。SVG 每个路径只写一次 fill-rule。
- PDF.Vector 保持已验证原字体；独立页进入最终包时重新分配局部文档单元身份，避免回退页与 native 页冲突，不放宽 Writer 身份校验。
- 15 的文档/包/依赖门禁继续生效。依赖许可清单新增项来自已接纳可选包的实际 NuGet 元数据；未执行发布路径。正式发布记录仍为 not-reviewed。

## 本轮排除

Issue11 / PR15：静态确认私有 v1 口令 KEK 使用无盐、无成本参数的单次 SM3 派生。原文档披露此限制，但不构成抗离线猜测保护。按人类本轮直接选择，保留 PR15 独立未合并；新分支不含该 head 祖先、源码、测试、E2E 或 CI job。仅保留与其他可选包共用的默认精确 11 包打包白名单及失败即停测试。新口令协议另行设计、审查。

Issue02 / PR11 的远端仍为旧 `b29ed9d`，未包含本地最终集成与验收候选，不在本轮合并。其余待开发票 07、08、10、13、14、17、18、20、22 不在本轮范围。

## 最新发布工作流门禁修复

Codex 指出 tag 工作流未自行执行强制真实字体入口。现已在 `.github/workflows/publish-nuget.yml` 的候选版本 Build/Pack 之前加入既有 runner，并把记录纳入 always-upload。故障默认终止后续 Build/Pack/Publish，没有跳过条件或 continue-on-error。置于候选 Build 之前，避免 runner 普通测试重编译覆盖待打包的 tag 版本 DLL 元数据。

本地实际执行同一入口 **65/65、零跳过**，脚本测试 **20/20**；[原始 TRX、运行命令与日志](evidence/roadmap-merge-round-2026-10-15/release-font-gate/readback.json) 已持久保存。产品源码仍为 8018d70；下方 2,039 全套、实际包与 34 页视觉记录继续有效。不执行 tag/dispatch/NuGet push；正式发布记录仍 **not-reviewed**。最终 head 的 CI/双 bot/线程闭合后执行本轮 merge。

## 最新 SVG 评审修复与验收

运行源码 `8018d70be81ce2fe62bb6d588eca5b6e8b109c18`。Cursor 指出纯空白 run 的透明设置会被部分 alpha 覆盖：已把空白覆盖赋值移到颜色 alpha 之后，其他产品逻辑不变。

- 有效外部 OFD 使用 `xml:space="preserve"` 保留窄不换行空白；alpha 1/128/254、producer 有/无显式 run 六个回归修复前均失败、修复后均通过，正常 A 的半透明颜色不变。
- **完整回归 2,039 通过、12 可选未执行、零失败**。此前 65/65 强制字体入口保留明确 0a21fbe 基线；本次没有改变 Core、字体几何或 PDF 实现。
- 新 `0.1.0-merge.20261015.5` 的默认 11 包、SES 12 包、vector 13 包实际消费全部通过；默认 SES 验签仍 Unsupported/退出码 2。
- 当前 SVG PNG 目视通过；同一当前 OFD→PDF 的一页已实际在 macOS Preview 打开、查看正文并关闭。MIT 测试字体把 A 定义为矩形，PDF 仅见一枚灰矩形，空白区域无额外轮廓。SVG 渲染器显示灰 A；不把不同文字引擎的字形形状宣称为相同。该 PNG 渲染器修复前也不画缺字框，因此缺陷复现实证是 SVG opacity 错误及六项失败，而不是伪造像素差异。
- 原 33 页 Preview 记录仍按原始基线保留；新增此 1 页，选定范围共 **34 页**。不重复验收未改的 PDF 页面。
- 新 [SVG 修复读回](evidence/roadmap-merge-round-2026-10-15/svg-alpha-readback.json) 与 [验证归档](evidence/roadmap-merge-round-2026-10-15/svg-alpha-validation.tar.zst) 保存当前样例、页面图、修复前失败、原始全套 TRX、实际包清单及逐项解包哈希。PR22 最新 CI、双 bot 和有效线程须再闭合后才 merge。

## 排除 Issue11 后的首次验证（修复 SVG 之前）

运行源码 `0a21fbedaf284ace51e0e7d742eb0a4efd8b3529`；PR22 后续证据提交不改变产品源码。

- 编译成功；完整回归 **2,033 通过、12 项可选未执行、零失败**。比排除前少 66 项独立口令测试，其他十个测试项目数量一致。强制真实字体 **65/65、零跳过**（与全套重叠）；Python **20/20**，API 文档门通过，历史基线未扩大。
- 本轮新版本 `0.1.0-merge.20261015.4` 经隔离缓存实际 PackageReference 消费：默认精确 **11 包**、SES **12 包**、矢量扩展 **13 包** 全部通过。默认 SES CLI 仍 Unsupported/退出码 2；没有发布 NuGet。
- 全部非口令产品、测试、E2E、脚本源码与原已验收汇总逐文件字节相同。15 页新矢量产物按同分辨率 PNG 读回与实际 Preview 验收版本全部一致；沿用明确原始基线的 **33 页 Preview** 记录，不冒称排除后又做了一次 GUI 验收。所有任务窗口已关闭。
- 11 个接纳 head 均为本候选祖先；排除 PR11/Issue02 与 PR15/Issue11。32 个新验证载荷约 285 KiB 已解包逐项校验，见 [排除版验证与范围读回](evidence/roadmap-merge-round-2026-10-15/exclusion-readback.json) 和 [日志、原始 TRX、实际包清单](evidence/roadmap-merge-round-2026-10-15/exclusion-validation.tar.zst)。
- PR22 最新 CI/Codex/Cursor 终态与线程另实时核对后才合并。下方旧候选 2,099 计数及旧归档明确为历史证据。

## 原汇总候选的历史验证

以下记录明确来自排除前 17dea14 候选，包含其独立口令项目测试；不是排除版的全套计数。排除版保留其余产品源码逐文件字节一致，重新执行全套和包消费后另记结果。视觉产物、渲染源码及字体均未改动，实际 Preview 记录保留原始基线。

### 已实际完成

- 本轮集成编译成功；脚本测试 20/20 成功。
- 运行源码 `17dea14282ee58c00051e36a3257449102883d7b`：完整回归 2,099 项通过、12 项可选字体测试未执行；强制真实字体入口 65/65、零跳过。这 65 项与完整回归有重叠，不叠加为独立测试数。
- 最新 `0.1.0-merge.20261015.3`：默认精确 11 包、SES 可选 12 包及 vector 可选 13 包均完成隔离缓存的实际 PackageReference 消费；默认 SES 验签维持 Unsupported/退出码 2。旧版实际 typed 消费与 36 样例/55 页 document-tools 链路记录保留明确基线。
- 33 页实际 macOS Preview 验收通过（含 2 页仅本地原始诊断），所有本轮任务窗口关闭。先验收初始 31 页，再验收 2 页空格修复；新 vector 15 页重新生成后再次逐页重开。像素读回 15/15 与初始验收版本完全一致。
- PNG 自动渲染完整；独立 PNG 目视范围为绘图 10 页与空格 2 页，其余页面由实际 Preview 逐页验收，不冒称做过独立 PNG 目视。
- 公开 [验收记录及产物归档](evidence/roadmap-merge-round-2026-10-15/README.md) 已封存 111 个载荷并逐项解包校验。约 22.8 MiB，含 OFD/PDF/页面图、许可、原始 TRX/日志、包与依赖清单、哈希及环境。原始本地诊断的 OFD/PDF/SVG 不公开。
- CI 的 unused cmap4 与 SES 校验缓存问题已修；Linux dual 消费者为 `FixtureFace` 显式绑定同一份许可字体，不再依赖主机 Arial。Cursor 缺字空格反馈由上下文校验修复：无定位混合文本须真实覆盖/回退，单独定位空白保持不绘制，Writer 原子失败及既有像素回归通过。
- 最新 GitHub head 的 CI、Codex、Cursor 及线程闭合另在替代汇总 PR 实时核对；本记录不把本地通过替代远端门禁。
- 补充本轮 129 处新增 API XML 文档，不扩大 CS1591 历史基线。
- 原票历史证据继续在各自 `docs/evidence/`，其基线和验收范围不改写为本次集成运行结果。
