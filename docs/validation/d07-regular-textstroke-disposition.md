# D07：Regular-only TextStroke 有界处置

日期：2026-11-02（Asia/Shanghai）。结论：**本轮可行性审查关闭；原 Regular-only 默认 TextStroke 表头“量”字字腔视觉 FAIL 未修复**。未找到符合本轮边界的新确定性机制，不交付产品补丁。继续推荐显式物理 Regular/Bold 配置；该配置的通过不覆盖原字体失败。

## 基线与停止判断

干净工作树 fetch 后，以最新 `origin/main` `a9de8cae485fc8826c775f79b3bea8c996458c4d`（含 D06／PR102）建立 `codex/d07-regular-textstroke-disposition`，未使用旧 preview.10 验收分支。开发协调为 gpt-6.1-sol High；独立审阅为 gpt-6-astra High，独立本地证据核验为 gpt-6.1-terra High。

授权为最多 60 分钟可行性判断，不要求用满；先复用历史证据，不做字体／系数矩阵。只考虑不改字体、不破坏原文、不过度扩展排版的最小确定性修复。读取下列记录及当前 producer/writer 后，选择 **0 个候选、0 组新实验**，主动可行性审查到此停止；后续仅文档核验与 PR 门禁。

| 已有路线／变化 | 本轮判断 |
| --- | --- |
| PR78 的重复 fill、字腔规划、同行宽度、轮廓／clip／并集 | [N01](cn-en-release-n01.md)已逐项处置；最终候选仍有实际 Preview 可读性 FAIL。[复盘](pr78-and-stability-retrospective-2026-11-01.md)记载多轮复杂度扩张，无可直接提取的原字体光学修复；本轮不重开 PR78 或搬运整条方案 |
| 降低描边系数或重新选择像素门槛 | 不能从现有记录证明同时满足可读性和既定字重；属于已停止的参数探索，不为求通过改变阈值 |
| D02／D03 两项独立正确性修复 | 已分别合并为 PR98／PR99，只修导入 paint/CTM 和 SVG 声明字体族；两轮控制页明确保留“量”字 FAIL，不是本项的新机制 |
| 改用物理 Bold、D04 严格策略 | 是已验收的显式配置与缺 Bold 拒绝策略，不是原 Regular 字体修复；不自动换字体、不改变默认 TextStroke |

这是一项有界工程处置，不证明所有可能的原字体修复均不可行，也不将几何等价、文本完整或 CI 成功当作可读性通过。

## 复用证据与配置建议

- [PR78](https://github.com/whynpc9/ofdrw.net/pull/78)已于 `2026-11-02T03:10:45Z` 关闭、未合并，head 为 `2c4b48fbcd93618c131744c11be66cde83c84615`。关闭说明保留原字体可读性 FAIL、Find 失败和未解决审阅；原描述中的“当前未发布”属于关闭当时状态，preview.10 后续已发布，见[发布记录](preview10-publication-2026-11-02.md)。本轮不修改该 PR 或旧标签。
- [旧本地独立验收](cn-en-release-local-validation.md)的 `1e6f5eef` 集成矩阵共 34 页，原 Regular native/default 第 1 页“量”仍 FAIL；另设物理 Bold 的两模式清晰。历史 PNG／Preview 检查范围不能转移成当前 main 的新验收。
- [D02](d02-imported-paint.md)默认宿主配置的 PingFang SC 多种样式使用同一 regular 内嵌子集，普通字号“量”仍拥挤；[D03](d03-svg-declared-family.md)的 SVG/HTML 控制也保留 FAIL。这些配置与固定真实 Bold profile 不同。
- [D04](d04-cli-strict-bold.md)真实 Regular/Bold 的显式 native/default OFD→PDF 共四页，原文、导出和 PNG／Preview 已按该轮记录验收；七个粗体对象含“数量”，使用真实 weight 700 子集、无 Stroke 文本对象。该历史成功仅支持具名配置建议，不关闭 Regular-only 失败。

实际部署继续使用[受限物理字体 profile](../release-profile-cn-en.md)与[字体教程](../tutorials/cn-en-font-deployment.md)：API 显式配置 `NativeEmbeddedBoldBehavior=Throw`；D04 后续源码 CLI 可用 `--native-embedded-bold-behavior throw` 配合明确字体目录。严格策略只约束 Native 实际选中的嵌入粗体，不是全局字体隔离或 name-only 缺字体检测；直接 DOCX→PDF／dual-layer 不受该策略约束。已发布 preview.10 CLI 不含此开关，不宣称新修复已包含在该版本中。

## 本轮核验范围

本轮仅增加处置记录与路线图链接，没有产品、字体、fixture、验收阈值或发布变更。Terra 读取必要旧报告／收据，核对选定 18 个文件的 SHA-256，均与既有清单匹配；其中 14 项同时按清单核对大小，另 4 项仅记录实际大小。主线程独立重算这 18 项并读取 D04 与旧 main 的关键收据。范围包括 D04 两模式 OFD/PDF、两张表头 PNG，D02/D03 选定控制产物，旧 34 页 manifest 与 native 原字体／物理 Bold 产物，以及 N01 引用的两份旧状态收据。未复制整批页面、字体或源码。

核验清单位于本工作树 ignored `artifacts/d07/terra-evidence-validation.json`，SHA-256 `0e5d376608cfab31b2feb1c545e98b619401ffd9746ac0f56730bf656d23ea5a`。未递归核对全部包／字体／页面，也未新看旧截图。PR78 候选可达的 `current-status.json` 是较早的 GUI pending，最终 FAIL 依据历史 N01／复盘与 PR 描述，不称重新取得最终视觉原收据；旧 `b02162be` main 最终收据明确为 `quantityCounter=FAIL`、`overallVisualPass=false`。旧 34 页生成 manifest 中视觉字段为 `not-performed`，后续实际视觉范围取自独立验收报告，不把生成状态误写成通过。

| 项目 | 本轮结果 |
| --- | --- |
| 产品功能测试／本地包消费 | 未运行；文档处置不声称新产品回归通过 |
| 新 OFD／PDF／自动渲染 | 无新产物 |
| 新逐页 PNG 视觉复查 | 0 页；不将旧记录称为本轮视觉检查 |
| 新 macOS Preview 验收 | 0 页；未操作 GUI，无任务窗口需关闭 |
| 原文／重复文本／视觉 | 仅引用具名旧轮次各自结论，不把文本成功等同视觉成功 |
| 遗留限制 | Regular-only “量”字 FAIL 未修复；D06 Find FAIL／根因 INCONCLUSIVE、D05 原始 Unicode Copy 未完成均保持 |

## 明确重开条件

仅当新增证据定位到可达、独立且确定的产品绘制缺陷，或提出能在限定预算内验证的新最小机制时，另行明确授权有界任务；重复旧参数扫描、自动替换字体或恢复 PR78 整体光学路线均不构成触发。

任何实际候选须在实验前固定输入与字体字节、源码／环境、普通尺寸可读性和既定字重预期、原有像素门槛、退出条件。以同一原字体保持原文／GID／位置及局部样式，分别核对 OFD 原文与 PDF 无重复文本；显式 native/default 实际新产物须逐页 PNG 与真实 macOS Preview 检查“量”字及受影响页面。不得降低字重目标或改阈值求通过；共享路径变化仍需完整回归与精确本地包消费。未完成这些标准不得宣称修复或整体视觉通过。
