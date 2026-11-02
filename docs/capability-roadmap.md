# 后续范围与现行路线图

更新：2026-11-02。`0.1.0-preview.10` 已完成发布及公开源回读，详见[最终发布记录](validation/preview10-publication-2026-11-02.md)。本页列出保留的问题与候选方向；不自动启动实施、实验或发布。能力使用范围见[功能对照](feature-parity.md)和[分出口契约](capability-contract-v12.md)。

## 已有能力

Flow 段落/表格、命名区域、图片/Canvas/回调、文档工具、多 DocBody 显式选择、类型化对象、有限归档规则、签名保护范围、受限标准加密及固定字体查询已有有界交付，不再列为从零实施的待办。R02 的分组策略、R08 的各字体出口和 R22 的多个查询/输出增量各有独立边界；有源码不代表全部标准、脚本、阅读器或页面通过。

| 使用方向 | 维护入口 |
| --- | --- |
| DOCX/普通中英文与字体部署 | [转换契约](conversion-contracts.md)、[受限发布配置](release-profile-cn-en.md)、[字体教程](tutorials/cn-en-font-deployment.md) |
| 排版、表格、图片与 Canvas | [Flow](flow-layout.md)、[公共表格](public-tables.md)、[原生绘图](tutorials/16-native-graphics.md) |
| 文档、附件和多文档操作 | [文档工具](tutorials/15-document-tools.md)、[附件](tutorials/09-templates-attachments-tags.md)、[多 DocBody](multi-docbody.md) |
| 字体与塑形增量 | [字体出口表](validation/r08-four-outlet-ledger.md)、[预排版](text-prelayout.md)、[typed 输出](text-shaped-output.md)、[混排查询](text-mixed-paragraph.md) |
| 签名、加密与归档 | [签名保护范围](signature-scope-v12.md)、[可选加密](password-crypto.md)、[归档预检](archival-preflight.md)、[标准覆盖账本](archival-standard-coverage.md) |

## 遗留问题与独立候选

| 项目 | 当前决策与边界 | 依据 |
| --- | --- | --- |
| Regular-only 合成粗体 | D07 有界审查关闭，无新候选；默认 TextStroke 下“量”字字腔视觉 FAIL 未修复；继续推荐显式物理 Regular/Bold 与 D04 严格策略，不恢复整条光学增强研究 | [原失败与独立验收](validation/cn-en-release-local-validation.md)、[D07 处置与重开条件](validation/d07-regular-textstroke-disposition.md) |
| Preview Find / 原始 Copy | 页末日期/部分整词 Find 异常根因未定；D06 复用旧证据关闭本轮调查，保留 Find 失败；原始 Unicode 剪贴板 Copy 未完成；不以文本提取或 API 查询替代实际 UI | [N02 有界调查](validation/cn-en-release-n02.md)、[系统对照](validation/cn-en-preview-control-2026-11-02.md)、[D05 Copy 未完成](validation/d05-raw-copy-capability.md)、[D06 有界处置](validation/d06-preview-find-disposition.md) |
| PR78 / C1、C2 | PR78 已关闭未合并，保留分支与证据；C1/C2 已独立完成 D02/PR98、D03/PR99，不包含光学增强、不关闭 Regular-only 失败，不改写 preview.10 发布内容 | [N01 历史处置](validation/cn-en-release-n01.md)、[D02](validation/d02-imported-paint.md)、[D03](validation/d03-svg-declared-family.md) |
| N04 Skia / N05 Windows Audit025 | D08、D09 已分别关闭本轮主动取证，两项历史根因均 INCONCLUSIVE／未修复；继续被动观察，保留现有诊断、测试和原预算，新自然失败有可比较判别证据再限定范围，不追加循环压力实验 | [D08 处置](validation/d08-skia-passive-disposition.md)、[D09 处置](validation/d09-audit025-passive-disposition.md)、[Skia](validation/skia-font-read-evidence.md)、[分配边界](validation/audit025-allocation-boundary.md) |
| N09 完整未知外来映射 | 已取消，不等待自动续接；现有受限 R02 能力与失败证据保留，完整 R02 未验收 | [CGTransform](validation/r02-cgtransform-profile.md)、[GroupedLocal](validation/r02-grouped-textcode.md) |
| ImageSharp 告警 | 现有 NU1902/NU1903 仍披露；准入限制、测试通过或限定发布不等于依赖问题修复 | [发布说明](release-notes-0.1.0-preview.10.md)、[依赖声明](../THIRD-PARTY-NOTICES.md) |

## 未来议题

以下不属于已发布版本的前置，也没有默认排期：

- **通用 R22 / 新文字链路**：段落 bidi、自动断行与重塑形、样式/多字体最终行、分页、paint/writer/Flow 全链路。已有 run/span/字号/定位查询不等于该链路完成。保留[查询设计](architecture/r22-mixed-text-decision.md)、[共享定位消费者](architecture/r22-shared-position-consumer.md)与[剩余范围](validation/r22-remaining-scope.md)；Unicode 算法版本、字体/原生资产与许可须在具体选型时冻结。
- **RTL、Arabic、Indic 与复杂字体**：只按各有界 profile 验收，默认拒绝范围不放宽。既有 reader 反例、字符映射方法失败及独立几何/语义门继续保留，见[reader 方法验证](validation/r22-reader-replacement-candidate.md)和[双词检查点](validation/r22-two-word-export-validation.md)。
- **完整 R02、R20、扩展字体和部署组合**：非 BMP 索引、真正越段映射、自然外来互操作、全部归档条款、任意 CFF/变量/彩色字体、musl/AOT/trimming 等均不能由局部通过推定。新范围须有明确消费者、固定反例、最小契约和独立验收。

生产 SES/SM2 验章、可信证书链、合格 TSA 与法律效力继续由调用方接入合规提供方。核心 `SupportsBuiltInSesSm2Verification=false` 与 `SupportsGmT0099EncryptionEnvelope=false` 不因可选包而提升；`FullyValid` 仍要求显式注册的签名值验证器实际成功。

## 开工与验收

新增工作先核对准确源码与已有实现，只为确定缺陷或明确新范围建立任务；功能测试、包消费、自动渲染、PNG、实际 Preview、Copy/Find 和公开发布分别判定。涉及页面变更遵守根 `AGENTS.md` 与[发布门禁](release-engineering-gates.md)，不能靠旧页面、删样例或改阈值完成验收。

旧批次、工时、代理分工和已完成计划已从当前树移除；历史追溯见[文档索引](README.md#历史资料)。
