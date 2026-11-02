# 文档索引

安装与快速示例见[仓库 README](../README.md)，循序使用见[教程](tutorials/README.md)。当前公开预览版为 `0.1.0-preview.10`；下表中的实现契约不构成任意输入或标准认证承诺。

## 按用途阅读

| 用途 | 入口 |
| --- | --- |
| 安装、发布范围与证据 | [发布说明](release-notes-0.1.0-preview.10.md)、[最终发布记录](validation/preview10-publication-2026-11-02.md)、[受限中英文字体配置](release-profile-cn-en.md) |
| 能力边界与后续范围 | [功能对照](feature-parity.md)、[分出口契约](capability-contract-v12.md)、[现行路线图](capability-roadmap.md) |
| DOCX/OFD/PDF 与资源预算 | [转换契约](conversion-contracts.md)、[字体部署](tutorials/cn-en-font-deployment.md)、[字体子集与复用](font-subset-and-reuse.md) |
| 内容生成与操作 | [Flow](flow-layout.md)、[表格](public-tables.md)、[绘图](tutorials/16-native-graphics.md)、[文档工具](tutorials/15-document-tools.md)、[类型化对象](standard-object-editing.md)、[多 DocBody](multi-docbody.md) |
| 其他入口与输出 | [纯文本](plain-text-ofd.md)、[HTML](html-preview.md)、[图片](tutorials/15-image-conversion.md)、[PDF 矢量](tutorials/18-pdf-vector-mode.md)、[Skia](tutorials/17-skia-event-adapter.md) |
| 文字与字体查询 | [预排版](text-prelayout.md)、[typed 输出](text-shaped-output.md)、[混排](text-mixed-paragraph.md)、[各字体出口](validation/r08-four-outlet-ledger.md) |
| 签名、加密与归档 | [签名范围](signature-scope-v12.md)、[加密](password-crypto.md)、[SES](tutorials/19-ses-v5-35275-interop.md)、[归档预检](archival-preflight.md)、[标准覆盖](archival-standard-coverage.md) |
| 维护、测试与发布门 | [CI](ci-performance.md)、[仓库存储](repository-storage.md)、[发布门禁](release-engineering-gates.md)、[Preview 模板](preview-acceptance-template.md) |

## 验收资料

`validation/` 保留实际验证、失败与方法边界；`evidence/` 保留必要输入/输出、视觉记录、许可和带清单的 LFS 归档。每份报告只证明其具名源码、输入、配置和实际页范围；日期或包数不同的旧报告不能代表当前发布。图像、manifest、已签收 JSON 与原始哈希不会随文档整理重新生成或改写。

当前发布状态以最终发布记录为准。[早期中英文验收](validation/cn-en-release-local-validation.md)的 Regular-only 视觉 FAIL、Find 异常及 Copy 未完成保留；[N01](validation/cn-en-release-n01.md)保存 C1/C2 的有限提案。Q05 JSON 内的旧文件路径、源码清单中的历史路径均相对于它们记录的原提交解释。

## 历史资料

已完成的实施计划、旧调度状态、重复设计探针及发布前预检从当前树移除，通过固定公开产品提交 `8f27f92b51b70a9b01082da35ff1cfe989d46123` 追溯，避免继续充当待办入口：

- [2026-10 审查与实施计划](https://github.com/whynpc9/ofdrw.net/tree/8f27f92b51b70a9b01082da35ff1cfe989d46123/plans)、[早期 ticket 草稿](https://github.com/whynpc9/ofdrw.net/tree/8f27f92b51b70a9b01082da35ff1cfe989d46123/.scratch/capability-roadmap/issues)。
- [原能力开发计划](https://github.com/whynpc9/ofdrw.net/blob/8f27f92b51b70a9b01082da35ff1cfe989d46123/docs/capability-roadmap-development-plan.md)、[2026-11-01 后续计划](https://github.com/whynpc9/ofdrw.net/blob/8f27f92b51b70a9b01082da35ff1cfe989d46123/docs/next-steps-2026-11-01.md)、[preview.10 原候选计划](https://github.com/whynpc9/ofdrw.net/blob/8f27f92b51b70a9b01082da35ff1cfe989d46123/docs/next-steps-2026-11-02.md)。
- [PR23 最终验收与 LFS 恢复入口](evidence/pr23-integration/README.md)，以及各 `evidence/` 中原有验收与失败记录。

这些历史计划不授权新工作，里面的 open/pending/未发布等状态仅属于原时间点。历史发行说明继续保留，发布标签和已发布包不变；新运行产物留在 ignored `artifacts/`。
