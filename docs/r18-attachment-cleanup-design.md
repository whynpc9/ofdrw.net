# R18 附件物理清理设计检查点

源码基线：`070594e414fcb43b55a11a7927799afe797a3cfe`（2026-10-21）。R00 能力契约已合入。本文件描述本票可验收的窄边界；最终 API 和结果以实现及实测为准。

## 公共语义

- `OfdDocumentPackage.Attachments.Remove/RemoveAt/Clear` 继续只移除活动附件引用；原 ZIP 条目可能保留。
- 新增显式 `OfdDocumentPackage.RemoveAttachment(attachment, mode)`，`Logical` 与现有 List 移除等价；`Physical` 仅接纳该包 Reader 来源的内部附件，登记来源槽位并从活动列表移除。外部定位不删除，物理模式拒绝。
- 物理模式的成功含义是本次请求的原载荷路径确实从输出 ZIP 消失；若最后一个附件被移除，原描述符也被清除。共享路径、仍被引用或未知语义时整个写入拒绝，不能静默降为逻辑移除。

## 证明与写入顺序

Reader 捕获的 `OfdAttachmentSource` 是唯一来源证据：owner、描述符摘要、原对象身份/槽位、原 `FileLoc` 解出的实际路径。ID、Name 和 `Attach_*` 文件名均不是删除权限。保存前重验来源、待删除槽位未复活、路径仍在原包中、内部/外部属性未改；保存时从最终构造的条目和签名清理结果核对所有标准引用。以真实文件路径检查全部 XML 后缀和无后缀的 XML。对无法解析的疑似 XML、扩展引用、不明确的嵌套 `BaseLoc`、别名与多 DocBody 共享，物理清理拒绝。明确处理 Res 的 `BaseLoc` 与 `FontFile`/`MediaFile`、PageRes 的隐式位置、签名和批注的 `BaseLoc`/`FileLoc`；不能用普通字符串包含或归档预检无告警代替闭包。

本轮将 XML 形态的附件载荷判为不确定角色并拒绝物理删除；二进制附件仍须通过完整引用闭包检查。此限制避免附件路径同时充当不显式声明的 XML 控制条目。

删除在 `WriteWithResultAsync` 创建 ZIP/触碰目标流前完成；不改调用方 `PreservedEntries`。`OfdPackageWriteResult.RemovedEntries` 报真实删除路径，签名失效沿用既有 `SignaturesInvalidated` 与重签诊断。受控写入克隆及同路径 Split 要么保留来源请求并重验，要么在操作前明确拒绝，不能漏掉请求后写出貌似成功的包。合并/混合不继承原包删除权限。

## 首轮固定验收

复用 `Audit027RegressionTests` 和 `Pr23AttachmentSplitTests` 证明默认逻辑行为；针对物理模式增加同源独占正例、非规范路径、同名新对象、共享载荷、其它 XML/资源/批注/签名/跨 DocBody 引用、未知及损坏 XML、全删描述符、零字节失败输出、签名失效、再次读写和 Split 来源。对成功样本比较 ZIP 清单、保留条目字节、payload SHA-256 及 preflight。附件元数据票不改变页面呈现，视觉门记录为不适用；若实现触及页面/渲染则另按 `AGENTS.md` 完成逐页检查。
