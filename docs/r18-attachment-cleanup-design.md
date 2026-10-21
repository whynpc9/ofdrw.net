# R18 附件物理清理设计检查点

源码基线：`070594e414fcb43b55a11a7927799afe797a3cfe`（2026-10-21）。R00 能力契约已合入。本文件描述本票可验收的窄边界；最终 API 和结果以实现及实测为准。

## 公共语义

- `OfdDocumentPackage.Attachments.Remove/RemoveAt/Clear` 继续只移除活动附件引用；原 ZIP 条目可能保留。
- 新增显式 `OfdDocumentPackage.RemoveAttachment(attachment, mode)`，`Logical` 与现有 List 移除等价；`Physical` 仅接纳该包 Reader 来源的内部附件，登记来源槽位并从活动列表移除。外部定位不删除，物理模式拒绝。
- 物理模式的成功含义是本次请求的原载荷路径确实从输出 ZIP 消失；若最后一个附件被移除，原描述符也被清除。共享路径、仍被引用或未知语义时整个写入拒绝，不能静默降为逻辑移除。

## 证明与写入顺序

Reader 捕获的 `OfdAttachmentSource` 是唯一来源证据：owner、描述符摘要、原对象身份/槽位、原 `FileLoc` 解出的实际路径。ID、Name 和 `Attach_*` 文件名均不是删除权限。保存前重验来源、待删除槽位未复活、路径仍在原包中、内部/外部属性未改；保存时从最终构造的条目和签名清理结果核对所有标准引用。以真实文件路径检查全部 XML 后缀和无后缀的 XML。对无法解析的疑似 XML、扩展引用、不能证明所有权的嵌套 `BaseLoc`、别名与多 DocBody 共享，物理清理拒绝。明确处理 Res 的 `BaseLoc` 与 `FontFile`/`MediaFile`、PageRes 的隐式位置、签名和批注的 `BaseLoc`/`FileLoc`；不能用普通字符串包含或归档预检无告警代替闭包。

本轮将 XML 形态的附件载荷判为不确定角色并拒绝物理删除；二进制附件仍须通过完整引用闭包检查。XML 探测覆盖带 BOM 及无 BOM 的 UTF-8、UTF-16、UTF-32 常见编码，避免将多字节 XML 当成不透明二进制。路径标量若含前后空白则拒绝物理清理，避免本轮闭包比较与会调用 `Trim()` 的签名消费者分歧；保留 XML 中的 processing instruction 也拒绝，因为本轮无法证明其引用语义。已由 Reader 来源槽位认证的保留二进制附件可以使用 `.xml` 后缀；未认证的非 XML 控制条目不能借后缀或伪造附件名获准。此限制避免附件路径同时充当不显式声明的 XML 控制条目。

封闭标准字段集合包含仓库写入器实际输出的路径绘制属性 `Cap`、`Join`、`Rule` 与裁切 `Clip/Area/Path`，Reader 接纳的裁切 `Area/@Start`、图片 `ImageObject/@Substitution` 和文本方向属性 `ReadDirection`、`CharDirection`，标准模板引用 `Template/@TemplateID`，批注的 `Creator`、`LastModDate`、`Subtype`、`Print`、`NoZoom`、`NoRotate`、`ReadOnly`、`Remark`，以及导航 `Dest`、`Goto` 和 EMR 自定义标签 `EMRTags/Tag` 的实际字段。导入签名的 `Seal/BaseLoc` 子元素按签名清理器支持的路径形式扫描。写入器自有的精确 QName `TextObject` 斜体标记和取值规范的 `ImageObject` 插值标记也可通过；其它外部命名空间属性仍拒绝。`PageAnnot` 是可包含 `Annot` 子节点的标准根，不是路径标量。这些字段仍逐一扫描其值和祖先 `BaseLoc`，因此不能藏匿被删路径；正常页面模板、批注、导航与自定义标签不再因字段遗漏而被误判为未知扩展。任何新增字段仍需验证引用语义后才能准入。

R16 标准对象已实际合入后，封闭集合补纳 Reader/Writer 支持的 `DocInfo` 元数据、`Permissions`、`VPreferences` 与 `URI`、`GotoA`、`Sound`、`Movie` 动作字段。`DocBody/DocInfo/Cover` 是精确的标准二进制引用位置，保留正常封面，同时拒绝其直接或编码形式指向被删附件。最终输出 XML 中保留的 `GotoA/@AttachID` 若指向被删附件的原始 ID，即使动作以原始 XML 形式保留，也拒绝写入，避免留下悬空动作；已由调用方删除或改写的动作不会因导入快照误判。URI 动作的导航目标不按包内路径编码限制处理，但仍参与字面别名扫描。R16 的类型化 `Versions` 修改与附件清理可同次写入；版本描述文件仍按其实际路径扫描。

只有 Document 声明的资源表及页面或 `CommonData/TemplatePage` 位置旁置的隐式 `PageRes.xml` 中、结构合法的 `Fonts/Font/FontFile`（包含标准 `Serif`、`FixedWidth` 元数据）、`MultiMedias/MultiMedia/MediaFile` 或 Reader 接纳的 `MultiMedia/@MediaFile` 才能证明其它二进制条目属于标准资源；资源节点还需具备 Reader 接纳的 ID、字体名称和首个文件定位；同时存在属性与子元素时只认 Reader 优先读取的属性，伪造在其它 XML 位置的同名元素无效。附件描述符也只认 Reader 优先读取的第一个 `FileLoc` 子元素，缺少子元素时才认 `FileLoc` 属性；未被读取的重复子元素或属性不授予二进制所有权。`External=true` 的附件从不授予包内二进制所有权，即使其 `FileLoc` 恰好可解为 ZIP 条目。模板旁置资源按 Reader、资源目录和裁剪器相同的 `TemplatePage/@BaseLoc` 相对路径发现，且模板声明必须具备 Reader 接纳的非空 `ID`。所有祖先 `BaseLoc` 都参与待删路径别名检查；授予二进制所有权时，资源文件只采用 Reader 使用的资源表根 `Res/@BaseLoc`，附件 `FileLoc` 只采用描述符条目自身路径，不采用嵌套 `BaseLoc`。路径字段和祖先 `BaseLoc` 若包含 `%HH` 编码或反斜杠即拒绝；普通正文和元数据可包含这些字面字符，但仍参与完整路径别名扫描。所有会被 Reader 当作路径文本读取的节点（包括 `DocBody/Signatures` 和 `Document/Extensions`）只接纳纯标量内容；子元素分段拼接的路径继续拒绝，避免 Reader 与清理器对路径的解释不同。按 GB/T 33190—2016 第 19.1 节及规范附录 A.1、A.12，`DocBody/Versions` 是包含 `Version/@BaseLoc` 的容器，不是路径标量；物理清理保留合法版本容器，并扫描版本描述文件中 `DocVersion/FileList/File` 的路径，拒绝伪造的叶子文本和异常子结构。

嵌套于 `Document` 任意位置的 `Attachments`、`Annotations`、`CustomTags` 和 `Extensions` 声明也按路径标量检查；子元素拆分拒绝，因为结构检查器会读取整段聚合文本。删除在 `WriteWithResultAsync` 创建 ZIP/触碰目标流前完成；不改调用方 `PreservedEntries`。`OfdPackageWriteResult.RemovedEntries` 报真实删除路径，签名失效沿用既有 `SignaturesInvalidated` 与重签诊断。受控写入克隆及同路径 Split 要么保留来源请求并重验，要么在操作前明确拒绝，不能漏掉请求后写出貌似成功的包。合并/混合不继承原包删除权限。

若既有签名被本次写入完整失效并实际移除，签名描述中的标准 `SignedInfo/References/Reference/@FileRef` 仅作摘要范围记录，不构成输出包中的资源所有权；原签名中其它路径仍逐项检查。未被实际移除的签名描述及异常编码引用不得借此例外绕过闭包检查。

## 首轮固定验收

复用 `Audit027RegressionTests` 和 `Pr23AttachmentSplitTests` 证明默认逻辑行为；针对物理模式增加同源独占正例、非规范路径、同名新对象、共享载荷、其它 XML/资源/批注/签名/跨 DocBody 引用、未知及损坏 XML、全删描述符、零字节失败输出、签名失效、再次读写和 Split 来源。对成功样本比较 ZIP 清单、保留条目字节、payload SHA-256 及 preflight。附件元数据票不改变页面呈现，视觉门记录为不适用；若实现触及页面/渲染则另按 `AGENTS.md` 完成逐页检查。
