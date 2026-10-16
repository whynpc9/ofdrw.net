# Plan 027 后续实施：附件来源包络保留

状态：IMPLEMENTED / 验证进行中；本报告由任务027独立维护，不修改协调索引。

- 冻结起点：`codex/audit-integration@3247c76f95c9a3312cc26144e24c4984767af26f`。
- 分支：`codex/audit-followup-027`。
- 工作区：`/Users/wanghongyi/.codex/worktrees/audit-followup-027/ofdrw.net`。
- 授权：父会话已将原调查的“不实现”限制更新为最小生产修复，允许本地提交；不推送、不创建PR、不合并、不发布。
- 旧 [调查报告](027-attachment-investigation-result.md) 保留为历史调查，不把旧7case/99MATCH/30LOSS/5UNDECIDABLE重写为本次修复成绩。

## 本次复现

在冻结起点仅添加 `Audit027RegressionTests` 的两条公开 Reader→Writer 回归后运行过滤器。两条实际命中、均失败：canonical 活动根属性丢失；嵌套 `Metadata/List.xml` 的活动引用被重定向为 canonical 路径。证据：`artifacts/followup-027/tests/baseline.trx`、`logs/baseline.log`。随后相同断言在实现上通过。测试比较展开QName、非xmlns属性、所有节点种类/顺序、继承前缀及原始payload字节。

## 实现及所有权

`OfdAttachmentSource` 是内部 Reader 所有的来源状态：保留 Document/descriptor 路径、带继承命名空间的引用包络、描述符digest、根QName、按原始元素ordinal关联的对象身份及各建模属性基线。描述符XML仍留在 `PreservedEntries`，不增加可公开修改的 SourceXml 字段。附件对象的内部 SourceOwner 防止调用者直接把导入对象借给另一文档。Name/ID不是来源键。

Writer 在原始活动描述符的槽位上修改当前集合：保留根、附件、FileLoc、Document引用的未知内容和顺序；仅更新被编辑的建模属性。原始内部payload/FileLoc和描述符路径优先保留。新附件分配未占用路径；共享payload在同一路径有冲突编辑时拒绝。外部附件不抓取，Data不写入包。原MediaType/Format分别保留，未作规范优先级判断。

删除遵循当前集合：移除的原始槽位不恢复，清除所有原附件移除活动引用，旧描述符/载荷仍可能作为孤立条目留在ZIP。本来只含未知扩展、没有建模附件的活动描述符仍可no-op保留。克隆为被移除的槽位生成仅供来源身份使用的私有tombstone，不把它们加回集合。

新增状态进入 `OfdSerializedInventory`，由同一字段表覆盖存储、输出预算、源防变更和受控克隆；捕获XML/解析/哈希/路径和状态复制沿用materialization预算与取消检查。`OfdModelTraversal` 有附件来源时纳入库存预检。受控克隆复制包络、digest、槽位和属性基线；原/候选来源或基线变更均在输出前拒绝。

## 复制边界与拒绝

核查了 ControlledEdit、Merger.CopyAttachments、Splitter与Mixer的附件复制路径。Split/Mix/Merge只沿用已证明闭合的普通描述符typed复制：无扩展属性/节点/混合内容、已知标准属性和单一纯文本FileLoc。带不透明扩展、引用包络扩展或未证明可重建别名时拒绝；不猜Name/ID映射，也不把源Owner套到目标。Mixer原有临时package直接AddRange借用附件的调用点改为经过原所有者验证的typed复制。专项全项目回归首次暴露该调用点的37条拒绝；修复后1164条全部通过，零跳过，未修改既有断言。

文档路径/namespace迁移、导入附件内外部所有权切换、不透明锚点间重排、混合FileLoc内容改写、原始来源XML背后被修改等操作明确拒绝。物理擦除、跨文档扩展重映射、签名重写和规范MediaType优先级不在范围内。

## 验证

本次源码修改只涉及ZIP/XML元数据及来源状态，没有修改页面或转换/渲染算法。自动页面渲染、PNG视觉复查及macOS Preview不适用本次元数据修复；未取得GUI锁、未操作GUI，不宣称视觉门通过。

功能测试、全套、docs和最终commit证据在验收完成后追加。产物均在独立worktree的 `artifacts/followup-027/`，不提交二进制。
