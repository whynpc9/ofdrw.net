# PR #23 / 027：持久化 Split 保留附件来源

本次范围完成，可供 root 审阅；PR 统一集成、全套/docs/包与 CI/review closure 仍由 root 执行。

- Cursor P2：[完整评论](https://github.com/whynpc9/ofdrw.net/pull/23#discussion_r4291090880)，review `5543269374`、thread `PRRT_kwDORQ4ZKs6tEgOy`、comment `4291090880`。
- 冻结起点：`26dc12087ddff71af31c1b46d5eb216723a36fad`。
- 实现及测试提交：**`2acfaed0503b2d59f2a5ebc1161d804a4cf2fbf0`**。
- 独立受管工作区：`/Users/wanghongyi/.codex/worktrees/audit-027-pr23-split/ofdrw.net`。
- 分支：`codex/pr23-027-split-attachments`。
- 原027分支仍为 `470de1ebb98b2f9cf1448e0dfbc1763e527b3913`；未改原交付、旧证据、原027.json ready来源或共享计划索引。

## 真实复现

先在当前PR基线只新增三个公开API回归，使用 Writer → Reader → 修改附件Name/Data → Split第二页 → Writer → Reader/ArchivalPreflight。三种文档路径 `Doc_0/`、包根、`Nested/Doc/` 均实际命中且失败于 `OrphanAttachmentIndex`；不是测试零命中、静态猜测或编译失败。`tests/baseline.trx` 为 **3 failed / 0 passed / 0 skipped**，日志为 `logs/baseline-reproduction.log`。最初测试工具代码把报告成员误写为Results导致的编译错误只保存在 `logs/baseline.log`，修正为真实RuleResults后才得到上述缺陷证据。

产物在本工作区 `artifacts/pr23-027-split/baseline/`：三个原始输入、错误输出、实际ZIP清单和预检JSON，均在断言前保存。标准目录输入8个条目，Split错误输出10个条目，附件目录从3个条目变为6个：

```text
Doc_0/Attachs/Attachments.xml           # 旧列表仍存在，旧payload仍可见
Doc_0/Attachs/Attachments.xml.1         # Document.xml 指向这里
Doc_0/Attachs/Attach_1_original.bin     # 原始00 01 02 FF，没有被更新
Doc_0/Attachs/Attach_1_renamed.bin      # 编辑后的09 08 07 00 FF
Doc_0/Attachs/Attach_2_second.bin
Doc_0/Attachs/Attach_2_second.bin.1
```

预检真实报告旧 `Doc_0/Attachs/Attachments.xml` 为 OrphanAttachmentIndex。原因是Splitter虽然复制了原ZIP，却经CopyAttachments丢弃AttachmentSource及附件来源身份，Writer无法证明原列表/载荷归属，只能Fresh避碰。

## 修复取舍

只修改两个生产文件：

1. `OfdDocumentMerger.CopyAttachments` 新增内部可选 `retainedSourceMapping`。默认Merge/Mix行为保持重建闭合typed附件、不继承其他文档Owner。仅提供该映射的调用保留附件ID，按**原对象 → 实际克隆对象**注册映射，再调用现有 `OfdAttachmentSource.Clone`。新克隆对象、payload和映射及来源克隆沿用materialization计费与取消检查；重复所有者拒绝。
2. 持久化 `OfdDocumentSplitter` 的保留路径分支传入已经使用的真实对象映射，克隆附件来源；要求原/目标DocumentEntryPath一致。Reader捕获的描述符路径、payload路径、属性基线、digest、槽位与tombstone随现有克隆机制保留。来源库存、source guard与受控克隆继续使用既有字段表，无新绕过状态。

既有 `AssertCopyable` 先执行，包含未知根/附件/FileLoc/Document引用包络的来源仍拒绝Split/Mix/Merge；未放宽opaque迁移。允许的是现有闭合普通描述符的保留路径Split，包括原本可证明闭合的非canonical路径。没有根据Name/ID查找来源节点、推断引用或做跨文档重映射。

**没有修改 Writer Fresh、pruner 或预检。没有删除任何ZIP条目。** 因此未证明归属的conventional占位条目不被强占或删除。新附件遇到无关payload占位仍会合法选择后缀；这与正常Split错误复制出第二套活动列表的缺陷明确区分。

## 修复后实际ZIP与预检

同一组三个失败断言均通过；标准目录Split输出7个条目，附件仍只占原有3个路径：

```text
Doc_0/Attachs/Attachments.xml
Doc_0/Attachs/Attach_1_original.bin     # 已更新为09 08 07 00 FF
Doc_0/Attachs/Attach_2_second.bin
```

Document仍指向 `Attachs/Attachments.xml`。实际Reader回读Name为renamed.bin、Data为编辑后的5字节，附件ID保留；原始输入payload不变。预检在活动descriptor上记录两个LocalPayloadPresent/Pass，且零OrphanAttachmentIndex。再次Split/Writer/Reader不增长新后缀。根目录及嵌套目录均验证了相同结果。

`fixed/` 保存最新成功产物。`manifest.json` 保存每个OFD、ZIP清单、预检JSON的字节数/SHA-256及每个附件条目的解压内容hash；`tested-source.patch` 与提交 `26dc120..2acfaed` 的binary diff逐字相等，该断言已实际执行。

| 标准目录输出 | 字节 | SHA-256 |
| --- | ---: | --- |
| baseline/canonical-output.ofd | 2337 | `8dc69c602df32a63cf66587f402dbc2d7285563be40343155f86f44d7ffc27a9` |
| fixed/canonical-output.ofd | 1741 | `03103f0f682311546926d4792a7aa79b5ae9521cf77dc161275a08ffb4f510ec` |

体积减少596字节仅是本次产物观察，正确性依赖实际路径/解引用字节/来源和预检断言，不根据体积推断。

## 回归与命令

新 `Pr23AttachmentSplitTests` 共14条：3种文档路径的真实ZIP/预检/二次Split；单删/清空；Split后受控克隆及prepared write；4种opaque包络拒绝且同文档no-op仍保留未知属性；闭合的Metadata/List.xml原路径保存；来源不明conventional占位不覆盖；添加fresh附件遇到无关payload占位不覆盖；真实loopback监听端点在读/拆/写/回读/预检全链零连接、外部载荷不内嵌。

删除边界准确保留：单删后的活动列表只含当前附件，不复活已删槽位，旧非活动payload保留；Clear移除活动Document引用，仍保留原孤立列表/载荷，预检仍明确报告其OrphanAttachmentIndex。此既有逻辑删除契约不是物理擦除，修复没有隐藏该告警。来源不明的旧conventional占位控制也保留其真实预检告警，未宣称任意包均健康。

环境：SDK10.0.401、runtime10.0.12、macOS arm64，完整记录 `dotnet-info.log`。仓库离线可信字体准备及hash在 `fonts.json`。所有dotnet先设置可写任务DOTNET_CLI_HOME、禁firsttime/telemetry，使用 `--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false`，未清全局NuGet。

```bash
# 从独立worktree根执行；设置上述环境/flags，产物环境变量指向baseline或fixed。
export OFDRW_PR23_SPLIT_ARTIFACTS="$PWD/artifacts/pr23-027-split/fixed"
dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --no-restore --filter 'FullyQualifiedName~Pr23AttachmentSplitTests|FullyQualifiedName~Audit027RegressionTests|FullyQualifiedName~ArchivalPreflightTests' --logger 'trx;LogFileName=boundary-tests.trx' --results-directory artifacts/pr23-027-split/tests
dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --no-restore --logger 'trx;LogFileName=packaging.trx' --results-directory artifacts/pr23-027-split/tests
dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --logger 'trx;LogFileName=layout.trx' --results-directory artifacts/pr23-027-split/tests
git diff --check
```

| 验证 | 实际结果 | artifacts/pr23-027-split/证据 |
| --- | --- | --- |
| 修复前真实基线 | 3失败，0跳过 | tests/baseline.trx、baseline/、logs/baseline-reproduction.log |
| 三项原失败断言 | 3/3通过 | tests/fixed-first.trx |
| 联合边界 | **133/133，0失败/跳过**（新14 + 原027的38 + 预检81） | tests/boundary-tests.trx |
| 完整Packaging | **1260/1260，0失败/跳过**，新14实际均Passed | tests/packaging.trx、logs/packaging.log |
| 完整Layout | **129/129，0失败/跳过** | tests/layout.trx、logs/layout.log |
| git diff --check | exit0 | 提交前检查 |
| 原027分支 | 实测仍为470de1ebb98b2f9cf1448e0dfbc1763e527b3913 | acceptance-summary.json |

最终Packaging运行已修正新增测试的Assert.Single分析器建议，保留的告警来自已有测试/Reader。未运行solution全套、docs Rebuild或pack，不占共享validation.lock；这些整体验收按root指令留给合并后的统一门禁。没有并行同worktree dotnet。纯元数据/来源修复无新增视觉验收，未获取GUI锁、操作GUI或宣称视觉通过。

## 剩余门与权限边界

本地修复、专项、Packaging/Layout和实际ZIP/预检闭环均完成。剩余为root完整diff审阅、统一集成后的全套/docs/包/CI及当前head review closure；本交付不替代那些门。未push、发PR评论/resolve thread、合并PR/main、tag或publish。只有本独立分支的本地提交及本报告；原027 ready状态和共享索引保持原样。
