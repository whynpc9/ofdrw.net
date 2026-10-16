# Plan 027 调查结果：附件描述符无操作往返

- **状态：COMPLETE（调查完成，未实施修复）**。
- **I01 结论：CONFIRMED**，限定为公开 `OfdReader.ReadAsync → OfdPackageWriter.WriteAsync` 路径静默丢失当前活动附件描述符中的未知 XML 元数据。最小复现保持描述符和 FileLoc 路径不变，排除了仅由路径规范化造成差异的解释。
- 源码基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`；执行日期：2026-10-15。
- 工作区：`/Users/wanghongyi/.codex/worktrees/ofdrw-plan-027/ofdrw.net`；分支：`codex/plan-027-attachment-preservation-investigation`。
- 唯一跟踪交付为本报告。探针、输入/输出、逐字段差异与日志均在该工作区的 `artifacts/plan-027/`，由 `.gitignore` 排除。协调者维护索引；本轮未改生产代码。

## 契约边界

本调查没有把 CustomTags 的 F08 结论套用到附件。直接构造 ZIP/XML，经公开 API 读写，再沿输出 `Document.xml/Attachments` 引用找到**活动描述符**，比较 XML 语义和解引用后的原始载荷字节。

[开发计划第 2 节](../docs/capability-roadmap-development-plan.md)要求未知 XML 不得因强类型化而静默丢失。`OfdDocumentPackage.PreservedEntries` 的 XML 文档仅说明先复制条目、再替换重建条目，它本身不保证重建描述符内部所有未知字段。这里确认附件往返与上述保留目标存在可复现缺口，不宣称所有公开成员均已承诺任意扩展可编辑，也不作 OFD schema/外部阅读器兼容性判断。

[附件教程](../docs/tutorials/09-templates-attachments-tags.md)明确说明 Writer 生成 `Doc_0/Attachs/Attachments.xml` 和 `Attach_N_…` 载荷路径，因此**原始路径必须逐字稳定**的契约仍为 UNDECIDABLE。路径变化单独记录，不作为确认未知元数据缺陷的依据。无残留物理删除、任意未知外部引用迁移也不从本调查推导。

## 最小复现

`minimal-input.ofd` 只有五个 ZIP 条目：

```text
OFD.xml
Doc_0/Document.xml
Doc_0/Pages/Page_0/Content.xml
Doc_0/Attachs/Attachments.xml
Doc_0/Attachs/Attach_1_original.bin
```

包含一个空白页；`Document.xml` 的附件引用是 `Attachs/Attachments.xml`。附件 XML 输入如下：

```xml
<Attachments xmlns="http://www.ofdspec.org/2016">
  <Attachment ID="7" Name="original.bin" Format="application/octet-stream" VendorAttr="keep">
    <FileLoc>Attach_1_original.bin</FileLoc>
  </Attachment>
</Attachments>
```

公开调用只执行：

```csharp
OfdDocumentPackage model;
await using (var source = File.OpenRead(inputPath))
    model = await new OfdReader().ReadAsync(source);
await using (var destination = File.Create(outputPath))
    await new OfdPackageWriter().WriteAsync(model, destination);
```

活动输出为：

```xml
<ofd:Attachments xmlns:ofd="http://www.ofdspec.org/2016">
  <ofd:Attachment ID="7" Name="original.bin" Format="application/octet-stream">
    <ofd:FileLoc>Attach_1_original.bin</ofd:FileLoc>
  </ofd:Attachment>
</ofd:Attachments>
```

`VendorAttr` 为 **LOSS**；活动描述符路径、FileLoc 和载荷为 **MATCH**。前缀、缩进不是失败依据。载荷为 282 字节：`00..FF` 接 UTF-8 `\0附件 original payload\r\n`；SHA-256 为 `d328999289e52293a68051fba660a552e1853d5a73a72735f6c5c4aa6444c9c3`。完整构造代码在 `artifacts/plan-027/probe/Program.cs`；最小输入/输出包及解出的 XML 均已保存。

## 案例和结果

完整差异在 `artifacts/plan-027/results.json` 和 `results.md`：**7 个 case、134 个观察项：99 MATCH / 30 LOSS / 5 UNDECIDABLE**。探针退出 0 仅表示调查执行完整，不表示元数据保留通过。30 个 LOSS 行包含 23 个未知元数据差异、3 个 MediaType 模型变化和 4 个路径变化；后两类不混作未知 XML 缺陷。

比较使用展开的 QName（URI + local name），忽略前缀、属性顺序和 xmlns 声明形式。未知子节点比较 QName、属性、文本和顺序；日期比较 UTC 值，布尔/尺寸比较模型值。载荷比较解压后的字节及 SHA-256，不比较 ZIP 压缩流。

| 场景 / 复跑 case | 预期 | 实际 | 结论 |
|---|---|---|---|
| OpaqueDescriptorNoOp / `minimal-noop` | 原路径 no-op 保存保留未知属性 | VendorAttr=keep 消失；活动路径、FileLoc、payload 不变 | **LOSS** 元数据，其他 4 项 MATCH |
| OpaqueDescriptorNoOp / `canonical-noop` | 保留根、Attachment、FileLoc 和引用节点扩展 | 根普通属性、两个命名空间同名属性、未知前后子节点；附件普通/命名空间属性、未知前后子节点；FileLoc 扩展属性；Document 的 Attachments 引用属性均消失 | **LOSS** |
| 多命名空间 / 上述 case | 两套 URI 信息独立保留 | `{urn:plan027:vendor-one}`、`{urn:plan027:vendor-two}` 信息均消失，不只是前缀改变；含 w:Binary QName 字面量的未知子节点整段消失 | **LOSS** |
| 原描述符 / `relocated-noop` | 检查 Metadata/List.xml 是否仍是活动描述符 | 原 XML 字节仍在 ZIP；活动引用改到 Attachs/Attachments.xml，新活动列表扩展消失 | **LOSS** 活动元数据；路径稳定性契约 UNDECIDABLE |
| PayloadNoOp / 两个 noop | 当前 FileLoc 可解析，载荷字节不变 | 原始载荷和当前引用载荷均为上述 282 字节和同一 SHA-256 | **MATCH** |
| 单字段编辑 / `modeled-name-edit` | 仅改第一个附件 Name 为 changed.bin | Name 可读回；ID、Format、日期、Size、Visible、Usage、IsExternal 保留；未知信息仍丢失 | **MATCH** 编辑；**LOSS** 未知信息 |
| ExplicitRemoval / `remove-one` | 删除 ID=7 不被 opaque 恢复，保留 ID=8 | 活动列表、公开 Reader 再读均仅 1 个附件，无 ID=7；存留外部 ID=8 的扩展消失 | **MATCH** 删除；**LOSS** 存留节点扩展 |
| ExplicitRemoval / `remove-all` | Clear 后公开 Reader 再读为 0 | Document 移除附件引用；旧描述符和载荷仍在 ZIP，但不恢复为活动附件 | **MATCH** 逻辑删除；物理清理契约 UNDECIDABLE |
| 新建 / `fresh-add` | 无来源新附件可写入/取回 | 活动列表 1 个、Name=fresh.bin、payload 原字节匹配 | **MATCH** |
| ExternalAttachment / 所有存留外部附件的 case | 不抓取、不内嵌、定位不变 | `http://127.0.0.1:32767/never-fetch-plan027.bin` 原样写回，仍外部、Data.Length=0，没有 external.bin 载荷 | **MATCH** |
| 外部端点观察 / `ExternalAttachment` | 公开 Reader/Writer/再读期间零连接到引用端点 | 本地 TcpListener 成功监听该端口，全部调用后 Pending=False；探针没有网络客户端；Reader 源码仅查 ZIP 条目 | **MATCH** 本次引用端点；不是全机网络审计 |
| MediaType/Format 共存 / noop 与 edit | 记录实际模型值，不预设规范优先级 | 输入 MediaType=application/x-plan027、Format=application/octet-stream；Reader 优先前者，Writer 仅发 Format，再读 MediaType 变为后者 | **LOSS** 模型值；规范扩展/别名优先级契约未定案 |

`remove-one`、`remove-all` 都保留原内部 payload 条目；逻辑删除不表示物理擦除。`relocated-noop` 的旧描述符保留也不代表活动附件元数据保留。探针不执行附件内容。

## 来源和写入所有权

源码定位均针对本次基线：

1. [Reader](../src/Ofdrw.Net.Reader/Readers/OfdReader.cs) 第 116–124 行把原 ZIP 条目放入 PreservedEntries；第 392–432 行只投影 OfdAttachment 的建模属性、外部定位和 Data。未知 XML 未存入附件级包络。
2. [OfdAttachment](../src/Ofdrw.Net.Core/Models/OfdAttachment.cs) 第 5–30 行无描述符源路径、内部原 FileLoc、根/单节点 source envelope 或未知字段；模型和 raw 条目间没有附件级来源身份。
3. [Writer](../src/Ofdrw.Net.Packaging/OfdPackageWriter.cs) 第 365 行复制 PreservedEntries；第 1181–1225 行按 Attachments 集合从零创建描述符和固定路径载荷。同路径时覆盖原列表；其他路径时新建 canonical 列表。Data 字节保留，原 payload 可继续作为孤立条目存在。
4. Writer 第 510–511 行按集合数量生成/移除 Document 附件引用；[Document envelope writer](../src/Ofdrw.Net.Packaging/OfdDocumentEnvelopeWriter.cs) 第 67–68 行整体替换这个 singleton，引用自己的未知属性也丢失。
5. 当前集合控制逻辑删除，raw 文件不会在本轮公开 Reader 中复活。未来 preservation 不得无条件补回原列表节点，否则会破坏此删除语义。

## 后续最小修复建议（未实施）

建议单独立项，首个边界是“同一文档公开 Reader→Writer 保存未知附件信息”：

- `src/Ofdrw.Net.Core/Models/OfdAttachment.cs`、`OfdDocumentPackage.cs`：内部记录列表来源、根扩展、单附件与 FileLoc 来源包络/路径/身份和基线；typed 只拥有已知槽位，新附件没有来源并正常生成。
- `src/Ofdrw.Net.Reader/Readers/OfdReader.cs`：捕获活动描述符、原定位与来源身份，不发起外部读取，不用可变 Name/ID 猜测原节点。
- `src/Ofdrw.Net.Packaging/OfdPackageWriter.cs`、`OfdDocumentEnvelopeWriter.cs`：在来源 XML 上只更新拥有字段；保留根/节点/FileLoc/Document 引用扩展，按当前集合实施编辑、新增、删除；不恢复显式删除节点。
- 同步检查 `src/Ofdrw.Net.Core/Models/OfdSerializedInventory.cs`、`OfdOwnedEditCandidate.cs`、`OfdModelMaterialization.cs`、`OfdModelTraversal.cs` 的快照、克隆、限额/防变更机制。来源状态必须进入现有库存和预算，不能只加 SourceXml 绕开这些规则。
- 在 `tests/Ofdrw.Net.Packaging.Tests/` 加入最小 no-op、双命名空间包络、单字段编辑、单个/全部删除、fresh、外部不抓取及 payload 字节回归；保留公开/受控克隆路径的已有测试。

如必须迁移路径才能保存，需要先证明未知引用闭包；不得猜测重映射。首个修复可保留导入路径，或对无法证明安全的编辑明确拒绝。全面物理清理、Mix/Split/Merge 跨文档映射、规范 MediaType 优先级、签名策略和发布不纳入本轮最小修复。`OfdDocumentMerger.CopyAttachments`、Splitter/Mixer 也是来源复制调用点；新增状态如影响它们，应先确定保留/拒绝边界，再独立扩计划，不能把原文档身份直接套到另一文档。

## 执行命令和验收

环境：macOS 26.6.2 arm64、SDK **10.0.401**、.NET runtime **10.0.12**；global.json 是 10.0.100 + latestFeature。未清 NuGet 缓存。以下命令均在该工作区实际执行；`artifacts/plan-027/reproduce.sh` 汇总这些命令用于复跑，脚本本身未再次全量运行：

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-027/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-027/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
python3 scripts/check-public-api-docs.py

dotnet restore artifacts/plan-027/probe/Probe.csproj --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build artifacts/plan-027/probe/Probe.csproj -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet run --project artifacts/plan-027/probe/Probe.csproj -c Release --no-build --no-restore

git diff --check
git status --short
```

| 验证 | 实际结果 / artifacts 下日志 |
|---|---|
| 计划要求的基线 diff、status、工作树/暂存 diff | 均退出 0；开始时干净、报告不存在、无他人改动 |
| solution restore | 退出 0；logs/solution-restore.log |
| solution Release build | 退出 0，0 errors / 83 warnings；logs/solution-build.log；含既有 ImageSharp advisory/代码告警，未改依赖 |
| Packaging baseline tests | 退出 0；**1073 passed / 0 failed / 0 skipped / total 1073**；logs/packaging-tests.log；tests/_wanghongyideMacBook-Pro_2026-10-15_23_13_35_net10.0.trx |
| public API docs | 退出 0；272 historical symbols / 1 cleared；logs/public-api-docs.log；未更新 baseline |
| probe restore / build | 均退出 0；最终 build 0 warnings / 0 errors；logs/probe-restore.log、logs/probe-build.log |
| probe run | 退出 0；7 cases / 99 MATCH / 30 LOSS / 5 UNDECIDABLE；logs/probe-run.log、results.json、results.md |
| git diff --check | 退出 0；唯一跟踪交付为允许的 result report |
| 功能调查 | **完成**，确认元数据缺口，修复未实施 |
| 自动页面渲染 / PNG 复查 / macOS Preview | **不适用本计划**：仅调查 ZIP/XML 和 payload，没有修改/评价页面或转换渲染，不声称视觉验收通过 |

环境记录在 environment.json；文件尺寸/哈希总表在 manifest.json。输入构造和 ZIP 输入时间戳固定；Writer ZIP 时间戳可能随复跑改变，输出整包哈希仅标识本次产物。复跑依据 XML 语义和 payload 字节。

| 包名（artifacts/plan-027/） | 字节 | SHA-256 |
|---|---:|---|
| minimal-input.ofd | 1716 | `84852b30288c5c17b1750883314b381640621e29e871184e70d6559e7b9eaf1d` |
| minimal-noop-output.ofd | 1886 | `11366890c9b297e66a8789e02c9231cb2f1b8f24b0a110e3c2ba334669a50ba2` |
| canonical-input.ofd | 2556 | `6678688524cb64bfd29411452703622b71577a16256e7cab929f79e248742e0a` |
| canonical-noop-output.ofd | 2440 | `dd30985230a008a475f600b8e4f167f69ecaa92f8a1660359e85d8bcbc871292` |
| relocated-input.ofd | 2540 | `ef5e8346a230ab2be927488b377edc8ebf9c130e728cbcfea22b10f5aca98eb0` |
| relocated-noop-output.ofd | 3034 | `48a91ae6e6883e32136a60c27a0cbbecc9ec16b26586c23dac3a5e2df891234e` |

canonical 输出减少 116 字节，relocated 增加 494 字节；XML 生成、压缩、保留孤立原条目都会影响体积，不能据此判断语义完整性。

## 完成边界

全部计划命令和样例完成；无需 STOP。未改生产/测试/依赖、父工作区或索引，未推送/创建 PR/合并/发布。五个 UNDECIDABLE 为三个“原路径必须稳定”契约项与两个“删除后必须物理清理”契约项，仍保留未知。调查完成不表示缺陷已修复；索引应登记 **调查完成 / CONFIRMED / 修复未实施**。
