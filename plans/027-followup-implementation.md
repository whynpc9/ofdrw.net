# Plan 027 后续实施：附件来源包络保留

状态：COMPLETE（最小生产修复及本次功能验证完成，可供父协调者审阅集成）；本报告由任务027独立维护，不修改协调索引。

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

产物均在独立worktree的 `artifacts/followup-027/`，不提交二进制。


## 最终源码与命令

执行日期：2026-10-16（Asia/Shanghai）。环境实测 SDK **10.0.401**、runtime **10.0.12**、macOS arm64；完整 `dotnet --info` 见 `logs/environment.log`，源及字体环境见 `environment.json`。生产实现提交 `2ac46f139867e81edf92af9b58d3c1031a8d2cf5`，清除已被新附件Writer取代的旧私有helper后，验证冻结源码为 **`9a165122127b765e02b12699473c88ba80e28f71`**。本报告后续提交只有验收记录，不改变该已验证源码。

`artifacts/followup-027/validate.py` 通过原子mkdir取得共享validation.lock，记录任务/线程/UTC时间/唯一token，在finally核对自己token后释放；本次完成日志为 `RELEASED validation.lock`。取得前仅做专项测试和静态核查；持锁后无并行dotnet，docs Rebuild与其他build/test串行。未取得或操作GUI租约。

所有dotnet运行先设置可写 `DOTNET_CLI_HOME`、`DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1`、`DOTNET_CLI_TELEMETRY_OPTOUT=1`，build/test/restore使用 `--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false`，未清NuGet缓存。`fonts.json` 来自仓库固定hash的离线可信字体包，不依赖可跳过的系统字体发现。完整实际命令、源码SHA、起止时间和exit code见 `validation-commands.json`。

```bash
python3 scripts/prepare-trusted-fonts.py --output-dir artifacts/followup-027/trusted-fonts
# 将输出environment加入子进程环境；下列dotnet均加上述flags。
dotnet restore Ofdrw.Net.sln
dotnet build Ofdrw.Net.sln -c Release --no-restore
dotnet test Ofdrw.Net.sln -c Release --no-build --no-restore --logger trx --results-directory artifacts/followup-027/tests/full
python3 scripts/run-trusted-font-tests.py --font-output-dir artifacts/followup-027/trusted-fonts --results-directory artifacts/followup-027/tests/trusted-final
python3 -m unittest discover -s scripts/tests -v
python3 scripts/check-public-api-docs.py
bash scripts/run-converter-package-e2e.sh 0.1.0-preview.followup027 --packages-dir artifacts/followup-027/nuget --output-dir artifacts/followup-027/package-output
git diff --check
```

## 本次验收结果

| 项目 | 实际结果 | 证据（相对 artifacts/followup-027/） |
| --- | --- | --- |
| 原基线复现 | 实际2条命中，2失败，0跳过；canonical未知根属性丢失与nested活动路径改写 | tests/baseline.trx、logs/baseline.log |
| 027最终专项 | 38/38，零失败/跳过；全套TRX内也确认这38条全部Passed | tests/targeted-final-2.trx、acceptance-summary.json |
| Solution restore | exit 0 | logs/solution-restore.log |
| Release build | exit 0，0 errors / 85 warnings；含既有ImageSharp advisory及已有代码/分析器告警，依赖未变 | logs/solution-build.log |
| 全套 | **2850/2850，11项目，零失败/跳过** | tests/full/、full-summary.json、logs/solution-tests.log |
| 可信字体强制runner | **65/65，零失败/跳过**，离线固定字体准备成功 | tests/trusted-final/trusted-fonts.trx、summary.json、fonts.json |
| Python | **20/20** | logs/python-tests.log |
| API docs Rebuild | exit 0，272 historical symbols / 1 cleared；baseline未更新 | logs/public-api-docs.log |
| 本地全包消费E2E | exit 0，**11包**，版本0.1.0-preview.followup027；验证fresh消费包bytes及输出 | nuget/package-manifest.json、package-output/、logs/local-package-e2e.log |
| git diff --check | exit 0 | 最终提交前再次检查 |
| 自动页面渲染 / PNG目视 / macOS Preview | 本修复未修改页面，不适用新的视觉验收；包消费中的自动渲染比较不宣称人工视觉验收 | environment.json、acceptance-summary.json |

全套项目实际计数：Core85、Packaging1165（包含Reader与来源库存回归、新38条）、Pdf514、Signatures44、Docx278、CLI176、Layout94、SES37、Skia44、Pdf.Vector315、Crypto.Password98。没有借用父分支旧2812/独立516或旧调查观察项充当新修复证据。

38条新增实际回归覆盖：canonical/root文档与嵌套原路径no-op、根/附件/FileLoc/Document引用双命名空间与未知节点顺序、9种属性单字段编辑、单个/全部逻辑删除、同ID/Name新对象不借用被删来源、fresh附件、loopback外部端点零连接且不内嵌、受控克隆保留/单删/清空、原/候选来源/marker/槽位/基线的8组防变更、原始来源tamper、外来所有者、路径及内外部迁移拒绝、Split/Mix/Merge opaque拒绝与plain支持、opaque锚点重排拒绝、可变重复ID仍按对象槽位保存、存储/读取预算、opaque-only活动列表、外部定位纯文本编辑、共享payload冲突拒绝、重复Document引用与混合FileLoc编辑拒绝。所有断言基于本次公开读写产物或真实受控写入，过滤器均非零命中。

## 产物与哈希

`manifest.json` 记录本次生成的input/output和活动descriptor字节数/SHA-256；`acceptance-summary.json` 汇总本次门禁。两个样例的原始/输出payload均为282字节，SHA-256 **`d328999289e52293a68051fba660a552e1853d5a73a72735f6c5c4aa6444c9c3`**，并非压缩流哈希。

| 本次包 | 字节 | SHA-256 |
| --- | ---: | --- |
| packages/canonical-input.ofd | 1644 | `1ca130ea8b45e2dc17e0e33c3e17975a5276c8e6658277ba28a17d01dd66adc1` |
| packages/canonical-output.ofd | 2058 | `e2a2d2d14577ca1f98effe7bd7100be2dacf8e630b5f6856dcd4c814a12ee4d4` |
| packages/nested-input.ofd | 1732 | `2c26e5bec5a4ab4f7146442fbe5cb570700d0169cc434e88278e411b666a3cdc` |
| packages/nested-output.ofd | 2170 | `f7a638eff2df27c3cb60685916c60fbbdfb596fb8f26edd8004206b5078ea1b9` |

活动描述符两个样例均925字节，SHA-256 `8416176de0a011c356419a10e06055141d6b3605098a04d40afae5a51bb44e8f`。canonical增加414字节、nested增加438字节；Writer已有的控制/资源XML生成、保留命名空间及压缩会影响整包体积。此处根据语义和payload断言确认保存结果，未把体积增加当作保留正确性的证据。ZIP时间戳随运行变化，整包hash只标识本次产物。

## 交付边界

本次最小元数据修复与要求的功能门均完成，无未完成必需门。物理删除、跨文档不透明引用迁移、签名重写、规范MediaType优先级保持范围外；因无法证明安全而拒绝的操作均按文档和回归保留，不伪装成已支持。未修改共享plans/README、execution-state/status或父/main checkout；未推送、创建PR、合并目标分支、tag或publish。由父协调会话在全部任务完成后审阅并统一集成。
