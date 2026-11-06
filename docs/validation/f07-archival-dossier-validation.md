# F07 证据档案验证（2026-11-05）

状态：**证据追踪层已实现；F07 完整归档覆盖未完成，未认证、未发布。**

源码基线 `1fb3294a104426b004aafba18a5d114be3d2b5c0` 加本次工作区变更。具名 profile 为 `ofdrw-archival-dossier/1` / `ofdrw-archival-ledger/1`，目标范围与停止条件见[设计记录](../architecture/f07-archival-dossier.md)。本次没有归档机构的正式范围/级别、接收政策或外部核准；引用标准版本仍由目标 profile 显式登记，不能推定。

## 已执行

- Linux x64，SDK 10.0.401；Packaging 包的 netstandard2.0、netstandard2.1 均构建、打包成功；无新增依赖。
- Packaging 全套 **1695/1695**，零失败、零跳过，包括新增 **12** 个测试实例。覆盖完整账本及 17 个规则映射、源/结果字节身份和输入保留、缺失材料、摘要错配、旧来源证据、已知违规优先、累计预算三个阶段、数量预算、重复 ID、未知条款、读取中取消、profile 快照及 JSON 序列化。
- 以 `0.2.0-f07-validation.2` 打包 Core/Packaging，通过仅 PackageReference 的隔离消费者加载候选包；独立 package cache，不清空全局缓存。执行已有副本操作并收集其真实审计文件：105 项、证据完整性 Pass、完整标准 Unknown；源文件的 Permissions 违规仍在源报告中可见。
- 新增 API 的 XML 注释随 `EnforcePublicApiDocs=true` 的包构建检查，新增两份源码没有 CS1591。构建仍有历史源码告警；未将其描述为零告警。
- 文档 `npm ci`、`npm run build`、`npm run typecheck` 完成；Astro 检查为 0 errors、0 warnings、1 hint。页面生成、链接检查和 Pagefind（45 页）通过。`npm run check` 退出零但 **partial**：Nimbus 不 lint 45 个 `.md` 正文；不将它表述为全覆盖检查。未进行浏览器桌面/窄屏交互验收。
- 仓库存储检查与 diff whitespace 检查通过。无新二进制 fixture 或历史验收重签收。

运行产物位于忽略的 `artifacts/f07/`：`f07-packaging-complete.trx`、构建与文档日志、隔离 nupkg/cache、`consumer-output-final/source.ofd`、`copy.ofd` 及审计文件、`dossier.json`、`environment.json`。`manifest.json` 记录本次源码文件和关键运行产物大小/SHA-256。可复跑入口见 `scripts/validation/f07-dossier/README.md`。输入是代码生成的有限结构测试样例，不含真实附件、字体或签章；它们不构成条款字体、媒体、附件转换或 PKI 正例。

## 视觉与外部验收

- 功能测试：通过上述范围。
- 自动渲染检查：本次只读证据索引不生成或修改页面；已有副本操作只改变 Permissions，页面条目由原操作执行字节保留检查。未进行新的 OFD→PDF/PNG 渲染。
- PNG 视觉复查：未做；不据此对任何实际归档页面作外观通过结论。
- macOS Preview：未完成，Linux 无 GUI。独立阅读器、Copy/Find、OFD-A 编辑锁与未处理项 UI 均未验。
- 真实附件、资源/字体/媒体验证器、新转换 handler、归档政策及适用 PKI 独立核准：未完成，具体缺口继续保留在[覆盖账本](../content/reference/archival-standard-coverage.md)及设计记录。

这次测试证明索引和预算/身份/四态边界，不能把摘要匹配、本地 APF 或已有 AR/AC 子条件等同完整标准符合性。F07 尚不能关闭；完整符合性保持 Fail/Unknown，认证为 NotAssessed。

## 官方标准材料场景的补充验证

用户确认当前仅有官方标准，没有指定归档机构、接收政策或独立核准。新增显式外部缺失门禁：机构可为 null，外部基础材料缺失、版本/政策未评估及核准未认证分别报告；空文件或身份错配不填补材料角色。文件角色及版本声明齐全仍保持外部评估 Unknown。此更新未实现剩余资源/媒体/附件处理器，F07 整体状态不变。

补充后 Packaging 全套 **1699/1699**、零失败零跳过（新增测试累计 16 个实例）；文档生产构建及页面/链接/搜索检查通过。本轮仅文档正文变化，未重复未变更的站点类型检查；此前 Nimbus partial 边界仍适用。最新测试记录为 `artifacts/f07/f07-external-gates.trx`。

最新 Core/Packaging 候选 `0.2.0-f07-validation.3` 的两个目标框架打包通过；隔离消费者实际执行通过并断言机构未指定、三类外部材料缺失及外部评估 Unknown。产物保存在 `artifacts/f07/consumer-output-external-gates/`；源码、候选包及输出身份另存 `manifest-external-gates.json`，不覆盖前轮记录。

## PR #113 文档示例复核（2026-11-06）

Codex 对初始提交 `5b67218` 审阅完成且无 findings。自查修复使用示例中行尾注释遮住 `scope` 实参的问题，并将隔离消费者的默认候选版本与已验证的 `.3` 对齐。直接提取修复后的文档代码，以本地 `.3` NuGet 编译并将警告视为错误：0 warnings、0 errors；文档生产构建及链接/搜索检查通过。此轮未改变库源码，不重复声称执行新的完整功能回归。
