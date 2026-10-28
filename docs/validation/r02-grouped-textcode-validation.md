# R02 分组策略候选验收记录（2026-10-28）

本记录覆盖 `GroupedLocal` 的有界增量：按 XML 分组绑定映射，组合既有段内基数能力，保持默认 `LegacyConcatenatedGlobal`。对应 [PR #70](https://github.com/whynpc9/ofdrw.net/pull/70)。本增量不解决任意越段映射、补充平面索引单位、无定位复杂映射或任意业务原件互操作；PR #67 的审阅处理未恢复。

## 身份与功能结果

产品源码冻结于 `e45379a9d04fc0e9b82ad67bc5dd2b6aaa479081`，包含 main `0d3c3b5c`（PR #69）及 PR #48。后续 `ea88993f` 仅修正验收脚本的依赖缓存选择，`src/`、`tests/`、`e2e/` 和构建属性相对冻结源码无差异。包版本为 `0.1.0-r02.20261028.18`，一次打包得到 12 个包，各消费者使用同一份清单和包载荷。

| 检查 | 实际结果 |
|---|---|
| Release/API 文档门 | 通过；没有新增未注释公开 API |
| 全套测试 | 11 个程序集，4,353 通过、17 既有跳过、0 失败；其中 CGTransform 定向测试 94 项通过 |
| 广义精确包 E2E | 通过，包含 R22 的 AV/ffi 开关及 e + U+0301、a + U+0308 六个生成样例 |
| macOS 隔离包消费者 | 24 项出口检查通过：21 项接受、3 项预期写出前拒绝 |
| WSL Ubuntu 24.04 隔离包消费者 | 同样 24 项通过；.NET SDK 10.0.401、linux-x64 |
| Windows 原生隔离包消费者 | 同样 24 项通过；Windows 10.0.26200、.NET SDK 10.0.301、win-x64；由原生 `dotnet.exe` 执行 |
| 包身份 | 三平台实际加载的 Ofdrw.Net DLL SHA-256 均属于同一候选包；远端 12 个包及全部 OFD 输入哈希与本地一致 |
| 原 GID / 坐标 | 每个平台 7 组输入/输出、29 个可见字形通过独立 FontTools 与 PDF/SVG 路径核对；详见下节 |
| PNG 渲染 | 32 个 PDF/SVG 产物生成 34 张页面图；生成成功本身不构成目视验收 |
| macOS Preview / PNG 目视 | 指定 Terra High 子代理实际检查 34 张 PNG、17 份 PDF 的 19 页 Preview 正文；所查范围未见明显视觉缺陷，原生窗口已关闭 |
| Windows 原生阅读器 GUI | 未执行；原生 .NET 消费不能替代阅读器验收 |
| 最终 PR CI / SES / 审阅 | 以最终头提交的实际回执为准；本地结果不替代远端门 |

macOS 构建环境为 26.6.2、.NET SDK 10.0.401；独立 Go 工具为 go1.27.1 darwin/arm64。构建通过共享独占 launcher，使用单节点、关闭 build servers/共享编译。只读来源、生成输入、包隔离与几何脚本见 [`scripts/validation/r02-grouped`](../../scripts/validation/r02-grouped/README.md)。

## 独立输入与互操作矩阵

`xiaoqidun/ofdgo` 固定于 `7e3f2c468aaa56d461f117cc045413c1b457e75c`；Go module 与依赖校验和随复现脚本保存。使用其公开 `Editor.AddObject` / `SaveAs` 生成原件，源字符与字体选取由定向用例明确指定。其写出器会做保留 GID 槽位的字体子集；验收使用写出的原件和其中实际字体，没有修改 XML 或替换字体后冒充原件通过。

| 输入来源 / 用例 | 候选 PDF/SVG/HTML | 独立阅读器输出与边界 |
|---|---|---|
| ofdrw.net 定向输入：second-only | 接受；第二段位置 0 正确作用于 C，可见 AB / ED | ofdgo PDF 的嵌入子集轮廓与原 GID 一致，原点一致 |
| 两段位置 0：both-local-zero | 接受；可见 EB / FD，逻辑 ABCD | 同上；没有把第二段映射当作重叠 |
| 第一段 2:1、第二段 1:2：unequal-both | 接受；可见 fl 连字 / AB，逻辑 flA | 同上；源偏移 2 与输出字形偏移 1 分离 |
| 第二段普通字夹 2:1：mixed-second | 接受；可见 AB / Cfl 连字D，逻辑 ABCflD | 同上；三个第二段字形起点独立累积 |
| 公共 Go 写出器前三种可兼容原件：second-only、both-local-zero、mixed-second | 三种均在三个出口、三个平台接受 | 与相应定向输入的 GID、原点契约一致 |
| 公共 Go 写出器 unequal-both 原件 | 三个出口均在写出前拒绝 | 编辑器为 `fl → 一个字形` 自动补 `DeltaX="0"`；数量对应源字符间隔，而候选要求输出字形槽位间隔。保留该原件作为兼容边界 |
| 分组后的 Latin combining 回归 | source 测试接受第二段 `e` + U+0301 → GID 171，并保持分解 Unicode；保留 PR #69 单段完整源簇路线 | 属于单独的合并回归，不把它加入上述 29 字形互操作范围 |

对七组接受的候选输出，FontTools 从原固定字体独立解出完整轮廓，检查所有 SVG 段及 PDF 可见路径，最大 PDF 坐标误差约 `3.05×10^-5 mm`，SVG 误差在浮点舍入量级。独立 PDF 中的子集字形轮廓与对应原 GID 完全一致，原点最大误差约 `2.25×10^-6 mm`。没有只靠 `data-original-gid` 属性或 PNG 相似性判定正确。

独立阅读器的 PDF `/ActualText` 保存原逻辑文本（分组间插入换行），但 pypdf 提取路径读出可见字形对应的 Unicode：例如源 `AB/CD` 可提取为 `A B/ E D`，`fl/A` 可提取为 `ﬂ/ A B`。候选 PDF 的 PdfPig 检查得到精确源序列。两者是不同证据，不能据此宣称所有阅读器的复制或查找一致；Preview 的实际复制/查找结果另列。

## 实际页面与 Preview 行为

视觉验收使用 `gpt-6.1-terra`、High、独立上下文子代理，输入与 34 张页面图先核对 SHA-256。检查包括七份候选 PDF/SVG、四份独立阅读器 PDF/SVG、分组 combining 与 legacy 对照、native/default DOCX 两份各两页 PDF，以及 R22 的两个组合字符 PDF/SVG。15 张 SVG 页面图直接来自 SVG 离线渲染；19 张 PDF 页面图另行检查，并在 macOS Preview 中新打开对应 17 份 PDF 逐页查看。原始 GID 外观、两段定位、重音、DOCX 中英文和局部样式、表格与分页在所查页面未见明显缺陷。

DOCX 查看链路为 `generated-layout.docx → native/default OFD → PDF → Preview`；没有用直接 DOCX→PDF 代替。原生 GUI 操作结束于 `2026-10-28T05:39:28Z`，自有窗口/标签关闭后，CUA 返回 `noWindowsAvailable`。截图及逐项记录保存在证据包的 `visual-review/`。

| Preview 操作 | 实际结果 | 结论范围 |
|---|---|---|
| 候选 second-only / both-local-zero 复制 | `AB\nCD` | 源段文本保留，阅读器在组间加入换行 |
| 候选 unequal-both 复制 | `fl\nA` | 没有按可见 `ﬂ / AB` 替换逻辑源字符 |
| 候选 mixed-second 复制 | `AB\nCflD` | 映射前后普通字符和源 `fl` 保留 |
| 分组 combining、R22 e-acute / a-diaeresis 复制 | 剪贴板为预组合 `é`（U+00E9）/ `ä`（U+00E4）；AX 选中文本显示分解形式 | 不能将剪贴板结果记为原始分解码点完全保持；与 PdfPig 精确序列检查分开记录 |
| 独立 second-only / both-local-zero 复制 | `A B\nE D` / `E B\nF D` | 此阅读器路径复制出可见字形对应文字，未保留原逻辑串 |
| 候选 unequal-both 查找 `fl` | 显示一个匹配并有可见高亮 | 仅此查询取得明确匹配证据 |
| 候选 second/both 查找 `CD`，mixed 查找 `CflD`，combining 查找分解 `e` + U+0301，R22 查找分解 `a` + U+0308 | 同时出现 `Found on 1 page` 和 `Not found` | 结果不确定，不记查找通过 |
| 独立 second/both 查找 `CD` | `No Results Found` | 原逻辑串未找到 |

其它未取得有效输入或明确结果的查询不计通过。跨阅读器查找一致性、复制时的组合形式以及选择/高亮几何继续单独跟踪；页面目视通过不覆盖这些语义差异，也不替代 Windows 原生阅读器 GUI。

## 默认策略与剩余范围

本次保持旧调用默认值。对于同一分组输入，两个策略可能都成功但绘制不同：legacy 将第二段前的位置 0 解释成第一段的 A，grouped 解释成第二段的 C。调用者必须明确选择；此兼容决策不为 legacy 在分组输入上的语义正确性背书。

分组 PDF 使用已存在的原 GID 轮廓加逻辑文本路线，逻辑源段在段起点保留一次。可见字形按输出槽位定位；本记录不为任意阅读器的选择框、查找高亮或复杂源簇内文本坐标作保证。SVG/HTML 路径与逐段可访问文本分别验证。

未知业务外来文件、单个映射跨出所属 TextCode、非 BMP 计数、通用 shaping/bidi、一般无定位复杂映射仍需独立证据。Windows 原生阅读器的页面/复制/查找仍未验证。R22 生成端的其它能力继续由独立工作线交付。

## 复取材料

关键输入、实际输出、页面图、独立几何/包身份记录和视觉验收记录按仓库存储规则保存至[可按需下载的 LFS 证据包](../evidence/r02-grouped-textcode-20261028/README.md)；不归档构建目录、包缓存、第三方源码副本或全量包。公开规范引用及标准 PDF 哈希保存在[方案文档](r02-grouped-textcode-plan.md)。
