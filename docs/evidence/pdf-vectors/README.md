# 21 可选 PDF 矢量转换验收证据

最新已完成的运行时验证为 **`8478b84`**：主代理及独立 Low6 全套578/578、Python5/5、默认11/可选13实际包通过；2026-10-09 实看新增八个PDF/16页。**整体视觉门 FAIL / OPEN**：Indexed源与actualr4导出在Preview中蓝色外观有差异，未把它当作通过。原先35页保留原日期/哈希。后续修复、受影响页新包实看及新增文本精度评审仍待闭合。下列R2/R3为历史结果。

当前运行时冻结 **`b9c136b`**：主代理和独立 Low5 的 .NET **542/542**、Python **5/5**、默认11/可选13实际包消费及独立 observer 通过。首轮双 bot 的无绘制路径/奇异矩阵反馈和 Ubuntu `magick` 缺失已修复；所有失败历史保留。2026-10-08 新增八个 PDF / 八页实际 Preview，原 R2 十三个 PDF /27页保留原日期和哈希；共21个文件/35个已查看页面的明确范围。闭合单点填充的装置像素保真未建立，只验收布局与 page policy，不声称任意或逐像素保真。最新推送 head 的复审/CI/threads 待读回。

冻结运行时 `58a710731b93026e2efbd7ca0bf0420e5be0bf65`，起点 `d878aeda57c1da79917e11d0a35bcc028bc4c1fd`。Astra High 先做真实 PDF 探针和设计，主代理 GPT-6.1 Sol High 实施，GPT-6 Sol Low 在 `git archive` 独立副本运行并保留失败，未自行修补。功能和有限样例的 Preview 验收完成；最新 PR 的 Codex、Cursor、CI 与 review threads 仍需读回，不合并或发布。

- 全套 .NET 507/507、Python 5/5。Low 独立默认 11 包、可选 13 包及额外 PackageReference-only observer 均通过，assets 无 ProjectReference；默认 metapackage/CLI 不新增 Skia 或 vector 依赖。
- 同一真实 PDF 的 native 两页共 14 PathObject、27 可见 TextObject、0 ImageObject；原始 Unicode 与连续空格保留，仿射矩阵和绘制顺序实际检查。默认双层仍为两幅页面图和 110 个透明 TextObject。
- fallback 五页覆盖 clip、alpha/Multiply、原始 RGB 扫描页、真正空白页和可原生纯文字页。超出范围默认 Fail，显式整页回退丢弃已暂存矢量，不叠加可见文字冒充 native。
- 原始 R1 PDF SHA256 `40567d531b000b269f5c6a8e3af7d49b54cc0e6b8be66a2096d1fd0e6beae7a2` 未改。R1 Preview 图片硬块失败保留；生产修复仅命中严格证明的 raw RGB/8 单幅全页图片，保留原样本、分辨率和插值标志。absent/false/true 的 decoded RGB、PDF 字典和 mixed false/true/false 资源行为由实际包独立验证。

2026-10-08 主代理在独占 GUI 时段实看下表 13 个 PDF / 27 页，核对实际 Preview 文档 URL 和哈希；逐个关闭本票 13 个文档并释放 GUI。查看链路是实际 PDF → OFD → PDF → macOS Preview，DOCX 回归是 DOCX → Native/default OFD → PDF → Preview。PNG 只是辅助证据，不能代替此记录。Preview 对三个插值标志均显示平滑，不能靠外观证明标志；同一查看器中的源与产物位置、方向、颜色和形状一致。

| 归档目录 | PDF | 已检查页 |
| --- | --- | --- |
| previous-r1 | fallback-source.pdf | 1–5 |
| golden-original-source-repair/output | golden-r1-repaired.pdf | 1–5 |
| candidate | image-absent-source.pdf、image-absent.pdf | 各 1 |
| candidate | image-false-source.pdf、image-false.pdf | 各 1 |
| candidate | image-true-source.pdf、image-true.pdf | 各 1 |
| candidate | image-flags-mixed.pdf | 1–3 |
| candidate | vector.pdf、dual.pdf | 各 1–2 |
| docx-regression | baseline-native.pdf、baseline-default.pdf | 各 1–2 |

检查 CJK、比例英文、局部样式、表格底色/边框、对齐、分页、仿射文字、裁切/遮盖、cubic/v、填充孔洞、页脚和异常空白。扫描页没有叠加 native 文字；源中有意空白和裁切保留。DOCX 两种模式四页均为本次共享导出器生成。只对这些页作结论，不推断任意 PDF/Word 保真。

[acceptance.json](acceptance.json) 分开列出功能、辅助 PNG、实际 Preview、遗留限制和待办 reviews；[manifest.json](manifest.json) 绑定 [evidence.tar.zst](evidence.tar.zst) 的逐文件大小/SHA256。归档包括真实源 PDF、OFD/PDF/SVG/PNG、原始 R1 失败、Low1/2/3 报告和日志、两个可选 nupkg、13 包清单与实际 consumer assets、设计探针和已撤销 fixture 实验。设计样例的自制探针为仓库 MIT；真实样例使用随附 OFL 静态 Noto CJK 字体（SHA256 `3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa`），PdfPig/Skia 许可亦保留。无真实客户文档。

完整字体使 native OFD 约 11.60 MB，大于默认双层约 0.12 MB；源 PDF 约 21.67 MB，native 导出 PDF 约 79.85 KB。不是体积优化，issue05 子集/共享字体服务未实现。原样本回退将原来的大 DPI 图片改为 2×2/3×2 原网格；导出 PDF 约 1.5 KB。源/导出页框在 144 DPI 的画布最多相差一像素，辅助比较取共同视口且不缩放；原始数据/插值标志精确核验。各 renderer 插值算法、设备色彩及第三方 OFD hint 支持没有跨环境保证。

解压和复现（先按 AGENTS.md 配置 dotnet 环境及单节点参数，准备许可字体）：

```sh
mkdir -p artifacts/pdf-vectors/public-review
zstd -dc docs/evidence/pdf-vectors/evidence.tar.zst | tar -xf - -C artifacts/pdf-vectors/public-review
scripts/run-converter-package-e2e.sh 0.1.0-pdfvector.20261008.r2
scripts/run-pdf-vector-package-e2e.sh 0.1.0-pdfvector.20261008.r2 \
  artifacts/package-e2e/0.1.0-pdfvector.20261008.r2/packages artifacts/pdf-vectors/package-r2
scripts/run-graphics-e2e.sh artifacts/pdf-vectors/graphics-regression-r3
```

设计和 API 范围见 [探针](../../pdf-vector-probe-design.md)、[契约](../../pdf-vector-design-contract.md)、[图片修复设计](../../pdf-image-fallback-repair-design.md)、[教程](../../tutorials/18-pdf-vector-mode.md)。输入/工作预算不等于硬进程内存、时间或最终 ZIP 大小上限。所有失败证据保留；`production_release_accepted=false`。

## 首轮审查修复 R3

[review-r3-acceptance.json](review-r3-acceptance.json) 与 [review-r3-manifest.json](review-r3-manifest.json) 绑定新 [review-r3.tar.zst](review-r3.tar.zst)，旧 `evidence.tar.zst` 原样保留。新归档包括 actualr3全部样例、八页Preview记录、首轮两个bot/CI失败、真实旧nupkg复现、Astra18-case语义探针、542全套日志/TRX、Low4失败及Low5成功/source324blob核验/实际包和DLL哈希、验证器真实工具入口记录、许可。运行时及样例不在证据整理时修改。

安全 open singleton / butt 描边 no-op 消耗累计命令但不产生事件；混合正常内容保持native，无事件页默认Fail或显式一幅页面图。闭合 singleton fill 可能涉及设备像素，故明确 `DEGENERATE_POINT_FILL` 整页policy，不能默默删除；与其它段共存也保守拒绝。显式line/cubic含退化仍沿用19契约。奇异或转换后rank-collapse在producer创建事件前复用04数值判定，给出 `SINGULAR_SERIALIZED_MATRIX`；没有泛catch ArgumentException，也未修改19/04/共享renderer。

八个新看文件均在新归档 `candidate/`，每个页1：`no-op-source.pdf/no-op.pdf`、`no-content-source.pdf/no-content.pdf`、`closed-point-source.pdf/closed-point.pdf`、`singular-source.pdf/singular.pdf`。核验源与actualr3 OFD→PDF准确URL/哈希，全页实看，closed-point另在两次放大后的下方点区域复查；源/回退无明显点痕，字体/布局可读，预期144DPI栅格柔化保留。装置像素保真未建立，不以此宣称像素一致。八文档全部关闭、GUI已释放。原R2 normal vector/fallback/images各ZIP条目解压字节相同，dual仅`OFD.xml`元数据改变，原27页的实际观察范围不改日期或冒充新生成PDF已重看。

Low4的near-float观察断言把非零可逆的1e-11对角缩放当作rank-collapse，首次失败原样保留；主实际low4包证明原A native，两策略均正确。新的Low5将该合法案例保留为独立native检查，另用`1 1 1 1.000000001`证明float后rank-collapse，重新运行全部门。Low5首次observer命令字体路径错，在进入converter前停止；错误日志与正确绝对路径重新启动后的完整结果分别保留，没有修补产品或断言绕过失败。六个具名PNG另作Low辅助检查，不能代替Preview。

ImageMagick6 `compare/identify` 与7 `magick`工具选择兼容，阈值/裁切共同视口/不缩放保持不变；本机仅实际验证7及其standalone入口，不冒充Linux6，Ubuntu CI另验。最终PR评论记录最新head的双bot/checks/线程状态，不在文档中自称已包含自身commit。

```sh
mkdir -p artifacts/pdf-vectors/public-review-r3
zstd -dc docs/evidence/pdf-vectors/review-r3.tar.zst | tar -xf - -C artifacts/pdf-vectors/public-review-r3
python3 scripts/collect-pdf-vector-review-evidence.py b9c136bc6f701388fb6a58483a45e2055c22408f 0.1.0-pdfvector.20261008.r3
```

## 第二轮修复 R4 与保留的视觉失败

[review-r4-acceptance.json](review-r4-acceptance.json) 和 [review-r4-manifest.json](review-r4-manifest.json) 绑定 [review-r4.tar.zst](review-r4.tar.zst)（27,452,177字节、217个文件）；前两个归档原样保留。此轮为源码8478b84，路径精度损失在创建事件前按明确producer误差界进入page policy；合法Encoding/CID映射流也进入policy，严格间接引用的损坏/环/缺失仍为输入失败，不吞异常。Low6独立332个源码blob核验、578全套/5Python、11/13实际包和23条路径/3类资源observer通过。

八个本轮PDF：precision-source.pdf/precision.pdf各3页，encoding-stream-source.pdf/encoding-stream.pdf与cid-map-stream-source.pdf/cid-map-stream.pdf各2页，indexed-image-source.pdf/indexed-image.pdf各1页。几何、字体原文、空格、page policy范围检查通过；Indexed棋盘格布局完整，但源/目标蓝色外观差异为真实未通过项。全部八窗口已关闭并释放GUI。未宣称整体视觉通过、逐像素保真或发行通过。失败源PDF SHA256为49994fe3d745b9623ab098c98285c9a9cf6c39024edeba3da439714410913441；目标PDF SHA256为00c6c9cbfc1aa250055b066e667c935fbfaab85fc0b18c0b087d421d4d7b77b4。详细查因及修复实包证据将追加，不能以扩大限制文案关闭此门。
