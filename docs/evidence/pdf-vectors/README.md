# 21 可选 PDF 矢量转换验收证据

R11 writer-precision 修复候选正在验证：R10 exact-head Codex 新增两个 P2，极小正 width/font size 在三位写出精度归零时绕过整页 fallback。已将 04 原样 WriterValue 抽取到 Core 并保留 04 Positive 语义；21 对实际 float scalar 在原精度及硬预算/取消门之后、事件前分类为 STROKE_WRITER_PRECISION/TEXT_WRITER_PRECISION。R10 已完成的 28 页 Preview 和下列证据保持历史身份，不代表 R11 验收；新候选全套/真实包/独立验证/新增页面 Preview/最终 head 复审尚待闭合。

R10 运行时为 **`28fc29d`**，交付冻结源码为 **`6fa0a30`**（其差异仅为文档）。可选 Vector producer 的每个 VectorFont/page 私有复制原字体 bytes 并创建私有 face；首次有效文本逐块比对实际流与快照并 SHA，后续文本复用私有身份。公共 Text 仍每次完整验证，内部/公共路径共用每 run 校验，Append 资源认证保持；取消/事件门前置，不新增默认 quota，不改变 fallback/font 输出预算。

R10 主验证与独立 Low12 均完成 **641/641**（各九份 TRX）、Python **5/5**、默认 **11** 包及可选 **13** 包消费。独立结果包括原始/高 Index 图像、冻结九页文字观察器、八个 Tj/TJ 1/10/50/100-run 样例、真实主 R10 包测量，以及 21 个旧 OFD payload 和八个新压力 OFD 对照。主协调在 Low12 结束后复核了 70 项产物、365 个冻结源码文件和 13 个主 nupkg 的哈希。

2026-10-13 对本轮实际产物完成 **10 份 PDF /28 页 Preview**：normal 源与 native 导出各两页、文字精度源与导出各九页、high Index 源与导出各一页、Tj100/TJ100 两组源与导出各一页。逐页核对准确 URL/哈希、正文及整页，重绘精度第四页重叠字形簇和第五页真实空白；源/产物未见阻断差异。压力页本来是 12pt 字体/4pt baseline 的密集重叠列，不声称肉眼数出每个字形。误开目录的 GUI 事故、实际 AX URL 清理证据及临时 PostScript 归属未独立证实的限制另记，误开文件不计验收；本轮八份剩余目标逐个核对准确 URL 关闭，最后取消打开面板并释放 GUI。

[review-r10-acceptance.json](review-r10-acceptance.json)、[review-r10-manifest.json](review-r10-manifest.json) 绑定 [review-r10.tar.zst](review-r10.tar.zst)，包含主/Low12 日志、TRX、实际包与消费者清单、源码/包冻结、测量源码与完整诊断 diff、IL、OFD/PDF/PNG、逐页 Preview 记录及 GUI 事故证据。真实包观察与诊断副本计数明确分开：1/10/50/100 个 snapshots 复用同一私有身份；完整流次数由加载的包 IL 推导，副本每 binding 实测一次 21,659,544-byte 读取/SHA。中途取消注入仅为诊断副本，不冒称真实 nupkg 中途延迟测试。实际 consumer 的八个 DLL 副本均匹配包条目，runtime.json 明确记录三个加载程序集，不声称完整 AppDomain 清单。单进程计时无统计 SLA 或全转换硬内存/时间沙箱承诺。

R9 初次 CS0122 构建失败、公共每次流比较的额外耗时及修复、R10 首次把 21 个 OFD 误写成 20 的比较脚本失败均原样保留。旧归档哈希不变。最终提交的 CI、Codex/Cursor 复审和线程状态以 PR #18 当前 head 读回为准，不继承旧 head 结果；PR 保持开放，未合并、未发布。

R8 归档代码基线为 **`d8d197e`**，生产 `src` 与 **`deda6d2`** 相同。主验证 .NET **625/625**、Python **5/5**、默认 **11** 包和可选 **13** 包通过；独立 Low10 通过相同功能门后，因把 gzip 路径传给要求 raw PDF 的 observer 而停止。失败原样保留，Low11 使用相同源码、包和既有二进制补齐未执行观察门，未重新声称全套测试由 Low11 执行。原九页输入 SHA `d78e…`、完整字体 SHA `3012…` 均保持不变。

R8 归档时整票未闭环。2026-10-13 读回 R8 head 的六项远端检查均成功，Cursor 为 No findings；Codex 的字体重复哈希 P1 仍开放。Astra 在真实 R8 包上测量固定 21,659,544 字节字体的 1/10/50/100 次 Text 调用，并用明确分开的诊断源码副本计数：100 次完整读取/哈希共 2,165,954,400 字节。修复设计为每个 VectorFont/page 生命周期内复用经实际字体流验证的身份，该归档形成时 19 内部桥接待协调审查，未实施或验收性能修复；后续 R10 实施与验收见本文开头。

视觉记录按实际产物区分：R6 `deda6d2` 的四个 PDF /20 页于 2026-10-09 完成 Preview，旧轮次观察保留原日期和哈希。R8 的新导出 PDF 字节不同，**未做 R8 Preview**；相同 OFD payload、功能测试及辅助 PNG 均不替代新文件实看。后续 R10 已完成重新冻结、回归、真实包独立验证与受影响产物 Preview，范围见本文开头。PR #18 保持开放，未合并、未发布。

| 轮次 | 基线 | 保留的主要结果 |
| --- | --- | --- |
| R10 | 28fc29d | 主/Low12 各 641/5、11/13 包；真实包/独立观察；28 页 Preview，最终评审另记 |
| R8 / R9 设计 | d8d197e | 625/5、11/13 包；Low10 失败、Low11 补验；字体性能 P1；无 R8 Preview |
| R7 | abd85f3 | 主 optional13 /Low9 资源输出路径失败，修复后由 R8 重跑 |
| R6 | deda6d2 | 625/5、11/13 包、20 页 Preview；Linux 重生成冻结输入失败保留 |
| R5 | 8de5958 | 602/5、11/13 包、8 页 Preview；原调色板恢复；后续文字精度及高 Index 边界待修 |
| R4 | 8478b84 | 578/5、11/13 包、16 页实看；Indexed 蓝色外观失败保留 |
| R3 | b9c136b | 542/5、11/13 包、8 页 Preview；Low4 观察断言失败保留，Low5 独立通过 |
| R2 | 58a7107 | 507/5、11/13 包、27 页 Preview，原 R1 图片失败保留 |

下述样例和体积说明属于 R2 的历史验收范围。冻结运行时 `58a710731b93026e2efbd7ca0bf0420e5be0bf65`，起点 `d878aeda57c1da79917e11d0a35bcc028bc4c1fd`。Astra High 做探针和设计，主代理实施，Sol Low 在独立冻结副本验证；各轮后续状态见表及归档。

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

## 原调色板采样修复 R5

[review-r5-acceptance.json](review-r5-acceptance.json)、[review-r5-manifest.json](review-r5-manifest.json) 绑定 [review-r5.tar.zst](review-r5.tar.zst)（9,494,526字节/159文件）。Astra只读颜色诊断以原palette/index确定旧Poppler页面回退有23520/999600个边缘混色；旧OFD PNG、导出PDF保持这些混色，内部颜色未变，不能由此证明Preview内部蓝色差异的原因。诊断及失败都保留。R5仅在既有严格纯整页条件内把literal Indexed/DeviceRGB8调色板展开为RGB8，并保留原网格及源Interpolate，仍是非native图像页；未改默认renderer/字体/19链路。

真实R5与独立Low7包对原SHA49994fe3源全部999600像素检查，PNG/PDF字节均等于原palette展开；有效Interpolate=false。GS与CoreGraphics共同840x1190视口各0调色板差异；Poppler源与目标pixelsha相同，但两者都有其23520个采样边缘混色，不能称所有renderer逐像素保真。旧R4源/新R5产物的八页实看分别为indexed-image-source.pdf/indexed-image.pdf（原棋盘格）、indexed-absent/false/true-source.pdf及各对应实际OFD导出.pdf，各页1。准确Preview URL和哈希、全页重绘、棋盘格另Actual Size匹配重绘，蓝色/浅格/四边/方向一致，原finding不再可见；3x2三组都呈平滑渐变，不用外观证明flag。八窗口逐个核对本票URL关闭，自动打开面板Escape后noWindowsAvailable，GUI已释放。

Low7核验338源码blob前后未变，602/5全套、11/13实际包和独立原PDF observer均通过；首个observer命令把MSBuildflags交给exe，在打开源前失败，原日志与直接DLL运行分别保留。Pillow不可用，Low用实际ImageMagick共同视口无缩放比较；Main真实CoreGraphics另验。Main首次Python发现目录写错，在进入测试前失败，原日志与scripts/tests五项结果分别保留。14类无变化OFD解压payload与R4相同（dual仅OFD.xml元数据），旧观察保留原日期。

本轮范围是原样本0..hival，后续标准核查确认超过hival的合法Indexed值应夹取，当前fatal边界需要修正；不得把本次有限样例通过变成任意Indexed/PDF接受。新的文本matrix/advance/font-size P2也未解决。大诊断PDF副本可由保留脚本、R4原工件重现；公共增量归档保留脚本/哈希/原RGB/页面图，未复制其重复完整字体。所有产品工件与历史归档未替换。

## 文本精度与合法Indexed夹取 R6

[review-r6-acceptance.json](review-r6-acceptance.json) 和 [review-r6-manifest.json](review-r6-manifest.json) 绑定 [review-r6.tar.zst](review-r6.tar.zst)（21,277,567字节/205文件）。Astra真实旧r4 nupkg九页复现、原始source/model/PDF/PNG及设计、主625测试/九个独立TRX、Low8全套/346源码blob核验/真实包与三个observer、四PDF20页Preview和failed Linuxjob113681779442日志一并保存。Astra可选PIL contact-sheet不可用原日志保留；本轮不把源/consumer后续decimal舍入叫成producer1e-4pt界限。

2026-10-09在准确Preview URL/固定哈希下实看text-precision-source.pdf/text-precision.pdf各页1..9，indexed-image-clipped-source.pdf/indexed-image-clipped.pdf各页1。first3保留预期可见中glyph，page4源有意把2049glyph叠入两个字形簇，回退保留而无额外可见native overlay；page5源和导出切页后曾保留前页缓存，单次ZoomOut留在新比例后均完整重绘为空白（巨字在页外真实预期），没有将微小shape差称作肉眼已证。source1..3又重绘复查。6..9比例英文、shear中英、中文原文、A-double-space-B及字距可读且与源一致。高Index0/255/hival1两页棋盘格四边/方向/颜色一致。四本票文档URL逐个核对后CmdW关闭，最后Cancel自动Open面板并停止激活Preview，GUI释放；只读应用清单无windows字段，未据此推断全局0窗口。

Low8三个独立观察输入分别保留原49994fe3、合法高值525f7c78、原文本d78e；图像均检查999600采样，文本Fail sentinel/五回退四native/重复选择[5,0,5]/四个完整原字体资源通过。实际malformed palette五项，copied console旧标签写六项已在报告纠正；旧Low7高值误判历史未抹去。Low8固定TRX文件名覆盖了前程序集TRX，仅最后168项；全625计数有完整控制台日志及Main九份TRX，不冒称独立九TRX齐备。

当前LinuxCI的Frozen Astra text source changed说明harness重新生成输入未等于本应被冻结的原PDF；该失败先于源文件持久化，失败的九页内存字节未知，不能猜成已证实字体或renderer原因。后续直接消费压缩原PDF，继续要求解压后原size/SHA与原字体payload，保留此失败，并复核新冻结harness/实际包/Low/LinuxCI。生产src保持deda不变；新生成PDF如需验收须按其实际hash重开，不把model相同当新文件已看。此处只关闭最终运行时已实看记录缺失，不声称PR当前CI/全部线程已闭合或可发布。

## 冻结输入harness失败 R7

[review-r7-failed-acceptance.json](review-r7-failed-acceptance.json) 与 [review-r7-failed-manifest.json](review-r7-failed-manifest.json) 绑定 [review-r7-failed.tar.zst](review-r7-failed.tar.zst)（9,356,302字节/41文件），旧五归档不变。R7生产src与deda相同，把原d78e PDF以lossless gzip存入E2E，保留原size/SHA及原font3012断言；主625/5/default11通过，但主optional13及Low9干净消费者在缺少bin/testdata文件时失败。Low9遵守首个实际失败即停止：353源码blob、gzip两层hash、default11通过，full625/Python/observers/新Preview未跑，不把旧Low8结果当作新通过。

只读结构诊断证明None items和CopyToOutputDirectory存在，但实际消费者把资源扁平复制到bin根目录，代码期待testdata子目录。后续显式TargetPath修复需新freeze/真实包/Low/CI与实际新PDF验收；旧失败保持不变。主早前源码预检成功不能替代干净包消费。

从failedLinux artifact11597357047实际取到同一fixture生成器的same-source.pdf，ZIP entry大小/CRC验证通过。嵌入完整FontFile2与本机完全同3012a9...，但实际/W数值不同，如空格223.999741对224、A607.999084对608；证明该生成器输出有跨平台度量差异。失败九页内存输入在assert前未落盘，仍未知，不能冒称已恢复或精确比较其字节。suffix range请求不支持的失败保留；显式范围成功取到合法所需entry后，停止冗余392MB全量下载，47MBpartial明确不算verifiedartifact，未纳入公共归档。

## 资源输出修复 R8 与字体身份 R9 设计

[review-r8-acceptance.json](review-r8-acceptance.json)、[review-r8-manifest.json](review-r8-manifest.json) 绑定 [review-r8.tar.zst](review-r8.tar.zst)（27,192,342 字节/148 文件）。包含主九个 TRX/625 测试及 5 Python、11/13 包清单/assets、实际 R8 受影响 OFD/PDF/PNG、Low10 全套门及失败、Low11 使用相同候选的限定补验、冻结源码/输入/二进制哈希、原始及高值 Indexed 999600 采样观察、source9 原文/字体/重复页选择、辅助页图、当前 head 审查和检查快照。新文件未 Preview，不以 R6 实看记录冒充。

R8 显式 TargetPath 让干净 consumer 自动得到 bin/testdata 原 gzip，不再扁平复制。2026-10-13 远端读回确认 d8d197e 的 build-test、三平台 BuiltIn DOCX、package-e2e 和 Cursor 检查全部成功。历史 Linux 冻结输入失败和 R7 复制失败仍在前一归档；未来 head 必须重新读回，不能继承本次成功状态。

Low10 在真实包/full625/5 通过后误把 gzip 当 raw PDF 传给 TextObserver，source identity 在进入文字 converter 前拒绝，停止记录不改。Low11 修正调用参数，使用同一已编译 high/text observer；原始 Indexed 的独立 observer 此前未编译，获准只首次构建不变源码，再补齐三个观察门。输入、候选源、既有 DLL 和 nupkg 前后哈希不变；18 个未影响 OFD payload 相同，dual 只有 OFD.xml 元数据不同。Low11 报告引用 Low10 已通过门，不宣称重复执行全套，也不涵盖后续字体性能修复。

归档 font-identity-r9-design/report.md 是 Astra 的只读设计与有界真实 R8 包测量。真实包构造 1/10/50/100 次 Text 的时间分别为 8.18/73.09/339.47/670.76 ms；一次顺序测量，非统计基准。读取/hash 次数来自单独命名、附完整 diff 的诊断源码副本；未对真实 nupkg 插桩，真实包计数由加载 IL/调用链推导，不能冒称 profiler 数据。真实包转换 1/10/50/100 Tj、正常 27 run 和原 source9 四个 native 页保留全文及完整字体；取消/各预算失败保持 sentinel。没有执行 640 GB 工作负载，TJ 与中途 hash 取消延迟未实测。

设计建议仅给 Vector 程序集可见的私有拥有固定 face 的不可变身份绑定，首个有效 snapshot 分块校验 actual face stream 与原 bytes 后只哈希一次；每次事件仍验证文本、matrix、paint、advance、size，public Text 继续完整验证任意传入 font，Emit 每个 Append 仍独立验证实际 destination payload/style/ID。没有全局缓存、字体解析/子集服务或 04/05 改动。19 桥接必须先完成协调审查，修复、计数回归、独立新包、当前 head 评审及新 Preview 均属待办。
