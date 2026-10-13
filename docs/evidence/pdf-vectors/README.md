# 21 可选 PDF 矢量转换验收证据

R15 当前运行时 **`f5914e6d41d63da541da5d8e031e7d97090296cc`** 在原 float 精度门之外，按实际双精度毫米 CTM 核验 Writer 写出的正笔宽/字号总误差；原 `.0006 × 100000` 的 60pt→100pt 膨胀分别进入 `STROKE_WRITER_PRECISION` / `TEXT_WRITER_PRECISION` 整页策略。路径 miter 绝对界、文字映射/位置/前缀/基线界均保留；原硬预算/取消优先级不变。成功处理后、Append 前按同一毫米尺寸分类 `PAGE_WRITER_PRECISION`；无法写成正数的物理页尺寸对两策略均原子失败，私有异常仅排除这一不可栅格化类型。Core/04/19/Writer 语义未改，早期空页/unknown/raw-image及默认Invisible的既有分支保留。

主与独立 Low17 各 **789/789（九份 TRX）、Python5/5、默认11/可选13实际包消费**通过。349冻结源码/13主包身份不变，独立1061产物逐项size/SHA已核验。主四个、独立八个真实PackageReference消费者各加载八个与自身nupkg对应的DLL，RepositoryCommit均为f5914e6。Astra保存112个唯一PDF、312次转换按原SHA/options重放，仅10页尺寸及44标量精度策略结果获准改变，其余258不变；原R14 158、R13 80、R12 62+2 BOM原语义保持。34旧OFD payload保持，35当前主/Low OFD对应，仅dual的两项日期元数据归一化。基线13失败/13通过、初次full783通过/6旧预期失败、探针编译错误、Low首轮TRX文件名覆盖及后续同二进制九TRX补验均保留，不冒称首次全绿。

2026-10-13实际在macOS Preview逐页检查本次 **writer-accuracy-source.pdf / writer-accuracy.pdf 各4页，共8页**，精确AX URL、分别设置Single Page/Zoom All to Fit，检查前后SHA/size不变。查看链路为源PDF对照本次 **PDF→native/raster/raster/native OFD→PDF**。页2红色60pt线条厚度/位置、页3绿色60pt“中文 A”字号/间距/基线对应，未见旧100pt膨胀；页1/4正常native恢复，未见缺字/重叠/裁切/重影/额外空白。栅格页边缘/文字略软，红绿有轻微色调观感差异，明确保留为限制，不以总RMSE覆盖。大块下方白区是原布局。源21,664,058B SHA `a6a8029426f4971779aab9fe1b41f3db5adc93a966d6b8e8addbc1e4594ea617`；导出160,760B SHA `63a7fb051025655f2fc464af787e51821047ea5106d10fb94d04e81d5b552183`。两个极小页样例两策略均只有失败输入/日志/sentinel，无成功输出PDF，未计视觉通过。

封存前复用现有源/导出PNG及实际OFD载荷做有界只读色差核对：两嵌入PNG逐字节等于对应源Poppler PNG。共同viewport不重采样，页2区域XYWH(200,450,400,40)内5×5纯色邻域14256点的源/载荷/导出均RGB(186,59,64)；页3区域(120,370,360,130)绿色内部3539点均RGB(51,140,89)，内部差异及最大通道delta均0。未见这些内部区域颜色污染；源840×1190/导出841×1191与边缘采样可造成柔化，但未证明Preview色彩管理原因。没有同源dual直接控制，不宣称跨renderer像素/色彩保真或任意PDF保真。

两份准确URL文档均已关闭；末态实际Open面板点击Cancel后仍显示Open，取消/零窗口未确认，不再激活、不操作未知窗口，GUI已释放。首个GoTo helper因AX diff前缀安全拒绝，改用full AX后准确打开源，未误开文档。14条实际UI调用/8幅实际内联页截图的事件记录保留；session图片数据可截断，不冒称可用导出Preview截图，PNG仅为辅助。历史R14四页/R13二十页/R11八页/R10二十八页保留原身份，未冒称本轮重看。

[review-r15-acceptance.json](review-r15-acceptance.json)、[review-r15-manifest.json](review-r15-manifest.json)、[review-r15-readback.json](review-r15-readback.json)绑定本轮结果。运行 `python3 docs/evidence/pdf-vectors/verify-r15.py` 实际重建分卷并核对whole及所有成员size/SHA。原R12/R13/R14分卷未改。最终evidence-only head的Codex/Cursor、六CI及有效线程单独读回，不继承历史结果；PR18保持OPEN，未合并、未tag、未发布。

R15最终归档 **2624成员 / 49,629,544字节**，whole及单分卷SHA256 `88d1dd408e68adfebb42e390335e84ea29fbc8f6b73ccc3916ab360ca734d74f`；[review-r15.tar.zst.part01](review-r15.tar.zst.part01)已实际重建并逐成员核验。初始本地880d53…草稿在补充只读颜色归因前已核验，随后重收集为本最终版本，未公开该草稿。

以下R14及更早说明保留原验收身份。

R14 历史运行时 **`35fe18e01ef3cbf2f287b5d70018681f9f46deaa`** 修复已耗尽事件预算仍被路径/文字 profile 错误降格为整页回退的问题。真实/潜在路径绘制先检查容量，fill+stroke 一次要求两个槽位；非空 glyph 先检查已消费字符额度和事件容量，非空游程在 profile 前复核。只由成功 Add 计数，支持的空路径/move/空 show/Tf 不收费；路径构造、close、字体输入尝试及取消门保持原顺序，不预扫描未来字形。

主与独立 Low16 各 **763/763（九份 TRX）、Python5/5、默认11/可选13真实包消费**通过。345 冻结源码及13主包身份未变，Low16全部665产物逐项size/SHA复核；六个独立主包/本地包消费者各加载八个与自身nupkg匹配的Ofdrw DLL，两套13包的RepositoryCommit均为35fe18e。Astra原158次转换使用49个唯一PDF输入，原SHA/options在实际主包和Low包重放，**仅36次批准的事件/已消费字符预算情况变为 InvalidDataException，sentinel040506和position0保持，其余122次结果/诊断不变**。R13保存80次及R12保存62+2 BOM保持原语义。33个旧OFD payload与R13相等，34个当前主/Low OFD payload相等，仅dual的CreationDate/ModDate归一化。

2026-10-13在macOS Preview实看新源PDF及本次实际OFD→PDF导出，**两PDF共4页**，各自明确设置Single Page/Zoom All to Fit，精确AX URL匹配，检查前后SHA/size保持。第1页容量充足时按PATH_FLOAT_PRECISION整页栅格回退，中英蓝标题/浅蓝填块/红色短窄竖标记保留，文字/边缘略软；第2页正常native中英标题、填色及蓝边框对应，尾随支持空操作无额外可见内容。未见明显缺字、裁切、重影或额外空白页；大块留白是源布局。容量2的Fail/RasterizePage均为硬失败、目标未动，没有失败PDF可作视觉验收；容量3的Fail保留原精度诊断，显式RasterizePage才产生上述实际raster/native产物。源21,662,719B，导出138,583B；实际OFD11,610,178B，大小差异包含字体/图像压缩，不代表任意PDF更小或保真。

两份准确URL文档均已关闭。末态AX出现任务Open面板，Cancel及Escape尝试均返回no change，因此**未确认面板取消或零窗口**，随后停止激活并释放GUI；未知窗口未操作。源关闭后下一快捷键曾因Computer Use not active拒绝，独占时段内重新绑定一次继续准确导出，未误开文档。五次Preview截图观察实际内联显示（含导出首次连续模式及最终单页复查），归档保存24条真实UI调用/AX/错误与图片块摘要；截断payload不冒充可用导出截图，四张PNG只作辅助。

[review-r14-acceptance.json](review-r14-acceptance.json)、[review-r14-manifest.json](review-r14-manifest.json)、[review-r14-readback.json](review-r14-readback.json)绑定本轮限定结论及不可变归档。运行 `python3 docs/evidence/pdf-vectors/verify-r14.py` 重建并核对whole SHA和每个成员size/SHA；各分卷低于90MiB。实际main/Low OFD、新源/导出PDF及PNG、49个保存输入/158次转换记录、原失败探针、测试日志、包/源码身份、4页Preview记录保留；派生SVG字体base64和cache省略，旧输入/PDF/PNG继续由原R12/R13归档提供。首次未push证据草稿遗留主Preview的pending字段，被独立审查发现后修正并重收集；原草稿和纠正记录保留，不影响产品/视觉文件身份。

默认Invisible对XML非法FFFE/FFFF仍不支持，原ArgumentException及sentinel保持；显式None是调用者放弃回退页搜索层的选择，没有自动禁用隐藏文本层。R13的20页保持7afd身份，R11八页/R10二十八页保持各自历史身份，未冒称本轮重看。最终head的Codex/Cursor、六项CI和有效线程另行读回，PR18保持OPEN，不合并、不tag、不发布。

R14最终归档共 **1156文件 / 48,775,678字节**，whole SHA256 `af8dacc7cebd9358d2feae91cada0af128ad494424dc042a6aceb573c8245bb4`；本轮单分卷 [review-r14.tar.zst.part01](review-r14.tar.zst.part01) 与whole同SHA/大小。已实际重建并逐项核验全部成员，先前草稿cd786…仅本地保留，不是最终公开分卷。

以下R13及更早说明保留原验收身份。

R13 历史运行时 **`7afd12042e9099093d9974472d1431686304a162`** 已完成 R12 按实际字形使用字体及 R13 XML 原始文字分类的功能和限定视觉验收。最终 evidence-only head 的双 bot、CI 和线程状态仍以 PR #18 当前 head 单独读回，不继承运行时或历史 head 的结果；PR 开放、未合并、未发布。

R12 只在非空字形 bytes 实际绘制时加载字体，并只输出成功文字事件引用的字体；仅 Tf、空 Tj/TJ 和纯数值 TJ 不嵌入字体，数值 TJ 的文字位置/缩放及 q/Q 恢复照常执行。读取尝试使用独立页预算，实际输出字体仍计入既有全局预算。实际 R12 包和 R13 包复放保存的 62 个转换及两个 BOM 控制，输入 SHA 不变、全部结果与资源身份一致。R12 `967d2f0` 的 **691/691、Python5、默认11/可选13包** 主/Low14 原始功能证据完整保留在独立分卷中；R12 当时视觉 pending、后被 R13 取代，不能回写为 R12 已验收。

R13 在原有硬预算/取消门后、SnapshotText 前将 U+FFFE/U+FFFF 原始文字分类为 `TEXT_XML_CHARACTER` 的 NotSupportedException。**Fail** 保持输出 sentinel；显式 **RasterizePage + Compatibility.TextLayerMode=None** 成功输出 native/raster/native 三页，中页单幅整页图、没有文字和路径，前后正常页继续 native。**默认 Invisible 文本层不能表示这些 XML 1.0 禁止字符，仍真实抛出 ArgumentException 并保持输出 sentinel，不支持这个组合**；没有自动关闭文本层、替换字符或泛化捕获 XML 异常。None 栅格页失去搜索文字。有效 BMP 邻界、原控制字符、真实 ToUnicode 代理对、硬预算与取消的既有行为另有回归。

主与独立 Low15 各 **715/715（九份TRX）、Python5/5、默认11及可选13包消费**通过。343 个冻结源码路径、13 个主 nupkg 及 Low15 646 项产物逐项 size/SHA 复核；主/独立实际 PackageReference 消费者加载的八个 Ofdrw DLL 各自匹配对应真实包。Astra 设计保存原始 80 次转换/20 个 PDF，主/Low15 用相同 PDF bytes/options 在真实主包和独立包重放，只有获准的八个 FFFE/FFFF策略组合变化；其他72个结果保持。31个旧 OFD payload 与 R12 保持，33个主/独立 R13 OFD 对照保持（仅 dual 的 CreationDate/ModDate 两项元数据归一化）。初次缺样例字形的22项失败、修正后24项通过、设计探针初次失败和真实网络 push 失败均保留，不冒称首次全绿。

2026-10-13 在 macOS Preview 实看最终 **六份PDF/20页**，每份精确核对 AX document URL，使用 Single Page / Zoom All to Fit，对照源和本轮实际 OFD→PDF 导出；六份检查前后大小/SHA均未变化。font-usage 源及 native 导出各页1..4：前两页路径-only无意外字形，后两页中英文/比例字体、浅蓝填色蓝框、绿色行及第4页 numeric TJ/qQ 红A位置对应。FFFE/FFFF 源及显式None回退导出各页1..3：中页原Z轮廓及蓝矩形保留，前后中英文正文恢复，无明显缺字、重影、裁切、重叠或额外空白页；栅格中页边缘较源向量略软。XML 样例的大块留白及靠下正文来自原始 PDF。只对这些页作结论，不声称像素一致或任意PDF/Word保真。默认Invisible失败组合没有成功产物，不计视觉通过。

六个文档均按精确 URL 保护关闭；最后 Cmd-W 后观察返回 `noWindowsAvailable`，因此没有执行 Cancel，也没有为补 Cancel 再激活 Preview；随后只有非激活 app inventory，已通知 GUI 释放，未知窗口不操作，不声称 inventory 证明全局零窗口。实际截图通过 cua_repl 内联观察；session 图片 payload 截断，归档保存真实 UI 调用、AX、错误、时间及图片块计数，**不冒称导出了可用 Preview 截图文件**。页面 PNG 是辅助渲染。历史 R10 28页/R11 8页保留原 runtime 范围，没有冒充本轮重看。

[review-r13-acceptance.json](review-r13-acceptance.json) 和 [review-r13-manifest.json](review-r13-manifest.json) 绑定最终分卷；[review-r12-functional-manifest.json](review-r12-functional-manifest.json) 绑定未改写的 R12 完整功能归档。R13 包含实际 OFD、六份已验 PDF、PNG、原始80次输入/输出、主/独立日志与失败、真实包/消费者加载身份及20页实际Preview记录；R12完整原始62+2输入/失败/产物另存其原归档。只省略派生SVG的字体base64副本和构建/cache目录。各分卷均小于90MiB，完整 archive SHA 与逐文件清单可复查。

按原始顺序重建并逐文件核验（需要 Python3 和 zstd，不解包执行未知文件）：

```sh
python3 docs/evidence/pdf-vectors/verify-r12-functional.py
python3 docs/evidence/pdf-vectors/verify-r13.py
cat docs/evidence/pdf-vectors/review-r12-functional.tar.zst.part01 docs/evidence/pdf-vectors/review-r12-functional.tar.zst.part02 > /tmp/r12-functional.tar.zst
cat docs/evidence/pdf-vectors/review-r13.tar.zst.part01 docs/evidence/pdf-vectors/review-r13.tar.zst.part02 > /tmp/r13-final.tar.zst
shasum -a 256 /tmp/r12-functional.tar.zst /tmp/r13-final.tar.zst
zstd -dc /tmp/r13-final.tar.zst | tar -tf -
```

| 归档/顺序 | 字节 | SHA256 |
| --- | ---: | --- |
| R12 whole | 178544984 | `a7cb01d96fcc7026e77473ce61b99a0e555bfc91301a8a7146dc96717425199b` |
| [review-r12-functional.tar.zst.part01](review-r12-functional.tar.zst.part01) / 1 | 90000000 | `2d425660a9b3395190c888264a6525e7a3f06c81d02a02fc43ae924f8255fb79` |
| [review-r12-functional.tar.zst.part02](review-r12-functional.tar.zst.part02) / 2 | 88544984 | `5fc09fbc6e46cd4686a19c03e17e3e9816631ded9b77faba2795b32d6c40be1c` |
| R13 whole | 109299339 | `091e7ac59e04e353fb36a8d6e18dc1bf5df66ff32edf2d0a1d1b63288a411805` |
| [review-r13.tar.zst.part01](review-r13.tar.zst.part01) / 1 | 90000000 | `a574eb7377873eb05b676af3ed9233d3eddf5cf52fa694e27fa77856bb14ce6b` |
| [review-r13.tar.zst.part02](review-r13.tar.zst.part02) / 2 | 19299339 | `1216ccbdeae1be55d9b53fa7e672c8a513c96b4783f8d10c292abc8fc0d3ca07` |

以下为保留原身份的历史验收记录；其中的当时待办以本轮说明及最新 PR 状态为准。

R11 运行时 **`6120498`** 已修复极小正 width/font size 在三位写出精度归零时绕过整页 fallback 的两个 P2：04 原样 WriterValue 抽取到 Core，04 Positive 语义不变；21 对实际 float scalar 在原精度及硬预算/取消门之后、事件前分类为 STROKE_WRITER_PRECISION/TEXT_WRITER_PRECISION。实际 612 包与旧 R10 包使用相同 18 个输入/两种策略复核，六组 writer-zero 案例由异常绕过修复为 Fail 的明确不支持或正确整页回退；正边界及 w0/Tf0 控制保持。

主验证与独立 Low13 各 **661/661**（九份 TRX）、Python **5/5**、默认 **11** 包和可选 **13** 包通过。Low13 核验 348 个源码文件、13 个主 nupkg，并记录八个实际加载 DLL 与包条目匹配；主协调复核最新 205 项独立产物 SHA/大小及补跑九份 TRX。Low13 首次 shell 包装器在测试后写入 zsh readonly status 导致包装器失败；原日志保留，修正包装器后用既有二进制重新执行测试，九份新 TRX 和 dotnet exit0 单列。旧 29 个 OFD 的主 R10→R11 payload 保持（28 个全部 ZIP 条目同字节，dual 仅 OFD.xml 日期元数据）；主 R11 与 Low13 的 30 个 OFD 亦仅 dual 日期不同。

2026-10-13 实际 macOS Preview 逐页查看新增 **两份 PDF/8 页**：四页源与本轮实际 OFD 导出各页1..4，同 Single Page/Zoom All to Fit。第2页保留 local .0004×CTM5000 的红线，第3页保留 local .0004×CTM40000 的绿色 A；两页整页栅格化边缘比源向量柔和，未见明显位置、颜色、正文或边框丢失。第1/4页正常 native 文本、浅蓝填色及蓝色边框对应，无明显缺字、裁切、重影或多余空白。源21,664,865B SHA47faf94e…、导出170,450B SHAc21c1f1a…，检查前后哈希未变。仅关闭 AX URL 精确匹配的两个文档，最后取消任务 Open 面板并停止激活 Preview，GUI 已释放。历史 R10 28 页保留原运行时范围，未将 R11 重生成的旧 PDF 冒充已重看。截图经 cua_repl 内联实际观察，独立 PNG 仅为辅助。

[review-r11-acceptance.json](review-r11-acceptance.json)、[review-r11-manifest.json](review-r11-manifest.json) 绑定 [review-r11.tar.zst](review-r11.tar.zst)，保存旧实际包复现、Astra设计与预算顺序修订、主/Low13 日志及失败、实际消费者/包/加载DLL哈希、冻结源码、新旧OFD对照、8页实际Preview记录及UI调用证据。八个旧归档不替换。非零 scalar 量化误差仍按既有04契约，不作任意PDF保真或硬内存/时间承诺。

R11 exact-head Codex 新增“仅 Tf 选择、没有字形绘制仍加载并嵌入字体”的 P2，尚待后续修复及复验；因此此归档只证明上述 R11 修复与限定视觉范围，不代表整票闭环。最终 head 的 CI、Codex/Cursor 及线程另行读回；PR #18 开放、未合并、未发布。

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
| R13 | 7afd120 | 主/Low15 各715/5、11/13包；六PDF20页Preview；默认Invisible XML禁止字符组合不支持 |
| R12 | 967d2f0 | 主/Low14 各691/5、11/13包；完整功能分卷；视觉pending后被R13取代 |
| R11 | 6120498 | 主/Low13 各 661/5、11/13 包；8 页 Preview；新字体计量 P2 待修 |
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
