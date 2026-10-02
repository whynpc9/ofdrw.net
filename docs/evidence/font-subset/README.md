# Issue 05 字体子集证据

源码基线：`df71c7f20e0c45e9cba9cb78f2d2a91d61413ec9`（PR #12）；05 分支 `codex/font-subset-reuse`，stacked base `codex/document-tools`。测试源码与字体的 SHA-256 在包内 manifest/source-freeze 中记录；字体来源均为固定官方 commit，保留 OFL 原许可。不包含专有系统字体。

| 本次结果 | 范围 |
| --- | --- |
| 功能回归 | 当前全套 520/520；独立 R6 定向 72/72，R7 实际控制字符渲染探针通过 |
| 包消费 | 11 个 `0.1.0-issue05.20261003.21` 本地包，干净目录/缓存验证；包含本票真实字体样例 |
| 自动渲染 | full/subset、native/default 各两页 PDF；全部 Poppler exit 0 且无 stderr；full/subset 与 native/default 逐像素相等 |
| SVG | 8 个页面由真实 Chromium 渲染，FontFaceSet loaded 且 CSS 平台字体 `isCustomFont=true`；保留页面截图和字体记录 |
| PNG 目视 | 检查 full/subset PDF 1–2 页、native/default DOCX→本次 OFD→PDF 1–2 页；subset SVG Chromium 1–2 页，无缺字、裁切、重叠或样式扩散；其他 SVG 截图供复查 |
| macOS Preview | **通过**：2026-10-03 独占时段逐页实际查看 full/subset、native/default 四份 8 页；关闭全部 05 窗口后释放 UI；检查范围及每份 SHA 见 `preview-acceptance.json` |

统一样例（相同 source fonts、两个页面、样式、Unicode 文本和矢量矩形）：

| 载荷 | full bytes | subset bytes | 倍率 / 字形 |
| --- | ---: | ---: | --- |
| LXGW WenKai | 25,575,676 | 963,136 | 26.6x；46867 GID 中保留 140 个轮廓闭包 |
| Noto Sans | 569,208 | 171,292 | 3.32x；3748 GID 中保留 57 个闭包 |
| OFD | 12,998,912 | 397,966 | 32.66x，下降约 96.94% |

四个样式资源绑定两份唯一字体载荷，减少两份重复内容。全量对照也按内容复用，所以体积差主要证明真实子集收益。hmtx 不变；独立直接解析记录 glyf：LXGW `24,450,815 → 31,176`，Noto `401,868 → 6,624`。HarfBuzz 抽样 GID、cluster、advance、offset 与原字体相同。Unicode 文本仍在 OFD/SVG，不是栅格或路径化文本。

产物可从 `evidence.tar.zst` 解开：

```sh
zstd -d evidence.tar.zst -c | tar -xf -
```

包含本次 full/subset OFD 与 PDF、native/default DOCX 衍生 OFD/PDF、subset/native/default SVG、源 DOCX、页面 PNG、环境/哈希/检查记录和测试日志。较大的 full SVG 可由 full OFD 使用公开 SDK 重现，不重复打包。`manifest.json` 记录公开证据内每个文件的 SHA-256 与体积。

DocumentTools 首轮 CI 失败与两轮定位日志保留：空白包会生成没有文字 repertoire 的子集，随后新增水印/注释必须重新绑定原 face，或对明确待注入未知绘制的编辑样例选择 Full。本次修复按此契约执行，不放宽 coverage。再跑 36 样例 / 55 PDF 页通过。

已保留最初样例标题过长及修正后的记录；当前两页页标已拆行，重生成后再次检查。非 BMP P2 的失败轮次和独立 R2 复验亦保留。范围限制详见[字体子集契约](../../font-subset-and-reuse.md)：CFF/变量/未知表/RTL 全量保护；非 BMP PDF 与未建模 CGTransform PDF/SVG 明确失败；已有未知 OFD 资源不裁剪。Preview 仅覆盖列出的 8 页；发布/merge 未授权，不关闭 issue #3。当前 head 的 CI/review 闭环单独记录，不以这些样例推断任意复杂文档保真。

第二轮修复：default-ignorable 控制不要求 cmap 轮廓，PDF 绘制跳过可省略控制符而不改变 OFD 原文或显式 Delta 槽位；方向控制/UVS PDF 明确拒绝。保留 R6 的可见 tofu 失败和 R7 的修复页面。当时新增 ZWJ/ZWNJ 两页未获 Preview 时段；R11 已改为明确拒绝该 PDF 语义，原 8 页实际验收不受这些无控制符的 guard/绘制分支影响。TTC 容器/face 双预算、OS/2-first 统一样式已有回归；非法 SourceXml 在两模式均拒绝（a6ae520 Full 真实探针为 XmlException/0 字节，并未成功写出）。

R8：绘制时保留已映射 Hangul filler 的真实字宽，普通文本与每 gap Delta 的实测像素/位置一致；全部方向格式控制、蒙古 FVS 语义 PDF 明确拒绝。新增 filler 的无/有 Delta 页面 PNG 复查通过，实际成功输出的 Preview 排解锁后独占时段；整体视觉门仍保持未完成。

R9：LRM/弃用零宽控制的普通与定位 PDF 成功且无 tofu，九方向格式控制维持明确拒绝。name-only local/host style probe 的物理字节同时注册与缓存 cmap；缺失 filler 省略，映射 filler 保留宽度。独立实际 name-only 24 项定向验证与像素对比通过；探针使用公开 `PdfFontRegistry.CreateResolver(host)` 组合约定。记录区分不符合约定的旧探针，不将它冒充成功证据。

R10：name-only 首选缺普通字符时，配置的宿主默认回退必须覆盖，独立实际 PDF 使用 LXGW 字形并可抽取；仍缺字则明确失败/零字节。RLM/ALM 语义维持 PDF 拒绝。Symbol3/0 format4 读取与 F000 映射、全量保护、选中/未选中 SVG 回归通过；浏览器截图显示文本，但 symbol 内嵌字体身份尚未独立证明，不宣称该扩展格式完整渲染验收。

R11：ZWJ/ZWNJ 的连接/连字语义 PDF 明确拒绝并保持零输出，R7/R8 的成功省略页仅保留为历史轮次，不代表最终契约。最终仍需 Preview 的成功输出是 filler、LRM/弃用零宽控制、名称字体/覆盖回退页面。合法未使用 format0 cmap 不阻断 PDF，实际选中拒绝；symbol 未映射普通字符及 read/edit 保全量并标识 coverage 未验证。名称回退按 default→Arial 探测，null、I/O、unsupported cmap、TTC face0 均有回归；只有实际字节覆盖且注册成功才绘制，预算失败不会返回已知缺字字体。

R11 最终补验：默认名称 getter 失败仍继续 Arial 的已验证覆盖探测；固定 Noto Sans 的 office/ffi 连字上下文，普通样例 PDF 成功，ZWJ/ZWNJ 两种受控变体均明确拒绝 PDF / 0 bytes，OFD 原文逐字符保留。全套 478/478，11 个 .13 包干净消费通过。最新版包重生的原 8 页与此前 Preview 所验产物像素一致；这项自动对比不替代新增成功路径的 Preview。

R12：symbol-only 不再作为已验证 Unicode 字体；在 OFD 保全量，实际选中时 PDF/SVG 统一明确拒绝未建模字符语义/0 字节，未选中可成功导出。R10/R11 选中 A 的 SVG 成功仅属历史行为，不代表最终支持。字体绑定从按组重复解析改为每个文本一次，按源 face 内容组汇总；跨页、相同字节别名、不同字节 face 和名称字体的 repertoire 不混入，并在包括不匹配项的每个元素前检查取消。原 14 个成功路径 Preview 队列不含 symbol，继续待独占时段。

R13：已探测的 name-only symbol 同样记录 unsupported，marker 在 optional catch 外保留；选中 A/A-space/AB 都明确拒绝，未使用不增加注册字节。默认候选 symbol 不可通过 F000 remap 半验证 A，必须跳至后续 Unicode Arial 候选；全为 symbol 则明确缺覆盖失败。三处 DocumentFontContext cmap 构造均检查此边界；未将普通未探测名称字体扩大为全量覆盖审计。

R14：实际选中的普通名称字体和无 FontResource 绑定文本惰性解析物理字节、验证 cmap/当前全文并注册内容身份；未使用普通名称字体不探测，空文本在 SourceXml/控制语义预检后跳过 XFont。符号语义拒绝 marker 不被 catch 吞掉；其他普通探测失败只可进入已验证 default/Arial 候选，不再未经检查地绘制原 family 或 Arial。物理字节以不可变、按内容复用的快照保留，有效粗体/斜体 alias 指向相同物理 face；缓存覆盖逐次重新验证当前文本。

计数更正：此前全套汇总沿用了多计 2 项的总数；原始日志和独立报告不重写。`test-count-audit.json` 按六个测试项目的日志结果逐项汇总：R10 实为 464，R11 最终实为 476，R12 实为 477，R13 实为 481；本轮 R14 为 **488/488**（Core5、Packaging305、PDF112、Signatures4、DOCX49、CLI13）。旧评论或历史段落中的更高总数已由这份审计更正，独立定向计数另列。

R15：宿主 TTC 的主名称探测不得任意选择 face0；按完整名/PostScript 名（name IDs4/6）唯一匹配，未命中时 family IDs1/16 仍要求唯一，歧义或不明名称进入已验证默认候选。默认单面集合可按明确默认策略采用唯一面，多面同样必须匹配；公开 Data.CollectionFaceIndex 不变。名称目录在展开 face 前有界扫描，输入256MiB、名称解码1MiB、选中 face64MiB，非法编码/边界明确拒绝。真实非零粗体面与不明名称 fallback 已有回归。

R14 包消费首轮失败：导入 Test.pdf 的默认透明语义层含无原字体的 CIDFont+F8/U+F06C；SDK 拒绝再导 PDF/0B。只验证栅格、非空白及页面尺寸的 upstream PDF 视觉 smoke 现显式 TextLayerMode=None；SDK 默认 Invisible 不变，语义正例与真实 PUA 负例、失败日志均保留。不将这个视觉 smoke 通过冒充无损语义再导出。

R16：LRM U+200E 对混合方向文字有语义，PDF 与 RLM/ALM 一样统一明确拒绝，包含纯 Latin 用例；OFD 原文逐字符保留，实际 Latin/混合方向负例均因 bidi policy 失败且 0B。此前 LRM 省略成功页只作历史证据，不再列入最终成功 Preview 清单。字体契约已整合为现行规则，删除旧的 TTC face0 和 LRM 省略说法；本轮全套492/492。

R17：PDF 对 default-ignorable 语义统一明确拒绝，只有四个 Hangul filler 是有界绘制例外（R17 当时实际 advance/Delta 样例仅 U+3164）；WORD JOINER、不可见数学运算符、SHY、ZWSP、CGJ、MVS、FEFF、旧方向控制均不再删除后成功输出。弃用不等于无语义，OFD 字符首位 FEFF 不自动认作编码 BOM。13 个语义用例各覆盖普通/定位文字，原 OFD Unicode 完整、PDF NotSupported/0B。本轮505/505；旧 LRM/弃用控制成功页只属历史，当前最终成功清单不再包含它们。

R18：修正 R17 的证据范围表述，原本只有 U+3164 的 advance/Delta 检查。现四个 filler（115F/1160/3164/FFA0）各补普通/定位两路：映射至真实 space glyph 时 B 坐标与 A-space-B 对照一致；未映射时抽取 AB，定位 A-filler-B 的每 gap Delta5/5 对照 AB Delta10，确认跳过字形仍保留原槽位。删除 AB 自比的无效断言，产品运行逻辑未改。

R19：已建模模板文字和注释外观文字也进入同一次用字绑定/coverage 预遍历，与正文及 PDF/SVG 实际渲染集合一致；模板/注释仍触发全量保留，不能因此跳过已知 Unicode 缺字检查。typed A 保全量/原字节，typed B 缺字在 ZIP 写出前明确失败/0B；Raw/CGTransform/保留条目的未知 GID 保护和显式 Full opt-out 均不变。本轮519/519。

R19 nested read/rebind 的 baseline-resaved.ofd 使用既有 document-tools 的固定 Noto Sans CJK SC wght400 静态产物（SHA3012a9...），嵌入字节与 scripts/install-ci-fonts.py 产物一致。归档 nested-r19/font-provenance.json 记录官方 commit/source hash，NotoCJK-OFL.txt 随该新增载荷附上；无专有字体。

R20：保留策略扫描本身在开始、每页、每元素检查取消。公共 Writer 原已有入口取消，不将此前预取消行为错误归功于这次修复；改进的是内部扫描中途/入口响应。确定性 poison-page 回归验证取消先于读取包内容，无墙钟阈值。未取消时保留判断和字形路径不变，本轮520/520。
