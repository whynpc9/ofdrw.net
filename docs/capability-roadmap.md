# Ofdrw.Net 能力路线图

更新时间：2026-09-15

本文件是当前仓库的 **TODO 清单**。它把与
[ofdrw/ofdrw](https://github.com/ofdrw/ofdrw) 的能力差距、
[issue #3](https://github.com/whynpc9/ofdrw.net/issues/3)、
以及无密码厂商前提下的加密/签章边界收成一份可执行计划。

已实现能力仍以 [功能对照](feature-parity.md) 为准。对照表描述现状；本文件描述下一步做什么、做到哪、不做什么。

状态：

- **待办**：尚未开始。
- **进行中**：已有分支或部分实现。
- **不做（核心 SDK）**：不放进默认包，也不对外宣称。
- **可选扩展**：独立包或显式开关，默认关闭，且不得提升 `FullyValid`。

优先级：

- **P0**：不碰国密就能覆盖上游日常用法。
- **P1**：缩小与 ofdrw 的核心差距（含 issue #3）。
- **P2**：无厂商可做的密码/档案上限；明确不是生产签章。

## 1. 密码与签章：务实边界

没有密码厂商时，天花板是 **结构正确的自签包 + 口令密文包 + 防篡改报告**，不是可归档、可诉讼、可过税局阅读器的电子签章。

| 层 | 含义 | 决策 |
| --- | --- | --- |
| 引用完整性 | 保护文件的 SM3 / SHA 是否被改 | **核心 SDK 已具备，保持** |
| 开源自签 / SES 解析 | 用开源 SM2 写出并验自己的 `SignedValue.dat` | **可选扩展包**，标注 *dev/interop* |
| GM/T 0099 口令加解密 | SM4 加密包内文件，口令包装文件密钥 | **可选扩展包**，不依赖厂商 |
| 生产电子签章 | 持证设备、制章系统、可信证书链、合格 TSA | **不内置**；继续 `IOfdSignatureProvider` / `IOfdSignedValueVerifier` |
| `FullyValid` | 密码学签名完全有效 | **仅在调用方注册了验证器且验证成功时为真**；自签测试容器默认不得把它变绿 |

对外只承诺三句：

1. **防篡改**：SM3 / SHA-1 / SHA-256 引用完整性，已具备。
2. **自签/口令**：可以做成可选包，明确 *dev/interop*。
3. **生产签章**：由调用方接密码机、云签或 CA。核心 SDK 不内置「看起来能验章」的 SES。

`OfdCryptographicCapabilities` 保持：

- `SupportsSm3ReferenceDigests = true`
- `SupportsPluggableSignedValueVerification = true`
- `SupportsBuiltInSesSm2Verification = false`（可选包另暴露自己的能力标志，不得改这个核心标志）
- `SupportsGmT0099EncryptionEnvelope = false`（口令加密若做成可选包，用独立类型声明，不假装核心已支持）

上游 ofdrw 的 SES 容器同样是测试实现，其文档要求有法律效力的签章使用获认证密码设备。本仓库对齐这一分层，而不是把测试容器升格成默认验章。

## 2. 不列入 TODO 的已具备项

这些不再开新任务，除非回归失败：

- OFD ZIP 打包与有界加载（路径穿越、条目数、解压炸弹）
- 常用读写闭环：页面、图层、模板、文字游程、路径、图片、字体嵌入、附件、未知 XML 保留
- 文本抽取；页重排 / 删除 / 物理裁剪；自包含合并
- DOCX → OFD/PDF（Native 与 DualLayer）
- PDF → OFD 双层（页面图 + 透明文字）；OFD → PDF / SVG（预览）
- 签名目录编排、保护引用、可插拔签名值、引用完整性与完全有效的区分
- CLI：转换、抽取、合并、重排、验签

## 3. 明确不做

| 项 | 原因 |
| --- | --- |
| 核心 SDK 内置 SES/SM2 并让 `FullyValid` 对自签为真 | 会被理解成生产验章 |
| 伪造或仿制数科/福昕/税控 `Seal.esl` 与厂商印章 ID | 无互认，且超出库的职责 |
| 宣称满足《电子签名法》或商用密码产品认证 | 需要持证设备与检测，不是算法齐了就成立 |
| 合格时间戳（TSA）内置服务 | 无厂商/TSA 合同就只有自写时间字段 |
| 加密 OFD 作为长期保存件 | OFD-H / 档案方向要求入库前解密 |
| 用直接 DOCX→PDF 结果代替 Native OFD 的视觉验收 | 见仓库 `AGENTS.md` |

## 4. P0：日常生成与文档工具

| ID | 项 | 来源 | 落地 | 完成标准 |
| --- | --- | --- | --- | --- |
| P0-01 | 公开流式布局：`Paragraph` / `Span`、折行、分页 | ofdrw-layout | 把 DOCX BuiltIn 的折行/分页抽到 `Ofdrw.Net.Layout` 公开 API | 不用手写坐标即可生成多页中英文段落；有样例与视觉验收 |
| P0-02 | 公开表格 / 单元格 | ofdrw `Cell` | 同样从 BuiltIn 表格提升为 Layout 元素 | 可生成带边框/底色/对齐的表；跨页策略有文档 |
| P0-03 | 区域占位（类表单回填） | ofdrw `AreaHolderBlock` | 命名区域 + 事后填文字/图片 | 先占位再回填，不打乱已分页版面 |
| P0-04 | OFD → PNG / JPEG | ofdrw `ImageExporter`；预览刚需 | 栅格化现有 SVG 或 PDF 页；CLI `ofd-to-image` | 指定页、指定 ppm；默认 PNG |
| P0-05 | 图片 → OFD | ofdrw `ImageConverter` | 一图一页居中 `ImageObject`；CLI `image-to-ofd` | PNG/JPEG；可设页尺寸与 ppm |
| P0-06 | 水印 | ofdrw-layout watermark | 文字/图片水印写入指定图层 | 合并与导出保留水印；有样例 |
| P0-07 | 按页拆分 | ofdrw `OFDMerger.add(pages)` | `split`：选出页生成新包 | CLI + API；资源与失效签名按现有合并约定处理 |
| P0-08 | 多文档页混合 Mix | ofdrw `addMix` | 多页叠成一页（模板、注释一并叠） | 页尺寸以第一页为准；图层顺序有文档 |
| P0-09 | 签名清理 | ofdrw `SignCleaner` | 独立 API/CLI：删除签名列表、签名值与外观 | 清完后 `verify-signatures` 报告无签名声明，而不是带着残章 |

建议批次：**P0-01/02 → P0-04/05 → P0-06/07/08/09**。布局是生成侧最大缺口；图片进出是最短路径；工具项可并行。P0-03 可紧跟表格之后。

## 5. P1：缩小与 ofdrw 的核心差距

含 [issue #3](https://github.com/whynpc9/ofdrw.net/issues/3) 点名的绘图 API 与字体体积。

| ID | 项 | 来源 | 落地 | 完成标准 |
| --- | --- | --- | --- | --- |
| P1-01 | 类 Graphics2D 绘图 API | issue #3；ofdrw-graphics2d / canvas | 公开 `OfdGraphics`、`OfdPen`、`OfdBrush`、`OfdFont`、`OfdGraphicsPath`、`OfdMatrix`（名称可微调，语义对齐 issue） | 能画线、矩形、路径、文字、变换，并落到 OFD `PathObject` / `TextObject`；有教程样例 |
| P1-02 | 字体子集化与复用 | issue #3；ofdrw-font | 按用字子集化嵌入；相同字体文件按内容身份复用，禁止无差别整份嵌入 | 中英多样例下 OFD 体积明显下降；缺字有回退策略；许可处理有说明 |
| P1-03 | Canvas `DrawContext` | ofdrw-layout canvas | 可在 P1-01 之上提供布局层画布，供页眉/票面框线使用 | 与流式布局可同页混用 |
| P1-04 | PDF → OFD 保留矢量 | ofdrw PDFConverter（Graphics2D 桥接） | PdfPig 文本/路径写入 OFD；页面图可继续当稳定视觉层 | 可选模式；默认可仍为双层。矢量模式有样例对比 |
| P1-05 | 关键字定位 | ofdrw reader keyword | 从 `TextObject` 重建字形框 | 返回页码与毫米坐标；供盖章和替换使用 |
| P1-06 | 文档内容替换 | ofdrw `DocContentReplace` | 依赖 P1-05 | 替换后版面不塌、摘要约定有说明 |
| P1-07 | 附件增删 | ofdrw-layout attachment | 对已有包增删 `Attachments.xml` 及文件 | 与合并选项 `IncludeAttachments` 一致 |
| P1-08 | 纯文本 → OFD | ofdrw `TextConverter` | 调用 P0-01 布局引擎 | 可设字号与页尺寸 |
| P1-09 | OFD → HTML | ofdrw `HTMLExporter` | 多页 SVG 包进单 HTML | 浏览器可预览；不承诺可选中全部文字 |
| P1-10 | 强类型对象补齐 | ofdrw-core | 按使用频率补大纲/书签、注释、动作、裁剪、组合对象；其余继续 `Raw` / `Preserved*` | 有读写往返测试；未知节点仍不丢 |
| P1-11 | 多 `DocBody` | ofdrw-reader | 读包不再只取第一个文档 | 有多样本；写入策略有文档 |
| P1-12 | SkiaSharp 绘图层 | ofdrw-graphics2d | 可选适配，主要服务 P1-04 | 不替代 P1-01 的 OFD 原语 API |

issue #3 的完成定义就是 **P1-01 + P1-02**。API 形状以 issue 中的类型名为准，实现可以架在 Skia 或自绘路径上，但调用方应能只依赖 OFD 原语，而不必引用 Skia。

## 6. P2：无厂商可做的密码 / 档案上限

全部为 **可选扩展**，默认不进入 `Ofdrw.Net.Converter` 元包，也不改变核心能力标志。

| ID | 项 | 落地 | 完成标准 | 禁止 |
| --- | --- | --- | --- | --- |
| P2-01 | GM/T 0099 口令加解密 | `Encryptions.xml`、`decryptseed.dat`、`entriesmap.dat`；SM4-CBC 加密选定条目；口令派生包装文件密钥 | 自加密自解密闭环；可只加密部分页；文档写明无厂商、非商用密码产品 | 当长期保存格式；当核心 SDK 已支持加密 |
| P2-02 | SES 结构解析 + 自签测试容器 | 独立 *dev/interop* 包：解析 SES V1/V4 常见字段；用开源 SM2 自签/自验 | 测试可生成带 `SignedValue.dat` 的包，并能用同一测试密钥验过；CLI 对自签包仍不得报 `FullyValid`，除非调用方显式注册该测试验证器 | 默认注册进 `OfdSignatureVerifier`；宣称阅读器互认或法律效力 |
| P2-03 | 骑缝 / 对开外观定位 | 依赖 P1-05；把印章图画到页边 | 几何正确；外观仍不是签名值 | 当作已盖具有效力的骑缝章 |
| P2-04 | 锁定签名 / 继续签的保护范围 | 引用列表排除 `Annots` / `Signs` | 签完可加注释而不必拆原摘要 | 无签名值提供者时假装已锁定 |
| P2-05 | OFD-A 检查器 | 对标 ofdrw-archive 的规则引擎（GB/T 42133） | 先做只读检查与违规报告；转换管道另议 | 宣称档案合规认证 |
| P2-06 | GM/T 0099 证书加密、完整性协议、多重加密 | 仅当 P2-01 稳定且有明确调用方 | 有互操作样本 | 无证书来源时用自签冒充 PKI |

生产验章、证书链、OCSP、合格时间戳、制章系统 **不设 TODO ID**。需要时通过已有扩展点由调用方接入，不在本仓库排期。

## 7. 工程债（与功能并行）

来自对照表里「进入生产评估前」仍有效的项，不挡 P0 开发，但发布前要过：

| ID | 项 | 完成标准 |
| --- | --- | --- |
| Q-01 | 发布许可与第三方声明 | 每次发 NuGet 核对 license 与 `THIRD-PARTY-NOTICES.md`；同批包消费 E2E |
| Q-02 | 视觉回归语料 | 票据/发票/模板/异常包；像素差阈值，而不是只查非空 |
| Q-03 | loader/reader 模糊测试与结构检查 | 坏包、路径穿越、超限 ZIP 有稳定拒绝 |
| Q-04 | 公开 API XML 文档 | 消除新增 API 的 `CS1591`；旧债按模块清 |
| Q-05 | 转换与布局变更的 Preview 视觉验收 | 遵循 `AGENTS.md`：Native OFD 产物，不得用 DOCX→PDF 代替 |

## 8. 建议落地顺序

1. **生成**：P0-01、P0-02（布局/表格）
2. **进出与工具**：P0-04、P0-05、P0-06、P0-07、P0-08、P0-09
3. **issue #3**：P1-01、P1-02（绘图 API + 字体子集化/复用）
4. **无厂商密码上限**：P2-01、P2-02（口令加密、自签测试容器）
5. **保真与检索**：P1-04、P1-05、P1-10
6. **其余 P1 / P2**：按调用方需求抽，不默认全做

同一批次内可以并行。跨批次不要提前把 P2-02 接到默认验签路径上。

## 9. 文档同步

做完对应项时：

- 更新 [功能对照](feature-parity.md) 的状态列，而不是只改本文件的表格勾选
- 布局/绘图补 [教程](tutorials/README.md) 的实现对照
- 签章/加密只允许改「可选包做了什么」和能力标志，不得删除两层验证的区分
- 涉及页面生成或转换时按 `AGENTS.md` 做视觉验收，并留下可复查产物说明

本文件不承诺任意复杂 Word / OFD 都能完整保真，也不构成对 GB/T 33190、GB/T 38540、GM/T 0099 或商用密码的合规认证。
