---
title: "已知问题与版本差异"
description: "核对 0.5.1、0.5.0、0.4.0、0.3.x、0.2.0 与 preview.10 的行为差异，并按版本、模式、字体和实际输出排查转换及阅读器问题。"
sidebar:
  order: 10
---

当前文档以 **0.5.1** 为基线。见[0.5.1 发布说明](/releases/051/)。

## 从 0.5.0 升级到 0.5.1

请一起升级同系列包。注释外观中的复合对象现在参与 OFD 文本提取；比对旧版提取结果时，留意注释中的文字。`/2016` 文档的 PDF 注释外观按源命名空间解析裁剪。Medical 评估纳入隐式同目录 `PageRes.xml`，含百分号转义的资源位置或无法定位的 DocRoot 可能改变范围报告。

## 从 0.4.0 升级到 0.5.0

同系列 NuGet 包继续采用精确的同批版本依赖，请一起升级所有 `Ofdrw.Net.*` 直接引用。默认发布集合由 12 包增至 14 包，新增的 Archival、Medical 仍需显式安装，不会由 Converter 传递带入。

| 项目 | 0.4.0 及更早版本 | 0.5.0 与升级操作 |
| --- | --- | --- |
| 新文档默认 `DocType` | 默认 `OFD-H` | 默认普通 `OFD`。需要 OFD-H 时显式安装 Medical 并应用 `MedicalDocumentProfile`；不要把标签当作标准符合性证明 |
| OFD-A 内容修改 | 编辑、拆分和合并会把源声明带到输出 | 改变内容的通用入口拒绝原始 `DocType="OFD-A"`，CLI 返回 `5` 且不写输出。先用 `derive-ofd` 或 `OfdOrdinaryDerivation.CreateAsync` 创建新普通 OFD；已签名或已加密输入不能派生，派生副本不继承归档结论或签名效力 |
| Native DOCX 字体 | 系统中文族等可只声明名称，由阅读器解析本机字体 | 默认 `Compatibility` 不变；可显式选择 `RequireEmbedded`，要求实际使用的字体都嵌入该族自身字节。缺字节、跨族替代或嵌入受限会失败；不适用于 DualLayer 或 DOCX→PDF，许可仍由调用方核对 |
| 新写出的自定义标签 | 使用历史 EMR 通道 | 新内容使用中性 `ofdrw-kv` 通道。旧 EMR 内容仍可读，未修改结构继续保留；该形态不代表已核验 GB/T 33190 符合性 |
| 归档与电子病历扩展 | 未随 0.4.0 交付 Archival、Medical 包 | 自 0.5.0 起提供两个显式可选包。Archival 只读评估且不生成 OFD-A；Medical 提供 OFD-H 写出 profile 和只读评估。两者都不会给出整部标准的整体 `Pass`，不等于合规认证 |

OFD-A 只读读取、提取、定位、验签、预检和格式导出不受编辑保护影响。底层模型赋值与 `OfdPackageWriter` 原样写出不经过全部通用入口保护；完整操作矩阵见[页面与文档工具](/guides/document-tools/)。Archival 和 Medical 的未交付扩展、评估边界及外部证据要求见各自指南。

## 从 0.3.1 升级到 0.4.0

同系列 NuGet 包请一起升级。依赖旧 CLI 宽松解析或解析人读终端输出的脚本，应按[CLI 参考](/reference/cli/)核对输入参数和转义。最终排版引擎的行数限制按实际接受的行数计算，长串仍存在长距离试排。

## 0.3.0 到 0.3.1 的历史差异

如果使用 0.3.0 导出骑缝章时在切片接缝看到相邻颜色，可用 0.3.1 重新导出。两类修复及验证边界见[0.3.1 发布说明](/releases/031/)。

## 0.2.0 到 0.3.0 的历史差异

| 项目 | 0.2.0 | 0.3.0 |
| --- | --- | --- |
| Flow、纯文本、DOCX Native 的最终文字排版 | 独立段落 writer；现有入口仍用兼容排版 | 可显式启用真实字体塑形与嵌入；默认仍是兼容排版 |
| 制表位、装饰与复杂字体输出 | 独立 writer 的范围有限 | 显式引擎支持制表位、合成粗斜体、装饰线、高亮、字距、静态 CFF 与增补字符；固定样例通过，任意输入未验 |
| PDF→OFD 的 RTL 透明文字层 | 默认拒绝未验证的 RTL 语义 | 默认仍拒绝；可显式选有界 `VisualLineProfile`，超出其条件继续拒绝 |
| 归档来源与证据档案 | 有限只读预检 | 新增证据追踪和外部缺失门禁；完整符合性及认证仍未完成 |
| Preview Find、原始 Copy、段落 PDF 选择几何 | 未验证为通过 | 仍未验证为通过 |
| ImageSharp 2.1.13 | 已披露五项依赖告警 | 告警仍在，未修复 |

### preview.10 到 0.2.0 的历史差异

| 项目 | preview.10 | 0.2.0 |
| --- | --- | --- |
| 较大 OFD 读取（约 1,000+ 文本对象） | 读取/转换报 "OFD traversal exceeds its reference budget" | 默认工作计数上限已提高 |
| OFD→PDF 签章外观失败 | 静默跳过 | 通过诊断报告 |
| Regular-only 汉字合成粗体 | 固定表头“量”存在历史视觉失败；发布验收限定真实 Regular/Bold | 本库 OFD→PDF/图片出口修复固定样例的字腔堵塞 |
| CLI 严格嵌入粗体 | 无对应开关，严格配置使用 API | 支持 `--native-embedded-bold-behavior throw` |
| 段落排版引擎 | 无独立排版和 writer 入口 | 新增独立 API，Flow/DOCX 继续使用原入口 |
| macOS Preview Find | 部分完整词 / 日期查找异常，根因未定 | 保留独立限制 |
| 原始 Unicode 剪贴板 Copy | 未完成字节级验收 | 未完成，不等同已证实缺陷 |
| ImageSharp 2.1.13 | 已披露依赖告警 | D10 评估未消除告警 |
| Skia / Windows Audit025 偶发问题 | 历史根因未定 | 保留诊断，被动观察新自然失败 |

[0.3.0 发布说明](/releases/030/)记录该版范围；[preview.10 发布说明](https://whynpc9.github.io/ofdrw.net/releases/010-preview10/)只描述原发布。后续结果以[路线图](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/roadmap.md)及各具名验收报告为准。

## 转换或显示异常时如何排查

1. **确认版本和入口。** 记录 NuGet/CLI 版本、转换方向、API 或完整命令，以及 Native / DualLayer、实际渲染引擎。按安装版本查阅对应发布说明。
2. **核对实际字体。** 记录 `FontDirectories`、fallback、Regular/Bold/Italic 文件和输出内嵌资源。字体名称相同不代表字体字节相同；先阅读[字体指南](https://whynpc9.github.io/ofdrw.net/guides/fonts/)。
3. **检查诊断。** DOCX 使用 `ConvertWithResultAsync` 查看 `Diagnostics`、`OriginalTextPreserved` 和页级映射状态；其他 API 检查返回结果及异常，而不是只看文件是否生成。
4. **保留最小复现。** 保存可共享的最小输入、配置、完整错误、退出码和实际输出；不要只保留截图，也不要直接覆盖唯一输入副本。
5. **分离出口和阅读器。** 分别检查 OFD 原文、OFD→PDF、SVG、PNG、HTML 及目标阅读器 UI。某个出口成功不能证明其他出口正确。
6. **逐页检查。** 关注缺字/乱码、英文断词、局部样式扩散、字号与基线、表格边框、对齐、分页、裁切、重叠、重影和异常空白页。

验证 Native DOCX 时，必须查看本次 **DOCX → Native OFD** 的产物；如经 PDF 查看，应继续使用该 OFD 导出的 PDF，不能用直接 DOCX→PDF 替代。

## 不应忽略的能力边界

任意复杂 Word 保真、完整 R02 外来字形映射、通用 RTL/Arabic/Indic、完整档案符合性、生产 PKI/签章法律效力及所有阅读器互操作仍不在普遍承诺内。具体拒绝和未验证范围见[能力契约](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/)。
