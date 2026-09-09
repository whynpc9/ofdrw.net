# OFD 格式教程

Last verified: 2026-09-09

这是一组面向开发者的 **OFD 格式（Spec）教程**，不是转换器 API 手册。读完后应能打开一份 `.ofd`、看懂常用 XML、并知道文字、路径、图片分别写在哪里。

规范依据是 [GB/T 33190-2016](https://openstd.samr.gov.cn/bzgk/gb/newGbInfo?hcno=3AF6682D939116B6F5EED53D01A9DB5D)《电子文件存储与交换格式 版式文档》。教程只归纳公开结构与实现对照，不转录标准正文或 Schema。正式合规以标准文本为准。

实现对照来自：

- 源项目 [ofdrw/ofdrw](https://github.com/ofdrw/ofdrw/tree/master)（Java，`ofdrw-core` / `ofdrw-pkg` 最接近规范对象模型）
- 本仓库 Ofdrw.Net（.NET，常用读写闭环，对象模型是子集，见 [功能对照](../feature-parity.md)）

## 怎么读

按编号顺序读。前 5 课建立容器和页面骨架；第 6–8 课覆盖绝大多数版式内容；第 9–10 课是发票、公文、病案里常见但不属于“先画一页字”的能力。做电子病历交付时，在第 10 课之后读 OFD-H 补充课（第 11–14 课）。

每课结构固定：

1. 这一课要能独立完成什么
2. 对应规范章节（只给章条号）
3. 最小可观察的包/XML 例子
4. ofdrw 与 Ofdrw.Net 分别落在哪个模块
5. 实际文件里常见的坑

动手时把任意 `.ofd` 复制一份改名为 `.zip` 解压即可。不要用含隐私的真实票据或病案当示例。本仓库 CLI 可生成一份无隐私样例：

```bash
dotnet run --project src/Ofdrw.Net.Cli -- convert path/to/input.docx artifacts/hello.ofd
unzip -l artifacts/hello.ofd
```

## 课程

| # | 课 | 优先覆盖 | 规范 |
| --- | --- | --- | --- |
| 1 | [打开 OFD：ZIP 容器与文件组织](01-open-the-package.md) | 先会拆包 | 第 5、6 章 |
| 2 | [命名空间与基础类型](02-basic-types.md) | 读懂属性值 | 7.1–7.3 |
| 3 | [主入口 OFD.xml](03-ofd-xml.md) | 找到文档 | 7.4 |
| 4 | [Document.xml、页树与资源入口](04-document-and-pages.md) | 找到每一页 | 7.5–7.7、7.9 |
| 5 | [坐标系、图层与图元](05-coordinates-and-layers.md) | 知道东西画在哪 | 8.1、8.5、7.7 |
| 6 | [文字 TextObject](06-text.md) | 最常用 | 第 11 章 |
| 7 | [路径 PathObject](07-paths.md) | 线框、表格边、底色 | 第 9 章 |
| 8 | [图像与资源](08-images-and-resources.md) | 图、字体文件 | 第 10 章、7.9、11.1 |
| 9 | [模板、大纲、附件与自定义标引](09-templates-attachments-tags.md) | 页眉底板、目录树、内嵌源文件 | 7.7–7.8、第 16、20 章 |
| 10 | [数字签名结构导读](10-signatures.md) | 发票/公文常见，先认结构 | 第 18 章 |

### OFD-H 补充（电子病历轮廓）

材料是《电子病历版式文档技术要求》**征求意见稿**（英文题 Technical specification for OFD used in Medical record），不是已发布的 GB 号。它在 GB/T 33190 与 GB/T 42133 之上加医疗场景约束。教程只归纳差异，不转录稿正文。

| # | 课 | 优先覆盖 | 稿中位置 |
| --- | --- | --- | --- |
| 11 | [OFD-H 是什么](11-ofd-h-overview.md) | `DocType=OFD-H` 从哪来 | 范围、引言、第 5 章 |
| 12 | [包轮廓](12-ofd-h-package-profile.md) | 单文档、短命名空间、目录 | 第 6 章 |
| 13 | [医疗内容与资源](13-ofd-h-medical-content.md) | 嵌字体、双层扫描、DICOM 外置 | 第 7 章 |
| 14 | [安全与软件](14-ofd-h-security-and-software.md) | 必签、锁定签名、阅读器只读 | 6.3、第 8–9 章 |

本仓库默认 `DocType` 为 `OFD-H`、默认命名空间为 `http://www.ofdspec.org`，与这篇轮廓一致；不等于已经满足嵌字、签章范围和阅读器行为。

## 源项目模块对照

ofdrw 把规范对象几乎全部代理成 DOM 类型；Ofdrw.Net 用更小的强类型模型，读包时把尚未建模的节点保存在 `SourceXml` / `Preserved*` 里。

| 规范对象 | ofdrw (Java) | Ofdrw.Net |
| --- | --- | --- |
| 基础类型 `ST_*` | `ofdrw-core` `org.ofdrw.core.basicType` | 写入时直接格式化为字符串 |
| `OFD.xml` / `Document.xml` / 页 | `ofdrw-core` `basicStructure` + `ofdrw-pkg` | `OfdPackageWriter` / `OfdReader` |
| 图元 | `pageDescription`、`text`、`graph`、`image` | `OfdTextElement`、`OfdPathElement`、`OfdImageElement`、`OfdRawElement` |
| 布局（段落、分页） | `ofdrw-layout` | 低层 `OfdDocumentBuilder`；无段落引擎 |
| 签章 | `ofdrw-sign`、`ofdrw-gm` | `Ofdrw.Net.Signatures`（摘要已实现，SES/SM2 为扩展点） |

## 刻意后置的内容

下面这些在规范里完整存在，但日常生成/阅读一份业务 OFD 很少先碰到。本系列不展开，需要时直接读标准对应章和 ofdrw-core 同名包：

- 渐变、Pattern、复合对象（8.3、第 13 章）
- 视频/音频与动作（第 12、14 章）
- 注释（第 15 章）
- 版本链（第 19 章）
- GM/T 0099 加密包（ofdrw-crypto；本仓库未实现）
- OFD-A 档案子集的加工细则（GB/T 42133；ofdrw-archive）。电子病历在档案方向上仍引用它，轮廓差异见第 11–14 课

## 相关文档

- [功能对照](../feature-parity.md)：本仓库已支持什么
- [转换、编辑与资源约定](../conversion-contracts.md)：DOCX/PDF 转换行为，不是格式课
- 源项目布局说明：[ofdrw-layout/doc/layout](https://github.com/ofdrw/ofdrw/blob/master/ofdrw-layout/doc/layout/README.md)
- 源项目签章入门：[ofdrw-sign/doc/quickstart](https://github.com/ofdrw/ofdrw/blob/master/ofdrw-sign/doc/quickstart/README.md)
