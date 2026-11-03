---
title: "已知问题与版本差异"
description: "核对 preview.10 与 main 的行为差异，并按版本、模式、字体和实际输出排查转换及阅读器问题。"
sidebar:
  order: 10
---

当前公开版本为 **0.1.0-preview.10**。下表将已发布包的行为与当前 `main` 分开，避免把源码修复误认为 NuGet 已更新。

## 版本差异与现有限制

| 项目 | preview.10 | 当前 main |
| --- | --- | --- |
| Regular-only 汉字合成粗体 | 固定表头“量”存在历史视觉失败；发布验收限定真实 Regular/Bold | D11 / PR108 修复及固定样例验收通过，尚未发布 |
| CLI 严格嵌入粗体 | 无对应开关，严格配置使用 API | 已增加限定开关，尚未发布 |
| macOS Preview Find | 部分完整词 / 日期查找异常，根因未定 | 保留独立限制 |
| 原始 Unicode 剪贴板 Copy | 未完成字节级验收 | 未完成，不等同已证实缺陷 |
| ImageSharp 2.1.13 | 已披露依赖告警 | D10 评估未消除告警 |
| Skia / Windows Audit025 偶发问题 | 历史根因未定 | 保留诊断，被动观察新自然失败 |

[preview.10 发布说明](https://whynpc9.github.io/ofdrw.net/releases/010-preview10/)只描述原发布。后续结果以[路线图](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/roadmap.md)及各具名验收报告为准。

## 转换或显示异常时如何排查

1. **确认版本和入口。** 记录 NuGet/CLI 版本、转换方向、API 或完整命令，以及 Native / DualLayer、实际渲染引擎。不要混用 `main` 文档和 preview.10 包行为。
2. **核对实际字体。** 记录 `FontDirectories`、fallback、Regular/Bold/Italic 文件和输出内嵌资源。字体名称相同不代表字体字节相同；先阅读[字体指南](https://whynpc9.github.io/ofdrw.net/guides/fonts/)。
3. **检查诊断。** DOCX 使用 `ConvertWithResultAsync` 查看 `Diagnostics`、`OriginalTextPreserved` 和页级映射状态；其他 API 检查返回结果及异常，而不是只看文件是否生成。
4. **保留最小复现。** 保存可共享的最小输入、配置、完整错误、退出码和实际输出；不要只保留截图，也不要直接覆盖唯一输入副本。
5. **分离出口和阅读器。** 分别检查 OFD 原文、OFD→PDF、SVG、PNG、HTML 及目标阅读器 UI。某个出口成功不能证明其他出口正确。
6. **逐页检查。** 关注缺字/乱码、英文断词、局部样式扩散、字号与基线、表格边框、对齐、分页、裁切、重叠、重影和异常空白页。

验证 Native DOCX 时，必须查看本次 **DOCX → Native OFD** 的产物；如经 PDF 查看，应继续使用该 OFD 导出的 PDF，不能用直接 DOCX→PDF 替代。

## 不应忽略的能力边界

任意复杂 Word 保真、完整 R02 外来字形映射、通用 RTL/Arabic/Indic、完整档案符合性、生产 PKI/签章法律效力及所有阅读器互操作仍不在普遍承诺内。具体拒绝和未验证范围见[能力契约](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/)。
