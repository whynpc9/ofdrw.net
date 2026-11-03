---
title: "已知问题与版本差异"
---

| 项目 | preview.10 | 当前 main |
| --- | --- | --- |
| Regular-only 汉字合成粗体 | 固定表头“量”存在历史视觉失败；发布验收限定真实 Regular/Bold | D11 / PR108 修复及固定样例验收通过，尚未发布 |
| CLI 严格嵌入粗体 | 无对应开关，严格配置使用 API | 已增加限定开关，尚未发布 |
| macOS Preview Find | 部分完整词 / 日期查找异常，根因未定 | 保留独立限制 |
| 原始 Unicode 剪贴板 Copy | 未完成字节级验收 | 未完成，不等同已证实缺陷 |
| ImageSharp 2.1.13 | 已披露依赖告警 | D10 评估未消除告警 |
| Skia / Windows Audit025 偶发问题 | 历史根因未定 | 保留诊断，被动观察新自然失败 |

[preview.10 发布说明](https://whynpc9.github.io/ofdrw.net/releases/010-preview10/)只描述原发布。后续结果以[路线图](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/roadmap.md)及各具名验收报告为准。

## 转换或显示异常

先确认包版本、转换方向、Native / DualLayer、实际物理字体和是否触发明确回退，再保留可复现的最小输入。PNG、PDF、SVG 和阅读器 UI 行为需要分别检查。任意复杂 Word、完整 R02 外来映射、通用 RTL/Arabic/Indic、完整档案符合性仍不在普遍承诺内。
