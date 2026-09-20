# 发布闸门：许可、视觉语料、模糊测试、API 文档

**What to build:** 发布前可重复执行的工程闸门：每次发 NuGet 核对许可与第三方声明；视觉回归不只查非空页；坏包稳定拒绝；新增公开 API 有 XML 文档。转换/布局变更仍按仓库视觉验收规则走 Native OFD。

**Blocked by:** None (can start immediately)

**Status:** ready-for-agent

- [ ] 发包核对 license 与第三方声明，并跑同批包消费 E2E
- [ ] 票据/模板/异常包语料带像素差阈值
- [ ] 坏包、路径穿越、超限 ZIP 稳定拒绝
- [ ] 新增 API 消除 CS1591；Preview 验收留下记录模板

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) Q-01–Q-05

## What to build

功能票可以并行，但预览版要进入生产评估前，这四件事要能挡发布：许可清单、视觉回归语料、加载器模糊/结构检查、公开 API 文档。视觉验收规则已经写在仓库代理说明里：Native OFD 产物，禁止用直接 DOCX→PDF 冒充。这张票把闸门做成可重复的检查和语料，而不是口头约定。

## Acceptance criteria

- [ ] 发布核对 license 与第三方声明；同批 NuGet 消费 E2E 仍作为发包前置
- [ ] 视觉回归语料覆盖票据/发票/模板/异常包一类无隐私样本；比较是像素差阈值，而不只是「页非空」
- [ ] 模糊或固定坏包集覆盖路径穿越、条目超限、解压炸弹；拒绝稳定、有测试
- [ ] 本轮新增的公开 API 无 CS1591；旧债按模块清的顺序写在文档里即可，不要求一次清完历史
- [ ] 提供 Preview 验收记录模板（源码基线、样例、模式、查看链路、已检查页、遗留问题）；不把自动化绿当成 Preview 完成

## Blocked by

- None (can start immediately)
