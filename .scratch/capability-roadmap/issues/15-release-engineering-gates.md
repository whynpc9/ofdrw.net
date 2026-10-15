# 发布闸门：许可、视觉语料、模糊测试、API 文档、Preview 验收记录

**What to build:** 发布前可重复执行的五道工程闸门：每次发 NuGet 核对许可与第三方声明；视觉回归不只查非空页；坏包稳定拒绝；新增公开 API 有 XML 文档；转换与布局变更留下 Native OFD 的 Preview 验收记录。

**Blocked by:** None (can start immediately)

**Priority:** Q

**Status:** in-review (PR #8; release candidate Preview still pending)

- [x] 发包核对 license 与第三方声明，并跑同批包消费 E2E
- [x] 票据/模板/异常包语料带像素差阈值
- [x] 坏包、路径穿越、超限 ZIP 稳定拒绝
- [x] 新增 API 消除 CS1591
- [x] Preview 验收留下记录模板（Native OFD，不得用 DOCX→PDF 代替）

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) Q-01–Q-05

## What to build

功能票可以并行，但预览版要进入生产评估前，这五件事要能挡发布：许可清单、视觉回归语料、加载器模糊/结构检查、公开 API 文档、Preview 验收记录。视觉验收规则已经写在仓库代理说明里：Native OFD 产物，禁止用直接 DOCX→PDF 冒充。这张票把闸门做成可重复的检查和语料，而不是口头约定。Q-05 的验收记录模板在范围内，不能只做前四道就当作完成。

## Acceptance criteria

- [x] 发布核对 license 与第三方声明；同批 NuGet 消费 E2E 仍作为发包前置
- [x] 视觉回归语料覆盖票据/发票/模板/异常包一类无隐私样本；比较是像素差阈值，而不只是「页非空」
- [x] 模糊或固定坏包集覆盖路径穿越、条目超限、解压炸弹；拒绝稳定、有测试
- [x] 本轮新增的公开 API 无 CS1591；旧债按模块清的顺序写在文档里即可，不要求一次清完历史
- [x] 提供 Preview 验收记录模板（源码基线、样例、模式、查看链路、已检查页、遗留问题）；不把自动化绿当成 Preview 完成

实现正在 PR #8 审查；勾选表示闸门代码与模板已提交，不表示发布候选的人工 Preview 已完成或允许发包。`docs/release-preview-acceptance.json` 保持 `not-reviewed`，直到实际候选被逐页查看。

## Blocked by

- None (can start immediately)
