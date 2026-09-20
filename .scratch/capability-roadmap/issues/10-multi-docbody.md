# 多 DocBody 读取与写入策略

**What to build:** 读包不再只处理第一个文档体。多样本能列出并选中各个 `DocBody`；写入多文档时的策略有文档（默认仍可只写一个）。

**Blocked by:** None (can start immediately)

**Status:** ready-for-agent

- [ ] 读包可枚举/选择非首个 DocBody
- [ ] 有多样本；写入策略有文档
- [ ] 单文档包行为与现在兼容

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P1-11

## What to build

规范允许一份 OFD 里多个文档体，当前读取以第一个为主。这张票让阅读 API 能发现并打开后续文档，而不是静默丢掉。写入侧先把「何时写多个、默认仍一个」写清楚，避免调用方误以为合并两个文件就会自动变成两个 DocBody。

## Acceptance criteria

- [ ] 对含多个 DocBody 的样本，API 能列出全部文档并按选择读取页面/资源，而不是只返回第一份
- [ ] 只有一个 DocBody 的包，行为与现网一致
- [ ] 文档说明写入策略（默认单文档、如何显式追加）；没有文档就不假装已支持任意多文档写入
- [ ] 有多样本测试；更新功能对照

## Blocked by

- None (can start immediately)
