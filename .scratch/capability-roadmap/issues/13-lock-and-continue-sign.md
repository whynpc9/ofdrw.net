# 锁定签名与继续签的保护范围

**What to build:** 签名保护引用排除 `Annots` / `Signs`，签完可以加注释而不必拆掉原摘要。没有签名值提供者时，不得假装已经锁定。

**Blocked by:** None (can start immediately)

**Status:** ready-for-agent

- [ ] 锁定/继续签的引用列表排除 Annots 与 Signs
- [ ] 签完加注释后原保护引用仍能对上（按约定范围）
- [ ] 无 SignedValue 提供者时不宣称已锁定

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P2-04

## What to build

生产里常见「先签正文，再允许批注」。这张票调整签名编排的保护范围：把注释和签名目录排除在摘要之外，使继续签或事后加注不必作废原文摘要。这仍然只是引用列表策略，不是密码学签章。没有注册 `IOfdSignatureProvider` 时，不得生成看起来已经锁章的包。

## Acceptance criteria

- [ ] 可选的锁定/继续签模式生成的保护引用不含 Annots、Signs（及其约定载荷）
- [ ] 在该模式下签名编排完成后添加注释，原保护引用摘要仍匹配；改正文则不匹配
- [ ] 未提供签名值提供者时，API 拒绝「已锁定」语义，不写出假锁章
- [ ] 文档说明这是保护范围，不是已盖具有效力的章；更新功能对照

## Blocked by

- None (can start immediately)
