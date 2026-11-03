---
title: "签名结构与完整性"
---

安装 `Ofdrw.Net.Signatures`。核心提供标准签名目录、保护条目引用、签名值 provider / verifier 扩展，以及 SM3、SHA-1、SHA-256 摘要检查。

逐引用摘要一致只能证明受保护条目未发生该摘要意义下的变化。签名值验证、证书信任、印章外观和可信时间戳属于独立检查，需要合适的实现和信任策略。

编辑页面或附件后必须重新评估签名保护范围。关系及扩展参数见[签名参考](https://whynpc9.github.io/ofdrw.net/reference/signature-scope-v12/)。测试用途可读[SES 扩展](https://whynpc9.github.io/ofdrw.net/guides/ses-interop/)，加密另见[可选口令包](https://whynpc9.github.io/ofdrw.net/reference/password-crypto/)。
