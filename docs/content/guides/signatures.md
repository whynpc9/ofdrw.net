---
title: "签名结构与完整性"
description: "使用 OfdSignatureVerifier 检查受保护条目摘要，理解签名值 provider、verifier 与生产信任之间的差距。"
sidebar:
  order: 150
  label: "签名与完整性"
---

安装 `Ofdrw.Net.Signatures`。核心包提供标准签名目录、保护条目引用、签名值 provider / verifier 扩展点，以及 SM3、SHA-1、SHA-256 摘要检查。

## 检查签名报告

```csharp
using Ofdrw.Net.Signatures.Verification;

await using var input = File.OpenRead("signed.ofd");
var report = await new OfdSignatureVerifier().VerifyAsync(input);

Console.WriteLine($"引用摘要完整：{report.ReferenceIntegrityValid}");
Console.WriteLine($"完整验证通过：{report.FullyValid}");
foreach (var issue in report.Issues)
    Console.WriteLine(issue);
```

CLI 对应命令：

```sh
ofdrw verify-signatures -i signed.ofd
```

默认 verifier 没有厂商签名值实现时，仍可以检查受保护条目的摘要，但通常不会据此给出 `FullyValid=true`。

## 正确理解验证结果

- `ReferenceIntegrityValid=true`：列出的受保护条目在相应摘要算法下未发生变化。
- `FullyValid=true`：除引用摘要外，还要求已注册的签名值 verifier 成功，以及已声明的 `Relative` 基签名关系有效。
- 两者都**不自动证明**所有业务内容都被保护，也不等于证书链、吊销状态、可信时间戳、印章外观或法律效力已经验证。

签署时通过 `OfdSignatureOptions.ProtectFile` 和 `SignWithResultAsync` 核对实际 `ProtectedFiles` / `UnprotectedFiles`。保护范围、续签关系及扩展参数见[签名参考](https://whynpc9.github.io/ofdrw.net/reference/signature-scope-v12/)。

## 编辑与生产集成

编辑页面、附件或签名清单后必须重新评估保护范围。写包导致原签名失效时，应对最终输出重新签署，而不是保留“原签名仍有效”的假设。

生产签名值生成与验证需要调用方提供合适的 `IOfdSignatureProvider` / `IOfdSignedValueVerifier` 及独立信任策略。测试用途可阅读[SES 扩展](https://whynpc9.github.io/ofdrw.net/guides/ses-interop/)；测试自签成功不证明厂商互认或生产 PKI 信任。可选口令加密另见[口令加密包](https://whynpc9.github.io/ofdrw.net/reference/password-crypto/)。
