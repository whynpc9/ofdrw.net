# SES 结构解析与自签测试容器（dev/interop 可选包）

**What to build:** 独立 *dev/interop* 包：解析 SES V1/V4 常见字段，用开源 SM2 自签自验 `SignedValue.dat`。默认验签路径和 CLI 对自签包不得报 `FullyValid`，除非调用方显式注册该测试验证器。

**Blocked by:** None (can start immediately)

**Status:** ready-for-agent

- [ ] 测试可生成带 SignedValue 的包，并用同一测试密钥验过
- [ ] 默认 `OfdSignatureVerifier` / CLI 不把自签报成 FullyValid
- [ ] 不宣称阅读器互认或法律效力

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P2-02

## What to build

天花板是结构正确的自签包，不是可过税局阅读器的生产签章。这张票只做可选测试容器：能写出并解析常见 SES 字段，用开源 SM2 对自己的包验过去。核心仍是可插拔 `IOfdSignatureProvider` / `IOfdSignedValueVerifier`。

禁止：默认注册进 `OfdSignatureVerifier`；把 `SupportsBuiltInSesSm2Verification` 改为 true；伪造数科/福昕/税控印章；宣称《电子签名法》或商用密码认证。合格 TSA 不内置。

## Acceptance criteria

- [ ] 可选包能生成含 `SignedValue.dat` 的 OFD，并用同一测试密钥验证成功
- [ ] 未显式注册该验证器时，CLI `verify-signatures` 对自签包不得给出 `FullyValid`；引用完整性仍可单独为真
- [ ] 核心 `SupportsBuiltInSesSm2Verification` 保持 false；可选包用自己的能力标志
- [ ] 文档标注 *dev/interop*，明确无阅读器互认、无法律效力
- [ ] 不修改、不关闭 issue #3

## Blocked by

- None (can start immediately)
